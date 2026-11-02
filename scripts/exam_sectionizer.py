#!/usr/bin/env python3
"""Segment a Gaokao-style English exam and emit a routing manifest.

Reads references/module-registry.json as the single source of truth for
section patterns, module titles and IV-form policy, so the manifest route
can never drift from the blueprint documents.

Also produces a *heuristic* genre guess per section. The guess is advisory:
the agent must confirm or override it and record the decision in QA.

Supports TXT, MD and DOCX with the standard library only. PDF and image
inputs must be converted to text first.
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

REGISTRY_PATH = Path(__file__).resolve().parent.parent / "references" / "module-registry.json"

GENRE_TOKENS = {
    "narrative": [
        "said", "asked", "replied", "shouted", "whispered", "told", "felt", "realized",
        "remembered", "decided", "began", "started", "walked", "ran", "looked", "turned",
        "smiled", "cried", "laughed", "grabbed", "pulled", "opened", "noticed", "stood",
    ],
    "expository": [
        "is", "are", "refers", "means", "known", "called", "consists", "includes",
        "process", "system", "research", "study", "data", "results", "shows", "suggests",
        "according", "percent", "million", "billion", "species", "structure", "function",
    ],
    "argumentative": [
        "however", "but", "therefore", "thus", "hence", "moreover", "furthermore",
        "argue", "claim", "should", "must", "cannot", "fail", "wrong", "flawed",
        "evidence", "conclusion", "in fact", "rather than", "instead", "worse",
    ],
    "practical": [
        "please", "contact", "visit", "click", "register", "available", "free",
        "open", "hours", "price", "ticket", "book", "join", "apply", "deadline",
        "further information", "for more", "note that", "if you",
    ],
}

TIME_MARKERS = (
    "yesterday", "last ", "that day", "next morning", "the following day",
    "years later", "at first", "finally", "after that", "from then on",
)


def load_registry(path: Path = REGISTRY_PATH) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def read_docx(path: Path) -> str:
    with zipfile.ZipFile(path) as archive:
        xml = archive.read("word/document.xml")
    root = ET.fromstring(xml)
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    paragraphs = []
    for para in root.findall(".//w:p", ns):
        parts = [node.text or "" for node in para.findall(".//w:t", ns)]
        if parts:
            paragraphs.append("".join(parts))
    return "\n".join(paragraphs)


def read_input(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix in {".txt", ".md"}:
        return path.read_text(encoding="utf-8-sig")
    if suffix == ".docx":
        return read_docx(path)
    raise ValueError(f"Unsupported input type: {suffix}. Convert PDF/image to UTF-8 text first.")


def normalize(text: str) -> str:
    text = text.replace("\u00a0", " ").replace("\u3000", " ")
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    return re.sub(r"[ \t]+", " ", text).strip()


def guess_genre(text: str) -> dict:
    """Advisory genre guess from surface cues. Never authoritative."""
    lowered = text.lower()
    words = re.findall(r"[A-Za-z']+", lowered)
    total = max(len(words), 1)
    scores: dict[str, float] = {}

    for genre, tokens in GENRE_TOKENS.items():
        hits = 0
        for token in tokens:
            if " " in token:
                hits += lowered.count(token)
            else:
                hits += words.count(token)
        scores[genre] = hits / total * 100

    narrative_bonus = sum(1 for marker in TIME_MARKERS if marker in lowered)
    if re.search(r'[.!?]\s*["\u201c]', text):
        narrative_bonus += 2
    first_person = len(re.findall(r"\b(I|we|my|our|me|us)\b", text))
    past_forms = len(re.findall(r"\b[A-Za-z]{3,}ed\b", text))
    ambiguous = len(re.findall(r"\b(?:was|were|is|are)\s+\w+ed\b", text))
    scores["narrative"] += (narrative_bonus * 1.2 + first_person * 0.6 + max(past_forms - ambiguous, 0) * 0.8) / total * 100

    ranked = sorted(scores.items(), key=lambda item: item[1], reverse=True)
    top_id, top_score = ranked[0]
    second_id, second_score = ranked[1] if len(ranked) > 1 else ("", 0.0)
    gap = (top_score - second_score) / max(top_score, 1e-6)
    confidence = round(min(0.95, max(0.0, gap * 0.9 + min(top_score, 5) / 20)), 2)

    return {
        "guess": top_id,
        "confidence": confidence,
        "runner_up": second_id,
        "scores": {key: round(value, 2) for key, value in ranked},
        "confirmed": None,
        "note": "启发式猜测，须读原文确认或改写；确认后把 confirmed 填为该体裁 id",
    }


PART_RE = re.compile(r"^第\s*([一二三四五六七八九十]+)\s*部分\s*[：:]?\s*(.*)$")
SUB_RE = re.compile(r"^第\s*([一二])\s*节")
LETTER_RE = re.compile(r"^([A-DＡ-Ｄ])$")
CN_NUMERAL = {"一": "1", "二": "2", "三": "3", "四": "4", "五": "5"}


def split_lines(text: str) -> tuple[list[str], list[int]]:
    lines = text.split("\n")
    offsets = []
    position = 0
    for line in lines:
        offsets.append(position)
        position += len(line) + 1
    return lines, offsets


def _match_role(title: str, part_keywords: dict) -> str | None:
    for role, keywords in part_keywords.items():
        if any(keyword in title for keyword in keywords):
            return role
    return None


def find_passages(lines: list[str], start: int, end: int, letters: list[str], min_chars: int) -> list[tuple[str, int]]:
    """Bare A/B/C/D lines inside 第一节 of 阅读 mark the four passages."""
    expected = [letter.upper() for letter in letters]
    found: list[tuple[str, int]] = []
    cursor = 0
    for index in range(start, end):
        token = lines[index].strip().replace("Ａ", "A").replace("Ｂ", "B").replace("Ｃ", "C").replace("Ｄ", "D")
        if cursor >= len(expected) or token != expected[cursor]:
            continue
        tail = "".join(lines[index + 1: min(index + 40, end)])
        if len(tail) < min_chars:
            continue
        found.append((expected[cursor], index))
        cursor += 1
    return found


def parse_structural(lines: list[str], registry: dict) -> list[tuple[int, str]]:
    """Standard Gaokao layout: 第N部分 + 第N节 + bare passage letters."""
    rules = registry.get("structural_rules")
    if not rules:
        return []
    part_keywords = rules["part_keywords"]
    slots = rules["sub_slots"]
    letters = rules.get("passage_letters", ["A", "B", "C", "D"])
    min_chars = rules.get("min_passage_chars", 120)

    parts: list[tuple[int, str]] = []
    for index, raw in enumerate(lines):
        match = PART_RE.match(raw.strip())
        if not match:
            continue
        role = _match_role(match.group(2), part_keywords)
        if role:
            parts.append((index, role))
    if not parts:
        return []

    marks: list[tuple[int, str]] = []
    for position, (start, role) in enumerate(parts):
        end = parts[position + 1][0] if position + 1 < len(parts) else len(lines)
        if role == "listening":
            continue
        subs = []
        for index in range(start, end):
            match = SUB_RE.match(lines[index].strip())
            if match:
                subs.append((index, CN_NUMERAL.get(match.group(1), match.group(1))))
        sub_map = slots.get(role, {})

        if role == "reading":
            first_start, first_end = start, end
            for order, (index, number) in enumerate(subs):
                if number == "1":
                    first_start = index
                    first_end = subs[order + 1][0] if order + 1 < len(subs) else end
            for letter, index in find_passages(lines, first_start, first_end, letters, min_chars):
                marks.append((index, f"reading-{letter.lower()}"))
            for index, number in subs:
                if number == "2" and sub_map.get("2"):
                    marks.append((index, sub_map["2"]))
        else:
            for index, number in subs:
                target = sub_map.get(number)
                if target:
                    marks.append((index, target))
    return marks


def find_markers(lines: list[str], offsets: list[int], registry: dict) -> tuple[list[dict], str]:
    explicit: list[tuple[int, str, str]] = []
    for section in registry["sections"]:
        for index, line in enumerate(lines):
            if re.search(section["pattern"], line, flags=re.IGNORECASE):
                explicit.append((index, section["id"], line.strip()[:24]))
                break

    structural = [(index, section_id, "结构解析") for index, section_id in parse_structural(lines, registry)]

    source = "explicit"
    if len({sid for _, sid, _ in structural}) >= 3:
        source = "structural"
        covered = {sid for _, sid, _ in structural}
        picked = structural + [item for item in explicit if item[1] not in covered]
    else:
        picked = explicit or structural

    labels = {section["id"]: section["label"] for section in registry["sections"]}
    picked.sort(key=lambda item: item[0])

    markers = []
    seen = set()
    for index, section_id, marker in picked:
        if section_id in seen or section_id not in labels:
            continue
        seen.add(section_id)
        markers.append(
            {
                "id": section_id,
                "label": labels[section_id],
                "start": offsets[index],
                "marker": marker,
                "line": index,
            }
        )
    return markers, source


def pick_variant(form: str | None, section_id: str, registry: dict, preferred: str | None = None) -> dict:
    """IV-B has three sub-forms; the section decides which one is the default."""
    variants = (registry.get("iv_variants") or {}).get(form or "", [])
    if not variants:
        return {}
    if preferred:
        for variant in variants:
            if variant["id"] == preferred:
                return variant
    for variant in variants:
        if section_id in variant.get("default_for", []):
            return variant
    return variants[0]


def resolve_route(section: dict, registry: dict, genre_id: str | None, word_count: int) -> dict:
    modules = list(section["modules"])
    policy = section["iv_policy"]
    resolved = {
        "modules": modules,
        "module_titles": [],
        "iv_form": None,
        "iv_reason": "",
        "iv_variant": None,
        "iv_variant_title": None,
        "iv_variant_reason": "",
        "deviations": [],
    }

    if policy["type"] == "none":
        resolved["iv_reason"] = policy.get("reason", "该板块不设 IV")
    elif policy["type"] == "section_specific":
        resolved["iv_reason"] = "板块自带四模块结构，不使用 IV-A~E 代号"
    elif policy["type"] == "fixed":
        form = policy["form"]
        override_genre = policy.get("override_when_genre_not")
        if override_genre and genre_id and genre_id != override_genre:
            form = policy.get("override_form", form)
            resolved["deviations"].append(
                f"体裁为 {genre_id}，与默认 {policy['form']} 的记叙前提不符，改用 {form}"
            )
        if form not in modules:
            modules.append(form)
        resolved["iv_form"] = form
        resolved["modules"] = modules
        resolved["iv_reason"] = f"题型本质决定，固定 {form}"
        variant = pick_variant(form, section["id"], registry, policy.get("preferred_variant"))
        if variant:
            resolved["iv_variant"] = variant["id"]
            resolved["iv_variant_title"] = variant["title"]
            resolved["iv_variant_reason"] = policy.get("variant_reason", variant["desc"])
    elif policy["type"] == "by_genre":
        min_words = policy.get("min_words", 0)
        if min_words and word_count and word_count < min_words:
            resolved["iv_reason"] = (
                f"原文约 {word_count} 词 < {min_words}，篇级改造收益低于时间成本，本板块暂不设 IV"
            )
            resolved["deviations"].append(resolved["iv_reason"])
        else:
            form = None
            for genre in registry["genres"]:
                if genre["id"] == genre_id:
                    form = genre["default_iv"]
                    break
            if not form:
                form = "IV-A"
                resolved["deviations"].append("体裁未确认，按说明文默认 IV-A；请在 QA 中记录最终选择")
            modules.append(form)
            resolved["modules"] = modules
            resolved["iv_form"] = form
            resolved["iv_reason"] = f"体裁 {genre_id or '未定'} → {form}"
            variant = pick_variant(form, section["id"], registry)
            if variant:
                resolved["iv_variant"] = variant["id"]
                resolved["iv_variant_title"] = variant["title"]
                resolved["iv_variant_reason"] = variant["desc"]

    for module_id in resolved["modules"]:
        module = registry["modules"].get(module_id)
        if not module:
            raise KeyError(f"Registry missing module id: {module_id}")
        resolved["module_titles"].append(module["title"])
    return resolved


def sectionize(text: str, registry: dict) -> dict:
    lines, offsets = split_lines(text)
    markers, source = find_markers(lines, offsets, registry)
    by_id = {section["id"]: section for section in registry["sections"]}
    sections = []

    for index, marker in enumerate(markers):
        end = markers[index + 1]["start"] if index + 1 < len(markers) else len(text)
        content = text[marker["start"]:end].strip()
        section_def = by_id[marker["id"]]

        body = content[len(marker["marker"]):] if content.startswith(marker["marker"]) else content
        word_count = len(re.findall(r"[A-Za-z']+", body))
        genre = guess_genre(body)
        route = resolve_route(section_def, registry, genre["guess"], word_count)

        sections.append(
            {
                "id": marker["id"],
                "label": marker["label"],
                "marker": marker["marker"],
                "line": marker.get("line"),
                "start": marker["start"],
                "end": end,
                "char_count": len(content),
                "word_count": word_count,
                "preview": re.sub(r"\s+", " ", content[:160]),
                "genre": genre,
                "stars": section_def.get("default_stars", 2),
                "route": route["modules"],
                "module_titles": route["module_titles"],
                "iv_form": route["iv_form"],
                "iv_reason": route["iv_reason"],
                "iv_variant": route["iv_variant"],
                "iv_variant_title": route["iv_variant_title"],
                "iv_variant_reason": route["iv_variant_reason"],
                "deviations": route["deviations"],
            }
        )

    expected = [section["id"] for section in registry["sections"]]
    found = [item["id"] for item in sections]
    warnings = []
    if not sections:
        warnings.append("No known section heading detected; segment manually.")
    if len(sections) < 4:
        warnings.append("Fewer than four sections detected; verify that the source is complete and headings survived extraction.")
    for section in sections:
        if section["char_count"] < 80:
            warnings.append(f"{section['label']} is very short; verify its boundary or source completeness.")
        if section["genre"]["confidence"] < 0.45:
            warnings.append(
                f"{section['label']} 体裁猜测偏弱（{section['genre']['guess']} vs {section['genre']['runner_up']}），必须读原文确认"
            )

    return {
        "schema_version": 2,
        "registry_version": registry["registry_version"],
        "detection_source": source,
        "detected_sections": sections,
        "detected_ids": found,
        "not_detected_ids": [item for item in expected if item not in found],
        "warnings": warnings,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--registry", type=Path, default=REGISTRY_PATH)
    args = parser.parse_args()

    try:
        registry = load_registry(args.registry)
        text = normalize(read_input(args.input))
        result = sectionize(text, registry)
        result["source"] = str(args.input.resolve())
        result["source_chars"] = len(text)
    except Exception as exc:  # concise CLI failure
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2

    payload = json.dumps(result, ensure_ascii=False, indent=2)
    if args.output:
        args.output.write_text(payload + "\n", encoding="utf-8")
    else:
        print(payload)
    return 0 if result["detected_sections"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
