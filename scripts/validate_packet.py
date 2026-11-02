#!/usr/bin/env python3
"""Run structural, closure and leakage checks on a paired redevelopment packet.

Checks implemented (v2):

  1. section parity against the manifest
  2. module parity against the registry titles carried by the manifest
  3. item-number closure: every student item number has a teacher answer
  4. teacher/student stem parity
  5. teacher answer-block completeness (答/辨/错/换 or 参/忌/换)
  6. student answer leakage in the body text
  7. deep leakage: headers, footers, comments, footnotes, endnotes,
     hidden text, coloured text, highlighting
  8. Word formatting rules: no built-in heading styles, no TOC field,
     no auto-numbering, no bold answer lines
  9. blank presence and adapted-cloze blank count

Titles are read from references/module-registry.json so that a renamed
module can never silently pass parity.
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
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}

FORBIDDEN_STUDENT = [
    r"参考答案",
    r"答案解析",
    r"二次开发答案",
    r"参考范文",
    r"【答案与点拨】",
    r"【参考表达】",
    r"【改编说明】",
    r"(?<!使用)说明：",
    r"正确(?:的)?顺序(?:应)?为\s*[A-EＡ-Ｅ]",
    r"故选\s*[A-DＡ-Ｄ]",
]

HEADING_STYLE_PATTERN = re.compile(r"^(heading\s*[1-9]|标题\s*[1-9]|title|subtitle|toc\s*heading)$", re.I)
ANSWER_LABEL_STARTS = ("答：", "辨：", "错：", "换：", "参：", "忌：")
ITEM_NUMBER = re.compile(r"^\s*(?:(\d{1,3})\s*[.．、)）]|[（(]\s*(\d{1,3})\s*[）)])")
ROMAN_HEADING = re.compile(r"^\s*(I{1,3}|IV|V)\s*[.．、]\s*(.+)$")
BLANK = re.compile(r"_{3,}")
NUMBERED_BLANK = re.compile(r"_{2,}\s*\d{1,3}\s*_{2,}")


# --------------------------------------------------------------------------
# registry
# --------------------------------------------------------------------------

def load_registry(path: Path = REGISTRY_PATH) -> dict:
    return json.loads(path.read_text(encoding="utf-8"))


def module_title_map(registry: dict) -> dict[str, dict]:
    return registry["modules"]


def resolve_module_titles(section: dict, registry: dict) -> list[str]:
    """Prefer manifest module_titles; fall back to module ids."""
    titles = section.get("module_titles")
    if titles:
        return list(titles)
    modules = registry["modules"]
    return [modules[mid]["title"] for mid in section.get("route", []) if mid in modules]


def module_answer_mode(title: str, registry: dict) -> str:
    for module in registry["modules"].values():
        if module["title"] == title:
            return module.get("answer_mode", "closed")
    return "closed"


# --------------------------------------------------------------------------
# docx reading
# --------------------------------------------------------------------------

def docx_parts(path: Path) -> dict[str, bytes]:
    parts: dict[str, bytes] = {}
    with zipfile.ZipFile(path) as archive:
        for name in archive.namelist():
            if not name.endswith(".xml"):
                continue
            if name.startswith("word/") or name == "docProps/core.xml":
                try:
                    parts[name] = archive.read(name)
                except KeyError:  # pragma: no cover
                    continue
    return parts


def part_text(xml_bytes: bytes) -> str:
    try:
        root = ET.fromstring(xml_bytes)
    except ET.ParseError:
        return ""
    chunks = []
    for node in root.iter(f"{{{W_NS}}}t"):
        if node.text:
            chunks.append(node.text)
    return "\n".join(chunks)


def parse_document(xml_bytes: bytes) -> list[dict]:
    """Return paragraph records with text, style flags and run properties."""
    root = ET.fromstring(xml_bytes)
    paragraphs = []
    for para in root.iter(f"{{{W_NS}}}p"):
        text = "".join(node.text or "" for node in para.iter(f"{{{W_NS}}}t"))
        ppr = para.find(f"{{{W_NS}}}pPr")
        style = None
        outline = False
        numbered = False
        if ppr is not None:
            pstyle = ppr.find(f"{{{W_NS}}}pStyle")
            if pstyle is not None:
                style = pstyle.get(f"{{{W_NS}}}val")
            if ppr.find(f"{{{W_NS}}}outlineLvl") is not None:
                outline = True
            if ppr.find(f"{{{W_NS}}}numPr") is not None:
                numbered = True

        runs = []
        for run in para.iter(f"{{{W_NS}}}r"):
            rpr = run.find(f"{{{W_NS}}}rPr")
            info = {"text": "".join(node.text or "" for node in run.iter(f"{{{W_NS}}}t"))}
            if rpr is not None:
                bold = rpr.find(f"{{{W_NS}}}b")
                info["bold"] = bold is not None and bold.get(f"{{{W_NS}}}val") not in {"0", "false"}
                vanish = rpr.find(f"{{{W_NS}}}vanish")
                info["vanish"] = vanish is not None and vanish.get(f"{{{W_NS}}}val") not in {"0", "false"}
                color = rpr.find(f"{{{W_NS}}}color")
                info["color"] = color.get(f"{{{W_NS}}}val") if color is not None else None
                highlight = rpr.find(f"{{{W_NS}}}highlight")
                info["highlight"] = highlight.get(f"{{{W_NS}}}val") if highlight is not None else None
            else:
                info.update({"bold": False, "vanish": False, "color": None, "highlight": None})
            if info["text"]:
                runs.append(info)

        paragraphs.append(
            {
                "text": text,
                "style": style,
                "outline": outline,
                "numbered": numbered,
                "runs": runs,
            }
        )
    return paragraphs


def read_docx(path: Path) -> tuple[list[dict], dict[str, bytes]]:
    parts = docx_parts(path)
    xml = parts.get("word/document.xml")
    if xml is None:
        raise ValueError(f"{path} has no word/document.xml")
    return parse_document(xml), parts


def read_plain(path: Path) -> tuple[list[dict], dict[str, bytes]]:
    text = path.read_text(encoding="utf-8-sig")
    paragraphs = [{"text": line, "style": None, "outline": False, "numbered": False, "runs": []} for line in text.splitlines()]
    return paragraphs, {}


def load_edition(path: Path) -> tuple[list[dict], dict[str, bytes], bool]:
    if path.suffix.lower() == ".docx":
        paragraphs, parts = read_docx(path)
        return paragraphs, parts, True
    if path.suffix.lower() in {".txt", ".md"}:
        paragraphs, parts = read_plain(path)
        return paragraphs, parts, False
    raise ValueError(f"Unsupported packet type: {path.suffix}")


# --------------------------------------------------------------------------
# helpers
# --------------------------------------------------------------------------

def norm(text: str) -> str:
    return re.sub(r"\s+", "", text)


def text_of(paragraphs: list[dict]) -> str:
    return "\n".join(item["text"] for item in paragraphs)


def strip_number(text: str) -> tuple[int | None, str]:
    match = ITEM_NUMBER.match(text)
    if not match:
        return None, text
    number = int(match.group(1) or match.group(2))
    return number, text[match.end():]


def find_paragraph(paragraphs: list[dict], needle: str, start: int = 0) -> int:
    target = norm(needle)
    if not target:
        return -1
    for index in range(start, len(paragraphs)):
        if target in norm(paragraphs[index]["text"]):
            return index
    return -1


def is_module_heading(text: str, registry: dict) -> bool:
    """True when the paragraph is actually a module heading, not a section label."""
    compact = norm(text)
    return any(norm(module["title"]) in compact for module in registry["modules"].values())


def find_section_paragraph(paragraphs: list[dict], label: str, registry: dict, start: int = 0) -> int:
    """Locate a section label without colliding with module headings or summaries.

    Two hazards are handled here:
      * "语法填空" must not match the module heading "IV. 语法填空改编";
      * a summary line such as "本卷处理范围：阅读A—D篇、七选五、完形填空……"
        mentions many labels at once and is not a section heading.

    Pass 1 accepts a paragraph that starts with the label. Pass 2 accepts a
    paragraph that contains it, carries no module title, and mentions exactly
    one section label.
    """
    target = norm(label)
    if not target:
        return -1
    all_labels = [norm(item["label"]) for item in registry["sections"] if norm(item["label"])]

    for index in range(start, len(paragraphs)):
        if norm(paragraphs[index]["text"]).startswith(target):
            return index

    for index in range(start, len(paragraphs)):
        compact = norm(paragraphs[index]["text"])
        if target not in compact:
            continue
        if is_module_heading(paragraphs[index]["text"], registry):
            continue
        if sum(1 for other in all_labels if other in compact) > 1:
            continue
        return index
    return -1


def section_ranges(paragraphs: list[dict], sections: list[dict], registry: dict) -> list[tuple[dict, int, int]]:
    marks = []
    for section in sections:
        index = find_section_paragraph(paragraphs, section["label"], registry)
        if index >= 0:
            marks.append((index, section))
    marks.sort(key=lambda item: item[0])
    ranges = []
    for position, (index, section) in enumerate(marks):
        end = marks[position + 1][0] if position + 1 < len(marks) else len(paragraphs)
        ranges.append((section, index, end))
    return ranges


def module_ranges(paragraphs: list[dict], start: int, end: int, titles: list[str]) -> list[tuple[str, int, int]]:
    marks = []
    for title in titles:
        index = find_paragraph(paragraphs, title, start)
        if 0 <= index < end:
            marks.append((index, title))
    marks.sort(key=lambda item: item[0])
    ranges = []
    for position, (index, title) in enumerate(marks):
        stop = marks[position + 1][0] if position + 1 < len(marks) else end
        ranges.append((title, index, stop))
    return ranges


def item_numbers(paragraphs: list[dict], start: int, end: int) -> tuple[list[int], list[tuple[int, str]]]:
    numbers = []
    items = []
    for index in range(start, end):
        number, rest = strip_number(paragraphs[index]["text"])
        if number is None:
            continue
        numbers.append(number)
        items.append((number, rest))
    return numbers, items


# --------------------------------------------------------------------------
# checks
# --------------------------------------------------------------------------

def check_section_parity(manifest_sections, student_paras, teacher_paras, errors, warnings, parity):
    student_text = norm(text_of(student_paras))
    teacher_text = norm(text_of(teacher_paras))
    for section in manifest_sections:
        label = section["label"]
        key = norm(label)
        in_student = key in student_text
        in_teacher = key in teacher_text
        parity.append({"id": section["id"], "label": label, "student": in_student, "teacher": in_teacher})
        if not in_student:
            errors.append(f"学生版缺少已识别板块：{label}")
        if not in_teacher:
            errors.append(f"教师版缺少已识别板块：{label}")


def check_module_parity(manifest_sections, student_paras, teacher_paras, registry, errors, parity):
    student_text = norm(text_of(student_paras))
    teacher_text = norm(text_of(teacher_paras))
    for section in manifest_sections:
        titles = resolve_module_titles(section, registry)
        recorded = []
        for title in titles:
            key = norm(title)
            in_student = key in student_text
            in_teacher = key in teacher_text
            recorded.append({"title": title, "student": in_student, "teacher": in_teacher})
            if not in_student:
                errors.append(f"学生版缺少板块：{section['label']} / {title}")
            if not in_teacher:
                errors.append(f"教师版缺少板块：{section['label']} / {title}")
        parity.append({"id": section["id"], "label": section["label"], "modules": recorded})


def check_numbering(manifest_sections, student_paras, teacher_paras, registry, errors, warnings, report):
    s_ranges = section_ranges(student_paras, manifest_sections, registry)
    t_ranges = {section["id"]: (start, end) for section, start, end in section_ranges(teacher_paras, manifest_sections, registry)}

    for section, s_start, s_end in s_ranges:
        titles = resolve_module_titles(section, registry)
        t_pair = t_ranges.get(section["id"])
        s_modules = module_ranges(student_paras, s_start, s_end, titles)
        t_modules = {}
        if t_pair:
            t_modules = {title: (a, b) for title, a, b in module_ranges(teacher_paras, t_pair[0], t_pair[1], titles)}

        for title, m_start, m_end in s_modules:
            s_nums, s_items = item_numbers(student_paras, m_start, m_end)
            if not s_nums:
                continue
            if s_nums != list(range(1, len(s_nums) + 1)):
                warnings.append(f"{section['label']} / {title} 学生版题号不连续：{s_nums}")

            entry = {"section": section["label"], "module": title, "student_items": len(s_nums)}
            t_range = t_modules.get(title)
            if not t_range:
                errors.append(f"教师版缺少板块区间，无法核对题号：{section['label']} / {title}")
                report.append({**entry, "status": "missing-teacher-block"})
                continue

            t_nums, _ = item_numbers(teacher_paras, t_range[0], t_range[1])
            entry["teacher_items"] = len(t_nums)
            missing = sorted(set(s_nums) - set(t_nums))
            extra = sorted(set(t_nums) - set(s_nums))
            if missing:
                errors.append(f"{section['label']} / {title} 教师版缺答案编号：{missing}")
            if extra:
                warnings.append(f"{section['label']} / {title} 教师版多出编号：{extra}（确认是否为补充说明）")
            entry["missing"] = missing
            entry["extra"] = extra
            entry["status"] = "ok" if not missing else "gap"
            report.append(entry)

            # stem parity: compare every non-blank chunk of the student stem
            # against the teacher block, with blanks removed on both sides.
            # Robust to a teacher that inlines answers, strict about rewriting.
            teacher_key_text = norm(BLANK.sub("", text_of(teacher_paras[t_range[0]:t_range[1]])))
            for number, rest in s_items:
                chunks = [norm(chunk) for chunk in BLANK.split(rest)]
                chunks = [chunk for chunk in chunks if len(chunk) >= 6]
                missing = [chunk for chunk in chunks if chunk not in teacher_key_text]
                if missing:
                    errors.append(
                        f"{section['label']} / {title} 第 {number} 题题干在教师版中缺失或被改写"
                        f"（未找到：{missing[0][:18]}…）"
                    )


def check_answer_blocks(manifest_sections, student_paras, teacher_paras, registry, errors, warnings):
    labels = registry["answer_labels"]
    t_ranges = {section["id"]: (start, end) for section, start, end in section_ranges(teacher_paras, manifest_sections, registry)}

    for section in manifest_sections:
        t_pair = t_ranges.get(section["id"])
        if not t_pair:
            continue
        titles = resolve_module_titles(section, registry)
        for title, m_start, m_end in module_ranges(teacher_paras, t_pair[0], t_pair[1], titles):
            mode = module_answer_mode(title, registry)
            spec = labels[mode]
            block = norm(text_of(teacher_paras[m_start:m_end]))
            if norm(spec["block"]) not in block:
                errors.append(f"教师版 {section['label']} / {title} 缺少答案区标记 {spec['block']}")
                continue
            missing = [label for label in spec["required"] if norm(label) not in block]
            if missing:
                errors.append(
                    f"教师版 {section['label']} / {title} 缺三件套要素 {missing}；"
                    f"禁止只给答案（{''.join(spec['required'])}）"
                )


def check_text_leakage(student_paras, errors, report):
    body = text_of(student_paras)
    hits = []
    for pattern in FORBIDDEN_STUDENT:
        match = re.search(pattern, body, flags=re.IGNORECASE)
        if match:
            hits.append({"pattern": pattern, "match": match.group(0)})
            errors.append(f"学生版正文出现答案泄漏标记：{match.group(0)}")
    report.extend(hits)


def check_deep_leakage(student_parts, student_is_docx, errors, warnings, report):
    if not student_is_docx:
        warnings.append("输入为纯文本，页眉页脚/批注/隐藏文字/彩色字体等深层泄漏检查已跳过")
        return

    sensitive_parts = [
        name for name in student_parts
        if re.match(r"word/(header\d*|footer\d*|comments|footnotes|endnotes)\.xml$", name)
    ]
    for name in sensitive_parts:
        text = part_text(student_parts[name])
        if not text.strip():
            continue
        for pattern in FORBIDDEN_STUDENT:
            match = re.search(pattern, text, flags=re.IGNORECASE)
            if match:
                errors.append(f"学生版 {name} 中存在答案泄漏：{match.group(0)}")
                report.append({"part": name, "issue": "forbidden-text", "match": match.group(0)})
        if name.endswith("comments.xml"):
            errors.append("学生版带有批注（comments.xml 非空），批注可承载答案")
            report.append({"part": name, "issue": "comments-present"})

    xml = student_parts.get("word/document.xml")
    if not xml:
        return
    root = ET.fromstring(xml)
    colours = set()
    highlights = set()
    hidden = 0
    for run in root.iter(f"{{{W_NS}}}r"):
        rpr = run.find(f"{{{W_NS}}}rPr")
        if rpr is None:
            continue
        vanish = rpr.find(f"{{{W_NS}}}vanish")
        if vanish is not None and vanish.get(f"{{{W_NS}}}val") not in {"0", "false"}:
            hidden += 1
        colour = rpr.find(f"{{{W_NS}}}color")
        if colour is not None:
            value = (colour.get(f"{{{W_NS}}}val") or "").upper()
            if value and value not in {"AUTO", "000000", "FFFFFF00"}:
                colours.add(value)
        highlight = rpr.find(f"{{{W_NS}}}highlight")
        if highlight is not None:
            value = (highlight.get(f"{{{W_NS}}}val") or "").lower()
            if value not in {"", "none"}:
                highlights.add(value)

    if hidden:
        errors.append(f"学生版存在 {hidden} 处隐藏文字（vanish），可承载答案")
        report.append({"issue": "hidden-text", "count": hidden})
    red = [value for value in colours if value.startswith("FF00") or value in {"FF0000", "C00000", "E36C0A"}]
    if red:
        errors.append(f"学生版存在红色系字体 {sorted(red)}，按约定视为答案标记")
        report.append({"issue": "red-font", "values": sorted(red)})
    others = sorted(colours - set(red))
    if others:
        warnings.append(f"学生版存在非黑色字体 {others}，确认不是答案强调")
    if highlights:
        warnings.append(f"学生版存在高亮 {sorted(highlights)}，确认不是答案强调")


def check_formatting(student_paras, teacher_paras, student_is_docx, teacher_is_docx, errors, warnings, report):
    for name, paragraphs, is_docx in (("学生版", student_paras, student_is_docx), ("教师版", teacher_paras, teacher_is_docx)):
        heading_hits = []
        outline_hits = []
        numbered_hits = []
        for index, para in enumerate(paragraphs):
            style = para.get("style") or ""
            if style and HEADING_STYLE_PATTERN.match(style.strip()):
                heading_hits.append((index, para["text"][:24], style))
            if para.get("outline"):
                outline_hits.append((index, para["text"][:24]))
            if para.get("numbered"):
                numbered_hits.append((index, para["text"][:24]))

        if heading_hits:
            errors.append(
                f"{name}使用了 Word 内置标题样式 {sorted({item[2] for item in heading_hits})}"
                f"（共 {len(heading_hits)} 段）；层级只能用加粗+顶格"
            )
            report.append({"edition": name, "issue": "heading-style", "samples": heading_hits[:3]})
        if outline_hits:
            errors.append(f"{name}存在大纲级别（outlineLvl）{len(outline_hits)} 处，等同于标题样式")
        if numbered_hits:
            warnings.append(f"{name}存在自动编号（numPr）{len(numbered_hits)} 处，改样式后会重排")

    if student_is_docx:
        bold_answer = [
            para["text"][:24]
            for para in student_paras
            if para["text"].strip().startswith(ANSWER_LABEL_STARTS)
        ]
        if bold_answer:
            errors.append(f"学生版出现答案行 {bold_answer[:2]}，答案只能进教师版")
    for para in teacher_paras:
        text = para["text"].strip()
        if not text.startswith(ANSWER_LABEL_STARTS):
            continue
        runs = [run for run in para.get("runs", []) if run["text"].strip()]
        if runs and all(run.get("bold") for run in runs):
            warnings.append(f"教师版答案行整段加粗：{text[:20]}（约定答案行不加粗）")


def check_toc_field(parts, name, errors, report):
    xml = parts.get("word/document.xml")
    if not xml:
        return
    try:
        root = ET.fromstring(xml)
    except ET.ParseError:
        return
    for node in root.iter(f"{{{W_NS}}}instrText"):
        if node.text and "TOC" in node.text:
            errors.append(f"{name}存在目录域（TOC），会随样式变动而错乱")
            report.append({"edition": name, "issue": "toc-field"})
            return


def check_title_aliases(student_paras, teacher_paras, registry, errors, report):
    aliases = registry.get("forbidden_title_aliases", [])
    if not aliases:
        return
    for name, paragraphs in (("学生版", student_paras), ("教师版", teacher_paras)):
        for para in paragraphs:
            match = ROMAN_HEADING.match(para["text"])
            if not match:
                continue
            tail = match.group(2)
            for alias in aliases:
                if alias in tail:
                    errors.append(f"{name}板块标题使用了废弃别名「{alias}」：{para['text'][:30]}")
                    report.append({"edition": name, "issue": "legacy-title", "alias": alias})


def count_blanks(text: str) -> int:
    """Count answer blanks, treating "____1____" as one numbered blank."""
    masked = NUMBERED_BLANK.sub(" ", text)
    return len(NUMBERED_BLANK.findall(text)) + len(BLANK.findall(masked))


def check_blanks(manifest_sections, student_paras, teacher_paras, registry, errors, warnings):
    student_paras_text = text_of(student_paras)
    if not BLANK.search(student_paras_text):
        warnings.append("学生版未见作答横线（连续三个以上下划线）")

    s_ranges = {section["id"]: (start, end) for section, start, end in section_ranges(student_paras, manifest_sections, registry)}
    for section in manifest_sections:
        if section.get("iv_form") != "IV-A":
            continue
        pair = s_ranges.get(section["id"])
        if not pair:
            continue
        titles = resolve_module_titles(section, registry)
        for title, m_start, m_end in module_ranges(student_paras, pair[0], pair[1], titles):
            if not title.startswith("IV"):
                continue
            count = count_blanks(text_of(student_paras[m_start:m_end]))
            if count != 8:
                warnings.append(f"{section['label']} / {title} 空格数为 {count}，约定 8 空")


# --------------------------------------------------------------------------
# main
# --------------------------------------------------------------------------

def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--manifest", type=Path, required=True)
    parser.add_argument("--student", type=Path, required=True)
    parser.add_argument("--teacher", type=Path, required=True)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--registry", type=Path, default=REGISTRY_PATH)
    parser.add_argument("--fail-on-warning", action="store_true", help="把 warning 也视为失败")
    args = parser.parse_args()

    errors: list[str] = []
    warnings: list[str] = []
    try:
        registry = load_registry(args.registry)
        manifest = json.loads(args.manifest.read_text(encoding="utf-8"))
        student_paras, student_parts, student_is_docx = load_edition(args.student)
        teacher_paras, teacher_parts, teacher_is_docx = load_edition(args.teacher)
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2

    sections = manifest.get("detected_sections", [])
    if not sections:
        errors.append("manifest 中未检测到任何板块，无法核对")

    section_parity: list[dict] = []
    check_section_parity(sections, student_paras, teacher_paras, errors, warnings, section_parity)
    check_module_parity(sections, student_paras, teacher_paras, registry, errors, section_parity)

    numbering: list[dict] = []
    check_numbering(sections, student_paras, teacher_paras, registry, errors, warnings, numbering)

    check_answer_blocks(sections, student_paras, teacher_paras, registry, errors, warnings)

    leakage: list[dict] = []
    check_text_leakage(student_paras, errors, leakage)
    deep: list[dict] = []
    check_deep_leakage(student_parts, student_is_docx, errors, warnings, deep)

    formatting: list[dict] = []
    check_formatting(student_paras, teacher_paras, student_is_docx, teacher_is_docx, errors, warnings, formatting)
    check_toc_field(student_parts, "学生版", errors, formatting)
    check_toc_field(teacher_parts, "教师版", errors, formatting)
    check_title_aliases(student_paras, teacher_paras, registry, errors, formatting)

    check_blanks(sections, student_paras, teacher_paras, registry, errors, warnings)

    student_len = len(text_of(student_paras))
    teacher_len = len(text_of(teacher_paras))
    if teacher_len <= student_len:
        warnings.append("教师版不长于学生版，确认答案与分析是否真的补进去了")

    passed = not errors and not (args.fail_on_warning and warnings)
    report = {
        "schema_version": 2,
        "registry_version": registry["registry_version"],
        "status": "PASS" if passed else "FAIL",
        "manifest": str(args.manifest.resolve()),
        "student": str(args.student.resolve()),
        "teacher": str(args.teacher.resolve()),
        "section_parity": section_parity,
        "numbering_closure": numbering,
        "student_leakage": leakage,
        "deep_leakage": deep,
        "formatting": formatting,
        # Filled by the agent, not by this script: source ambiguities the
        # checker cannot judge, and whether any boundary was cut by hand.
        "ambiguities": [],
        "manual_segment": False,
        "omitted_sections": [],
        "iv_variants": {
            section["label"]: section.get("iv_variant_title")
            for section in sections
            if section.get("iv_variant_title")
        },
        "length": {"student_chars": student_len, "teacher_chars": teacher_len},
        "errors": errors,
        "warnings": warnings,
    }
    payload = json.dumps(report, ensure_ascii=False, indent=2)
    if args.output:
        args.output.write_text(payload + "\n", encoding="utf-8")
    else:
        print(payload)
    return 0 if passed else 1


if __name__ == "__main__":
    raise SystemExit(main())
