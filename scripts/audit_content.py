#!/usr/bin/env python3
"""内容层自审：查结构校验查不出的那几类错。

validate_packet.py 管的是形式（对齐、闭合、泄漏、排版）。这个脚本管内容：

  1. 原句回流 —— 句式层引的原句必须能在原文里逐字找到（归一化后子串匹配）
  2. 圈号完整 —— ①②③ 不许合并到一行，不许断号
  3. 构词自洽 —— 「辨」里声称的变换必须与源词、答案词的实际词形相符
  4. 答案对齐 —— 封闭题答案条数 = 题号数

用法：
    python scripts/audit_content.py content.json --exam 原卷.docx
    python scripts/audit_content.py content.json --exam 原卷.docx --report audit.json
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import unicodedata
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
CIRCLED = "①②③④⑤⑥⑦⑧⑨⑩"
CIRCLED_INDEX = {ch: i + 1 for i, ch in enumerate(CIRCLED)}
QUOTE_PAIRS = [("“", "”"), ("「", "」"), ("\"", "\"")]


def read_text(path: Path) -> str:
    if path.suffix.lower() == ".docx":
        with zipfile.ZipFile(path) as zf:
            root = ET.fromstring(zf.read("word/document.xml"))
        paras = ["".join(node.text or "" for node in p.iter(W + "t")) for p in root.iter(W + "p")]
        return "\n".join(paras)
    return path.read_text(encoding="utf-8", errors="ignore")


def normalize(text: str) -> str:
    """去掉空白、标点、全角差异，只留字母数字与汉字，便于跨排版比对。"""
    text = unicodedata.normalize("NFKC", text)
    text = text.replace("’", "'").replace("‘", "'").replace("“", '"').replace("”", '"')
    # 原卷常在生词后加中文注释，如「垃圾桶 lid（盖子）」；括号内容两边一致去掉再比。
    text = re.sub(r"[（(][^）)]*[）)]", "", text)
    return re.sub(r"[^0-9a-zA-Z\u4e00-\u9fff]", "", text).lower()


def source_sentence(item: str) -> str | None:
    """从『原句：……（A篇 第8段）』里取出句子本体。"""
    if "原句：" not in item:
        return None
    body = item.split("原句：", 1)[1]
    body = re.sub(r"（[A-D]篇[^）]*）\s*$", "", body).strip()
    return body or None


def intro_word(item: str) -> tuple[str, str] | None:
    """从『1. available（adj. 可获得的）→ ______（n. 可获得性）』取 (源词, 目标词形)。"""
    match = re.match(r"\s*\d+\.\s*([A-Za-z][A-Za-z'\-]*)", item)
    if not match:
        return None
    return match.group(1), ""


def answer_pair(answer: str, source: str) -> tuple[str, str] | None:
    """从『1. 答：availability　辨：…』取 (答案词, 辨正文)。"""
    match = re.search(r"答：([A-Za-z][A-Za-z'\-]*)", answer)
    if not match:
        return None
    target = match.group(1)
    rest = answer[match.end():]
    return target, rest


DERIVATION_RULES = [
    # (辨里的说法, 校验函数：源词, 目标词 → 是否自洽)
    (r"y\s*变\s*i", lambda s, t: s.endswith("y")),
    (r"去\s*(?:掉)?\s*词?尾?\s*e", lambda s, t: s.endswith("e")),
    (r"e\s*结尾保留", lambda s, t: s.endswith("e") and t.endswith("e")),
    (r"结尾\s*e\s*保留", lambda s, t: s.endswith("e") and t.endswith("e")),
    (r"双写", lambda s, t: len(s) >= 2 and s[-1] == s[-2] or (s[-1] * 2 in t)),
    (r"-able\s*变\s*-ability", lambda s, t: s.endswith("able") and t.endswith("ability")),
    (r"-ble\s*变\s*-bility", lambda s, t: s.endswith("ble") and t.endswith("bility")),
    (r"-ity\s*变\s*-ify", lambda s, t: s.endswith("ity") and t.endswith("ify")),
]


CLOZE_LIKE = ("完形填空", "语法填空")
ELLIPSIS = ("……", "…", "...")
TERMINATORS = set("。.?!；;”\"’'")


def tolerant_pattern(quote: str) -> str:
    """引句 → 容错正则：词间允许空白与生词注释括号（全角半角皆可）。"""
    tokens = [re.escape(tok) for tok in re.split(r"\s+", quote.strip()) if tok]
    return r"\s*(?:[（(][^）)]*[）)]\s*)?".join(tokens)


def check_sources(content: dict, exam: str, errors: list[str], warnings: list[str]) -> int:
    norm_exam = normalize(exam)
    checked = 0
    for section in content["sections"]:
        for block in section["blocks"]:
            for item in block.get("items", []):
                sentence = source_sentence(item)
                if not sentence:
                    continue
                checked += 1
                if normalize(sentence) not in norm_exam:
                    if section["label"] in CLOZE_LIKE:
                        warnings.append(
                            f"[原句回流] {section['label']}／{block['title']}：原文带空，引句已还原空位，需人工核对 —— {sentence[:40]}…"
                        )
                    else:
                        errors.append(
                            f"[原句回流] {section['label']}／{block['title']}：原文中找不到 —— {sentence[:60]}…"
                        )
                    continue

                # 句子完整性：引句自身已断句即通过；否则看原文该处是否也在此断句
                tail = sentence.rstrip()
                if tail.endswith(ELLIPSIS) or (tail and tail[-1] in TERMINATORS):
                    continue
                raw = re.search(tolerant_pattern(sentence), exam)
                if not raw:
                    continue
                after = exam[raw.end():].lstrip()
                if after and after[0] not in TERMINATORS:
                    errors.append(
                        f"[原句回流] {section['label']}／{block['title']}：引句在原文中未断句，须整句引用或补省略号 —— {sentence[:50]}…"
                    )
    return checked


def check_circled(content: dict, errors: list[str]) -> int:
    checked = 0
    for section in content["sections"]:
        for block in section["blocks"]:
            group: list[int] = []
            group_start = ""
            for item in block.get("items", []):
                found = [CIRCLED_INDEX[ch] for ch in item if ch in CIRCLED_INDEX]
                if len(set(found)) >= 2:
                    errors.append(
                        f"[圈号] {section['label']}／{block['title']}：一行出现多个圈号 {''.join(CIRCLED[i-1] for i in sorted(set(found)))} —— {item.strip()[:40]}"
                    )
                if re.match(r"\s*\d+\.\s", item):
                    if group:
                        checked += 1
                        expected = list(range(1, len(group) + 1))
                        if sorted(group) != expected:
                            errors.append(
                                f"[圈号] {section['label']}／{block['title']}：{group_start} 圈号应为 ①②③…，实际 {''.join(CIRCLED[i-1] for i in sorted(group))}"
                            )
                    group = []
                    group_start = item.strip()[:24]
                group.extend(found)
            if group:
                checked += 1
                expected = list(range(1, len(group) + 1))
                if sorted(group) != expected:
                    errors.append(
                        f"[圈号] {section['label']}／{block['title']}：{group_start} 圈号应为 ①②③…，实际 {''.join(CIRCLED[i-1] for i in sorted(group))}"
                    )
    return checked


def check_derivation(content: dict, errors: list[str]) -> int:
    checked = 0
    for section in content["sections"]:
        for block in section["blocks"]:
            items = block.get("items", [])
            answers = block.get("answers", [])
            for item, answer in zip(items, answers):
                pair = intro_word(item)
                if not pair:
                    continue
                source = pair[0]
                parsed = answer_pair(answer, source)
                if not parsed:
                    continue
                target, note = parsed
                for pattern, ok in DERIVATION_RULES:
                    if not re.search(pattern, note):
                        continue
                    checked += 1
                    if not ok(source, target):
                        errors.append(
                            f"[构词] {section['label']}／{block['title']}：{source} → {target}，"
                            f"但「辨」声称「{re.search(pattern, note).group(0)}」，与词形不符"
                        )
                    break
    return checked


def check_answer_count(content: dict, errors: list[str]) -> int:
    """按题号比对，不按行数——开放题的答案可能跨多行。"""
    checked = 0
    for section in content["sections"]:
        for block in section["blocks"]:
            numbered = sum(1 for item in block.get("items", []) if re.match(r"\s*\d+\.\s", item))
            if numbered == 0:
                continue
            checked += 1
            covered: set[str] = set()
            for answer in block.get("answers", []):
                match = re.match(r"\s*(\d+)\.", answer)
                if match:
                    covered.add(match.group(1))
            if len(covered) != numbered:
                errors.append(
                    f"[答案对齐] {section['label']}／{block['title']}：题号 {numbered} 个，答案覆盖 {len(covered)} 个"
                )
    return checked


def main() -> int:
    parser = argparse.ArgumentParser(description="内容层自审")
    parser.add_argument("content", type=Path)
    parser.add_argument("--exam", type=Path, help="原卷 docx 或 txt，用于原句回流比对")
    parser.add_argument("--report", type=Path, help="写出 JSON 报告")
    args = parser.parse_args()

    content = json.loads(args.content.read_text(encoding="utf-8"))
    errors: list[str] = []
    warnings: list[str] = []
    stats: dict[str, int] = {}

    if args.exam:
        stats["原句"] = check_sources(content, read_text(args.exam), errors, warnings)
    stats["圈号"] = check_circled(content, errors)
    stats["构词"] = check_derivation(content, errors)
    stats["答案对齐"] = check_answer_count(content, errors)

    print("内容自审：" + "　".join(f"{k} {v} 项" for k, v in stats.items()))
    for line in errors:
        print("  E -", line)
    for line in warnings:
        print("  W ·", line)
    if args.report:
        args.report.write_text(json.dumps({"stats": stats, "errors": errors, "warnings": warnings}, ensure_ascii=False, indent=2), encoding="utf-8")

    if errors:
        print(f"FAILED（{len(errors)} 项）")
        return 1
    print("PASS")
    return 0


if __name__ == "__main__":
    sys.exit(main())
