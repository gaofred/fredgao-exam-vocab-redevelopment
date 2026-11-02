#!/usr/bin/env python3
"""Render the student / teacher DOCX pair from a content model.

Implements the formatting rules in references/output-contract.md section 4 and
the rendering notes in section 7:

  * no Word built-in heading styles, no outline level, no TOC, no auto numbering
  * hierarchy is bold + flush left only
  * answer lines are not bold and sit right after the item
  * a whole-line underscore item becomes a paragraph bottom border, so it always
    spans the text width and can never wrap
  * "____1____" (numbered blank) is preserved verbatim
  * snapToGrid is disabled, otherwise 1.5 line spacing renders as ~2.5 lines
  * section notes and block notes go to the teacher edition only

Requires python-docx. If it is unavailable, render by hand following
references/output-contract.md section 7 instead of skipping the formatting rules.

Usage:
    python scripts/render_packet.py content.json --outdir out --stem "襄阳四中2026级高一10月_二次开发"
"""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

try:
    from docx import Document
    from docx.enum.section import WD_ORIENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm, Pt
except ImportError:  # pragma: no cover
    print(
        "ERROR: python-docx is required. Install with `pip install python-docx`, "
        "or render by hand following references/output-contract.md section 7.",
        file=sys.stderr,
    )
    raise SystemExit(2)

BODY_PT = 11.5
TITLE_PT = 16
SECTION_PT = 15
MODULE_PT = 13

# Inline blanks get a uniform width; numbered blanks ("____1____") are left
# untouched. The lookarounds must also exclude "_", otherwise `{3,}` backtracks
# to three underscores and rewrites the head of a numbered blank.
INLINE_BLANK = re.compile(r"(?<![\d_])_{3,}(?![\d_])")
RULE_ITEM_MIN = 10


def set_run_font(run, size: float, bold: bool = False) -> None:
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.name = "Times New Roman"
    rpr = run._element.get_or_add_rPr()
    fonts = rpr.find(qn("w:rFonts"))
    if fonts is None:
        fonts = rpr.makeelement(qn("w:rFonts"), {})
        rpr.append(fonts)
    fonts.set(qn("w:ascii"), "Times New Roman")
    fonts.set(qn("w:hAnsi"), "Times New Roman")
    fonts.set(qn("w:eastAsia"), "宋体")


def base_para(doc):
    para = doc.add_paragraph()
    fmt = para.paragraph_format
    fmt.space_before = Pt(0)
    fmt.space_after = Pt(0)
    fmt.line_spacing = 1.5
    fmt.first_line_indent = Pt(0)
    ppr = para._p.get_or_add_pPr()
    if ppr.find(qn("w:snapToGrid")) is None:
        snap = OxmlElement("w:snapToGrid")
        snap.set(qn("w:val"), "0")
        ppr.append(snap)
    return para


def add_para(doc, text: str = "", size: float = BODY_PT, bold: bool = False, align=None):
    para = base_para(doc)
    if align is not None:
        para.alignment = align
    run = para.add_run(text)
    set_run_font(run, size, bold)
    return para


def add_rule_line(doc):
    para = base_para(doc)
    ppr = para._p.get_or_add_pPr()
    pbdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "000000")
    pbdr.append(bottom)
    ppr.append(pbdr)
    run = para.add_run(" ")
    set_run_font(run, BODY_PT)
    return para


def is_rule(text: str) -> bool:
    stripped = text.strip()
    return bool(stripped) and set(stripped) <= {"_", " "} and len(stripped) >= RULE_ITEM_MIN


def setup_document() -> "Document":
    doc = Document()
    section = doc.sections[0]
    section.orientation = WD_ORIENT.PORTRAIT
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2.5)
    section.bottom_margin = Cm(2.5)
    section.left_margin = Cm(2.2)
    section.right_margin = Cm(2.2)

    style = doc.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(BODY_PT)
    style.element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    return doc


def build(content: dict, teacher: bool, path: Path, edition_label: str = "整卷二次开发") -> None:
    doc = setup_document()
    edition = "教师版" if teacher else "学生版"
    add_para(doc, f"{content['title']}　{edition_label}（{edition}）", TITLE_PT, True, WD_ALIGN_PARAGRAPH.CENTER)
    if content.get("source"):
        add_para(doc, f"材料来源：{content['source']}")
    if content.get("usage"):
        add_para(doc, "使用说明：")
        for line in content["usage"]:
            add_para(doc, f"· {line}")
    add_para(doc, "")

    for sec in content["sections"]:
        stars = f"（{sec['stars']}｜{sec['cefr']}）" if sec.get("cefr") else f"（{sec['stars']}）"
        add_para(doc, f"{sec['label']}{stars}", SECTION_PT, True)
        if teacher and (sec.get("genre") or sec.get("note")):
            meta = "　".join(part for part in (
                f"体裁：{sec['genre']}" if sec.get("genre") else "",
                f"说明：{sec['note']}" if sec.get("note") else "",
            ) if part)
            add_para(doc, meta)

        for block in sec["blocks"]:
            add_para(doc, f"{block['title']}（{block['stars']}）", MODULE_PT, True)
            if teacher and block.get("note"):
                add_para(doc, f"【改编说明】{block['note']}")
            for item in block.get("items", []):
                if is_rule(item):
                    add_rule_line(doc)
                else:
                    add_para(doc, INLINE_BLANK.sub("__________", item))
            for _ in range(int(block.get("lines", 0) or 0)):
                add_rule_line(doc)
            if teacher:
                label = "【参考表达】" if block.get("mode") == "open" else "【答案与点拨】"
                add_para(doc, label)
                for answer in block.get("answers", []):
                    add_para(doc, answer)
            add_para(doc, "")

    doc.save(path)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("content", type=Path, help="content.json produced in step 3")
    parser.add_argument("--outdir", type=Path, required=True)
    parser.add_argument("--stem", required=True, help="file stem, e.g. 襄阳四中2026级高一10月_二次开发")
    args = parser.parse_args()

    try:
        content = json.loads(args.content.read_text(encoding="utf-8"))
    except Exception as exc:
        print(f"ERROR: cannot read content model: {exc}", file=sys.stderr)
        return 2

    missing = [key for key in ("title", "sections") if key not in content]
    if missing:
        print(f"ERROR: content model missing required fields: {missing}", file=sys.stderr)
        return 2

    args.outdir.mkdir(parents=True, exist_ok=True)
    student = args.outdir / f"{args.stem}（学生版）.docx"
    teacher = args.outdir / f"{args.stem}（教师版）.docx"
    build(content, False, student)
    build(content, True, teacher)
    print(f"student: {student}")
    print(f"teacher: {teacher}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
