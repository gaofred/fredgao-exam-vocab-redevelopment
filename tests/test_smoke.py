#!/usr/bin/env python3
"""End-to-end smoke test.

Guards the things that broke in practice:

  * a numbered blank "____1____" must survive rendering verbatim;
  * the student edition must never contain answer markers;
  * the validator must PASS a clean pair and FAIL a leaked pair.

Run from the repo root: python tests/test_smoke.py
"""

from __future__ import annotations

import json
import re
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))

try:
    import render_packet as R
except ImportError:  # pragma: no cover
    print("SKIP: python-docx not installed")
    raise SystemExit(0)

SECTION_LABEL = "七选五"
MODULE_TITLE = "I. 词汇变形与拓展"

CONTENT = {
    "title": "冒烟测试卷",
    "source": "自动生成，仅用于测试",
    "usage": ["测试用例"],
    "sections": [
        {
            "label": SECTION_LABEL,
            "stars": "★★☆☆",
            "cefr": "B1–B2",
            "genre": "说明（测试）",
            "note": "这条备注只应出现在教师版",
            "blocks": [
                {
                    "title": MODULE_TITLE,
                    "stars": "★★☆☆",
                    "mode": "closed",
                    "items": [
                        "1. emit（v. 发出）→ ______（n.）",
                        "2. 语段：The system ____1____ (know) as evapotranspiration.",
                    ],
                    "answers": [
                        "1. 答：emission　辨：-sion 名词化　错：误作 emitment　换：reduce emissions",
                        "2. 答：known　辨：be known as 被动　错：误填 knowing　换：is known as",
                    ],
                }
            ],
        }
    ],
}

MANIFEST = {
    "schema_version": 2,
    "detected_sections": [
        {
            "id": "seven-choose-five",
            "label": SECTION_LABEL,
            "module_titles": [MODULE_TITLE],
            "route": ["vocab-family"],
            "iv_form": None,
        }
    ],
}


def read_docx_text(path: Path) -> str:
    import validate_packet as V

    paragraphs, _, _ = V.load_edition(path)
    return V.text_of(paragraphs)


def check(condition: bool, message: str, failures: list[str]) -> None:
    if condition:
        print(f"  ok   {message}")
    else:
        print(f"  FAIL {message}")
        failures.append(message)


def run_validator(manifest: Path, student: Path, teacher: Path) -> int:
    return subprocess.run(
        [
            sys.executable,
            str(ROOT / "scripts" / "validate_packet.py"),
            "--manifest", str(manifest),
            "--student", str(student),
            "--teacher", str(teacher),
            "--output", str(manifest.parent / "QA.json"),
        ],
        capture_output=True,
        text=True,
    ).returncode


def main() -> int:
    failures: list[str] = []
    with tempfile.TemporaryDirectory() as tmp:
        work = Path(tmp)
        student = work / "学生版.docx"
        teacher = work / "教师版.docx"
        R.build(CONTENT, False, student)
        R.build(CONTENT, True, teacher)

        student_text = read_docx_text(student)
        teacher_text = read_docx_text(teacher)

        print("渲染：")
        check("____1____" in student_text, "编号空 ____1____ 存在", failures)
        # 关键断言：加宽后的 "__________1__________" 里仍含 "____1____" 子串，
        # 所以必须直接检查数字两侧的下划线个数，否则抓不到这个 bug。
        widened = re.search(r"_{5,}\d|\d_{5,}", student_text + teacher_text)
        check(widened is None, f"编号空未被加宽（命中：{widened.group(0) if widened else '无'}）", failures)
        check("____1____" in teacher_text, "教师版同样保留编号空", failures)
        check("答：" not in student_text, "学生版无答案标记", failures)
        check("【答案与点拨】" not in student_text, "学生版无答案区标题", failures)
        check("只应出现在教师版" not in student_text, "学生版无教研备注", failures)
        check("答：" in teacher_text, "教师版有答案标记", failures)
        check("只应出现在教师版" in teacher_text, "教师版有教研备注", failures)

        manifest = work / "manifest.json"
        manifest.write_text(json.dumps(MANIFEST, ensure_ascii=False), encoding="utf-8")

        print("校验器：")
        clean = run_validator(manifest, student, teacher)
        check(clean == 0, f"干净样本判定 PASS（退出码 {clean}）", failures)

        leaked_content = json.loads(json.dumps(CONTENT))
        leaked_content["sections"][0]["blocks"][0]["items"].append("3. 【答案与点拨】1. 答：x")
        leaked_content["sections"][0]["blocks"][0]["answers"].append("3. 答：x　辨：x　错：x　换：x")
        leaked_student = work / "泄漏版.docx"
        leaked_teacher = work / "泄漏版教师.docx"
        R.build(leaked_content, False, leaked_student)
        R.build(leaked_content, True, leaked_teacher)
        leaked = run_validator(manifest, leaked_student, leaked_teacher)
        check(leaked == 1, f"含泄漏样本判定 FAIL（退出码 {leaked}）", failures)

    print()
    if failures:
        print(f"FAILED ({len(failures)} 项)")
        return 1
    print("PASS")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
