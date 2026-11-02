#!/usr/bin/env python3
"""Guard the single source of truth for module naming.

Fails when:
  * the registry is internally inconsistent (unknown module id, bad genre link);
  * a canonical module title is missing from references/section-blueprints.md;
  * a legacy alias reappears in any shipped document;
  * a script hard-codes a module title instead of reading the registry.

Run this after editing the registry or the docs, before delivery.
"""

from __future__ import annotations

import ast
import json
import re
import sys
from pathlib import Path

SKILL_ROOT = Path(__file__).resolve().parent.parent
REGISTRY_PATH = SKILL_ROOT / "references" / "module-registry.json"
DOCS = [
    SKILL_ROOT / "SKILL.md",
    SKILL_ROOT / "references" / "section-blueprints.md",
    SKILL_ROOT / "references" / "genre-routing.md",
    SKILL_ROOT / "references" / "quality-rules.md",
    SKILL_ROOT / "references" / "output-contract.md",
]
SCRIPTS = [
    SKILL_ROOT / "scripts" / "exam_sectionizer.py",
    SKILL_ROOT / "scripts" / "validate_packet.py",
    SKILL_ROOT / "scripts" / "render_packet.py",
]
VALID_ANSWER_MODES = {"closed", "open"}
# Scripts that resolve module titles themselves; render_packet only consumes the
# content model, so it is exempt from the registry-reference rule.
REGISTRY_CONSUMERS = {"exam_sectionizer.py", "validate_packet.py"}
# tolerate "### IV. xxx", "**IV. xxx**", "- IV. xxx"
HEADING_LINE = re.compile(r"^\s*(?:[#>*\-]+\s*)*\*{0,2}\s*(I{1,3}|IV|V)\s*[.．、]\s*(.+?)\s*\*{0,2}\s*$")


def code_literals(path: Path) -> str:
    """String literals that are not docstrings.

    A title mentioned inside a docstring or a comment is documentation;
    a title baked into a real string value is a hard-code.
    """
    source = path.read_text(encoding="utf-8")
    try:
        tree = ast.parse(source)
    except SyntaxError:
        return source
    docstrings = set()
    for node in ast.walk(tree):
        if isinstance(node, (ast.Module, ast.FunctionDef, ast.AsyncFunctionDef, ast.ClassDef)):
            doc = ast.get_docstring(node, clean=False)
            if doc:
                docstrings.add(doc)
    values = [
        node.value
        for node in ast.walk(tree)
        if isinstance(node, ast.Constant) and isinstance(node.value, str) and node.value not in docstrings
    ]
    return "\n".join(values)


def main() -> int:
    problems: list[str] = []
    notes: list[str] = []

    try:
        registry = json.loads(REGISTRY_PATH.read_text(encoding="utf-8"))
    except Exception as exc:
        print(f"ERROR: cannot load registry: {exc}", file=sys.stderr)
        return 2

    modules = registry["modules"]
    genres = {genre["id"]: genre for genre in registry["genres"]}
    sections = registry["sections"]

    # 1. internal consistency
    for genre_id, genre in genres.items():
        form = genre.get("default_iv")
        if form not in modules:
            problems.append(f"genres[{genre_id}].default_iv 指向不存在的模块：{form}")
    for module_id, module in modules.items():
        for field in ("title", "order", "roman", "answer_mode"):
            if field not in module:
                problems.append(f"modules[{module_id}] 缺字段 {field}")
        mode = module.get("answer_mode")
        if mode not in VALID_ANSWER_MODES:
            problems.append(f"modules[{module_id}].answer_mode 非法：{mode}")
        expected_prefix = f"{module['roman']}."
        if not module["title"].startswith(expected_prefix):
            problems.append(f"modules[{module_id}].title 与 roman 不一致：{module['title']}")

    seen_titles: dict[str, str] = {}
    for module_id, module in modules.items():
        title = module["title"]
        if title in seen_titles:
            problems.append(f"重复标题：「{title}」同时属于 {seen_titles[title]} 与 {module_id}")
        seen_titles[title] = module_id

    for section in sections:
        for module_id in section["modules"]:
            if module_id not in modules:
                problems.append(f"sections[{section['id']}] 引用不存在的模块：{module_id}")
        policy = section["iv_policy"]
        if policy["type"] == "fixed" and policy.get("form") not in modules:
            problems.append(f"sections[{section['id']}].iv_policy.form 不存在：{policy.get('form')}")
        if policy["type"] == "fixed":
            inline = policy.get("override_form")
            if inline and inline not in modules:
                problems.append(f"sections[{section['id']}].iv_policy.override_form 不存在：{inline}")
        if policy["type"] == "by_genre":
            for genre in genres.values():
                if genre["default_iv"] not in modules:
                    problems.append(f"体裁 {genre['id']} 的 default_iv 不在注册表中")

    # 2. canonical titles present in docs
    for doc in DOCS:
        if not doc.exists():
            problems.append(f"缺少文档：{doc.relative_to(SKILL_ROOT)}")
            continue
        text = doc.read_text(encoding="utf-8")
        compact = re.sub(r"\s+", "", text)
        if doc.name == "section-blueprints.md":
            for title in seen_titles:
                if re.sub(r"\s+", "", title) not in compact:
                    problems.append(f"{doc.relative_to(SKILL_ROOT)} 缺少板块标题：{title}")

    # 3. legacy aliases must not reappear
    aliases = registry.get("forbidden_title_aliases", [])
    for doc in DOCS:
        if not doc.exists():
            continue
        for line in doc.read_text(encoding="utf-8").splitlines():
            match = HEADING_LINE.match(line)
            if not match:
                continue
            tail = match.group(2)
            for alias in aliases:
                if alias in tail:
                    problems.append(
                        f"{doc.relative_to(SKILL_ROOT)} 板块标题使用废弃别名「{alias}」：{line.strip()[:40]}"
                    )

    # 4. scripts must not hard-code titles; only registry consumers must read it
    for script in SCRIPTS:
        if not script.exists():
            problems.append(f"缺少脚本：{script.relative_to(SKILL_ROOT)}")
            continue
        source = script.read_text(encoding="utf-8")
        literals = code_literals(script)
        for title in seen_titles:
            if title in literals:
                problems.append(f"{script.relative_to(SKILL_ROOT)} 硬编码了板块标题「{title}」，应改读注册表")
        if script.name in REGISTRY_CONSUMERS and "module-registry.json" not in source:
            problems.append(f"{script.relative_to(SKILL_ROOT)} 未引用 module-registry.json")

    # 5. every references/ file cited from SKILL.md must exist
    skill_text = (SKILL_ROOT / "SKILL.md").read_text(encoding="utf-8") if (SKILL_ROOT / "SKILL.md").exists() else ""
    cited = set(re.findall(r"references/([A-Za-z0-9_\-]+\.(?:md|json))", skill_text))
    for name in sorted(cited):
        if not (SKILL_ROOT / "references" / name).exists():
            problems.append(f"SKILL.md 引用了不存在的文件：references/{name}")
    orphans = [
        path.name
        for path in (SKILL_ROOT / "references").glob("*")
        if path.suffix in {".md", ".json"} and path.name not in cited
    ]
    for name in sorted(orphans):
        notes.append(f"未在 SKILL.md 引用的 reference：references/{name}")

    notes.append(f"registry_version={registry['registry_version']}；模块 {len(modules)} 个；板块 {len(sections)} 个")

    if problems:
        print("FAIL")
        for problem in problems:
            print(f"  - {problem}")
        for note in notes:
            print(f"  · {note}")
        return 1

    print("PASS")
    for note in notes:
        print(f"  · {note}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
