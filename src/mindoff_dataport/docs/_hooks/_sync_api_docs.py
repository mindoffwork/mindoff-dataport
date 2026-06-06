from __future__ import annotations

import ast
from pathlib import Path
import textwrap
from typing import Any

TOKEN = "{{ DATAPORT_PUBLIC_API }}"
SOURCE_RELPATH = "src/mindoff_dataport/__init__.py"

# Ordered list of the public API functions whose docstrings drive the API
# Reference page. Order here is the order rendered on the page.
PUBLIC_FUNCTIONS: list[tuple[str, str]] = [
    ("extract_template", "Template Extraction"),
    ("get_template_inputs", "Input Discovery"),
    ("compile_report_bundle", "Bundle Compilation"),
    ("export_report_bundle", "Bundle Export"),
]


def _clean_docstring(value: str | None) -> str:
    if not value:
        return ""
    return textwrap.dedent(value).strip()


def _collect_function_docstrings(source_path: Path) -> dict[str, str]:
    tree = ast.parse(source_path.read_text(encoding="utf-8-sig"))
    found: dict[str, str] = {}
    for node in tree.body:
        if isinstance(node, ast.FunctionDef):
            found[node.name] = _clean_docstring(ast.get_docstring(node))
    return found


def _build_api_markdown(config: Any) -> str | None:
    config_file = Path(config.get("config_file_path", "mkdocs.yml")).resolve()
    source_path = config_file.parent / SOURCE_RELPATH
    if not source_path.exists():
        return None

    docstrings = _collect_function_docstrings(source_path)
    lines: list[str] = []
    for index, (func_name, human_name) in enumerate(PUBLIC_FUNCTIONS, start=1):
        doc = docstrings.get(func_name)
        if not doc:
            continue
        lines.extend(
            [
                f"### {index}. {human_name}",
                "",
                doc,
                "",
            ]
        )
    return "\n".join(lines).strip()


def on_page_markdown(markdown: str, config: Any = None, **kwargs: Any) -> str:
    if TOKEN not in markdown:
        return markdown
    cfg = config or {}
    snippet = _build_api_markdown(cfg)
    if not snippet:
        return markdown
    return markdown.replace(TOKEN, snippet)
