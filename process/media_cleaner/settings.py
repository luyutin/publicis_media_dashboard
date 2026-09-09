"""Load editable field schema and Ollama prompts for the media cleaner."""

from __future__ import annotations

import json
from pathlib import Path

from process.template_schema import TEMPLATE_COLUMNS


CONFIG_DIR = Path(__file__).resolve().parent / "config"


def _read_text(filename: str) -> str:
    path = CONFIG_DIR / filename
    try:
        return path.read_text(encoding="utf-8").strip()
    except OSError as exc:
        raise RuntimeError(f"無法讀取未清理資料格式化設定檔：{path}") from exc


def _load_target_schema() -> tuple[list[str], dict[str, str]]:
    descriptions: dict[str, str] = {}
    path = CONFIG_DIR / "target_columns.txt"
    for line_number, raw_line in enumerate(_read_text(path.name).splitlines(), start=1):
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue
        parts = line.split("\t", 1)
        if len(parts) != 2 or not all(part.strip() for part in parts):
            raise RuntimeError(
                f"{path} 第 {line_number} 行格式錯誤；應為「欄位 key<TAB>欄位說明」"
            )
        key, description = (part.strip() for part in parts)
        if key in descriptions:
            raise RuntimeError(f"{path} 第 {line_number} 行有重複欄位：{key}")
        descriptions[key] = description
    if not descriptions:
        raise RuntimeError(f"{path} 沒有定義任何輸出欄位")
    missing = set(TEMPLATE_COLUMNS) - descriptions.keys()
    if missing:
        raise RuntimeError(f"{path} 缺少 template 欄位：{', '.join(sorted(missing))}")
    extra = descriptions.keys() - set(TEMPLATE_COLUMNS)
    if extra:
        raise RuntimeError(f"{path} 含有 template 未定義欄位：{', '.join(sorted(extra))}")
    ordered_descriptions = {
        column: descriptions[column] for column in TEMPLATE_COLUMNS
    }
    return list(TEMPLATE_COLUMNS), ordered_descriptions


def _load_aliases() -> dict[str, tuple[str, ...]]:
    path = CONFIG_DIR / "aliases.json"
    try:
        raw_aliases = json.loads(_read_text(path.name))
    except json.JSONDecodeError as exc:
        raise RuntimeError(f"{path} 不是有效的 JSON：第 {exc.lineno} 行") from exc
    if not isinstance(raw_aliases, dict):
        raise RuntimeError(f"{path} 最外層必須是 JSON object")

    aliases: dict[str, tuple[str, ...]] = {}
    for target in TARGET_DESCRIPTIONS:
        values = raw_aliases.get(target, [])
        if not isinstance(values, list) or not all(
            isinstance(value, str) and value.strip() for value in values
        ):
            raise RuntimeError(f"{path} 的 {target} 必須是字串陣列")
        aliases[target] = tuple(
            dict.fromkeys([target, *(value.strip() for value in values)])
        )
    extra = raw_aliases.keys() - TARGET_DESCRIPTIONS.keys()
    if extra:
        raise RuntimeError(
            f"{path} 使用了 target_columns.txt 未定義的欄位："
            f"{', '.join(sorted(extra))}"
        )
    return aliases


TARGET_COLUMNS, TARGET_DESCRIPTIONS = _load_target_schema()
# Output metadata owned by the unknown-data cleaner. It is intentionally not
# part of TEMPLATE_COLUMNS, which remains shared with the legacy formatter.
SOURCE_COLUMN = "Source"
ALIASES = _load_aliases()
OLLAMA_SYSTEM_PROMPT = _read_text("ollama_system_prompt.txt")
OLLAMA_USER_PROMPT = _read_text("ollama_user_prompt.txt")


def render_ollama_user_prompt(*, target_descriptions: str, options: str) -> str:
    """Fill the two supported placeholders in the editable user prompt."""
    try:
        return OLLAMA_USER_PROMPT.format(
            target_descriptions=target_descriptions,
            options=options,
        )
    except (KeyError, ValueError) as exc:
        raise RuntimeError(
            "ollama_user_prompt.txt 僅能使用 {target_descriptions} 與 {options} 變數"
        ) from exc
