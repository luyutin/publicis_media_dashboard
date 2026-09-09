"""Ask the LLM how many tables share a seeded header row, across its full width.

The rules locate an anchor. They do not constrain the LLM's table count, metric
order, width or date-column position. Layout proposals are validated against the
source grid before they replace rule-based boundaries.
"""

import json
import math
from copy import deepcopy
from collections import defaultdict
from pathlib import Path

from . import mapping
from .detection import bound_blocks
from .grid import clean_text, date_parts, coerce_excel_date, is_blank, is_summary, numeric
from .models import Candidate


LAYOUT_SCHEMA = {
    "type": "object", "required": ["table_count", "blocks"],
    "properties": {
        "table_count": {"type": "integer", "minimum": 0},
        "blocks": {"type": "array", "items": {
            "type": "object",
            "required": ["start_column", "end_column", "date_column", "confidence"],
            "properties": {
                "start_column": {"type": "integer", "minimum": 1},
                "end_column": {"type": "integer", "minimum": 1},
                "date_column": {"type": "integer", "minimum": 0},
                "label_row": {"type": "integer", "minimum": 0},
                "label_column": {"type": "integer", "minimum": 0},
                "confidence": {"type": "number", "minimum": 0, "maximum": 1},
            },
        }},
    },
}


def row_band(grid, row, seeds):
    """Every populated column is included, even outside all known seed tables."""
    start_row = max(1, min([row - 12] + [b.header_start for b in seeds]))
    end_row = min(grid.max_row, row + 8)
    rows = []
    for r in range(start_row, end_row + 1):
        cells = [[c, mapping._compact(grid.get(r, c))] for c in range(1, grid.max_column + 1)
                 if not is_blank(grid.get(r, c))]
        rows.append({"row": r, "cells": cells})
    return {
        "header_row": row, "first_column": 1, "last_column": grid.max_column,
        "header_cells": [[c, clean_text(grid.get(row, c))] for c in range(1, grid.max_column + 1)
                         if not is_blank(grid.get(row, c))],
        "date_value_columns": [c for c in range(1, grid.max_column + 1) if any(
            date_parts(grid.get(r, c)) for r in range(row + 1, end_row + 1)
        )],
        "rows": rows,
        "merged_ranges": [str(m) for m in grid.merges if m.min_row <= end_row and m.max_row >= start_row],
        "seed_ranges": [[b.start_column, b.end_column] for b in seeds],
    }


def call_ollama_layout(band, config, feedback=""):
    prompt = ("請分析以下完整橫列與上下文，辨識這一列的所有表格或欄位群組。\n"
              + json.dumps(band, ensure_ascii=False)
              + ("\n上次提案未通過驗證，請修正：" + feedback if feedback else ""))
    system = (Path(__file__).parent / "config" / "ollama_layout_prompt.txt").read_text(encoding="utf-8")
    return mapping.call_ollama_json(prompt, config, schema=LAYOUT_SCHEMA, system_prompt=system)


def _date_evidence(grid, row, col, aliases, end_row=None):
    values = [grid.get(r, col) for r in range(row + 1, min(grid.max_row, row + 8, end_row or grid.max_row) + 1)]
    if any(date_parts(value) for value in values):
        return True
    # General-format serials need an explicit date header, not just LLM confidence.
    header = grid.header(row, col)[0]
    return mapping.match_header(header, aliases) == "Date" and any(
        coerce_excel_date(value, epoch=grid.epoch) for value in values
    )


def validate_layout(grid, band, response, aliases):
    """Accept a complete non-overlapping partition; reject guessed coordinates."""
    if not isinstance(response, dict) or not isinstance(response.get("blocks"), list):
        return [], "回應缺少 blocks 陣列"
    items = response["blocks"]
    if type(response.get("table_count")) is not int or response["table_count"] != len(items):
        return [], "table_count 必須等於 blocks 的數量"
    if not items:
        return [], "未提出任何表格範圍，保留規則結果"
    row = band["header_row"]
    required = {c for c in range(1, grid.max_column + 1) if not is_blank(grid.get(row, c)) and any(
        not is_blank(grid.get(r, c)) for r in range(row + 1, min(grid.max_row, row + 8) + 1)
    )}
    result, covered, used_metrics = [], set(), set()
    for item in items:
        if not isinstance(item, dict):
            return [], "表格提案不是 object"
        item = {"label_row": 0, "label_column": 0, **item}
        keys = ("start_column", "end_column", "date_column", "label_row", "label_column")
        if any(type(item.get(key)) is not int for key in keys):
            return [], "範圍及日期、標題座標必須是整數"
        left, right, date_col, label_row, label_col = (item[key] for key in keys)
        confidence = item.get("confidence")
        if not isinstance(confidence, (int, float)) or not math.isfinite(confidence) or not 0.8 <= confidence <= 1:
            return [], "表格辨識信心不足"
        if not 1 <= left <= right <= grid.max_column or not 0 <= date_col <= grid.max_column:
            return [], "表格範圍或日期欄超出工作表"
        if date_col and not _date_evidence(grid, row, date_col, aliases):
            return [], f"欄 {date_col} 缺少日期證據"
        if not date_col and any(_date_evidence(grid, row, c, aliases) for c in range(left, right + 1)):
            return [], "範圍內有日期證據，卻未指定日期欄"
        columns = set(range(left, right + 1))
        metrics = columns - ({date_col} if date_col else set())
        if not metrics or metrics & used_metrics:
            return [], "表格沒有獨立欄位，或與其他表格重複使用資料欄"
        if not any(not is_blank(grid.get(row, c)) for c in metrics):
            return [], "提案的資料欄沒有表頭"
        used_metrics.update(metrics)
        covered.update(columns)
        if date_col:
            covered.add(date_col)
        caption, top = "", row
        if label_row or label_col:
            if not 1 <= label_row < row or not left <= label_col <= right:
                return [], "區塊標題必須引用本區塊上方的實際儲存格"
            caption = clean_text(grid.get(label_row, label_col))
            if not caption or numeric(caption) or date_parts(caption):
                return [], "標題座標沒有有效文字"
            top = label_row
        headers = [grid.header(row, c)[0] for c in range(1, grid.max_column + 1)]
        for c in range(left, right + 1):
            if caption:
                headers[c - 1] = caption + " | " + clean_text(grid.get(row, c))
        candidate = Candidate(
            sheet=grid.sheet, header_row=row, score=0, headers=headers, mapped={}, unmapped=[], data_rows=0,
            start_column=left, end_column=right, header_start=top, end_row=grid.max_row,
            shared_date_column=date_col or None, group_label=caption, method="llm-layout",
        )
        probe = deepcopy(candidate)
        mapping.rule_mapping(grid, probe, aliases)
        conflicts = [warning for warning in probe.warnings if warning.startswith("多欄對應")]
        if conflicts and not is_summary(caption):
            return [], (f"欄 {left}–{right} 的提案仍有重複指標：{'；'.join(conflicts)}。"
                        "請重新判斷是否包含多張表；每張表同一標準指標只能有一個來源，共用日期可重複引用。")
        result.append(candidate)
    missing = required - covered
    if missing:
        return [], f"仍有來源欄位未涵蓋：{sorted(missing)}；請包含彙總及非日期表格以利稽核"
    return result, ""


def refine_blocks(grid, seeds, aliases, config):
    if config is None or not config.enabled:
        return seeds
    groups = defaultdict(list)
    for block in seeds:
        groups[block.header_row].append(block)
    result = []
    for row, originals in sorted(groups.items()):
        # One dated seed is enough. Inspect the ENTIRE row, not just unknown fields.
        if not any(_date_evidence(grid, row, c, aliases, b.end_row) for b in originals
                   for c in set(range(b.start_column, b.end_column + 1)) | ({b.shared_date_column} if b.shared_date_column else set())):
            result.extend(originals)
            continue
        band = row_band(grid, row, originals)
        feedback, proposals = "", []
        for _ in range(2):
            try:
                response = call_ollama_layout(band, config, feedback)
                proposals, feedback = validate_layout(grid, band, response, aliases)
                if proposals:
                    break
            except RuntimeError as exc:
                config.failure = str(exc)
                feedback = (
                    f"LLM 結構辨識請求失敗（sheet={grid.sheet}, header_row={row}）："
                    f"{exc}；本列使用規則範圍，後續表頭列仍會繼續呼叫 Ollama"
                )
                break
        if proposals:
            result.extend(proposals)
        else:
            for block in originals:
                block.method = "llm-layout-failed+rules"
                block.warnings.append("LLM 結構提案未採用；使用規則範圍：" + feedback)
            result.extend(originals)
    result.sort(key=lambda b: (b.header_row, b.start_column))
    bound_blocks(grid, result)
    return result
