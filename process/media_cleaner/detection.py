"""Find rule-based seeds; LLM discovery can replace their horizontal boundaries."""

import re

from .grid import SheetGrid, clean_text, date_parts, is_blank, is_error, numeric, normalize
from .models import Candidate


def _header_text(value) -> bool:
    return isinstance(value, str) and not is_blank(value) and not is_error(value) and not date_parts(value) and not numeric(value)


def _sample_row(grid: SheetGrid, row: int, start: int, end: int):
    for below in range(row + 1, min(grid.max_row, row + 4) + 1):
        values = grid.row(below, start, end)
        if any(not is_blank(v) for v in values):
            return below, values
    return None, []


def discover_blocks(grid: SheetGrid, scan_rows: int = 0) -> list[Candidate]:
    """Find horizontal runs, split repeated date axes, then bound vertical tables.

    scan_rows=0 scans the entire sheet. A positive limit restricts header search,
    never the number of data rows. Non-date tables are retained as audit candidates.
    """
    limit = min(grid.max_row, scan_rows) if scan_rows > 0 else grid.max_row
    candidates = []
    occupied: dict[int, list[int]] = {}
    for (row, col) in grid.values:
        if row <= limit:
            occupied.setdefault(row, []).append(col)
    for row in sorted(occupied):
        columns = sorted(occupied[row])
        runs = []
        start = previous = columns[0]
        for col in columns[1:]:
            # A blank header cell with data below is an unnamed column, not a gap.
            separator = col > previous + 1 and any(
                all(is_blank(grid.get(r, gap)) for r in range(row, min(grid.max_row, row + 3) + 1))
                for gap in range(previous + 1, col)
            )
            if separator:
                runs.append((start, previous))
                start = col
            previous = col
        runs.append((start, previous))

        for start, end in runs:
            headers = grid.row(row, start, end)
            filled = [v for v in headers if not is_blank(v)]
            if len(filled) < 2 or sum(_header_text(v) for v in filled) / len(filled) < 0.75:
                continue
            sample_row, sample = _sample_row(grid, row, start, end)
            if (not sample or any(date_parts(v) for v in filled) or not any(
                _header_text(grid.get(row, col)) and (date_parts(grid.get(sample_row, col)) or numeric(grid.get(sample_row, col)))
                for col in range(start, end + 1)
            )):
                continue
            # Age/sex tables qualify structurally and will receive exclusion audits.
            axes = []
            origins = set()
            for col in range(start, end + 1):
                origin = grid.origins.get((sample_row, col), (sample_row, col))
                if date_parts(grid.get(sample_row, col)) and origin not in origins:
                    axes.append(col)
                    origins.add(origin)
            # Multiple distinct date headers can also denote start/end dates in ONE
            # table. Split only repeated labels or repeated metric sets between axes.
            splits = [start]
            for left_axis, right_axis in zip(axes, axes[1:]):
                left_label = clean_text(grid.get(row, left_axis)).casefold()
                right_label = clean_text(grid.get(row, right_axis)).casefold()
                left_metrics = {clean_text(grid.get(row, c)).casefold() for c in range(left_axis + 1, right_axis)}
                right_metrics = {clean_text(grid.get(row, c)).casefold() for c in range(right_axis + 1, end + 1)}
                if left_label == right_label or (left_metrics & right_metrics) - {""}:
                    splits.append(right_axis)
            for left, right in zip(splits, [s - 1 for s in splits[1:]] + [end]):
                labels = [""] * grid.max_column
                header_start = row
                for col in range(left, right + 1):
                    labels[col - 1], first = grid.header(row, col)
                    header_start = min(header_start, first)
                if sum(bool(label) for label in labels[left - 1:right]) < 2:
                    continue
                candidates.append(Candidate(
                    sheet=grid.sheet, header_row=row, score=0, headers=labels,
                    mapped={}, unmapped=[], data_rows=0, start_column=left,
                    end_column=right, end_row=grid.max_row, header_start=header_start,
                ))

    expanded = []
    for block in candidates:
        expanded.extend(_shared_axis_groups(grid, block))
    candidates = [b for b in expanded if not any(
        other.header_row > b.header_row >= other.header_start
        and other.start_column <= b.end_column and other.end_column >= b.start_column
        for other in expanded
    )]
    bound_blocks(grid, candidates)
    return candidates


def bound_blocks(grid, candidates):
    """Bound each table by later headers; internal blank rows do not end a table."""
    for index, block in enumerate(candidates, 1):
        block.block_id = f"table_{index}"
        next_headers = [b.header_start for b in candidates if b.header_row > block.header_row
                        and b.start_column <= block.end_column and b.end_column >= block.start_column]
        block.end_row = min(next_headers, default=grid.max_row + 1) - 1
        last_content = block.header_row
        for row in range(block.header_row + 1, block.end_row + 1):
            values = grid.row(row, block.start_column, block.end_column)
            if block.shared_date_column is not None:
                values.append(grid.get(row, block.shared_date_column))
            if any(not is_blank(v) for v in values):
                last_content = row
        block.end_row = last_content


def _shared_axis_groups(grid, block):
    """Unpivot repeated metric groups that share a single date column on the left.

    Group boundaries require repeated header sequences, not customer-specific
    positions. Merged group captions can be several rows above image/target rows.
    """
    from dataclasses import replace
    sample_row, _ = _sample_row(grid, block.header_row, block.start_column, block.end_column)
    if sample_row is None:
        return [block]
    axes = [c for c in range(block.start_column, block.end_column + 1)
            if any(date_parts(grid.get(r, c)) for r in range(block.header_row + 1, min(grid.max_row, block.header_row + 8) + 1))]
    if axes != [block.start_column]:
        return [block]
    first = block.start_column + 1
    labels = [normalize(grid.get(block.header_row, c)) for c in range(first, block.end_column + 1)]
    if not labels:
        return [block]
    width = next((n for n in range(2, len(labels) // 2 + 1)
                  if len(labels) % n == 0 and labels[:n] == labels[n:2*n]
                  and all(labels[i] == labels[i % n] for i in range(len(labels)))), None)
    if width is None:
        return [block]
    groups = []
    for start in range(first, block.end_column + 1, width):
        end = start + width - 1
        caption, top = "", block.header_row
        for row in range(block.header_row - 1, max(0, block.header_row - 12), -1):
            merged = next((m for m in grid.merges if m.min_col == start and m.max_col == end
                           and m.min_row <= row <= m.max_row), None)
            value = grid.get(row, start)
            if merged and _header_text(value) and not re.match(r"^\d{1,2}[/月]\d{1,2}", clean_text(value)):
                caption, top = clean_text(value), merged.min_row
                break
        headers = list(block.headers)
        for col in range(start, end + 1):
            leaf = clean_text(grid.get(block.header_row, col))
            headers[col - 1] = f"{caption} | {leaf}" if caption else leaf
        groups.append(replace(block, headers=headers, start_column=start, end_column=end,
                              header_start=top, shared_date_column=block.start_column,
                              group_label=caption, mapped={}, warnings=[], excluded_columns=[]))
    return groups
