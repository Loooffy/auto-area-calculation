from __future__ import annotations

import argparse
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Tuple

from openpyxl import load_workbook
from openpyxl.utils.cell import column_index_from_string, coordinate_from_string


@dataclass
class RowData:
    row: int
    multiplier: float
    decimal: float
    rounded: int
    sum_value: float


@dataclass
class GroupData:
    decimal_value: float
    rows: List[RowData]
    weight_sum: float


class SheetAdapter:
    def __init__(self, ws, is_xls: bool) -> None:
        self.ws = ws
        self.is_xls = is_xls
        self.title = ws.name if is_xls else ws.title

    def cell_value(self, row: int, col: int):
        if self.is_xls:
            return self.ws.cell_value(row - 1, col - 1)
        return self.ws.cell(row, col).value

    def cell_value_by_ref(self, ref: str):
        col_name, row = coordinate_from_string(ref.strip().upper())
        col = col_idx(col_name)
        return self.cell_value(row, col)


def col_idx(col_name: str) -> int:
    return column_index_from_string(col_name.strip().upper())


def to_float(value, field: str, row: int) -> float:
    if value is None:
        raise ValueError(f"Row {row}: {field} is empty.")
    try:
        return float(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"Row {row}: {field} is not numeric: {value!r}") from exc


def read_rows(
    ws: SheetAdapter,
    start_row: int,
    end_row: int,
    multiplier_col: str,
    decimal_col: str,
    rounded_col: str,
    sum_col: str,
) -> List[RowData]:
    m_col = col_idx(multiplier_col)
    d_col = col_idx(decimal_col)
    r_col = col_idx(rounded_col)
    s_col = col_idx(sum_col)

    rows: List[RowData] = []
    for r in range(start_row, end_row + 1):
        multiplier = to_float(ws.cell_value(r, m_col), "multiplier", r)
        decimal = to_float(ws.cell_value(r, d_col), "decimal", r)
        rounded_raw = ws.cell_value(r, r_col)
        rounded_num = to_float(rounded_raw, "rounded", r)
        rounded_int = int(round(rounded_num))
        if abs(rounded_num - rounded_int) > 1e-9:
            raise ValueError(
                f"Row {r}: rounded value should be integer-like, got {rounded_num!r}"
            )
        sum_value = to_float(ws.cell_value(r, s_col), "sum", r)
        rows.append(
            RowData(
                row=r,
                multiplier=multiplier,
                decimal=decimal,
                rounded=rounded_int,
                sum_value=sum_value,
            )
        )
    return rows


def build_groups(rows: List[RowData]) -> List[GroupData]:
    # 依「持分(decimal)」分組：相同持分的列必須一同 +1/-1。
    # decimal 是浮點數，用 round(..., 9) 當鍵避免浮點表示誤差把相同值拆開。
    grouped: Dict[float, List[RowData]] = {}
    for row in rows:
        grouped.setdefault(round(row.decimal, 9), []).append(row)

    groups: List[GroupData] = []
    for decimal_value, group_rows in grouped.items():
        weight_sum = sum(r.multiplier for r in group_rows)
        groups.append(
            GroupData(
                decimal_value=decimal_value,
                rows=group_rows,
                weight_sum=weight_sum,
            )
        )
    return groups


def solve(groups: List[GroupData], required_delta: float) -> Tuple[int, Dict[int, int]]:
    # DP state: sum_delta -> (cost, assignments_by_group_index)
    # cost = 被調整(±1)的列數總和；整組一起動，故每組調整的代價為該組列數。
    states: Dict[float, Tuple[int, Dict[int, int]]] = {0.0: (0, {})}
    for idx, g in enumerate(groups):
        next_states: Dict[float, Tuple[int, Dict[int, int]]] = {}
        row_count = len(g.rows)
        options = [
            (0, 0.0, 0),
            (-1, -g.weight_sum, row_count),
            (1, g.weight_sum, row_count),
        ]
        for sum_delta, (cost_so_far, choice_map) in states.items():
            for shift, delta_sum, delta_cost in options:
                new_sum = round(sum_delta + delta_sum, 10)
                new_cost = cost_so_far + delta_cost
                new_map = dict(choice_map)
                if shift != 0:
                    new_map[idx] = shift
                prev = next_states.get(new_sum)
                if prev is None or new_cost < prev[0]:
                    next_states[new_sum] = (new_cost, new_map)
        states = next_states

    target_key = round(required_delta, 10)
    if target_key not in states:
        raise ValueError(
            "No exact solution found under constraints. "
            "Try checking target or data range."
        )
    return states[target_key]


def build_result_text(
    rows: List[RowData],
    groups: List[GroupData],
    current_total_sum: float,
    target: float,
    changed_rows: int,
    assignments: Dict[int, int],
) -> str:
    lines: List[str] = []
    lines.append("")
    lines.append("調整明細:")
    lines.append("列, 增減, 戶數, 調整前, 調整後, 群組編號")

    group_no_by_idx = {idx: group_no for group_no, idx in enumerate(sorted(assignments), start=1)}
    detail_rows = [
        (r, assignments[idx], group_no_by_idx[idx])
        for idx in assignments
        for r in groups[idx].rows
    ]
    for r, shift, group_no in sorted(detail_rows, key=lambda x: x[0].row):
        lines.append(
            f"{r.row}, {shift:+d}, {r.multiplier:g}, {r.rounded}, "
            f"{r.rounded + shift}, {group_no}"
        )

    if not assignments:
        lines.append("(不需要調整，已經吻合 target)")
    lines.append("")
    lines.append(f"current_total_sum: {int(round(current_total_sum))}")
    lines.append(f"target: {int(round(target))}")
    lines.append(f"changed_rows: {changed_rows}")
    return "\n".join(lines)


def build_pre_solve_debug_text(
    args: argparse.Namespace,
    ws: SheetAdapter,
    rows: List[RowData],
    groups: List[GroupData],
    current_total_sum: float,
    required_delta: float,
) -> str:
    calc_sum_from_sum_col = sum(r.sum_value for r in rows)
    calc_sum_from_multiplier_rounded = sum(r.multiplier * r.rounded for r in rows)
    current_total_sum_int = int(round(current_total_sum))
    target_int = int(round(args.target))
    required_delta_int = int(round(required_delta))
    calc_sum_from_sum_col_int = int(round(calc_sum_from_sum_col))
    calc_sum_from_multiplier_rounded_int = int(round(calc_sum_from_multiplier_rounded))
    lines: List[str] = []
    lines.append("=== Pre-solve diagnostics ===")
    lines.append(f"file_name: {args.file_name}")
    lines.append(f"sheet: {ws.title}")
    lines.append(f"row_range: {args.start_row}..{args.end_row}")
    lines.append(
        "columns: "
        f"multiplier={args.multiplier_col}, "
        f"decimal={args.decimal_col}, "
        f"rounded={args.rounded_col}, "
        f"sum={args.sum_col}"
    )
    lines.append(f"total_sum_cell: {args.total_sum_cell}")
    lines.append(f"rows_loaded: {len(rows)}")
    lines.append(f"current_total_sum({args.total_sum_cell}): {current_total_sum_int}")
    lines.append(f"target: {target_int}")
    lines.append(f"required_delta(target-current): {required_delta_int}")
    lines.append(f"sum(rows.{args.sum_col}): {calc_sum_from_sum_col_int}")
    lines.append(f"sum(rows.multiplier*rounded): {calc_sum_from_multiplier_rounded_int}")
    lines.append(f"group_count(by decimal): {len(groups)}")
    lines.append("=== End diagnostics ===")
    return "\n".join(lines)


def run_solver(args: argparse.Namespace) -> str:
    file_ext = Path(args.file_name).suffix.lower()

    if file_ext == ".xlsx":
        wb = load_workbook(args.file_name, data_only=True)
        raw_ws = wb[args.sheet] if args.sheet else wb.active
        ws = SheetAdapter(raw_ws, is_xls=False)
    elif file_ext == ".xls":
        try:
            import xlrd  # type: ignore
        except ImportError as exc:
            raise ImportError(
                "Reading .xls requires xlrd. Install it with: pip install xlrd"
            ) from exc
        wb = xlrd.open_workbook(args.file_name)
        if args.sheet:
            raw_ws = wb.sheet_by_name(args.sheet)
        else:
            raw_ws = wb.sheet_by_index(0)
        ws = SheetAdapter(raw_ws, is_xls=True)
    else:
        raise ValueError(
            f"Unsupported file extension: {file_ext}. Use .xlsx or .xls."
        )

    rows = read_rows(
        ws,
        args.start_row,
        args.end_row,
        args.multiplier_col,
        args.decimal_col,
        args.rounded_col,
        args.sum_col,
    )
    groups = build_groups(rows)

    total_sum_cell_value = ws.cell_value_by_ref(args.total_sum_cell)
    current_total_sum = float(total_sum_cell_value)
    required_delta = args.target - current_total_sum

    debug_text = build_pre_solve_debug_text(
        args=args,
        ws=ws,
        rows=rows,
        groups=groups,
        current_total_sum=current_total_sum,
        required_delta=required_delta,
    )

    changed_rows, assignments = solve(groups, required_delta)
    result_text = build_result_text(
        rows=rows,
        groups=groups,
        current_total_sum=current_total_sum,
        target=args.target,
        changed_rows=changed_rows,
        assignments=assignments,
    )
    return f"{debug_text}\n{result_text}"
