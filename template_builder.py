from __future__ import annotations

from collections.abc import Sequence
from dataclasses import dataclass
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import (
    column_index_from_string,
    get_column_letter,
    range_boundaries,
)
from openpyxl.worksheet.worksheet import Worksheet

from csv_import import PublicAreaData, UnitData
from formula_utils import (
    LayoutConfig,
    data_row_formulas,
    rewrite_formula,
    shift_formula_columns,
)
from parking_balance import (
    MAX_TOLERANCE,
    ShareAdjustment,
    allocation_area,
    compute_share_adjustments,
    excel_round,
)
from share_tweak import build_share_rows, solve_share_adjustments
from template_schema import (
    COLUMN_WIDTHS,
    HEADER_FORMULAS,
    HEADER_STATIC,
    MERGED_RANGES,
    PARKING_END_ROW,
    PARKING_FORMULAS,
    PARKING_START_ROW,
    PARKING_STATIC,
    PUBLIC_AREA_END_ROW,
    PUBLIC_AREA_START_ROW,
    PUBLIC_AREA_SUMMARY_COL,
    PUBLIC_AREA_SUMMARY_ROW,
    SHEET_TITLE,
    TOTAL_FORMULAS,
    TOTAL_ROW,
    TOTAL_STATIC,
)

INPUT_COLUMNS = ("A", "B", "C", "D")
HEADER_ROWS = (1, 2, 3, 4)
PARKING_NOTE_ROW = PARKING_START_ROW


@dataclass(frozen=True)
class AdjustmentStatus:
    """單項持分微調的結果（供建立後提示是否需人工處理）。"""

    name: str  # 項目名稱，如「車位持分」「公設計算」「土地持分」
    target: int  # 目標總持分
    actual: int  # 微調後實際加總
    completed: bool  # 是否與目標完全吻合


@dataclass(frozen=True)
class BuildResult:
    """create_template 的回傳結果。"""

    path: Path
    adjustments: tuple[AdjustmentStatus, ...] = ()

    @property
    def incomplete_adjustments(self) -> tuple[AdjustmentStatus, ...]:
        return tuple(item for item in self.adjustments if not item.completed)


def apply_column_widths(ws: Worksheet, config: LayoutConfig) -> None:
    for col_letter, width in COLUMN_WIDTHS.items():
        dest = get_column_letter(config.col(column_index_from_string(col_letter)))
        ws.column_dimensions[dest].width = width
    if config.include_porch:
        ws.column_dimensions["E"].width = 13.0
        weighted = get_column_letter(config.weighted_porch_col or 14)
        ws.column_dimensions[weighted].width = 13.0
        # 「主+附+門廊」比原來的「主+附」長，合計欄加寬。
        ws.column_dimensions["F"].width = 18.0
        ws.column_dimensions[
            get_column_letter(config.col(13))
        ].width = 18.0
    if config.terrace_col is not None:
        ws.column_dimensions[get_column_letter(config.terrace_col)].width = 12.0


BODY_FONT = Font(name="DFKai-SB", size=13)
SHARE_LABEL_FONT = Font(name="DFKai-SB", size=10)
THIN_SIDE = Side(style="thin", color="FF000000")
THIN_BORDER = Border(
    left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE
)
AREA_NUMBER_FORMAT = "0.00"
MULTILINE_ROW_HEIGHT = 32
# 各戶面積欄：主建物到坪（C–H），以及加權後的對應欄（K–P）。
AREA_COLUMNS = frozenset({3, 4, 5, 6, 7, 8, 11, 12, 13, 14, 15, 16})
PARKING_COUNT_LAST_ROW = 66
PARKING_PING_ROW = 68
PARKING_RATIO_BLOCK_LAST_ROW = 75
PARKING_CALC_LAST_ROW = 78
PARKING_GRID_FIRST_COL = 7
PUBLIC_BLOCK_FIRST_COL = 2
PUBLIC_BLOCK_LAST_COL = 5


def _merged_row_span(ws: Worksheet, row: int, col: int) -> int:
    for merged in ws.merged_cells.ranges:
        if (
            merged.min_row <= row <= merged.max_row
            and merged.min_col <= col <= merged.max_col
        ):
            return merged.max_row - merged.min_row + 1
    return 1


def _is_numeric_value(value: object) -> bool:
    if isinstance(value, bool):
        return False
    if isinstance(value, (int, float)):
        return True
    return isinstance(value, str) and value.startswith("=")


def _paint(
    ws: Worksheet,
    row: int,
    col: int,
    *,
    horizontal: str,
    vertical: str,
    number_format: str | None = None,
) -> None:
    cell = ws.cell(row, col)
    wrap = isinstance(cell.value, str) and "\n" in cell.value
    cell.font = BODY_FONT
    cell.alignment = Alignment(
        horizontal=horizontal, vertical=vertical, wrap_text=wrap
    )
    cell.border = THIN_BORDER
    if number_format is not None:
        cell.number_format = number_format
    if wrap and _merged_row_span(ws, row, col) == 1:
        height = ws.row_dimensions[row].height
        if height is None or height < MULTILINE_ROW_HEIGHT:
            ws.row_dimensions[row].height = MULTILINE_ROW_HEIGHT


def _paint_box(
    ws: Worksheet,
    min_row: int,
    max_row: int,
    min_col: int,
    max_col: int,
    *,
    horizontal: str,
    vertical: str,
    number_format: str | None = None,
) -> None:
    for row in range(min_row, max_row + 1):
        for col in range(min_col, max_col + 1):
            _paint(
                ws,
                row,
                col,
                horizontal=horizontal,
                vertical=vertical,
                number_format=number_format,
            )


def _area_columns(config: LayoutConfig) -> frozenset[int]:
    columns = {config.col(col) for col in AREA_COLUMNS}
    if config.include_porch:
        columns.add(5)
        if config.weighted_porch_col is not None:
            columns.add(config.weighted_porch_col)
    if config.terrace_col is not None:
        columns.add(config.terrace_col)
    return frozenset(columns)


def _style_main_table(ws: Worksheet, config: LayoutConfig) -> None:
    for col in range(2, config.col(8) + 1):
        _paint(ws, 1, col, horizontal="center", vertical="center")

    areas = _area_columns(config)
    public_header = (3, config.col(6))
    for row in range(2, 5):
        for col in range(1, config.max_column + 1):
            number_format = AREA_NUMBER_FORMAT if (row, col) == public_header else None
            _paint(
                ws,
                row,
                col,
                horizontal="center",
                vertical="center",
                number_format=number_format,
            )

    for row in range(config.data_start, config.total_row + 1):
        for col in range(1, config.max_column + 1):
            if col <= 2:
                _paint(ws, row, col, horizontal="center", vertical="center")
                continue
            number_format = AREA_NUMBER_FORMAT if col in areas else None
            _paint(
                ws,
                row,
                col,
                horizontal="right",
                vertical="bottom",
                number_format=number_format,
            )


def _public_label_columns(config: LayoutConfig) -> frozenset[int]:
    # 有門廊時右側樓別併入插入的空白欄（E），面積改在 F。
    columns = {2, 4}
    if config.include_porch:
        columns.add(5)
    return frozenset(columns)


def _style_public_area_cell(
    ws: Worksheet, row: int, col: int, config: LayoutConfig
) -> None:
    value = ws.cell(row, col).value
    # 樓別欄（B、D；有門廊時含併入的 E）置中；面積欄靠右。
    if col in _public_label_columns(config):
        number_format = AREA_NUMBER_FORMAT if _is_numeric_value(value) else None
        _paint(
            ws,
            row,
            col,
            horizontal="center",
            vertical="center",
            number_format=number_format,
        )
        return
    _paint(
        ws,
        row,
        col,
        horizontal="right",
        vertical="bottom",
        number_format=AREA_NUMBER_FORMAT,
    )


def _style_lower_section(
    ws: Worksheet, config: LayoutConfig, parking_size_count: int = 4
) -> None:
    grid_first = config.col(PARKING_GRID_FIRST_COL)
    public_last_col = config.col(PUBLIC_BLOCK_LAST_COL)
    size_first, size_last, grid_last = _parking_style_bounds(
        config, parking_size_count
    )
    calc_last = size_last

    note_row = config.map_row(PARKING_NOTE_ROW)
    for col in (config.col(9), config.col(19), config.col(22)):
        _paint(ws, note_row, col, horizontal="right", vertical="bottom")
    _paint(
        ws,
        config.map_row(PARKING_NOTE_ROW + 1),
        config.col(9),
        horizontal="right",
        vertical="bottom",
    )

    summary_row = config.map_row(PUBLIC_AREA_SUMMARY_ROW)
    header_row = config.map_row(PARKING_SIZE_HEADER_ROW)
    for col in range(PUBLIC_BLOCK_FIRST_COL, public_last_col + 1):
        if col >= PUBLIC_AREA_SUMMARY_COL:
            _paint(
                ws,
                summary_row,
                col,
                horizontal="right",
                vertical="bottom",
                number_format=AREA_NUMBER_FORMAT,
            )
        else:
            _paint(ws, summary_row, col, horizontal="center", vertical="center")
    _paint_box(
        ws,
        header_row,
        header_row,
        PUBLIC_BLOCK_FIRST_COL,
        public_last_col,
        horizontal="center",
        vertical="center",
    )

    public_first = config.public_data_row(0)
    public_last_row = config.map_row(PUBLIC_AREA_END_ROW) + config.public_area_extra_rows
    for row in range(public_first, public_last_row + 1):
        for col in range(PUBLIC_BLOCK_FIRST_COL, public_last_col + 1):
            _style_public_area_cell(ws, row, col, config)

    # 無尺寸欄時標題跨到資料欄（grid_last）；有尺寸時不含合計。
    title_last = grid_last if parking_size_count <= 0 else calc_last
    _paint_box(
        ws,
        summary_row,
        summary_row,
        grid_first,
        title_last,
        horizontal="center",
        vertical="center",
    )
    _paint_box(
        ws,
        header_row,
        header_row,
        grid_first,
        grid_last,
        horizontal="center",
        vertical="center",
    )
    for src_row in range(PUBLIC_AREA_START_ROW, PARKING_COUNT_LAST_ROW + 1):
        if _is_dropped_template_floor_row(src_row, config):
            continue
        row = config.map_row(src_row)
        _paint(ws, row, grid_first, horizontal="center", vertical="center")
        for col in range(grid_first + 1, grid_last + 1):
            _paint(ws, row, col, horizontal="right", vertical="bottom")
    for row in iter_parking_floor_rows(config):
        _paint(ws, row, grid_first, horizontal="center", vertical="center")
        for col in range(grid_first + 1, grid_last + 1):
            _paint(ws, row, col, horizontal="right", vertical="bottom")
    for src_row in range(PARKING_COUNT_LAST_ROW + 1, PARKING_PING_ROW + 1):
        row = config.map_row(src_row)
        _paint(ws, row, grid_first, horizontal="center", vertical="center")
        for col in range(size_first, size_last + 1):
            _paint(
                ws,
                row,
                col,
                horizontal="right",
                vertical="bottom",
                number_format=AREA_NUMBER_FORMAT,
            )

    if parking_size_count > 0:
        proportion_row = config.map_row(70)
        # 標題列含左側表頭欄（車位尺寸／實際面積…），避免單欄時文字凸出。
        _paint_box(
            ws,
            proportion_row,
            proportion_row,
            grid_first,
            size_last,
            horizontal="center",
            vertical="center",
        )
        for src_row in range(PARKING_SIZE_LABEL_ROW, PARKING_CALC_LAST_ROW + 1):
            row = config.map_row(src_row)
            _paint(ws, row, grid_first, horizontal="center", vertical="center")
            if src_row == PARKING_SHARE_ROW:
                ws.cell(row, grid_first).font = SHARE_LABEL_FONT
            number_format = (
                AREA_NUMBER_FORMAT if src_row == PARKING_ACTUAL_AREA_ROW else None
            )
            horizontal = "center" if src_row == PARKING_SIZE_LABEL_ROW else "right"
            vertical = "center" if src_row == PARKING_SIZE_LABEL_ROW else "bottom"
            for col in range(size_first, size_last + 1):
                _paint(
                    ws,
                    row,
                    col,
                    horizontal=horizontal,
                    vertical=vertical,
                    number_format=number_format,
                )

    for src_row in range(PUBLIC_RATIO_ROW, PARKING_RATIO_BLOCK_LAST_ROW + 1):
        row = config.map_row(src_row)
        _paint(ws, row, 2, horizontal="center", vertical="center")
        area_format = AREA_NUMBER_FORMAT if src_row == PARKING_RATIO_BLOCK_LAST_ROW else None
        _paint(
            ws,
            row,
            3,
            horizontal="right",
            vertical="bottom",
            number_format=area_format,
        )
        _paint(ws, row, 4, horizontal="right", vertical="bottom")


def apply_appearance(
    ws: Worksheet, config: LayoutConfig, parking_size_count: int = 4
) -> None:
    """套用標楷體、細框線、對齊與面積欄小數格式。不改底色。"""
    _style_main_table(ws, config)
    _style_lower_section(ws, config, parking_size_count)


def shift_range_rows(
    min_row: int, max_row: int, config: LayoutConfig
) -> tuple[int, int]:
    if min_row >= config.source_parking_start:
        return config.map_row(min_row), config.map_row(max_row)
    return min_row, max_row


def apply_merged_cells(
    ws: Worksheet,
    config: LayoutConfig,
    skip_ranges: set[str] | None = None,
) -> None:
    skip = skip_ranges or set()
    for range_str in MERGED_RANGES:
        if range_str in skip:
            continue
        min_col, min_row, max_col, max_row = range_boundaries(range_str)
        min_row, max_row = shift_range_rows(min_row, max_row, config)
        ws.merge_cells(
            start_row=min_row,
            start_column=config.col(min_col),
            end_row=max_row,
            end_column=config.col(max_col),
        )
    if config.include_porch:
        _apply_porch_merges(ws, config)
    if config.terrace_col is not None:
        ws.merge_cells(
            start_row=2,
            start_column=config.terrace_col,
            end_row=3,
            end_column=config.terrace_col,
        )


def _apply_porch_merges(ws: Worksheet, config: LayoutConfig) -> None:
    """門廊欄的表頭合併，以及公設右欄被插入欄切開後把空白欄併回樓別。"""
    ws.merge_cells(start_row=2, start_column=5, end_row=3, end_column=5)
    weighted = config.weighted_porch_col or 14
    ws.merge_cells(start_row=2, start_column=weighted, end_row=4, end_column=weighted)
    header_row = config.map_row(PARKING_SIZE_HEADER_ROW)
    ws.merge_cells(
        start_row=header_row, start_column=4, end_row=header_row, end_column=5
    )
    first = config.public_data_row(0)
    last = config.map_row(PUBLIC_AREA_END_ROW) + config.public_area_extra_rows
    for row in range(first, last + 1):
        ws.merge_cells(start_row=row, start_column=4, end_row=row, end_column=5)


def set_formula(ws: Worksheet, row: int, col: int, formula: str, config: LayoutConfig) -> None:
    shifted = shift_formula_columns(
        formula, config.include_porch, config.include_terrace
    )
    ws.cell(row, config.col(col)).value = rewrite_formula(shifted, config)


# 來源座標上的「主+附」表頭，有門廊時改稱「主+附+門廊」。
_PORCH_HEADER_TEXT = {
    (2, 5): "主+附+門廊",
    (2, 13): "主+附+門廊",
}


def build_header(ws: Worksheet, config: LayoutConfig) -> None:
    for (row, col), value in HEADER_STATIC.items():
        if config.include_porch and (row, col) in _PORCH_HEADER_TEXT:
            value = _PORCH_HEADER_TEXT[(row, col)]
        ws.cell(row, config.col(col)).value = value
    if config.include_porch:
        ws.cell(2, 5).value = "門廊"
        ws.cell(4, 5).value = "㎡"
        weighted = config.weighted_porch_col or 14
        ws.cell(2, weighted).value = "門廊"
    if config.terrace_col is not None:
        ws.cell(2, config.terrace_col).value = "備註：露台"
        ws.cell(4, config.terrace_col).value = "㎡"
    for (row, col), formula in HEADER_FORMULAS.items():
        set_formula(ws, row, col, formula, config)
    ws.cell(1, config.col(9)).value = None
    ws.cell(1, config.col(10)).value = None


def build_data_rows(
    ws: Worksheet,
    config: LayoutConfig,
    unit_data: list[UnitData] | None = None,
) -> None:
    for offset in range(config.units):
        row = config.data_start + offset
        input_columns = ("A", "B", "C", "D", "E") if config.include_porch else INPUT_COLUMNS
        for col in input_columns:
            ws[f"{col}{row}"].value = None
        for col, formula in data_row_formulas(
            row,
            config.total_row,
            config.include_porch,
            config.include_terrace,
        ).items():
            ws[f"{col}{row}"].value = formula

        if unit_data is not None and offset < len(unit_data):
            unit = unit_data[offset]
            ws[f"A{row}"] = unit.unit_count
            ws[f"B{row}"] = unit.label
            ws[f"C{row}"] = unit.main_area
            ws[f"D{row}"] = unit.balcony_area
            if config.include_porch and unit.porch_area is not None:
                ws[f"E{row}"] = unit.porch_area
            # 露台欄只在有任一戶 > 0 時出現；單戶空白／0 仍顯示 0。
            if config.terrace_col is not None:
                terrace = unit.terrace_area if unit.terrace_area is not None else 0.0
                ws.cell(row, config.terrace_col).value = terrace


def build_total_row(ws: Worksheet, config: LayoutConfig) -> None:
    row = config.total_row
    for col, text in TOTAL_STATIC.items():
        ws.cell(row, col).value = text
    for col, formula in TOTAL_FORMULAS.items():
        set_formula(ws, row, col, formula, config)
    if config.include_porch:
        # 門廊合計對到加權門廊合計；加權門廊是各列戶數 × 門廊的加總。
        weighted = get_column_letter(config.weighted_porch_col or 14)
        ws.cell(row, 5).value = rewrite_formula(f"={weighted}{TOTAL_ROW}", config)
        ws.cell(row, config.weighted_porch_col or 14).value = rewrite_formula(
            f"=sum({weighted}5:{weighted}53)", config
        )


def build_parking_section(ws: Worksheet, config: LayoutConfig) -> None:
    row_count = PARKING_END_ROW - PARKING_START_ROW + 1
    for offset in range(row_count):
        src_row = PARKING_START_ROW + offset
        if _is_dropped_template_floor_row(src_row, config):
            continue
        dst_row = config.map_row(src_row)

        for (static_row, col), text in PARKING_STATIC.items():
            if static_row == src_row:
                ws.cell(dst_row, config.col(col)).value = text

        for (formula_row, col), formula in PARKING_FORMULAS.items():
            if formula_row == src_row:
                set_formula(ws, dst_row, col, formula, config)


PUBLIC_AREA_SLOT_ROWS = PUBLIC_AREA_END_ROW - PUBLIC_AREA_START_ROW + 1
# 範本「停車位面積、類別、數量」預留的樓層列數（B1–B6）。
PARKING_TEMPLATE_FLOOR_COUNT = 6


def public_area_rows_needed(public_areas: list[PublicAreaData]) -> int:
    left_items = [item for item in public_areas if item.side == "left"]
    right_items = [item for item in public_areas if item.side == "right"]
    left_rows = sum(2 if item.is_per_floor else 1 for item in left_items)
    right_rows = len(right_items)
    return max(left_rows, right_rows)


def compute_public_area_extra_rows(
    public_areas: list[PublicAreaData],
    parking_floor_count: int = PARKING_TEMPLATE_FLOOR_COUNT,
) -> int:
    """公設列不夠時要再插入的列數。

    停車位樓層和公設寫在同一列。樓層變少時公設可用列跟著變少，
    樓層變多時中間多出來的列也可以拿來放公設。
    """
    slots = PUBLIC_AREA_SLOT_ROWS + (
        parking_floor_count - PARKING_TEMPLATE_FLOOR_COUNT
    )
    return max(0, public_area_rows_needed(public_areas) - max(slots, 0))


def public_area_row_bounds(
    config: LayoutConfig,
    public_areas: list[PublicAreaData],
) -> tuple[int, int]:
    left_items = [item for item in public_areas if item.side == "left"]
    right_items = [item for item in public_areas if item.side == "right"]
    left_rows = sum(2 if item.is_per_floor else 1 for item in left_items)
    right_rows = len(right_items)
    start = config.public_data_row(0)
    left_last = config.public_data_row(left_rows - 1) if left_rows else start - 1
    right_last = config.public_data_row(right_rows - 1) if right_rows else start - 1
    return left_last, right_last


def clear_public_area_data(
    ws: Worksheet,
    config: LayoutConfig,
    public_areas: list[PublicAreaData],
) -> None:
    for offset in range(public_area_rows_needed(public_areas)):
        dst_row = config.public_data_row(offset)
        for col in range(2, config.col(5) + 1):
            ws.cell(dst_row, col).value = None


def apply_public_area_merges(
    ws: Worksheet,
    config: LayoutConfig,
    public_areas: list[PublicAreaData],
) -> None:
    left_items = [item for item in public_areas if item.side == "left"]
    offset = 0
    for item in left_items:
        if item.is_per_floor:
            label_row = config.public_data_row(offset)
            per_floor_row = config.public_data_row(offset + 1)
            ws.merge_cells(
                start_row=label_row,
                start_column=3,
                end_row=per_floor_row,
                end_column=3,
            )
            offset += 2
        else:
            offset += 1


def update_public_area_total_formula(
    ws: Worksheet,
    config: LayoutConfig,
    public_areas: list[PublicAreaData],
) -> None:
    left_last, right_last = public_area_row_bounds(config, public_areas)
    summary_row = config.map_row(PUBLIC_AREA_SUMMARY_ROW)
    area_start = config.public_data_row(0)
    parts: list[str] = []
    if left_last >= area_start:
        parts.append(f"sum(C{area_start}:C{left_last})")
    if right_last >= area_start:
        right_col = get_column_letter(config.col(5))
        parts.append(f"sum({right_col}{area_start}:{right_col}{right_last})")
    if parts:
        ws.cell(summary_row, PUBLIC_AREA_SUMMARY_COL).value = "=" + "+".join(parts)


def fill_public_areas(
    ws: Worksheet,
    config: LayoutConfig,
    public_areas: list[PublicAreaData],
) -> None:
    left_items = [item for item in public_areas if item.side == "left"]
    right_items = [item for item in public_areas if item.side == "right"]

    offset = 0
    for item in left_items:
        label_row = config.public_data_row(offset)
        ws.cell(label_row, 2).value = item.floor_label
        if item.is_per_floor:
            per_floor_row = config.public_data_row(offset + 1)
            ws.cell(per_floor_row, 2).value = item.area
            multiplier = item.floor_multiplier
            if multiplier is not None:
                ws.cell(label_row, 3).value = f"=B{per_floor_row}*{multiplier}"
            offset += 2
        else:
            ws.cell(label_row, 3).value = item.area
            offset += 1

    offset = 0
    for item in right_items:
        dst_row = config.public_data_row(offset)
        ws.cell(dst_row, 4).value = item.floor_label
        ws.cell(dst_row, config.col(5)).value = item.area
        offset += 1


PUBLIC_RATIO_ROW = 73
PUBLIC_RATIO_COL = 3

# 車位尺寸區（原始列號）
PARKING_SIZE_HEADER_ROW = 59  # 停車位類別表頭（如 230*550）
PARKING_SIZE_LABEL_ROW = 71  # 車位尺寸（標題文字，如 230*550）
PARKING_ACTUAL_AREA_ROW = 72  # 實際面積（寬×長，公尺）
PARKING_COUNT_ROW = 73  # 車位數
PARKING_RATIO_ROW = 74  # 車位尺寸比例
PARKING_SHARE_ROW = 77  # 車位持分/100000（框1）
PARKING_SIZE_COLS = (8, 9, 10, 11)

# 各戶資料列：公設計算「總」(R) 與土地持分「總」(U) 欄
PUBLIC_SHARE_COL = 18  # R：公設整數持分（=round(Q,0)）
LAND_SHARE_COL = 21  # U：土地整數持分（=round(T,0)）
BUILDING_TOTAL_SHARE = 100_000  # I3：建物總持分基準
LAND_TARGET_TOTAL = 100_000  # T4：土地目標總持分

# 微調過的持分以黃色底色標記
ADJUST_FILL = PatternFill(start_color="FFFFFF00", end_color="FFFFFF00", fill_type="solid")
# 車位尺寸各欄會出現內容的列（清除未使用欄位時一併清空）
PARKING_COLUMN_ROWS = (59, 66, 67, 68, 71, 72, 73, 74, 77)
# 跨欄合併、寬度需隨車位尺寸種類數調整的列（小車位數倍數／基準車位持分／車位總持分）
PARKING_SPAN_MERGE_ROWS = (75, 76, 78)
PARKING_SPAN_MERGE_RANGES = ("H75:K75", "H76:K76", "H78:K78")


# 車位尺寸比例的算法。未指定時沿用原本的面積比。
PARKING_RATIO_AREA = "area"
PARKING_RATIO_FIXED = "fixed"
PARKING_FLAT = "flat"
PARKING_MECHANICAL = "mechanical"
RATIO_NUMBER_FORMAT = "0.00"
# 範本預留的尺寸欄；超過這個數量時往右加欄，合計欄跟著右移。
PARKING_TEMPLATE_SIZE_COUNT = len(PARKING_SIZE_COLS)
PARKING_FLOOR_FIRST_ROW = 60
PARKING_FLOOR_LAST_ROW = 65


@dataclass
class ParkingSpace:
    """試算表上的一個車位欄。

    width、length 以公分為單位（例如 230、550）。
    share_ratio 有值時，持分計算直接用這個比例，不再用面積重算。
    ratio_source 為 area 時，非基準欄寫成相對基準面積的公式；fixed 時寫死數字。
    """

    width: float
    length: float
    count: int = 0
    note: str | None = None
    kind: str = PARKING_FLAT
    share_ratio: float | None = None
    ratio_source: str = "legacy"
    area_factor: float = 1.0
    floor_counts: tuple[int, ...] = ()

    @property
    def label(self) -> str:
        text = f"{_format_dimension(self.width)}*{_format_dimension(self.length)}"
        if self.note:
            text += f"\n({self.note})"
        return text

    @property
    def actual_area(self) -> float:
        """實際面積（平方公尺）。"""
        return round((self.width / 100) * (self.length / 100), 4)


@dataclass
class ParkingSizeInput:
    """一個共用尺寸，以及它的平面、機械數量。

    flat_counts／mechanical_counts 有給時，依樓層順序記錄數量；
    某尺寸＋平面／機械若各層合計為 0，該欄不輸出。
    沒給時沿用 flat_count／mechanical_count 的整棟合計，數量為 0 的欄不輸出。
    """

    width: float
    length: float
    flat_count: int = 0
    mechanical_count: int = 0
    note: str | None = None
    flat_counts: tuple[int, ...] | None = None
    mechanical_counts: tuple[int, ...] | None = None


def _format_dimension(value: float) -> str:
    return f"{value:g}"


def _column_note(note: str | None, kind: str, show_kind: bool) -> str | None:
    parts: list[str] = []
    if note:
        parts.append(note)
    if show_kind:
        parts.append("平面" if kind == PARKING_FLAT else "機械")
    if not parts:
        return None
    return "、".join(parts)


def _uses_floor_counts(sizes: Sequence[ParkingSizeInput]) -> bool:
    return any(
        size.flat_counts is not None or size.mechanical_counts is not None
        for size in sizes
    )


def _aligned_floor_counts(
    sizes: Sequence[ParkingSizeInput], kind: str
) -> list[tuple[int, ...]]:
    """把各尺寸的樓層數量補成同一長度；缺漏的一側以 0 補齊。"""
    length: int | None = None
    for size in sizes:
        for counts in (size.flat_counts, size.mechanical_counts):
            if counts is None:
                continue
            if length is None:
                length = len(counts)
            elif len(counts) != length:
                raise ValueError("各尺寸的樓層數量不一致。")
    if length is None:
        raise ValueError("沒有樓層數量。")

    attr = "flat_counts" if kind == PARKING_FLAT else "mechanical_counts"
    aligned: list[tuple[int, ...]] = []
    for size in sizes:
        counts = getattr(size, attr)
        if counts is None:
            aligned.append(tuple(0 for _ in range(length)))
        else:
            aligned.append(tuple(int(count) for count in counts))
    return aligned


def expand_parking_sizes(
    sizes: list[ParkingSizeInput],
    mode: str = PARKING_RATIO_AREA,
    mechanical_ratio: float | None = None,
) -> list[ParkingSpace]:
    """把共用尺寸展開成試算表欄：先平面、再機械，並寫入車位尺寸比例。

    只有平面、且依實際面積時，維持原本的欄位（不標「平面」、比例留給試算表公式）。
    平面與機械都有數量時才使用 mechanical_ratio。
    有樓層數量時，以「尺寸＋平面／機械」為單位：各層合計 > 0 才輸出該欄；
    某一層為 0、其他層有數時仍保留該欄並寫 0。
    """
    if _uses_floor_counts(sizes):
        flat_rows = _aligned_floor_counts(sizes, PARKING_FLAT)
        mechanical_rows = _aligned_floor_counts(sizes, PARKING_MECHANICAL)
        has_flat = any(count > 0 for counts in flat_rows for count in counts)
        has_mech = any(count > 0 for counts in mechanical_rows for count in counts)
        show_kind = has_mech
        columns: list[ParkingSpace] = []
        if has_flat:
            for size, counts in zip(sizes, flat_rows, strict=True):
                total = sum(counts)
                if total <= 0:
                    continue
                columns.append(
                    ParkingSpace(
                        width=size.width,
                        length=size.length,
                        count=total,
                        note=_column_note(None, PARKING_FLAT, show_kind),
                        kind=PARKING_FLAT,
                        floor_counts=counts,
                    )
                )
        if has_mech:
            for size, counts in zip(sizes, mechanical_rows, strict=True):
                total = sum(counts)
                if total <= 0:
                    continue
                columns.append(
                    ParkingSpace(
                        width=size.width,
                        length=size.length,
                        count=total,
                        note=_column_note(None, PARKING_MECHANICAL, show_kind),
                        kind=PARKING_MECHANICAL,
                        floor_counts=counts,
                    )
                )
        _apply_parking_ratios(columns, mode, mechanical_ratio, has_flat, has_mech)
        return columns

    has_flat = any(size.flat_count > 0 for size in sizes)
    has_mech = any(size.mechanical_count > 0 for size in sizes)
    show_kind = has_mech
    columns = []
    for kind, count_of in (
        (PARKING_FLAT, lambda size: size.flat_count),
        (PARKING_MECHANICAL, lambda size: size.mechanical_count),
    ):
        for size in sizes:
            count = count_of(size)
            if count <= 0:
                continue
            columns.append(
                ParkingSpace(
                    width=size.width,
                    length=size.length,
                    count=count,
                    note=_column_note(size.note, kind, show_kind),
                    kind=kind,
                )
            )
    _apply_parking_ratios(columns, mode, mechanical_ratio, has_flat, has_mech)
    return columns


def parking_floor_count_of(config: LayoutConfig) -> int:
    return PARKING_TEMPLATE_FLOOR_COUNT + config.parking_floor_extra_rows


def iter_parking_floor_rows(config: LayoutConfig) -> list[int]:
    """停車位樓層在試算表上的實際列號，合計列緊接在最後一層下面。"""
    count = parking_floor_count_of(config)
    if count <= 0:
        return []
    first = config.map_row(PARKING_FLOOR_FIRST_ROW)
    return [first + offset for offset in range(count)]


def _is_dropped_template_floor_row(src_row: int, config: LayoutConfig) -> bool:
    """樓層少於六層時，範本多出來的 B 列不寫入，位置讓給往上移的合計。"""
    if config.parking_floor_extra_rows >= 0:
        return False
    kept_last = PARKING_FLOOR_FIRST_ROW + parking_floor_count_of(config) - 1
    return PARKING_FLOOR_FIRST_ROW <= src_row <= PARKING_FLOOR_LAST_ROW and src_row > kept_last


def displayed_parking_spaces(
    parking_spaces: Sequence[ParkingSpace | None],
    *,
    keep_zeros: bool = False,
) -> list[ParkingSpace]:
    """試算表要留下的車位欄。預設只留數量 > 0 的欄。

    keep_zeros 僅供相容舊呼叫；為 True 時仍會留下合計為 0 的欄。
    """
    if keep_zeros:
        return [space for space in parking_spaces if space is not None]
    return used_parking_spaces(list(parking_spaces))


def _apply_parking_ratios(
    columns: list[ParkingSpace],
    mode: str,
    mechanical_ratio: float | None,
    has_flat: bool,
    has_mech: bool,
) -> None:
    if not columns:
        return
    if mode == PARKING_RATIO_AREA and not has_mech:
        return
    both = has_flat and has_mech
    if both and mechanical_ratio is None:
        raise ValueError("平面與機械都有數量時，需要機械位比例。")

    if mode == PARKING_RATIO_FIXED and not both:
        for column in columns:
            column.share_ratio = 1.0
            column.ratio_source = PARKING_RATIO_FIXED
        return
    if mode == PARKING_RATIO_FIXED:
        for column in columns:
            column.ratio_source = PARKING_RATIO_FIXED
            column.share_ratio = (
                1.0 if column.kind == PARKING_FLAT else float(mechanical_ratio)
            )
        return
    if mode != PARKING_RATIO_AREA:
        raise ValueError(f"未知的車位比例算法：{mode}")

    base_area = columns[0].actual_area
    if base_area == 0:
        raise ValueError("車位實際面積不可為 0。")
    for idx, column in enumerate(columns):
        column.ratio_source = PARKING_RATIO_AREA
        factor = 1.0
        if both and column.kind == PARKING_MECHANICAL:
            factor = float(mechanical_ratio)
        column.area_factor = factor
        if idx == 0:
            column.share_ratio = 1.0
        else:
            column.share_ratio = excel_round(
                column.actual_area / base_area * factor, 2
            )


def used_parking_spaces(
    parking_spaces: list[ParkingSpace | None],
) -> list[ParkingSpace]:
    """取得實際有車位（數量 > 0）的尺寸，依原順序排列。"""
    return [
        space
        for space in parking_spaces
        if space is not None and space.count > 0
    ]


def _parking_style_bounds(
    config: LayoutConfig, parking_size_count: int
) -> tuple[int, int, int]:
    """尺寸欄與合計欄的實際欄號。有幾欄就畫幾欄，合計緊接其右。"""
    size_first = config.col(PARKING_SIZE_COLS[0])
    if parking_size_count <= 0:
        # 無尺寸欄時合計仍占原第一尺寸欄位置，避免樣式往左畫到樓別。
        return size_first, size_first - 1, size_first
    size_last = size_first + parking_size_count - 1
    return size_first, size_last, size_last + 1


def apply_parking_spaces(
    ws: Worksheet,
    config: LayoutConfig,
    parking_spaces: list[ParkingSpace | None],
    floors: Sequence[str] | None = None,
) -> None:
    """依使用者輸入改寫車位尺寸、實際面積、車位數與尺寸比例。

    只留數量 > 0 的欄。沒有樓層清單時，車位數寫成合計數字；
    有樓層清單時，各層格子寫數量（單層可為 0），車位數維持合計公式。
    合計列緊接在最後一個樓層下面；合計欄緊接最後一個有車位尺寸右側。
    平面欄在前、機械欄在後。超過範本四欄時往右加欄。
    """
    used = displayed_parking_spaces(parking_spaces)
    header_row = config.map_row(PARKING_SIZE_HEADER_ROW)
    label_row = config.map_row(PARKING_SIZE_LABEL_ROW)
    area_row = config.map_row(PARKING_ACTUAL_AREA_ROW)
    count_row = config.map_row(PARKING_COUNT_ROW)
    ratio_row = config.map_row(PARKING_RATIO_ROW)
    sum_row = config.map_row(66)
    clear_rows = [config.map_row(row) for row in PARKING_COLUMN_ROWS]
    start_col = config.col(PARKING_SIZE_COLS[0])
    template_total_col = start_col + PARKING_TEMPLATE_SIZE_COUNT

    # 先清掉範本預留的尺寸欄與合計欄，再依實際欄數重寫。
    for idx in range(PARKING_TEMPLATE_SIZE_COUNT + 1):
        col = start_col + idx
        for row in clear_rows:
            ws.cell(row, col).value = None
        for floor_row in iter_parking_floor_rows(config):
            ws.cell(floor_row, col).value = None

    for idx, space in enumerate(used):
        col = start_col + idx
        ws.cell(header_row, col).value = space.label
        ws.cell(label_row, col).value = space.label
        ws.cell(area_row, col).value = space.actual_area
        letter = get_column_letter(col)
        if floors is None:
            ws.cell(count_row, col).value = space.count
        else:
            ws.cell(count_row, col).value = f"={letter}{sum_row}"
        _write_parking_ratio(
            ws, col, idx, space, start_col, ratio_row, area_row
        )
        # 尺寸欄已整段清空，面積㎡／坪／持分公式一律重寫（樓層數量稍後覆寫）。
        _write_extra_size_formulas(ws, config, col, start_col)

    if used:
        _place_parking_total_column(ws, config, start_col, len(used))
        # 實際欄數少於範本時，舊合計欄可能還留著表頭文字。
        if len(used) < PARKING_TEMPLATE_SIZE_COUNT:
            ws.cell(header_row, template_total_col).value = None
        for col in range(start_col, start_col + len(used) + 1):
            ws.column_dimensions[get_column_letter(col)].width = 13
    else:
        # 無車位時不保留「大小車位面積比例」整塊（含殘留 0／#DIV/0!）。
        _clear_parking_ratio_block(ws, config)

    if floors is not None:
        _write_parking_floors(ws, config, used, floors, start_col)


def _write_parking_ratio(
    ws: Worksheet,
    col: int,
    idx: int,
    space: ParkingSpace,
    start_col: int,
    ratio_row: int,
    area_row: int,
) -> None:
    if space.ratio_source == "legacy":
        if idx == 0:
            # 最左欄為尺寸比例基準（其餘欄沿用試算表的面積公式）。
            ws.cell(ratio_row, col).value = 1
        return

    cell = ws.cell(ratio_row, col)
    cell.number_format = RATIO_NUMBER_FORMAT
    if space.ratio_source == PARKING_RATIO_FIXED or idx == 0:
        cell.value = 1 if space.share_ratio is None else space.share_ratio
        return

    col_letter = get_column_letter(col)
    base_letter = get_column_letter(start_col)
    if space.area_factor == 1:
        cell.value = f"=round({col_letter}{area_row}/${base_letter}${area_row},2)"
    else:
        cell.value = (
            f"=round({col_letter}{area_row}/${base_letter}${area_row}"
            f"*{space.area_factor:.2f},2)"
        )


def _write_extra_size_formulas(
    ws: Worksheet, config: LayoutConfig, col: int, start_col: int
) -> None:
    """寫入尺寸欄的樓層合計、面積㎡／坪與持分公式。"""
    letter = get_column_letter(col)
    base_letter = get_column_letter(start_col)
    floor_rows = iter_parking_floor_rows(config)
    for row in floor_rows:
        ws.cell(row, col).value = None

    sum_row = config.map_row(66)
    area_m2_row = config.map_row(67)
    ping_row = config.map_row(68)
    share_row = config.map_row(PARKING_SHARE_ROW)
    ratio_row = config.map_row(PARKING_RATIO_ROW)
    base_share_row = config.map_row(76)
    public_row = config.map_row(58)
    public_col = get_column_letter(config.col(4))
    ws.cell(sum_row, col).value = _vertical_sum(letter, floor_rows)
    ws.cell(area_m2_row, col).value = (
        f"=round(${public_col}${public_row}*{letter}{share_row}/100000,2)"
    )
    ws.cell(ping_row, col).value = f"=round({letter}{area_m2_row}*0.3025,2)"
    ws.cell(share_row, col).value = (
        f"=round(${base_letter}${base_share_row}*{letter}{ratio_row},0)"
    )


def _place_parking_total_column(
    ws: Worksheet,
    config: LayoutConfig,
    start_col: int,
    used_count: int,
) -> None:
    """合計欄移到尺寸欄右側，並讓小車位數倍數、車位總持分加總所有尺寸欄。"""
    total_col = start_col + used_count
    total_letter = get_column_letter(total_col)
    first_letter = get_column_letter(start_col)
    last_letter = get_column_letter(start_col + used_count - 1)
    header_row = config.map_row(PARKING_SIZE_HEADER_ROW)
    ws.cell(header_row, total_col).value = "合計"

    floor_rows = iter_parking_floor_rows(config)
    for row in floor_rows:
        ws.cell(row, total_col).value = (
            f"=SUM({first_letter}{row}:{last_letter}{row})"
        )
    sum_row = config.map_row(66)
    ws.cell(sum_row, total_col).value = _vertical_sum(total_letter, floor_rows)

    count_row = config.map_row(PARKING_COUNT_ROW)
    ratio_row = config.map_row(PARKING_RATIO_ROW)
    share_row = config.map_row(PARKING_SHARE_ROW)
    multiple_parts: list[str] = []
    total_parts: list[str] = []
    for offset in range(used_count):
        letter = get_column_letter(start_col + offset)
        multiple_parts.append(f"{letter}{count_row}*{letter}{ratio_row}")
        total_parts.append(f"{letter}{count_row}*{letter}{share_row}")
    ws.cell(config.map_row(75), start_col).value = "=" + "+".join(multiple_parts)
    ws.cell(config.map_row(78), start_col).value = "=" + "+".join(total_parts)


def _vertical_sum(letter: str, rows: Sequence[int]) -> str | int:
    if not rows:
        return 0
    return f"=sum({letter}{rows[0]}:{letter}{rows[-1]})"


def _write_parking_floors(
    ws: Worksheet,
    config: LayoutConfig,
    columns: Sequence[ParkingSpace],
    floors: Sequence[str],
    start_col: int,
) -> None:
    """把樓層名稱和各層數量寫進停車位表，並讓合計公式只加總這些樓層。"""
    if len(floors) != parking_floor_count_of(config):
        raise ValueError("樓層數量和版面不一致。")
    rows = iter_parking_floor_rows(config)
    label_col = config.col(PARKING_GRID_FIRST_COL)
    for index, row in enumerate(rows):
        ws.cell(row, label_col).value = floors[index]
        for col_index, space in enumerate(columns):
            if len(space.floor_counts) != len(floors):
                raise ValueError("車位欄的樓層數量和樓層清單不一致。")
            ws.cell(row, start_col + col_index).value = space.floor_counts[index]

    column_count = len(columns)
    total_col = start_col + column_count
    sum_row = config.map_row(66)
    for col_index in range(column_count):
        letter = get_column_letter(start_col + col_index)
        ws.cell(sum_row, start_col + col_index).value = _vertical_sum(letter, rows)

    if column_count:
        first_letter = get_column_letter(start_col)
        last_letter = get_column_letter(start_col + column_count - 1)
        for row in rows:
            ws.cell(row, total_col).value = (
                f"=SUM({first_letter}{row}:{last_letter}{row})"
            )
        ws.cell(sum_row, total_col).value = _vertical_sum(
            get_column_letter(total_col), rows
        )
    else:
        for row in rows:
            ws.cell(row, total_col).value = 0
        ws.cell(sum_row, total_col).value = 0


def _clear_parking_ratio_block(ws: Worksheet, config: LayoutConfig) -> None:
    """車位全為 0 時清掉「大小車位面積比例」區塊的標題、列標與公式。"""
    label_col = config.col(PARKING_GRID_FIRST_COL)
    start_col = config.col(PARKING_SIZE_COLS[0])
    last_col = start_col + PARKING_TEMPLATE_SIZE_COUNT
    for src_row in range(70, PARKING_CALC_LAST_ROW + 1):
        row = config.map_row(src_row)
        for col in range(label_col, last_col + 1):
            ws.cell(row, col).value = None


def apply_parking_span_merges(
    ws: Worksheet,
    config: LayoutConfig,
    used_count: int,
) -> None:
    """依實際有車位的尺寸欄數合併標題、比例列與持分列。

    停車位標題＝樓別＋尺寸（不含合計；無尺寸時含資料欄）。
    大小車位面積比例＝左側表頭欄＋尺寸欄（單欄也合併，避免標題凸出）。
    持分列僅在尺寸 ≥ 2 欄時合併。
    """
    start_col = config.col(PARKING_SIZE_COLS[0])
    label_col = start_col - 1
    title_row = config.map_row(58)
    proportion_row = config.map_row(70)
    if used_count <= 0:
        # 無尺寸欄時資料欄在 start_col（全 0），標題仍跨樓別＋該欄。
        ws.merge_cells(
            start_row=title_row,
            start_column=label_col,
            end_row=title_row,
            end_column=start_col,
        )
        return
    end_col = start_col + used_count - 1
    ws.merge_cells(
        start_row=title_row,
        start_column=label_col,
        end_row=title_row,
        end_column=end_col,
    )
    # 範本標題寫在第一尺寸欄；合併後改放到表頭欄（合併區左上角）。
    proportion_title = ws.cell(proportion_row, start_col).value
    if proportion_title is not None:
        ws.cell(proportion_row, label_col).value = proportion_title
        if label_col != start_col:
            ws.cell(proportion_row, start_col).value = None
    ws.merge_cells(
        start_row=proportion_row,
        start_column=label_col,
        end_row=proportion_row,
        end_column=end_col,
    )
    if used_count >= 2:
        for src_row in PARKING_SPAN_MERGE_ROWS:
            row = config.map_row(src_row)
            ws.merge_cells(
                start_row=row,
                start_column=start_col,
                end_row=row,
                end_column=end_col,
            )


def apply_share_adjustments(
    ws: Worksheet,
    config: LayoutConfig,
    used: list[ParkingSpace],
    unit_data: list[UnitData] | None,
    public_areas: list[PublicAreaData] | None,
    public_ratio: float | None,
) -> ShareAdjustment | None:
    """以每尺寸 ±1 微調車位持分，讓車位總持分（框2）盡量接近概算（框3）。

    調整後的持分公式改寫為 round($H$76*H74,0)+1 或 -1 的形式，並標上黃色底色。
    回傳計算結果（含微調後車位總持分），資料不足時回傳 None。
    """
    adjustment = compute_share_adjustments(
        unit_data,
        public_areas,
        public_ratio,
        used,
        include_porch=config.include_porch,
    )
    if adjustment is None:
        return None
    if not adjustment.has_adjustment:
        return adjustment

    share_row = config.map_row(PARKING_SHARE_ROW)
    for idx, delta in enumerate(adjustment.deltas):
        if delta == 0:
            continue
        col = config.col(PARKING_SIZE_COLS[0]) + idx
        cell = ws.cell(share_row, col)
        formula = cell.value
        if isinstance(formula, str) and formula.startswith("="):
            cell.value = f"{formula}{'+' if delta > 0 else '-'}{abs(delta)}"
        cell.fill = ADJUST_FILL
    return adjustment


def _apply_row_share_shifts(
    ws: Worksheet,
    config: LayoutConfig,
    shifts: dict[int, int],
    col: int,
) -> None:
    """將各戶整數持分列依 shift（±1）改寫公式並標上黃色底色。"""
    for offset, shift in shifts.items():
        if shift == 0:
            continue
        row = config.data_start + offset
        cell = ws.cell(row, col)
        formula = cell.value
        if isinstance(formula, str) and formula.startswith("="):
            cell.value = f"{formula}{'+' if shift > 0 else '-'}{abs(shift)}"
        cell.fill = ADJUST_FILL


def apply_unit_share_adjustments(
    ws: Worksheet,
    config: LayoutConfig,
    unit_data: list[UnitData],
    parking_total: int,
) -> list[AdjustmentStatus]:
    """以每列 ±1 微調公設計算(R)與土地持分(U)，讓實際加總「精準命中」各自目標。

    公設目標 = 100000 - 車位總持分；土地目標 = 100000。同面積（取相同）的列會
    一起調整；只有能精準命中目標時才套用並標上黃色底色，否則維持原狀。
    回傳兩項（公設、土地）的完成狀態。
    """
    weights = [unit.unit_count for unit in unit_data]
    areas = [allocation_area(unit, config.include_porch) for unit in unit_data]
    total_area = sum(w * a for w, a in zip(weights, areas))
    if total_area == 0:
        return []

    public_target = BUILDING_TOTAL_SHARE - parking_total
    public_rows = build_share_rows(weights, areas, public_target, total_area)
    public_result = solve_share_adjustments(public_rows, public_target)
    _apply_row_share_shifts(ws, config, public_result.shifts, config.col(PUBLIC_SHARE_COL))

    land_rows = build_share_rows(weights, areas, LAND_TARGET_TOTAL, total_area)
    land_result = solve_share_adjustments(land_rows, LAND_TARGET_TOTAL)
    _apply_row_share_shifts(ws, config, land_result.shifts, config.col(LAND_SHARE_COL))

    return [
        AdjustmentStatus(
            name="公設計算",
            target=public_result.target,
            actual=public_result.adjusted_total,
            completed=public_result.completed,
        ),
        AdjustmentStatus(
            name="土地持分",
            target=land_result.target,
            actual=land_result.adjusted_total,
            completed=land_result.completed,
        ),
    ]


def create_template(
    output_path: Path,
    units: int = 10,
    unit_data: list[UnitData] | None = None,
    public_areas: list[PublicAreaData] | None = None,
    sheet_name: str | None = None,
    public_ratio: float | None = None,
    parking_spaces: list[ParkingSpace | None] | None = None,
    parking_floors: Sequence[str] | None = None,
    include_porch: bool = False,
    include_terrace: bool = False,
) -> BuildResult:
    if unit_data is not None:
        units = len(unit_data)
    floor_count = (
        len(parking_floors)
        if parking_floors is not None
        else PARKING_TEMPLATE_FLOOR_COUNT
    )
    extra_rows = (
        compute_public_area_extra_rows(public_areas, floor_count)
        if public_areas
        else 0
    )
    config = LayoutConfig(
        units=units,
        public_area_extra_rows=extra_rows,
        parking_floor_extra_rows=floor_count - PARKING_TEMPLATE_FLOOR_COUNT,
        include_porch=include_porch,
        include_terrace=include_terrace,
    )
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name or SHEET_TITLE

    apply_column_widths(ws, config)
    build_header(ws, config)
    build_data_rows(ws, config, unit_data)
    build_total_row(ws, config)
    build_parking_section(ws, config)
    parking_count: int | None = None
    displayed: list[ParkingSpace] | None = None
    if parking_spaces is not None or parking_floors is not None:
        apply_parking_spaces(
            ws, config, list(parking_spaces or []), floors=parking_floors
        )
        displayed = displayed_parking_spaces(list(parking_spaces or []))
        parking_count = len(displayed)
    if public_ratio is not None:
        ws.cell(config.map_row(PUBLIC_RATIO_ROW), PUBLIC_RATIO_COL).value = public_ratio
    if public_areas:
        clear_public_area_data(ws, config, public_areas)
        fill_public_areas(ws, config, public_areas)
    skip_ranges: set[str] | None = None
    if parking_count is not None:
        # 標題／比例／持分合併一律依實際欄數重做，避免撐滿範本四欄。
        skip_ranges = set(PARKING_SPAN_MERGE_RANGES)
        skip_ranges.update({"G58:K58", "H70:K70"})
    apply_merged_cells(ws, config, skip_ranges=skip_ranges)
    if parking_count is not None:
        apply_parking_span_merges(ws, config, parking_count)
    if public_areas:
        apply_public_area_merges(ws, config, public_areas)
        update_public_area_total_formula(ws, config, public_areas)
    adjustments: list[AdjustmentStatus] = []
    if displayed:
        parking_adjustment = apply_share_adjustments(
            ws,
            config,
            displayed,
            unit_data,
            public_areas,
            public_ratio,
        )
        if parking_adjustment is not None:
            adjustments.append(
                AdjustmentStatus(
                    name="車位持分",
                    target=parking_adjustment.target_total,
                    actual=parking_adjustment.adjusted_total,
                    # 車位容忍誤差 20 以內即視為完成（公設／土地才需完全精準）。
                    completed=(
                        abs(
                            parking_adjustment.adjusted_total
                            - parking_adjustment.target_total
                        )
                        <= MAX_TOLERANCE
                    ),
                )
            )
            if unit_data:
                adjustments.extend(
                    apply_unit_share_adjustments(
                        ws, config, unit_data, parking_adjustment.adjusted_total
                    )
                )

    apply_appearance(
        ws,
        config,
        parking_size_count=parking_count if parking_count is not None else 4,
    )
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return BuildResult(path=output_path, adjustments=tuple(adjustments))
