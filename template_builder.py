from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import range_boundaries
from openpyxl.worksheet.worksheet import Worksheet

from csv_import import PublicAreaData, UnitData
from formula_utils import LayoutConfig, data_row_formulas, rewrite_formula
from parking_balance import MAX_TOLERANCE, ShareAdjustment, compute_share_adjustments
from share_tweak import build_share_rows, solve_share_adjustments
from template_schema import (
    COLUMN_WIDTHS,
    HEADER_FORMULAS,
    HEADER_STATIC,
    MAX_COLUMN,
    MERGED_RANGES,
    PARKING_END_ROW,
    PARKING_FORMULAS,
    PARKING_NOTE_PLACEHOLDER,
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
PARKING_NOTE_COL = 2


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


def apply_column_widths(ws: Worksheet) -> None:
    for col_letter, width in COLUMN_WIDTHS.items():
        ws.column_dimensions[col_letter].width = width


def shift_range_rows(
    min_row: int, max_row: int, config: LayoutConfig
) -> tuple[int, int]:
    if min_row >= config.source_parking_start:
        base_min = min_row - config.rows_removed
        base_max = max_row - config.rows_removed
        if min_row > PUBLIC_AREA_END_ROW:
            extra = config.public_area_extra_rows
            return base_min + extra, base_max + extra
        return base_min, base_max
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
            start_column=min_col,
            end_row=max_row,
            end_column=max_col,
        )


def set_formula(ws: Worksheet, row: int, col: int, formula: str, config: LayoutConfig) -> None:
    ws.cell(row, col).value = rewrite_formula(formula, config)


def build_header(ws: Worksheet, config: LayoutConfig) -> None:
    for (row, col), value in HEADER_STATIC.items():
        ws.cell(row, col).value = value
    for (row, col), formula in HEADER_FORMULAS.items():
        set_formula(ws, row, col, formula, config)
    ws.cell(1, 9).value = None
    ws.cell(1, 10).value = None


def build_data_rows(
    ws: Worksheet,
    config: LayoutConfig,
    unit_data: list[UnitData] | None = None,
) -> None:
    for offset in range(config.units):
        row = config.data_start + offset
        for col in INPUT_COLUMNS:
            ws[f"{col}{row}"].value = None
        for col, formula in data_row_formulas(row, config.total_row).items():
            ws[f"{col}{row}"].value = formula

        if unit_data is not None and offset < len(unit_data):
            unit = unit_data[offset]
            ws[f"A{row}"] = unit.unit_count
            ws[f"B{row}"] = unit.label
            ws[f"C{row}"] = unit.main_area
            ws[f"D{row}"] = unit.balcony_area


def build_total_row(ws: Worksheet, config: LayoutConfig) -> None:
    row = config.total_row
    for col, text in TOTAL_STATIC.items():
        ws.cell(row, col).value = text
    for col, formula in TOTAL_FORMULAS.items():
        set_formula(ws, row, col, formula, config)


def build_parking_section(ws: Worksheet, config: LayoutConfig) -> None:
    row_count = PARKING_END_ROW - PARKING_START_ROW + 1
    for offset in range(row_count):
        src_row = PARKING_START_ROW + offset
        dst_row = config.map_row(src_row)

        if src_row == PARKING_NOTE_ROW:
            ws.cell(dst_row, PARKING_NOTE_COL).value = PARKING_NOTE_PLACEHOLDER

        for (static_row, col), text in PARKING_STATIC.items():
            if static_row == src_row:
                ws.cell(dst_row, col).value = text

        for (formula_row, col), formula in PARKING_FORMULAS.items():
            if formula_row == src_row:
                set_formula(ws, dst_row, col, formula, config)


PUBLIC_AREA_SLOT_ROWS = PUBLIC_AREA_END_ROW - PUBLIC_AREA_START_ROW + 1


def public_area_rows_needed(public_areas: list[PublicAreaData]) -> int:
    left_items = [item for item in public_areas if item.side == "left"]
    right_items = [item for item in public_areas if item.side == "right"]
    left_rows = sum(2 if item.is_per_floor else 1 for item in left_items)
    right_rows = len(right_items)
    return max(left_rows, right_rows)


def compute_public_area_extra_rows(public_areas: list[PublicAreaData]) -> int:
    return max(0, public_area_rows_needed(public_areas) - PUBLIC_AREA_SLOT_ROWS)


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
        for col in (2, 3, 4, 5):
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
        parts.append(f"sum(E{area_start}:E{right_last})")
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
        ws.cell(dst_row, 5).value = item.area
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


@dataclass
class ParkingSpace:
    """單一車位尺寸與數量。

    width、length 以公分為單位（例如 230、550）。
    """

    width: float
    length: float
    count: int = 0
    note: str | None = None

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


def _format_dimension(value: float) -> str:
    return f"{value:g}"


def used_parking_spaces(
    parking_spaces: list[ParkingSpace | None],
) -> list[ParkingSpace]:
    """取得實際有車位（數量 > 0）的尺寸，依原順序排列。"""
    return [
        space
        for space in parking_spaces
        if space is not None and space.count > 0
    ]


def apply_parking_spaces(
    ws: Worksheet,
    config: LayoutConfig,
    parking_spaces: list[ParkingSpace | None],
) -> None:
    """依使用者輸入改寫車位尺寸、實際面積與車位數。

    僅保留數量 > 0 的尺寸，並靠左壓縮排列；數量為 0 或未填的尺寸欄位會被完全清空，
    以免閱讀者看到沒有車位的尺寸欄位而混淆。第一個（最左）欄為比例基準（=1）。
    """
    used = used_parking_spaces(parking_spaces)

    header_row = config.map_row(PARKING_SIZE_HEADER_ROW)
    label_row = config.map_row(PARKING_SIZE_LABEL_ROW)
    area_row = config.map_row(PARKING_ACTUAL_AREA_ROW)
    count_row = config.map_row(PARKING_COUNT_ROW)
    ratio_row = config.map_row(PARKING_RATIO_ROW)
    clear_rows = [config.map_row(row) for row in PARKING_COLUMN_ROWS]

    for idx, col in enumerate(PARKING_SIZE_COLS):
        if idx < len(used):
            space = used[idx]
            ws.cell(header_row, col).value = space.label
            ws.cell(label_row, col).value = space.label
            ws.cell(area_row, col).value = space.actual_area
            ws.cell(count_row, col).value = space.count
            if idx == 0:
                # 最左欄為尺寸比例基準（試算表對其餘欄使用 =該欄面積/$H$72）。
                ws.cell(ratio_row, col).value = 1
        else:
            for row in clear_rows:
                ws.cell(row, col).value = None


def apply_parking_span_merges(
    ws: Worksheet,
    config: LayoutConfig,
    used_count: int,
) -> None:
    """將小車位數倍數／基準車位持分／車位總持分依車位尺寸種類數合併。

    有幾種尺寸（數量 > 0）就只合併幾格；僅 1 種時不需合併。
    """
    if used_count < 2:
        return
    start_col = PARKING_SIZE_COLS[0]
    end_col = start_col + used_count - 1
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
        unit_data, public_areas, public_ratio, used
    )
    if adjustment is None:
        return None
    if not adjustment.has_adjustment:
        return adjustment

    share_row = config.map_row(PARKING_SHARE_ROW)
    for idx, delta in enumerate(adjustment.deltas):
        if delta == 0:
            continue
        col = PARKING_SIZE_COLS[0] + idx
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
    areas = [unit.main_area + unit.balcony_area for unit in unit_data]
    total_area = sum(w * a for w, a in zip(weights, areas))
    if total_area == 0:
        return []

    public_target = BUILDING_TOTAL_SHARE - parking_total
    public_rows = build_share_rows(weights, areas, public_target, total_area)
    public_result = solve_share_adjustments(public_rows, public_target)
    _apply_row_share_shifts(ws, config, public_result.shifts, PUBLIC_SHARE_COL)

    land_rows = build_share_rows(weights, areas, LAND_TARGET_TOTAL, total_area)
    land_result = solve_share_adjustments(land_rows, LAND_TARGET_TOTAL)
    _apply_row_share_shifts(ws, config, land_result.shifts, LAND_SHARE_COL)

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
) -> BuildResult:
    if unit_data is not None:
        units = len(unit_data)
    extra_rows = compute_public_area_extra_rows(public_areas) if public_areas else 0
    config = LayoutConfig(units=units, public_area_extra_rows=extra_rows)
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name or SHEET_TITLE

    apply_column_widths(ws)
    build_header(ws, config)
    build_data_rows(ws, config, unit_data)
    build_total_row(ws, config)
    build_parking_section(ws, config)
    parking_count: int | None = None
    if parking_spaces:
        apply_parking_spaces(ws, config, parking_spaces)
        parking_count = len(used_parking_spaces(parking_spaces))
    if public_ratio is not None:
        ws.cell(config.map_row(PUBLIC_RATIO_ROW), PUBLIC_RATIO_COL).value = public_ratio
    if public_areas:
        clear_public_area_data(ws, config, public_areas)
        fill_public_areas(ws, config, public_areas)
    skip_ranges = (
        set(PARKING_SPAN_MERGE_RANGES) if parking_count is not None else None
    )
    apply_merged_cells(ws, config, skip_ranges=skip_ranges)
    if parking_count is not None:
        apply_parking_span_merges(ws, config, parking_count)
    if public_areas:
        apply_public_area_merges(ws, config, public_areas)
        update_public_area_total_formula(ws, config, public_areas)
    adjustments: list[AdjustmentStatus] = []
    if parking_spaces:
        parking_adjustment = apply_share_adjustments(
            ws,
            config,
            used_parking_spaces(parking_spaces),
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

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return BuildResult(path=output_path, adjustments=tuple(adjustments))
