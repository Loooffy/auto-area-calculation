from __future__ import annotations

from pathlib import Path

import openpyxl
import pytest

from csv_import import PublicAreaData, UnitData
from formula_utils import LayoutConfig, map_source_column, rewrite_formula
from parking_balance import compute_share_adjustments, excel_round, total_unit_area
from share_tweak import ShareRow, solve_share_adjustments
from template_builder import (
    LAND_SHARE_COL,
    PARKING_RATIO_AREA,
    PARKING_RATIO_FIXED,
    PUBLIC_SHARE_COL,
    AdjustmentStatus,
    BuildResult,
    ParkingSizeInput,
    ParkingSpace,
    create_template,
    expand_parking_sizes,
)
from template_schema import MERGED_CELL_COUNT


@pytest.fixture
def config() -> LayoutConfig:
    return LayoutConfig(units=10)


@pytest.fixture
def template_path(tmp_path: Path) -> Path:
    output = tmp_path / "template.xlsx"
    create_template(output, units=10)
    return output


@pytest.fixture
def ws(template_path: Path):
    wb = openpyxl.load_workbook(template_path, data_only=False)
    yield wb.active
    wb.close()


def test_rewrite_total_row_reference(config: LayoutConfig) -> None:
    assert rewrite_formula("=sum(A5:A53)", config) == "=sum(A5:A14)"
    assert rewrite_formula("=$Q$4*E5/$E$54", config) == "=$Q$4*E5/$E$15"


def test_rewrite_parking_reference(config: LayoutConfig) -> None:
    assert rewrite_formula("=D58", config) == "=D19"
    assert rewrite_formula("=I3-H78", config) == "=I3-H39"
    assert rewrite_formula("=E54*C73/C74", config) == "=E15*C34/C35"


def test_merged_cell_count(ws) -> None:
    assert len(ws.merged_cells.ranges) == MERGED_CELL_COUNT


def test_data_rows_empty_inputs(ws) -> None:
    for row in range(5, 15):
        for col in ("A", "B", "C", "D"):
            assert ws[f"{col}{row}"].value is None


def test_data_rows_have_formulas(ws) -> None:
    assert ws["E5"].value == "=C5+D5"
    assert ws["Q5"].value == "=round($Q$4*E5/$E$15,2)"
    assert ws["R5"].value == "=round(Q5,0)"
    assert ws["U5"].value == "=round(T5,0)"


def test_total_row(ws) -> None:
    assert ws["A15"].value == "=sum(A5:A14)"
    assert ws["K15"].value == "=sum(K5:K14)"
    assert ws["E15"].value == "=M15"
    assert ws["W15"].value == "=round(F15/G15*100,2)"


def test_header_formulas(ws) -> None:
    assert ws["F3"].value == "=D19"
    assert ws["Q4"].value == "=I3-H39"


def test_parking_formula_shifted(ws) -> None:
    assert ws["D19"].value == "=sum(C21:C30)+sum(E21:E29)"
    assert ws["C36"].value == "=round(E15*C34/C35,2)"
    assert ws["H39"].value == "=H34*H38+I34*I38+J34*J38+K34*K38"


def test_parking_size_ratio_formulas(ws) -> None:
    assert ws["H35"].value == "1"
    assert ws["I35"].value == "=round(I33/$H$33,2)"
    assert ws["J35"].value == "=round(J33/$H$33,2)"
    assert ws["K35"].value == "=round(K33/$H$33,2)"


def test_parking_share_reference(ws) -> None:
    assert ws["I16"].value == "=H39"


def test_date_cleared(ws) -> None:
    assert ws["I1"].value is None
    assert ws["J1"].value is None


def test_no_original_unit_data(ws) -> None:
    for row in range(5, 15):
        assert ws[f"B{row}"].value is None


def test_parking_numeric_inputs_cleared(ws) -> None:
    assert ws["C21"].value is None
    assert ws["H21"].value is None
    assert ws["C34"].value is None
    assert ws["D36"].value == 100_000


def test_main_title_preserved(ws) -> None:
    assert ws["B1"].value == "各戶面積及停車位分算明細"


def test_appearance_matches_reference_sheet(ws) -> None:
    assert ws.column_dimensions["A"].width == 13
    assert ws.column_dimensions["B"].width == 17.63
    assert ws.column_dimensions["D"].width == 14.63
    assert ws.column_dimensions["F"].width == 15
    assert ws.column_dimensions["G"].width == 15.38
    assert ws.column_dimensions["W"].width == 13

    title = ws["B1"]
    assert title.font.name == "DFKai-SB"
    assert title.font.size == 13
    assert title.alignment.horizontal == "center"
    assert title.alignment.vertical == "center"
    assert title.border.left.style == "thin"

    header = ws["A2"]
    assert header.font.name == "DFKai-SB"
    assert header.alignment.horizontal == "center"
    assert header.alignment.vertical == "center"
    assert header.alignment.wrap_text is True
    assert header.border.bottom.style == "thin"

    assert ws["A5"].alignment.horizontal == "center"
    assert ws["B5"].alignment.horizontal == "center"
    area = ws["E5"]
    assert area.alignment.horizontal == "right"
    assert area.alignment.vertical == "bottom"
    assert area.number_format == "0.00"
    assert area.border.right.style == "thin"
    assert ws["P5"].number_format == "0.00"
    assert ws["R5"].number_format == "General"
    assert ws["R5"].alignment.horizontal == "right"
    assert ws["I3"].number_format == "General"
    assert ws["F3"].number_format == "0.00"
    assert ws["F3"].alignment.horizontal == "center"

    assert ws["B15"].alignment.horizontal == "center"
    assert ws["A15"].alignment.horizontal == "center"
    assert ws["G15"].number_format == "0.00"
    assert ws["I15"].number_format == "General"

    config = LayoutConfig(units=10)
    note_cell = ws.cell(config.map_row(55), 2)
    assert note_cell.value is None
    assert note_cell.border.left.style is None
    assert not any(
        merged.min_row == config.map_row(55)
        and merged.min_col <= 2 <= merged.max_col
        for merged in ws.merged_cells.ranges
    )
    footnote = ws.cell(config.map_row(80), 2)
    assert footnote.value is None
    assert footnote.border.left.style is None
    assert ws.row_dimensions[config.map_row(59)].height == 32
    share_label = ws.cell(config.map_row(77), 7)
    assert share_label.value == "車位持分/100000"
    assert share_label.font.size == 10
    assert share_label.font.name == "DFKai-SB"


def test_per_floor_public_area_uses_area_format(tmp_path: Path) -> None:
    output = tmp_path / "public.xlsx"
    create_template(
        output,
        unit_data=[UnitData(label="1F-A1", main_area=10, balcony_area=1, unit_count=1)],
        public_areas=[
            PublicAreaData(
                floor_label="6~26F*21",
                area=343.2,
                is_per_floor=True,
                side="left",
            )
        ],
    )
    wb = openpyxl.load_workbook(output)
    ws = wb.active
    config = LayoutConfig(units=1)
    label = ws.cell(config.public_data_row(0), 2)
    per_floor = ws.cell(config.public_data_row(1), 2)
    total = ws.cell(config.public_data_row(0), 3)
    assert label.alignment.horizontal == "center"
    assert per_floor.value == 343.2
    assert per_floor.alignment.horizontal == "center"
    assert per_floor.alignment.vertical == "center"
    assert per_floor.number_format == "0.00"
    assert total.number_format == "0.00"
    wb.close()


def test_parking_labels(ws) -> None:
    assert ws["G19"].value == "停車位面積、類別、數量"
    assert ws["G27"].value == "合計"
    assert ws["G39"].value == "車位總持分"
    assert ws["H31"].value == "大小車位面積比例"


def _sample_balance_inputs() -> tuple[
    list[UnitData], list[PublicAreaData], float, list[ParkingSpace]
]:
    units = [UnitData(label="x", main_area=8653.64, balcony_area=0.0, unit_count=1)]
    public_areas = [
        PublicAreaData(floor_label="1F", area=9159.2, is_per_floor=False, side="left")
    ]
    spaces = [
        ParkingSpace(230, 550, 24),
        ParkingSpace(250, 550, 109),
        ParkingSpace(250, 600, 5),
        ParkingSpace(350, 600, 2, "無障礙"),
    ]
    return units, public_areas, 33.88, spaces


def test_compute_share_adjustments_hits_target() -> None:
    units, public_areas, ratio, spaces = _sample_balance_inputs()
    adjustment = compute_share_adjustments(units, public_areas, ratio, spaces)
    assert adjustment is not None
    assert adjustment.target_total == 51588
    assert adjustment.base_total == 51607
    assert adjustment.deltas == (-1, 0, 1, 0)
    assert adjustment.adjusted_total == 51588


def test_tolerance_allows_near_miss_adjustment() -> None:
    units = [UnitData(label="x", main_area=8653.64, balcony_area=0.0, unit_count=1)]
    public_areas = [
        PublicAreaData(floor_label="1F", area=9159.2, is_per_floor=False, side="left")
    ]
    spaces = [ParkingSpace(230, 550, 20), ParkingSpace(250, 550, 50)]

    # 無法精準命中：差距必須剛好為 0 時不調整。
    exact = compute_share_adjustments(units, public_areas, 33.88, spaces, tolerance=0)
    assert exact is not None
    assert exact.has_adjustment is False
    assert exact.deltas == (0, 0)

    # 放寬到 20 以內：選出差距最小的組合。
    relaxed = compute_share_adjustments(units, public_areas, 33.88, spaces, tolerance=20)
    assert relaxed is not None
    assert relaxed.has_adjustment is True
    assert relaxed.deltas == (0, 1)
    assert abs(relaxed.target_total - relaxed.adjusted_total) == 2


def test_no_adjustment_when_target_unreachable() -> None:
    units, public_areas, ratio, spaces = _sample_balance_inputs()
    # max_delta=0 代表完全不能調整，框2 ≠ 框3 時應回傳全 0（不改動）。
    adjustment = compute_share_adjustments(
        units, public_areas, ratio, spaces, max_delta=0
    )
    assert adjustment is not None
    assert adjustment.has_adjustment is False
    assert adjustment.deltas == (0, 0, 0, 0)
    assert adjustment.adjusted_total == adjustment.base_total


def test_compute_share_adjustments_requires_inputs() -> None:
    units, public_areas, ratio, spaces = _sample_balance_inputs()
    assert compute_share_adjustments(units, public_areas, None, spaces) is None
    assert compute_share_adjustments(units, [], ratio, spaces) is None
    assert compute_share_adjustments(None, public_areas, ratio, spaces) is None


def test_share_adjustment_applied_to_template(tmp_path: Path) -> None:
    units, public_areas, ratio, spaces = _sample_balance_inputs()
    output = tmp_path / "adjusted.xlsx"
    create_template(
        output,
        unit_data=units,
        public_areas=public_areas,
        public_ratio=ratio,
        parking_spaces=spaces,
    )
    wb = openpyxl.load_workbook(output)
    ws = wb.active
    share_row = LayoutConfig(units=1).map_row(77)

    h_cell = ws.cell(share_row, 8)
    j_cell = ws.cell(share_row, 10)
    assert h_cell.value.endswith("-1")
    assert j_cell.value.endswith("+1")
    assert h_cell.fill.patternType == "solid"
    assert h_cell.fill.fgColor.rgb == "FFFFFF00"
    assert j_cell.fill.fgColor.rgb == "FFFFFF00"

    for col in (9, 11):
        cell = ws.cell(share_row, col)
        assert not str(cell.value).endswith(("+1", "-1"))
        assert cell.fill.patternType is None
    wb.close()


def test_solve_share_adjustments_prefers_single_group() -> None:
    rows = [
        ShareRow(offset=0, weight=1, decimal=10.0, rounded=10),
        ShareRow(offset=1, weight=1, decimal=10.0, rounded=10),
        ShareRow(offset=2, weight=1, decimal=20.0, rounded=20),
    ]
    result = solve_share_adjustments(rows, target=41)
    assert result.base_total == 40
    assert result.adjusted_total == 41
    assert result.shifts == {2: 1}


def test_solve_share_adjustments_groups_equal_decimals() -> None:
    rows = [
        ShareRow(offset=0, weight=1, decimal=10.0, rounded=10),
        ShareRow(offset=1, weight=1, decimal=10.0, rounded=10),
        ShareRow(offset=2, weight=1, decimal=20.0, rounded=20),
    ]
    result = solve_share_adjustments(rows, target=42)
    # 只能靠面積相同的兩列一起 +1 才剛好命中。
    assert result.adjusted_total == 42
    assert result.shifts == {0: 1, 1: 1}


def test_solve_share_adjustments_exact_only() -> None:
    rows = [ShareRow(offset=0, weight=5, decimal=10.0, rounded=10)]
    # 基礎 50，目標 52，只能 ±5 無法精準命中 → 不調整、completed=False。
    result = solve_share_adjustments(rows, target=52)
    assert result.base_total == 50
    assert result.adjusted_total == 50
    assert result.shifts == {}
    assert result.completed is False


def test_solve_share_adjustments_no_negative_share() -> None:
    rows = [ShareRow(offset=0, weight=1, decimal=0.0, rounded=0)]
    # rounded 為 0 不允許 -1，目標 -3 時只能維持 0。
    result = solve_share_adjustments(rows, target=-3)
    assert result.adjusted_total == 0
    assert result.shifts == {}


def test_unit_share_adjustments_applied_to_template(tmp_path: Path) -> None:
    areas = [55.87, 60.69, 30.0, 30.0, 30.0, 48.9, 72.34, 72.34]
    units = [
        UnitData(label=f"u{i}", main_area=a, balcony_area=0.0, unit_count=1)
        for i, a in enumerate(areas)
    ]
    public_areas = [
        PublicAreaData(floor_label="1F", area=9159.2, is_per_floor=False, side="left")
    ]
    spaces = [ParkingSpace(230, 550, 24), ParkingSpace(250, 550, 109)]

    output = tmp_path / "unit_adjusted.xlsx"
    create_template(
        output,
        unit_data=units,
        public_areas=public_areas,
        public_ratio=33.88,
        parking_spaces=spaces,
    )
    wb = openpyxl.load_workbook(output)
    ws = wb.active
    config = LayoutConfig(units=len(units))

    # 收集每列公設(R)持分的微調尾碼與面積。
    shift_by_area: dict[float, set[str]] = {}
    any_highlight = False
    for offset, area in enumerate(areas):
        cell = ws.cell(config.data_start + offset, PUBLIC_SHARE_COL)
        value = str(cell.value)
        suffix = value[-2:] if value.endswith(("+1", "-1")) else ""
        shift_by_area.setdefault(area, set()).add(suffix)
        if suffix:
            assert cell.fill.patternType == "solid"
            assert cell.fill.fgColor.rgb == "FFFFFF00"
            any_highlight = True
        else:
            assert cell.fill.patternType is None

    # 面積相同的列必須有一致的調整（同組一起動）。
    for suffixes in shift_by_area.values():
        assert len(suffixes) == 1
    assert any_highlight
    wb.close()


def test_build_result_reports_all_adjustments(tmp_path: Path) -> None:
    units, public_areas, ratio, spaces = _sample_balance_inputs()
    output = tmp_path / "result.xlsx"
    result = create_template(
        output,
        unit_data=units,
        public_areas=public_areas,
        public_ratio=ratio,
        parking_spaces=spaces,
    )
    assert isinstance(result, BuildResult)
    assert result.path == output
    names = {item.name for item in result.adjustments}
    assert names == {"車位持分", "公設計算", "土地持分"}


def test_parking_within_tolerance_is_completed(tmp_path: Path) -> None:
    # 此組車位最佳調整僅能逼近到差 5（< 容忍誤差 20），應視為完成、不列入未完成。
    units = [
        UnitData(label=f"u{i}", main_area=a, balcony_area=0.0, unit_count=2)
        for i, a in enumerate((38.97, 37.39, 47.98))
    ]
    public_areas = [
        PublicAreaData(floor_label="1F", area=10486.84, is_per_floor=False, side="left")
    ]
    spaces = [ParkingSpace(230, 550, 10), ParkingSpace(250, 550, 10)]

    output = tmp_path / "tolerance.xlsx"
    result = create_template(
        output,
        unit_data=units,
        public_areas=public_areas,
        public_ratio=33.88,
        parking_spaces=spaces,
    )
    parking = next(item for item in result.adjustments if item.name == "車位持分")
    assert 0 < abs(parking.actual - parking.target) <= 20
    assert parking.completed is True
    assert result.incomplete_adjustments == ()


def test_build_result_incomplete_filter() -> None:
    result = BuildResult(
        path=Path("x.xlsx"),
        adjustments=(
            AdjustmentStatus(name="車位持分", target=100, actual=100, completed=True),
            AdjustmentStatus(name="公設計算", target=871, actual=870, completed=False),
        ),
    )
    incomplete = result.incomplete_adjustments
    assert len(incomplete) == 1
    assert incomplete[0].name == "公設計算"


def test_map_source_column_inserts_porch_columns() -> None:
    assert map_source_column(4, True) == 4
    assert map_source_column(5, True) == 6
    assert map_source_column(12, True) == 13
    assert map_source_column(13, True) == 15
    assert map_source_column(23, True) == 25
    assert map_source_column(5, False) == 5


def test_map_source_column_inserts_terrace_before_public_percent() -> None:
    # 露台插在公設%（欄 23）左側；停車位尺寸欄不受影響且保持連續。
    assert map_source_column(22, include_terrace=True) == 22
    assert map_source_column(23, include_terrace=True) == 24
    assert map_source_column(8, include_terrace=True) == 8
    size_cols = [map_source_column(col, include_terrace=True) for col in (8, 9, 10, 11)]
    assert size_cols == list(range(8, 12))
    assert map_source_column(23, include_porch=True, include_terrace=True) == 26
    assert map_source_column(8, include_porch=True, include_terrace=True) == 9
    assert map_source_column(13, include_porch=True, include_terrace=True) == 15
    porch_terrace_sizes = [
        map_source_column(col, include_porch=True, include_terrace=True)
        for col in (8, 9, 10, 11)
    ]
    assert porch_terrace_sizes == list(range(9, 13))
    assert LayoutConfig(include_terrace=True).terrace_col == 23
    assert LayoutConfig(include_porch=True, include_terrace=True).terrace_col == 25


def test_total_unit_area_includes_porch_when_enabled() -> None:
    units = [UnitData(label="a", main_area=10, balcony_area=1, porch_area=2, unit_count=3)]
    assert total_unit_area(units) == 33
    assert total_unit_area(units, include_porch=True) == 39


def _merged(ws, min_col: int, min_row: int, max_col: int, max_row: int) -> bool:
    return any(
        (
            merged.min_col,
            merged.min_row,
            merged.max_col,
            merged.max_row,
        )
        == (min_col, min_row, max_col, max_row)
        for merged in ws.merged_cells.ranges
    )


def test_porch_layout_shifts_columns_and_allocation_base(tmp_path: Path) -> None:
    output = tmp_path / "porch.xlsx"
    create_template(
        output,
        unit_data=[
            UnitData(label="1F-A", main_area=10, balcony_area=2, porch_area=1.5, unit_count=1),
            UnitData(label="1F-B", main_area=20, balcony_area=0, porch_area=0, unit_count=1),
        ],
        public_areas=[
            PublicAreaData(floor_label="屋突1", area=50, is_per_floor=False, side="right")
        ],
        include_porch=True,
    )
    wb = openpyxl.load_workbook(output, data_only=False)
    ws = wb.active

    assert ws["E2"].value == "門廊"
    assert ws["F2"].value == "主+附+門廊"
    assert ws["N2"].value == "門廊"
    assert ws["O2"].value == "主+附+門廊"
    assert ws["E4"].value == "㎡"
    assert ws["E5"].value == 1.5
    assert ws["E6"].value == 0
    assert ws["F5"].value == "=C5+D5+E5"
    assert ws["N5"].value == "=A5*E5"
    assert ws["O5"].value == "=A5*F5"
    assert ws["H5"].value == "=F5+G5"
    assert ws["S5"].value == "=round($S$4*F5/$F$7,2)"
    assert ws["T5"].value == "=round(S5,0)"
    assert ws["E7"].value == "=N7"
    assert ws["F7"].value == "=O7"
    assert ws["N7"].value == "=sum(N5:N6)"

    config = LayoutConfig(units=2, include_porch=True)
    public_row = config.public_data_row(0)
    assert ws.cell(public_row, 4).value == "屋突1"
    assert ws.cell(public_row, 5).value is None
    assert ws.cell(public_row, 6).value == 50
    assert _merged(ws, 4, public_row, 5, public_row)
    summary = ws.cell(config.map_row(58), 4).value
    assert summary == f"=sum(F{public_row}:F{public_row})"
    wb.close()


def test_porch_template_moves_parking_columns(tmp_path: Path) -> None:
    output = tmp_path / "porch-parking.xlsx"
    create_template(output, units=10, include_porch=True)
    wb = openpyxl.load_workbook(output, data_only=False)
    ws = wb.active
    assert ws["G3"].value == "=D19"
    assert ws["S4"].value == "=J3-I39"
    assert ws["J16"].value == "=I39"
    assert ws["D19"].value == "=sum(C21:C30)+sum(F21:F29)"
    assert ws["C36"].value == "=round(F15*C34/C35,2)"
    assert ws.cell(38, 9).value == "=round($I$37*I35,0)"
    assert _merged(ws, 5, 2, 5, 3)
    assert _merged(ws, 14, 2, 14, 4)
    assert _merged(ws, 4, 20, 5, 20)
    wb.close()


def test_terrace_layout_left_of_public_percent(tmp_path: Path) -> None:
    output = tmp_path / "terrace.xlsx"
    create_template(
        output,
        unit_data=[
            UnitData(label="1F-A", main_area=10, balcony_area=2, terrace_area=5, unit_count=1),
            UnitData(label="2F-A", main_area=20, balcony_area=0, terrace_area=0, unit_count=1),
        ],
        include_terrace=True,
    )
    wb = openpyxl.load_workbook(output, data_only=False)
    ws = wb.active
    # 露台緊貼公設%左側：W=露台，X=公設%
    assert ws["W2"].value == "備註：露台"
    assert ws["W4"].value == "㎡"
    assert ws["X2"].value == "公設%"
    assert ws["W5"].value == 5
    assert ws["W6"].value == 0
    assert ws["I2"].value == "公設持分"
    assert ws["I4"].value == "/100000"
    assert ws["H4"].value == "坪"
    assert ws["I5"].value == "=R5"
    assert ws["G5"].value == "=E5+F5"
    assert _merged(ws, 23, 2, 23, 3)
    wb.close()


def test_porch_and_terrace_layout(tmp_path: Path) -> None:
    output = tmp_path / "porch-terrace.xlsx"
    create_template(
        output,
        unit_data=[
            UnitData(
                label="1F-A",
                main_area=10,
                balcony_area=2,
                porch_area=1.5,
                terrace_area=3,
                unit_count=1,
            ),
            UnitData(
                label="1F-B",
                main_area=20,
                balcony_area=0,
                porch_area=0,
                terrace_area=0,
                unit_count=1,
            ),
        ],
        include_porch=True,
        include_terrace=True,
    )
    wb = openpyxl.load_workbook(output, data_only=False)
    ws = wb.active
    assert ws["E2"].value == "門廊"
    assert ws["F2"].value == "主+附+門廊"
    # 門廊讓公設%右移兩欄；露台仍緊貼其左（Y=露台，Z=公設%）
    assert ws["Y2"].value == "備註：露台"
    assert ws["Z2"].value == "公設%"
    assert ws["Y5"].value == 3
    assert ws["Y6"].value == 0
    assert ws["J2"].value == "公設持分"
    assert ws["J5"].value == "=T5"
    assert ws["F5"].value == "=C5+D5+E5"
    # 加權門廊不受露台影響（露台在欄 23 之後）
    assert ws["N2"].value == "門廊"
    assert ws["N5"].value == "=A5*E5"
    assert ws["O2"].value == "主+附+門廊"
    assert _merged(ws, 25, 2, 25, 3)
    assert _merged(ws, 14, 2, 14, 4)
    wb.close()


def _sizes_with_both_kinds() -> list[ParkingSizeInput]:
    return [
        ParkingSizeInput(230, 550, flat_count=10, mechanical_count=20),
        ParkingSizeInput(250, 550, flat_count=5, mechanical_count=0),
        ParkingSizeInput(250, 600, flat_count=0, mechanical_count=0),
        ParkingSizeInput(350, 600, flat_count=2, mechanical_count=1, note="無障礙"),
    ]


def test_area_mode_multiplies_mechanical_ratio() -> None:
    columns = expand_parking_sizes(
        _sizes_with_both_kinds(), PARKING_RATIO_AREA, mechanical_ratio=0.6
    )
    assert [column.kind for column in columns] == [
        "flat",
        "flat",
        "flat",
        "mechanical",
        "mechanical",
    ]
    assert [column.count for column in columns] == [10, 5, 2, 20, 1]
    assert columns[0].share_ratio == 1
    assert columns[1].share_ratio == excel_round(13.75 / 12.65, 2)
    assert columns[3].share_ratio == 0.6
    assert columns[4].label == "350*600\n(無障礙、機械)"
    assert columns[2].label == "350*600\n(無障礙、平面)"


def test_area_mode_only_flat_keeps_legacy_labels() -> None:
    columns = expand_parking_sizes(
        [ParkingSizeInput(350, 600, flat_count=2, note="無障礙")],
        PARKING_RATIO_AREA,
    )
    assert len(columns) == 1
    assert columns[0].share_ratio is None
    assert columns[0].label == "350*600\n(無障礙)"


def test_area_mode_only_mechanical_drops_type_ratio() -> None:
    columns = expand_parking_sizes(
        [
            ParkingSizeInput(230, 550, mechanical_count=10),
            ParkingSizeInput(250, 550, mechanical_count=4),
        ],
        PARKING_RATIO_AREA,
        mechanical_ratio=0.6,
    )
    assert [column.share_ratio for column in columns] == [
        1,
        excel_round(13.75 / 12.65, 2),
    ]
    assert columns[0].label == "230*550\n(機械)"


def test_fixed_mode_uses_one_mechanical_ratio() -> None:
    columns = expand_parking_sizes(
        _sizes_with_both_kinds(), PARKING_RATIO_FIXED, mechanical_ratio=0.6
    )
    assert [column.share_ratio for column in columns] == [1, 1, 1, 0.6, 0.6]


def test_fixed_mode_only_mechanical_uses_base_ratio() -> None:
    columns = expand_parking_sizes(
        [
            ParkingSizeInput(230, 550, mechanical_count=10),
            ParkingSizeInput(350, 600, mechanical_count=2, note="無障礙"),
        ],
        PARKING_RATIO_FIXED,
        mechanical_ratio=0.6,
    )
    assert [column.share_ratio for column in columns] == [1, 1]
    assert columns[1].label == "350*600\n(無障礙、機械)"


def test_share_adjustment_uses_explicit_ratio() -> None:
    units, public_areas, ratio, _spaces = _sample_balance_inputs()
    spaces = [
        ParkingSpace(230, 550, 10, share_ratio=1.0),
        ParkingSpace(230, 550, 10, share_ratio=0.6),
    ]
    adjustment = compute_share_adjustments(units, public_areas, ratio, spaces)
    assert adjustment is not None
    assert adjustment.base_shares[0] != adjustment.base_shares[1]


def test_area_mode_writes_mechanical_ratio_formula(tmp_path: Path) -> None:
    output = tmp_path / "mixed.xlsx"
    columns = expand_parking_sizes(
        _sizes_with_both_kinds(), PARKING_RATIO_AREA, mechanical_ratio=0.6
    )
    create_template(output, units=49, parking_spaces=columns)
    wb = openpyxl.load_workbook(output, data_only=False)
    ws = wb.active
    assert ws["H74"].value == 1
    assert ws["I74"].value == "=round(I72/$H$72,2)"
    assert ws["K74"].value == "=round(K72/$H$72*0.60,2)"
    assert ws["H71"].value == "230*550\n(平面)"
    assert ws["L59"].value == "350*600\n(無障礙、機械)"
    assert ws["M59"].value == "合計"
    assert ws["H75"].value == "=H73*H74+I73*I74+J73*J74+K73*K74+L73*L74"
    assert _merged(ws, 8, 75, 12, 75)
    assert _merged(ws, 7, 58, 12, 58)
    assert _merged(ws, 7, 70, 12, 70)
    assert ws["G70"].value == "大小車位面積比例"
    wb.close()


def test_fixed_mode_writes_ratio_numbers(tmp_path: Path) -> None:
    output = tmp_path / "fixed.xlsx"
    columns = expand_parking_sizes(
        [
            ParkingSizeInput(230, 550, flat_count=4, mechanical_count=3),
        ],
        PARKING_RATIO_FIXED,
        mechanical_ratio=0.6,
    )
    create_template(output, units=49, parking_spaces=columns)
    wb = openpyxl.load_workbook(output, data_only=False)
    ws = wb.active
    assert ws["H74"].value == 1
    assert ws["I74"].value == 0.6
    assert ws["I74"].number_format == "0.00"
    assert ws["H59"].value == "230*550\n(平面)"
    assert ws["I59"].value == "230*550\n(機械)"
    wb.close()


def test_single_parking_size_merges_proportion_title_with_labels(
    tmp_path: Path,
) -> None:
    """僅一欄尺寸時，大小車位面積比例仍含左側表頭欄合併。"""
    output = tmp_path / "one-size.xlsx"
    columns = expand_parking_sizes(
        [ParkingSizeInput(230, 550, flat_count=10, mechanical_count=0)]
    )
    create_template(output, units=10, parking_spaces=columns)
    wb = openpyxl.load_workbook(output, data_only=False)
    ws = wb.active
    config = LayoutConfig(units=10)
    title_row = config.map_row(58)
    proportion_row = config.map_row(70)
    assert ws.cell(proportion_row, 7).value == "大小車位面積比例"
    assert _merged(ws, 7, proportion_row, 8, proportion_row)
    assert _merged(ws, 7, title_row, 8, title_row)
    wb.close()


def test_floor_counts_drop_all_zero_size_columns() -> None:
    flat_only = [
        ParkingSizeInput(230, 550, flat_counts=(3, 0), mechanical_counts=(0, 0)),
        ParkingSizeInput(250, 550, flat_counts=(0, 0), mechanical_counts=(0, 0)),
        ParkingSizeInput(250, 600, flat_counts=(0, 0), mechanical_counts=(0, 0)),
        ParkingSizeInput(350, 600, flat_counts=(0, 0), mechanical_counts=(0, 0)),
    ]
    columns = expand_parking_sizes(flat_only, PARKING_RATIO_AREA)
    assert [column.count for column in columns] == [3]
    assert columns[0].label == "230*550"
    assert columns[0].floor_counts == (3, 0)
    assert all(column.kind == "flat" for column in columns)

    both = [
        ParkingSizeInput(230, 550, flat_counts=(1, 0), mechanical_counts=(2, 0)),
        ParkingSizeInput(250, 550, flat_counts=(0, 0), mechanical_counts=(0, 0)),
        ParkingSizeInput(250, 600, flat_counts=(0, 0), mechanical_counts=(0, 0)),
        ParkingSizeInput(350, 600, flat_counts=(0, 0), mechanical_counts=(0, 1)),
    ]
    columns = expand_parking_sizes(both, PARKING_RATIO_FIXED, mechanical_ratio=0.6)
    assert [column.kind for column in columns] == ["flat", "mechanical", "mechanical"]
    assert [column.count for column in columns] == [1, 2, 1]
    assert columns[0].label == "230*550\n(平面)"
    assert columns[1].label == "230*550\n(機械)"
    assert columns[2].label == "350*600\n(機械)"


def test_fewer_parking_floors_move_total_row_up(tmp_path: Path) -> None:
    output = tmp_path / "two-floors.xlsx"
    columns = expand_parking_sizes(
        [
            ParkingSizeInput(230, 550, flat_counts=(3, 0), mechanical_counts=(0, 0)),
            ParkingSizeInput(250, 550, flat_counts=(0, 4), mechanical_counts=(0, 0)),
            ParkingSizeInput(250, 600, flat_counts=(0, 0), mechanical_counts=(0, 0)),
            ParkingSizeInput(350, 600, flat_counts=(0, 0), mechanical_counts=(0, 0)),
        ]
    )
    create_template(
        output,
        units=49,
        parking_spaces=columns,
        parking_floors=["B1", "地下二"],
    )
    wb = openpyxl.load_workbook(output, data_only=False)
    ws = wb.active
    assert ws["G60"].value == "B1"
    assert ws["G61"].value == "地下二"
    assert ws["G62"].value == "合計"
    assert ws["H60"].value == 3
    assert ws["H61"].value == 0
    assert ws["I60"].value == 0
    assert ws["I61"].value == 4
    assert ws["H62"].value == "=sum(H60:H61)"
    assert ws["J59"].value == "合計"
    assert ws["J60"].value == "=SUM(H60:I60)"
    assert ws["K59"].value is None
    assert ws["L59"].value is None
    assert ws["H69"].value == "=H62"
    assert ws["G63"].value == "面積 ㎡"
    assert ws["H59"].value == "230*550"
    assert ws["I59"].value == "250*550"
    assert _merged(ws, 7, 58, 9, 58)
    wb.close()


def test_extra_parking_floors_extend_the_sum(tmp_path: Path) -> None:
    output = tmp_path / "eight-floors.xlsx"
    floors = [f"B{index}" for index in range(1, 9)]
    counts = tuple(1 if index == 0 else 0 for index in range(8))
    columns = expand_parking_sizes(
        [
            ParkingSizeInput(
                230, 550, flat_counts=counts, mechanical_counts=tuple(0 for _ in floors)
            ),
            ParkingSizeInput(
                250, 550, flat_counts=tuple(0 for _ in floors), mechanical_counts=counts
            ),
            ParkingSizeInput(
                250,
                600,
                flat_counts=tuple(0 for _ in floors),
                mechanical_counts=tuple(0 for _ in floors),
            ),
            ParkingSizeInput(
                350,
                600,
                flat_counts=tuple(0 for _ in floors),
                mechanical_counts=tuple(0 for _ in floors),
            ),
        ],
        PARKING_RATIO_AREA,
        mechanical_ratio=0.6,
    )
    create_template(output, units=49, parking_spaces=columns, parking_floors=floors)
    wb = openpyxl.load_workbook(output, data_only=False)
    ws = wb.active
    assert ws["G67"].value == "B8"
    assert ws["G68"].value == "合計"
    assert ws["H60"].value == 1
    assert ws["H67"].value == 0
    assert ws["I60"].value == 1
    assert ws["I67"].value == 0
    assert ws["H68"].value == "=sum(H60:H67)"
    assert ws["J59"].value == "合計"
    assert ws["J60"].value == "=SUM(H60:I60)"
    assert ws["K59"].value is None
    assert ws["L59"].value is None
    assert ws["H75"].value == "=H68"
    assert ws["H59"].value == "230*550\n(平面)"
    assert ws["I59"].value == "250*550\n(機械)"
    assert _merged(ws, 7, 58, 9, 58)
    wb.close()


def test_fewer_parking_floors_still_fit_public_areas(tmp_path: Path) -> None:
    output = tmp_path / "public-floors.xlsx"
    public_areas = [
        PublicAreaData(
            floor_label=f"{index}F",
            area=float(index),
            is_per_floor=False,
            side="left",
        )
        for index in range(1, 11)
    ]
    create_template(
        output,
        units=49,
        public_areas=public_areas,
        parking_spaces=[],
        parking_floors=["B1", "B2", "B3"],
    )
    wb = openpyxl.load_workbook(output, data_only=False)
    ws = wb.active
    assert [ws.cell(row, 2).value for row in range(60, 70)] == [
        f"{index}F" for index in range(1, 11)
    ]
    assert ws["G62"].value == "B3"
    assert ws["G63"].value == "合計"
    wb.close()


def test_blank_parking_counts_clear_sample_sizes(tmp_path: Path) -> None:
    output = tmp_path / "blank.xlsx"
    create_template(output, units=49, parking_spaces=[])
    wb = openpyxl.load_workbook(output, data_only=False)
    ws = wb.active
    assert ws["H59"].value is None
    assert ws["K71"].value is None
    # 標題跨樓別＋資料欄，避免文字凸出
    assert _merged(ws, 7, 58, 8, 58)
    assert ws["G58"].value == "停車位面積、類別、數量"
    # 無車位時不保留大小車位面積比例表
    assert ws["H70"].value is None
    assert ws["G71"].value is None
    assert ws["G76"].value is None
    assert ws["H75"].value is None
    assert ws["H76"].value is None
    assert ws["H78"].value is None
    wb.close()


def test_blank_parking_with_porch_and_terrace_clears_all_size_columns(
    tmp_path: Path,
) -> None:
    """門廊+露台時尺寸欄不受露台切開；全空車位清掉範本預設，露台在公設%左。"""
    output = tmp_path / "blank-porch-terrace.xlsx"
    create_template(
        output,
        unit_data=[
            UnitData(
                label="1F-A",
                main_area=10,
                balcony_area=2,
                porch_area=1.5,
                terrace_area=3,
                unit_count=1,
            ),
        ],
        parking_spaces=[],
        parking_floors=["B1", "B2"],
        include_porch=True,
        include_terrace=True,
    )
    wb = openpyxl.load_workbook(output, data_only=False)
    ws = wb.active
    config = LayoutConfig(
        units=1,
        include_porch=True,
        include_terrace=True,
        parking_floor_extra_rows=-4,
    )
    header_row = config.map_row(59)
    label_row = config.map_row(71)
    area_row = config.map_row(72)
    start = config.col(8)
    assert [start + offset for offset in range(4)] == [
        config.col(col) for col in (8, 9, 10, 11)
    ]
    for offset in range(4):
        col = start + offset
        assert ws.cell(header_row, col).value is None
        assert ws.cell(label_row, col).value is None
        assert ws.cell(area_row, col).value is None
    assert config.terrace_col == 25
    assert ws.cell(2, 25).value == "備註：露台"
    assert ws.cell(2, 26).value == "公設%"
    wb.close()


def test_create_without_source_file(tmp_path: Path) -> None:
    output = tmp_path / "out.xlsx"
    create_template(output, units=20)
    assert output.exists()
    wb = openpyxl.load_workbook(output, data_only=False)
    ws = wb.active
    assert ws["A25"].value == "=sum(A5:A24)"
    wb.close()
