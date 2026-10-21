from __future__ import annotations

from pathlib import Path

import openpyxl
import pytest

from csv_import import PublicAreaData, UnitData
from formula_utils import LayoutConfig, rewrite_formula
from parking_balance import compute_share_adjustments
from share_tweak import ShareRow, solve_share_adjustments
from template_builder import (
    LAND_SHARE_COL,
    PUBLIC_SHARE_COL,
    AdjustmentStatus,
    BuildResult,
    ParkingSpace,
    create_template,
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


def test_parking_note_placeholder(ws) -> None:
    assert "___" in str(ws["B16"].value)


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


def test_create_without_source_file(tmp_path: Path) -> None:
    output = tmp_path / "out.xlsx"
    create_template(output, units=20)
    assert output.exists()
    wb = openpyxl.load_workbook(output, data_only=False)
    ws = wb.active
    assert ws["A25"].value == "=sum(A5:A24)"
    wb.close()
