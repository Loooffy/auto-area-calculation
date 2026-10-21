from __future__ import annotations

from pathlib import Path

import openpyxl
import pytest

from csv_import import (
    PublicAreaData,
    UnitData,
    parse_import_csv,
    parse_unit_count,
    parse_units_csv,
)
from template_builder import (
    compute_public_area_extra_rows,
    create_template,
    public_area_rows_needed,
)

SAMPLE_CSV = Path(__file__).resolve().parent.parent / "試算面積.csv"

FULL_PUBLIC_CSV = """\
,主建物,陽台
1F-A,55.87,0
公設,樓別,面積,樓別,面積
,1F,456.52,屋突1,97.51
,2F,96.47,屋突2,95.78
,3F,104.87,屋突3,95.78
,4~9F*6,104.87,B1,1071.58
,10~14F*5,104.87,B2,1071.58
,15~19F*5,104.87,B3,1071.58
,20F,104.87,B4,1071.58
,,,B5,1071.58
,,,B6,1071.58
"""


def _write_csv(tmp_path: Path, content: str, name: str = "sample.csv") -> Path:
    path = tmp_path / name
    path.write_text(content, encoding="utf-8-sig")
    return path


def test_parse_sample_csv() -> None:
    data = parse_import_csv(SAMPLE_CSV)
    units = data.units
    assert len(units) == 49
    assert units[0].label == "1F-A"
    assert units[0].main_area == 55.87
    assert units[0].balcony_area == 0.0
    assert units[3].label == "1F-D"
    assert units[4].label == "2F-A1"
    assert units[4].main_area == 62.8
    assert units[4].balcony_area == 4.98
    assert units[0].unit_count == 1
    assert units[4].unit_count == 1


def test_parse_public_areas(tmp_path: Path) -> None:
    data = parse_import_csv(_write_csv(tmp_path, FULL_PUBLIC_CSV))
    assert len(data.public_areas) == 16
    by_label = {p.floor_label: p for p in data.public_areas}
    assert by_label["1F"].area == 456.52
    assert by_label["1F"].side == "left"
    assert by_label["1F"].is_per_floor is False
    assert by_label["4~9F*6"].area == 104.87
    assert by_label["4~9F*6"].is_per_floor is True
    assert by_label["4~9F*6"].floor_multiplier == 6
    assert by_label["10~14F*5"].area == 104.87
    assert by_label["20F"].area == 104.87
    assert by_label["屋突1"].area == 97.51
    assert by_label["屋突1"].side == "right"
    assert by_label["屋突3"].area == 95.78
    assert by_label["B1"].area == 1071.58
    assert by_label["B6"].area == 1071.58


def test_parse_public_areas_legacy_three_columns(tmp_path: Path) -> None:
    csv_text = """\
,主建物,陽台
1F-A,10,0
公設,樓別,面積
,1F,100
,屋突1,50
,B1,200
"""
    data = parse_import_csv(_write_csv(tmp_path, csv_text))
    assert len(data.public_areas) == 3
    assert data.public_areas[0].floor_label == "1F"
    assert data.public_areas[0].side == "left"
    assert data.public_areas[1].floor_label == "屋突1"
    assert data.public_areas[1].side == "right"
    assert data.public_areas[2].floor_label == "B1"
    assert data.public_areas[2].side == "right"


def test_parse_unit_count() -> None:
    assert parse_unit_count("1F-A") == 1
    assert parse_unit_count("2F-A1") == 1
    assert parse_unit_count("4~9F-A1") == 6
    assert parse_unit_count("10~14F-A1") == 5
    assert parse_unit_count("15~19F-A5") == 5
    assert parse_unit_count("20F-A9") == 1


def test_sample_csv_floor_ranges() -> None:
    units = parse_units_csv(SAMPLE_CSV)
    by_label = {u.label: u.unit_count for u in units}
    assert by_label["4~9F-A1"] == 6
    assert by_label["10~14F-A1"] == 5
    assert by_label["15~19F-A1"] == 5


def test_skip_header_row(tmp_path: Path) -> None:
    csv_path = tmp_path / "units.csv"
    csv_path.write_text(
        ",主建物,陽台\n1F-A,10,2\n1F-B,20,3\n",
        encoding="utf-8-sig",
    )
    units = parse_units_csv(csv_path)
    assert len(units) == 2
    assert units[1].label == "1F-B"


def test_parse_semicolon_delimited_csv(tmp_path: Path) -> None:
    csv_path = tmp_path / "semicolon.csv"
    csv_path.write_text(
        ";主建物;陽台\n1F-A;55.87;0\n1F-B;60.69;0\n",
        encoding="utf-16",
    )
    units = parse_units_csv(csv_path)
    assert len(units) == 2
    assert units[0].label == "1F-A"
    assert units[0].main_area == 55.87


def test_import_into_template(tmp_path: Path) -> None:
    data = parse_import_csv(SAMPLE_CSV)

    output = tmp_path / "filled.xlsx"
    create_template(
        output,
        unit_data=data.units,
        public_areas=data.public_areas,
    )

    wb = openpyxl.load_workbook(output, data_only=False)
    ws = wb.active

    assert ws["B5"].value == "1F-A"
    assert ws["C5"].value == 55.87
    assert ws["D5"].value == 0.0
    assert ws["A5"].value == 1

    row_4_9 = 5 + next(i for i, u in enumerate(data.units) if u.label == "4~9F-A1")
    assert ws[f"A{row_4_9}"].value == 6
    assert ws[f"B{row_4_9}"].value == "4~9F-A1"

    last_row = 5 + len(data.units) - 1
    assert ws[f"B{last_row}"].value == "20F-A9"

    total_row = last_row + 1
    assert ws[f"A{total_row}"].value == f"=sum(A5:A{last_row})"

    # 公設面積（依 CSV 動態填入）
    assert ws["C60"].value == 456.52
    assert ws["C61"].value == 96.47
    assert ws["C62"].value == 104.87
    assert ws["B63"].value == "4~9F*6"
    assert ws["B64"].value == 104.87
    assert ws["C63"].value == "=B64*6"
    assert ws["B65"].value == "10~14F*5"
    assert ws["B66"].value == 104.87
    assert ws["C65"].value == "=B66*5"
    assert ws["B67"].value == "15F"
    assert ws["C67"].value == 104.87
    assert ws["B68"].value == "16~20F*5"
    assert ws["B69"].value == 104.87
    assert ws["C68"].value == "=B69*5"
    assert ws["B70"].value == "21~26F*6"
    assert ws["B71"].value == 104.87
    assert ws["C70"].value == "=B71*6"
    assert ws["B72"].value == "27F"
    assert ws["C72"].value == 97.51
    assert ws["B73"].value == "28~30F*3"
    assert ws["B74"].value == 37.02
    assert ws["C73"].value == "=B74*3"

    # 右欄
    assert ws["D60"].value == "屋突1"
    assert ws["E60"].value == 97.51
    assert ws["D61"].value == "屋突2"
    assert ws["D62"].value == "B1"
    assert ws["D68"].value == "B7"

    for row in range(69, 75):
        assert ws[f"D{row}"].value is None
        assert ws[f"E{row}"].value is None

    # 公設比區塊下移至原列 73 + 8（公設多列展開）
    assert "公" in str(ws["B81"].value)
    assert ws["D58"].value == "=sum(C60:C77)+sum(E60:E68)"

    wb.close()


def test_import_full_public_template(tmp_path: Path) -> None:
    data = parse_import_csv(SAMPLE_CSV)
    public_data = parse_import_csv(_write_csv(tmp_path, FULL_PUBLIC_CSV, "public.csv"))

    output = tmp_path / "filled.xlsx"
    create_template(
        output,
        unit_data=data.units,
        public_areas=public_data.public_areas,
    )

    wb = openpyxl.load_workbook(output, data_only=False)
    ws = wb.active

    assert ws["B67"].value == "15~19F*5"
    assert ws["B68"].value == 104.87
    assert ws["C67"].value == "=B68*5"
    assert ws["B69"].value == "20F"
    assert ws["C69"].value == 104.87
    assert ws["E68"].value == 1071.58
    assert ws["D58"].value == "=sum(C60:C69)+sum(E60:E68)"

    wb.close()


def test_empty_csv_raises(tmp_path: Path) -> None:
    csv_path = tmp_path / "empty.csv"
    csv_path.write_text(",主建物,陽台\n", encoding="utf-8-sig")
    with pytest.raises(ValueError, match="無有效各戶資料"):
        parse_units_csv(csv_path)


EXTENDED_PUBLIC_CSV = """\
,主建物,陽台
1F-A,10,0
公設,樓別,面積
,1F,100
,2F,100
,3F,100
,4~9F*6,100
,10~14F*5,100
,15F,100
,16~20F*5,100
,21~26F*6,100
,27F,100
,28~30F*3,100
,屋突1,50
,B1,200
,B2,200
,B3,200
,B4,200
,B5,200
,B6,200
,B7,200
"""


def test_public_area_expansion_shifts_ratio_section(tmp_path: Path) -> None:
    from formula_utils import LayoutConfig

    data = parse_import_csv(_write_csv(tmp_path, EXTENDED_PUBLIC_CSV))
    assert public_area_rows_needed(data.public_areas) == 15
    extra = compute_public_area_extra_rows(data.public_areas)
    assert extra == 5

    output = tmp_path / "expanded.xlsx"
    create_template(
        output,
        unit_data=data.units,
        public_areas=data.public_areas,
    )

    config = LayoutConfig(units=len(data.units), public_area_extra_rows=extra)
    wb = openpyxl.load_workbook(output, data_only=False)
    ws = wb.active

    assert ws[f"B{config.public_data_row(13)}"].value == "28~30F*3"
    assert ws[f"B{config.public_data_row(14)}"].value == 100

    ratio_row = config.map_row(73)
    assert "公" in str(ws[f"B{ratio_row}"].value)
    assert config.public_data_row(14) < ratio_row

    wb.close()
