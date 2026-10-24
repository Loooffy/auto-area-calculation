from __future__ import annotations

import csv
import re
from dataclasses import dataclass
from io import StringIO
from pathlib import Path

FLOOR_RANGE_PATTERN = re.compile(r"^(\d+)~(\d+)F")
PUBLIC_FLOOR_LABEL_PATTERN = re.compile(r"^(\d+)~(\d+)F\*(\d+)$")


@dataclass(frozen=True)
class UnitData:
    """各戶匯入資料。"""

    label: str
    main_area: float
    balcony_area: float
    unit_count: int = 1
    porch_area: float | None = None


@dataclass(frozen=True)
class PublicAreaData:
    """公設面積匯入資料。"""

    floor_label: str
    area: float
    is_per_floor: bool
    side: str = "left"

    @property
    def floor_multiplier(self) -> int | None:
        match = PUBLIC_FLOOR_LABEL_PATTERN.match(self.floor_label)
        return int(match.group(3)) if match else None


@dataclass(frozen=True)
class ImportData:
    units: list[UnitData]
    public_areas: list[PublicAreaData]
    include_porch: bool = False


def parse_unit_count(label: str) -> int:
    """從層數及編號推算戶數。

    例：4~9F-A1 → 4~9 樓共 6 層 → 戶數 6；2F-A1 → 戶數 1。
    """
    match = FLOOR_RANGE_PATTERN.match(label)
    if not match:
        return 1
    start_floor, end_floor = int(match.group(1)), int(match.group(2))
    if end_floor < start_floor:
        raise ValueError(f"樓層範圍無效：{label!r}")
    return end_floor - start_floor + 1


def normalize_public_floor_label(label: str) -> str:
    """正規化公設樓別文字（去除空白、統一 * 格式）。"""
    text = label.strip()
    text = re.sub(r"\s*\*\s*", "*", text)
    return text


def is_per_floor_public_label(label: str) -> bool:
    """是否為多層樓格式（如 4~9F*6）。"""
    return PUBLIC_FLOOR_LABEL_PATTERN.match(label) is not None


def classify_public_side(label: str) -> str:
    """依樓別文字判斷公設區左右欄（屋突、地下室為右欄）。"""
    if label.startswith("屋突") or re.match(r"^B\d+$", label):
        return "right"
    return "left"


def _normalize_area_header(header: str) -> str:
    return header.replace("(㎡)", "").replace("㎡", "").strip()


def _is_unit_header(col1: str, col2: str, col3: str) -> bool:
    if _normalize_area_header(col1) != "主建物" or _normalize_area_header(col2) != "陽台":
        return False
    extra = _normalize_area_header(col3)
    return extra in ("", "門廊")


def _is_public_header(cells: list[str]) -> bool:
    if not cells or cells[0] != "公設":
        return False
    if len(cells) < 3:
        return False
    return cells[1] == "樓別" and _normalize_area_header(cells[2]) == "面積"


def _parse_public_pair(label: str, area_text: str, side: str) -> PublicAreaData:
    floor_label = normalize_public_floor_label(label)
    try:
        area = float(area_text)
    except ValueError as exc:
        raise ValueError(f"無法解析公設面積：{floor_label!r}, {area_text!r}") from exc
    return PublicAreaData(
        floor_label=floor_label,
        area=area,
        is_per_floor=is_per_floor_public_label(floor_label),
        side=side,
    )


def _decode_csv_text(path: Path) -> str:
    raw = path.read_bytes()
    if raw.startswith(b"\xff\xfe") or raw.startswith(b"\xfe\xff"):
        return raw.decode("utf-16")
    return raw.decode("utf-8-sig")


def _detect_delimiter(text: str) -> str:
    sample = "\n".join(text.splitlines()[:10])
    if sample.count(";") > sample.count(","):
        return ";"
    return ","


def _read_csv_rows(path: Path) -> list[list[str]]:
    text = _decode_csv_text(path)
    delimiter = _detect_delimiter(text)
    return list(csv.reader(StringIO(text), delimiter=delimiter))



def parse_import_csv(path: Path) -> ImportData:
    """讀取試算表 CSV（各戶面積 + 公設面積）。

    各戶區段：層數及編號, 主建物(㎡), 陽台(㎡) [, 門廊(㎡)]
    表頭含「門廊」時這個案子才有門廊欄；空白視為 0。
    公設區段：, 樓別, 面積(㎡) [, 樓別, 面積(㎡)]
    左欄為地上樓層，右欄為屋突／地下室；多層如 4~9F*6 的數值為單層面積。
    """
    units: list[UnitData] = []
    public_areas: list[PublicAreaData] = []
    section: str | None = None
    include_porch = False
    public_has_right_column = False

    for row in _read_csv_rows(path):
        if not row or all(not cell.strip() for cell in row):
            continue

        cells = [cell.strip() for cell in row]
        while len(cells) < 5:
            cells.append("")

        col0, col1, col2, col3, col4 = cells[:5]

        if _is_unit_header(col1, col2, col3):
            section = "units"
            include_porch = _normalize_area_header(col3) == "門廊"
            continue
        if _is_public_header(cells):
            section = "public"
            public_has_right_column = (
                len(cells) >= 5
                and cells[3] == "樓別"
                and _normalize_area_header(cells[4]) == "面積"
            )
            continue

        if section == "units":
            if not col0:
                continue
            try:
                main_area = float(col1)
                balcony_area = float(col2)
                porch_area: float | None = None
                if include_porch:
                    porch_area = float(col3) if col3 else 0.0
            except ValueError as exc:
                values = (col0, col1, col2, col3) if include_porch else (col0, col1, col2)
                raise ValueError(f"無法解析各戶數值：{values!r}") from exc
            units.append(
                UnitData(
                    label=col0,
                    main_area=main_area,
                    balcony_area=balcony_area,
                    unit_count=parse_unit_count(col0),
                    porch_area=porch_area,
                )
            )
        elif section == "public":
            if col1 and col2:
                side = "left" if public_has_right_column else classify_public_side(
                    normalize_public_floor_label(col1)
                )
                public_areas.append(_parse_public_pair(col1, col2, side))
            if public_has_right_column and col3 and col4:
                public_areas.append(_parse_public_pair(col3, col4, "right"))

    if not units:
        raise ValueError(f"CSV 無有效各戶資料：{path}")

    return ImportData(
        units=units, public_areas=public_areas, include_porch=include_porch
    )


def parse_units_csv(path: Path) -> list[UnitData]:
    """讀取各戶面積（相容舊介面）。"""
    return parse_import_csv(path).units
