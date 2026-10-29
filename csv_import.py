from __future__ import annotations

import csv
import re
from dataclasses import dataclass
from io import StringIO
from pathlib import Path

FLOOR_RANGE_PATTERN = re.compile(r"^(\d+)~(\d+)F")
PUBLIC_FLOOR_LABEL_PATTERN = re.compile(r"^(\d+)~(\d+)F\*(\d+)$")

# 各戶：{a}F-{戶} 或 {a}~{b}F-{戶}；戶＝單一字母，或字母+數字
UNIT_LABEL_PATTERN = re.compile(r"^(\d+)(?:~(\d+))?F-([A-Za-z]\d*)$")
PUBLIC_SINGLE_LABEL_PATTERN = re.compile(r"^(\d+)F$")

UNIT_LABEL_HINT = "{a}F-{戶} 或 {a}~{b}F-{戶}（戶為單一字母或字母+數字）"
PUBLIC_LABEL_HINT = "{a}F、{a}~{b}F*{層數}，或屋突／地下室名稱（如 B1、地下一）"


@dataclass(frozen=True)
class UnitData:
    """各戶匯入資料。"""

    label: str
    main_area: float
    balcony_area: float
    unit_count: int = 1
    porch_area: float | None = None
    terrace_area: float | None = None


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
    include_terrace: bool = False


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


def is_basement_public_label(label: str) -> bool:
    """公設樓別是否為地下室。

    屋突，以及含 F 的地上樓層（含寫錯的樓層寫法）都不是地下室。
    其餘文字（B1、b1、地下一）視為地下室，並區分大小寫。
    """
    return not label.startswith("屋突") and "F" not in label


def basement_floor_labels(public_areas: list[PublicAreaData]) -> list[str]:
    """依出現順序取出地下室樓層。屋突與地上樓層不算。"""
    return [
        item.floor_label
        for item in public_areas
        if is_basement_public_label(item.floor_label)
    ]


def classify_public_side(label: str) -> str:
    """依樓別文字判斷公設區左右欄（屋突、地下室為右欄）。"""
    if label.startswith("屋突") or is_basement_public_label(label):
        return "right"
    return "left"


def _unit_label_error(label: str) -> str | None:
    match = UNIT_LABEL_PATTERN.fullmatch(label)
    if not match:
        return f"「{label}」應為 {UNIT_LABEL_HINT}"
    end_text = match.group(2)
    if end_text is not None:
        start_floor = int(match.group(1))
        end_floor = int(end_text)
        if end_floor < start_floor:
            return f"「{label}」樓層範圍無效：結束樓層小於起始樓層"
    return None


def _public_label_error(label: str) -> str | None:
    if label.startswith("屋突") or is_basement_public_label(label):
        return None

    single = PUBLIC_SINGLE_LABEL_PATTERN.fullmatch(label)
    if single:
        return None

    ranged = PUBLIC_FLOOR_LABEL_PATTERN.fullmatch(label)
    if ranged:
        start_floor = int(ranged.group(1))
        end_floor = int(ranged.group(2))
        multiplier = int(ranged.group(3))
        if end_floor < start_floor:
            return f"公設「{label}」樓層範圍無效：結束樓層小於起始樓層"
        expected = end_floor - start_floor + 1
        if multiplier != expected:
            correct = f"{start_floor}~{end_floor}F*{expected}"
            return (
                f"公設「{label}」層數應為 {expected}"
                f"（{start_floor}~{end_floor} 共 {expected} 層），"
                f"正確寫法為 {correct}"
            )
        return None

    return f"公設「{label}」應為 {PUBLIC_LABEL_HINT}"


def _expected_public_label_for_unit(label: str) -> str | None:
    """由已通過格式檢查的各戶標籤，推導對應公設樓別。"""
    match = UNIT_LABEL_PATTERN.fullmatch(label)
    if not match:
        return None
    start_floor = int(match.group(1))
    end_text = match.group(2)
    if end_text is None:
        return f"{start_floor}F"
    end_floor = int(end_text)
    return f"{start_floor}~{end_floor}F*{end_floor - start_floor + 1}"


def validate_import_labels(
    units: list[UnitData], public_areas: list[PublicAreaData]
) -> None:
    """檢查各戶／公設標籤格式，以及有公設時與各戶樓層的對應。

    收集所有問題後一次拋出，方便一次修正。
    """
    errors: list[str] = []

    for unit in units:
        message = _unit_label_error(unit.label)
        if message:
            errors.append(message)

    for item in public_areas:
        message = _public_label_error(item.floor_label)
        if message:
            errors.append(message)

    if public_areas:
        public_labels = {item.floor_label for item in public_areas}
        missing: list[str] = []
        seen: set[str] = set()
        for unit in units:
            if _unit_label_error(unit.label):
                continue
            expected = _expected_public_label_for_unit(unit.label)
            if expected is None or expected in seen:
                continue
            seen.add(expected)
            if expected not in public_labels:
                missing.append(expected)
        for expected in missing:
            errors.append(f"各戶樓層在公設中找不到對應「{expected}」")

    if errors:
        raise ValueError("輸入面積格式錯誤：\n" + "\n".join(f"- {e}" for e in errors))


def _normalize_area_header(header: str) -> str:
    return header.replace("(㎡)", "").replace("㎡", "").strip()


def _normalize_terrace_header(header: str) -> str:
    """露台表頭：露台／備註：露台（全半形冒號皆可）。"""
    text = _normalize_area_header(header).replace("：", ":")
    if text in ("露台", "備註:露台"):
        return "露台"
    return text


UNIT_HEADER_HINT = (
    "欄標題應為：第 2 欄「主建物」、第 3 欄「陽台」，"
    "其後可加「門廊」，以及「露台」或「備註：露台」。"
)
PUBLIC_HEADER_HINT = (
    "公設表頭應為：第 1 欄「公設」、第 2 欄「樓別」、第 3 欄「面積」，"
    "右欄可再接第 4 欄「樓別」、第 5 欄「面積」。"
)
_UNIT_HEADER_KEYWORDS = ("主建", "陽台", "陽臺", "門廊", "露台")


def _quoted_cell(text: str) -> str:
    return f"「{text}」" if text else "「空白」"


def _column_mismatch(index: int, expected: str, actual: str) -> str:
    if actual:
        return f"第 {index} 欄{_quoted_cell(actual)}應為{expected}"
    return f"第 {index} 欄應為{expected}，目前是「空白」"


def _unit_header_problems(col1: str, col2: str, col3: str, col4: str = "") -> list[str]:
    problems: list[str] = []
    if _normalize_area_header(col1) != "主建物":
        problems.append(_column_mismatch(2, "「主建物」", col1))
    if _normalize_area_header(col2) != "陽台":
        problems.append(_column_mismatch(3, "「陽台」", col2))
    extra = _normalize_area_header(col3)
    if extra == "":
        return problems
    if extra == "門廊":
        if _normalize_terrace_header(col4) not in ("", "露台"):
            problems.append(_column_mismatch(5, "「露台」或「備註：露台」", col4))
        return problems
    if _normalize_terrace_header(col3) != "露台":
        problems.append(_column_mismatch(4, "「門廊」或「露台」", col3))
    return problems


def _is_unit_header(col1: str, col2: str, col3: str, col4: str = "") -> bool:
    return not _unit_header_problems(col1, col2, col3, col4)


def _looks_like_unit_header(*cells: str) -> bool:
    return any(keyword in cell for cell in cells for keyword in _UNIT_HEADER_KEYWORDS)


def _public_header_problems(cells: list[str]) -> list[str]:
    problems: list[str] = []
    col0 = cells[0] if cells else ""
    col1 = cells[1] if len(cells) > 1 else ""
    col2 = cells[2] if len(cells) > 2 else ""
    if col0 != "公設":
        problems.append(_column_mismatch(1, "「公設」", col0))
    if col1 != "樓別":
        problems.append(_column_mismatch(2, "「樓別」", col1))
    if _normalize_area_header(col2) != "面積":
        problems.append(_column_mismatch(3, "「面積」", col2))
    col3 = cells[3] if len(cells) > 3 else ""
    col4 = cells[4] if len(cells) > 4 else ""
    if col3 or col4:
        if col3 != "樓別":
            problems.append(_column_mismatch(4, "「樓別」", col3))
        if _normalize_area_header(col4) != "面積":
            problems.append(_column_mismatch(5, "「面積」", col4))
    return problems


def _is_number(text: str) -> bool:
    if not text:
        return False
    try:
        float(text)
    except ValueError:
        return False
    return True


def _looks_like_public_header(cells: list[str]) -> bool:
    if _is_public_header(cells):
        return False
    col0, col1, col2, col3 = cells[0], cells[1], cells[2], cells[3]
    if _is_number(col1) and _is_number(col2):
        return False
    if "公設" in col0 and not _is_number(col1):
        return True
    if col1 in ("樓別", "樓層") or col3 in ("樓別", "樓層"):
        return True
    normalized = _normalize_area_header(col2)
    return "面積" in normalized or "坪" in normalized


def _row_preview(cells: list[str]) -> str:
    shown = [cell for cell in cells if cell]
    return "、".join(shown) if shown else "空白"


def _format_header_error(
    kind: str, row_index: int, cells: list[str], problems: list[str], hint: str
) -> str:
    lines = "\n".join(f"- {item}" for item in problems)
    return f"{kind}（第 {row_index} 列：{_row_preview(cells)}）：\n{lines}\n{hint}"


def units_include_porch(units: list[UnitData]) -> bool:
    """任一戶門廊 > 0 才在試算表開門廊欄。"""
    return any(
        unit.porch_area is not None and unit.porch_area > 0 for unit in units
    )


def units_include_terrace(units: list[UnitData]) -> bool:
    """任一戶露台 > 0 才在試算表開「備註：露台」欄。"""
    return any(
        unit.terrace_area is not None and unit.terrace_area > 0 for unit in units
    )


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

    各戶區段：層數及編號, 主建物(㎡), 陽台(㎡) [, 門廊(㎡)] [, 露台(㎡)]
    表頭含「門廊」時讀取門廊；空白視為 0。任一戶門廊 > 0 時 include_porch 才為 True。
    表頭含「露台」或「備註：露台」時讀取露台；空白視為 0。
    任一戶露台 > 0 時 include_terrace 才為 True（試算表才開備註欄）。
    公設區段：, 樓別, 面積(㎡) [, 樓別, 面積(㎡)]
    左欄為地上樓層，右欄為屋突／地下室；多層如 4~9F*6 的數值為單層面積。
    不含 F、也不是屋突的樓別（B1、b1、地下一）視為地下室，單欄時歸右欄。
    含 F 但格式不對的樓層仍視為錯誤，不當成地下室。
    表頭文字不符時，錯誤會指出列號、欄位、目前文字與應有文字。
    載入後檢查標籤格式；有公設時並檢查各戶樓層是否都有對應公設列。
    """
    units: list[UnitData] = []
    public_areas: list[PublicAreaData] = []
    section: str | None = None
    has_porch_column = False
    has_terrace_column = False
    public_has_right_column = False
    unit_header_found = False
    unit_header_reported = False
    header_errors: list[str] = []
    first_content: tuple[int, list[str]] | None = None

    for row_index, row in enumerate(_read_csv_rows(path), start=1):
        if not row or all(not cell.strip() for cell in row):
            continue

        cells = [cell.strip() for cell in row]
        while len(cells) < 6:
            cells.append("")

        col0, col1, col2, col3, col4 = cells[:5]
        if first_content is None:
            first_content = (row_index, cells)

        if _is_unit_header(col1, col2, col3, col4):
            section = "units"
            unit_header_found = True
            has_porch_column = _normalize_area_header(col3) == "門廊"
            if has_porch_column:
                has_terrace_column = _normalize_terrace_header(col4) == "露台"
            else:
                has_terrace_column = _normalize_terrace_header(col3) == "露台"
            continue
        if (
            not unit_header_found
            and _looks_like_unit_header(col1, col2, col3, col4)
            and not _looks_like_public_header(cells)
        ):
            header_errors.append(
                _format_header_error(
                    "欄標題不符",
                    row_index,
                    cells,
                    _unit_header_problems(col1, col2, col3, col4),
                    UNIT_HEADER_HINT,
                )
            )
            unit_header_reported = True
            continue
        if _is_public_header(cells) or _looks_like_public_header(cells):
            problems = _public_header_problems(cells)
            if problems:
                header_errors.append(
                    _format_header_error(
                        "公設表頭不符",
                        row_index,
                        cells,
                        problems,
                        PUBLIC_HEADER_HINT,
                    )
                )
                continue
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
            if (
                not _is_number(col1)
                and not _is_number(col2)
                and not any(char.isdigit() for char in col0)
            ):
                header_errors.append(
                    _format_header_error(
                        "公設表頭不符",
                        row_index,
                        cells,
                        _public_header_problems(cells),
                        PUBLIC_HEADER_HINT,
                    )
                )
                continue
            try:
                main_area = float(col1)
                balcony_area = float(col2)
                porch_area: float | None = None
                terrace_area: float | None = None
                if has_porch_column:
                    porch_area = float(col3) if col3 else 0.0
                    if has_terrace_column:
                        terrace_area = float(col4) if col4 else 0.0
                elif has_terrace_column:
                    terrace_area = float(col3) if col3 else 0.0
            except ValueError as exc:
                values: tuple[str, ...]
                if has_porch_column and has_terrace_column:
                    values = (col0, col1, col2, col3, col4)
                elif has_porch_column or has_terrace_column:
                    values = (col0, col1, col2, col3)
                else:
                    values = (col0, col1, col2)
                raise ValueError(f"無法解析數值：{values!r}") from exc
            try:
                unit_count = parse_unit_count(col0)
            except ValueError:
                # 格式錯誤改由 validate_import_labels 一次彙整回報
                unit_count = 1
            units.append(
                UnitData(
                    label=col0,
                    main_area=main_area,
                    balcony_area=balcony_area,
                    unit_count=unit_count,
                    porch_area=porch_area,
                    terrace_area=terrace_area,
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

    if not unit_header_found and not unit_header_reported:
        if first_content is None:
            header_errors.insert(0, f"找不到各戶表頭。\n{UNIT_HEADER_HINT}")
        else:
            row_index, cells = first_content
            if _is_public_header(cells) or _looks_like_public_header(cells):
                header_errors.insert(
                    0,
                    f"找不到各戶表頭（第 {row_index} 列是公設表頭）。\n{UNIT_HEADER_HINT}",
                )
            else:
                col1, col2, col3, col4 = cells[1:5]
                header_errors.insert(
                    0,
                    _format_header_error(
                        "欄標題不符",
                        row_index,
                        cells,
                        _unit_header_problems(col1, col2, col3, col4),
                        UNIT_HEADER_HINT,
                    ),
                )

    if header_errors:
        raise ValueError("\n\n".join(header_errors))

    if not units:
        raise ValueError(f"CSV 無有效各戶資料：{path}")

    validate_import_labels(units, public_areas)

    return ImportData(
        units=units,
        public_areas=public_areas,
        include_porch=units_include_porch(units),
        include_terrace=units_include_terrace(units),
    )


def parse_units_csv(path: Path) -> list[UnitData]:
    """讀取各戶面積（相容舊介面）。"""
    return parse_import_csv(path).units
