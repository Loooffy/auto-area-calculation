from __future__ import annotations

import re
from dataclasses import dataclass

from openpyxl.utils import column_index_from_string, get_column_letter

from template_schema import MAX_COLUMN

# 有門廊時插入的來源欄：E（主+附之前）與 M（加權主+附之前）。
# 大於等於該欄的來源欄各右移一格，兩次插入累加。
_PORCH_INSERT_AT = (5, 13)

CELL_REF_PATTERN = re.compile(
    r"(?<![A-Z0-9_])"
    r"(\$?)([A-Z]{1,3})(\$?)(\d+)"
    r"(?![A-Z0-9_])"
)


@dataclass(frozen=True)
class LayoutConfig:
    """來源試算表與新試算表的列配置。"""

    source_data_start: int = 5
    source_data_end: int = 53
    source_total_row: int = 54
    source_parking_start: int = 55
    source_parking_end: int = 81
    source_public_area_start_row: int = 60
    source_public_area_end_row: int = 69
    public_area_extra_rows: int = 0
    units: int = 10
    include_porch: bool = False

    @property
    def data_start(self) -> int:
        return self.source_data_start

    @property
    def data_end(self) -> int:
        return self.data_start + self.units - 1

    @property
    def total_row(self) -> int:
        return self.data_end + 1

    @property
    def parking_start(self) -> int:
        return self.total_row + 1

    @property
    def parking_end(self) -> int:
        span = self.source_parking_end - self.source_parking_start
        return self.parking_start + span + self.public_area_extra_rows

    @property
    def rows_removed(self) -> int:
        return self.source_data_end - self.data_end

    def public_data_row(self, offset: int) -> int:
        """公設資料第 offset 列（0 起算）的實際列號（連續排列）。"""
        return self.map_row(self.source_public_area_start_row) + offset

    def col(self, source_col: int) -> int:
        """來源欄號；有門廊時平移到插入欄位之後的位置。"""
        return map_source_column(source_col, self.include_porch)

    @property
    def max_column(self) -> int:
        return self.col(MAX_COLUMN)

    def map_row(self, old_row: int) -> int:
        if old_row == self.source_total_row:
            return self.total_row
        if self.source_parking_start <= old_row <= self.source_parking_end:
            base = old_row - self.rows_removed
            if old_row > self.source_public_area_end_row:
                return base + self.public_area_extra_rows
            return base
        if self.source_data_start <= old_row <= self.source_data_end:
            return min(old_row, self.data_end)
        return old_row


def map_source_column(col: int, include_porch: bool) -> int:
    """把無門廊版面的欄號對應到實際欄號。"""
    if not include_porch:
        return col
    return col + sum(1 for insert_at in _PORCH_INSERT_AT if col >= insert_at)


def shift_formula_columns(formula: str, include_porch: bool) -> str:
    """有門廊時，將公式裡的欄位字母依插入位置右移。列號不變。"""
    if not include_porch or not formula.startswith("="):
        return formula

    def replace_ref(match: re.Match[str]) -> str:
        col_abs, col, row_abs, row_str = match.groups()
        new_col = get_column_letter(
            map_source_column(column_index_from_string(col), True)
        )
        return f"{col_abs}{new_col}{row_abs}{row_str}"

    return CELL_REF_PATTERN.sub(replace_ref, formula)


def rewrite_formula(formula: str, config: LayoutConfig) -> str:
    """將公式中的列參照依新版面重寫。"""

    def replace_ref(match: re.Match[str]) -> str:
        col_abs, col, row_abs, row_str = match.groups()
        old_row = int(row_str)
        new_row = config.map_row(old_row)
        return f"{col_abs}{col}{row_abs}{new_row}"

    if not formula.startswith("="):
        return formula
    return CELL_REF_PATTERN.sub(replace_ref, formula)


def data_row_formulas(
    row: int, total_row: int, include_porch: bool = False
) -> dict[str, str]:
    """各戶資料列公式模板。有門廊時欄位改為主+附+門廊。"""
    formulas = {
        "E": f"=C{row}+D{row}",
        "F": f"=round($F$3*R{row}/$I$3,2)",
        "G": f"=E{row}+F{row}",
        "H": f"=round(G{row}*0.3025,2)",
        "I": f"=R{row}",
        "J": f"=U{row}",
        "K": f"=A{row}*C{row}",
        "L": f"=A{row}*D{row}",
        "M": f"=A{row}*E{row}",
        "N": f"=A{row}*F{row}",
        "O": f"=A{row}*G{row}",
        "P": f"=A{row}*H{row}",
        "Q": f"=round($Q$4*E{row}/$E${total_row},2)",
        "R": f"=round(Q{row},0)",
        "S": f"=A{row}*R{row}",
        "T": f"=round(E{row}/$E${total_row}*$T$4,2)",
        "U": f"=round(T{row},0)",
        "V": f"=A{row}*U{row}",
        "W": f"=round(F{row}/G{row}*100,2)",
    }
    if not include_porch:
        return formulas

    shifted = {
        get_column_letter(map_source_column(column_index_from_string(col), True)): (
            shift_formula_columns(formula, True)
        )
        for col, formula in formulas.items()
    }
    # E 是新的門廊輸入欄，合計改掛在平移後的 F，不能再把公式裡的 E 一起右移。
    shifted["F"] = f"=C{row}+D{row}+E{row}"
    shifted["N"] = f"=A{row}*E{row}"
    return shifted
