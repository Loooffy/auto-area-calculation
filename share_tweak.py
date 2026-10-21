from __future__ import annotations

"""各戶「公設計算」與「土地持分」整數持分的 ±1 微調（DP 求最佳解）。

試算表中（原始列號）：

- 公設計算：Q=round(Q4·E/E54, 2)（取，decimal），R=round(Q, 0)（總，整數持分），
  S=戶數·R。框1 公設目標總持分 = Q4 = I3 - H78 = 100000 - 車位總持分；
  框2 公設持分實際加總 = S54 = Σ(戶數·R)。
- 土地持分：T=round(E/E54·T4, 2)（取），U=round(T, 0)（總），V=戶數·U。
  框3 土地目標總持分 = T4 = 100000；框4 土地持分實際加總 = V54 = Σ(戶數·U)。

調整規則：每個整數持分只能 ±1，且「持分(decimal)相同」的列必須一起調整（同面積
的戶別持分需一致）。公設與土地的實際加總必須與目標「完全吻合」（誤差 0），因此
只在能精準命中目標時才套用；若 DP 找不到可精準命中的組合，則完全不調整（shifts
為空、completed=False），交由人工處理。因列數可能很多，改用 DP（背包式）求解。

DP 設計參考 tweak_core.py：以「相同 decimal」分組，整組一起 +1/-1，調整一組的
代價為該組列數；狀態為累計的加總增減量，於可精準命中目標的解中取代價最小者。
"""

from dataclasses import dataclass
from typing import Sequence

from parking_balance import excel_round


@dataclass(frozen=True)
class ShareRow:
    """單一戶別列的微調輸入。"""

    offset: int  # 相對於資料起始列的位移（0 起算）
    weight: int  # 戶數（調整 1 點持分對加總的影響量）
    decimal: float  # 取（未取整持分），分組依據
    rounded: int  # 總（整數持分）


@dataclass(frozen=True)
class ShareTweakResult:
    """公設／土地持分微調結果。"""

    target: int  # 目標總持分（框1 或框3）
    base_total: int  # 微調前實際加總（框2 或框4）
    adjusted_total: int  # 微調後實際加總
    shifts: dict[int, int]  # offset -> 增減量（僅含 ±1，0 不列入）

    @property
    def has_adjustment(self) -> bool:
        return bool(self.shifts)

    @property
    def completed(self) -> bool:
        """微調後是否與目標完全吻合（誤差 0）。"""
        return self.adjusted_total == self.target


def build_share_rows(
    weights: Sequence[int],
    areas: Sequence[float],
    target: int,
    total_area: float,
) -> list[ShareRow]:
    """依各戶面積與目標總持分，算出每列的取(decimal)與總(rounded)。

    decimal = round(target · 面積 / 總面積, 2)，rounded = round(decimal, 0)，
    對應試算表的 Q/R（公設）或 T/U（土地）。
    """
    rows: list[ShareRow] = []
    for offset, (weight, area) in enumerate(zip(weights, areas)):
        decimal = excel_round(target * area / total_area, 2)
        rounded = int(excel_round(decimal, 0))
        rows.append(
            ShareRow(offset=offset, weight=weight, decimal=decimal, rounded=rounded)
        )
    return rows


def solve_share_adjustments(rows: Sequence[ShareRow], target: int) -> ShareTweakResult:
    """以 DP 求每組 ±1 微調，使實際加總「精準命中」target。

    僅在能精準命中時才回傳對應 shifts；否則 shifts 為空（不調整）。
    """
    base_total = sum(row.weight * row.rounded for row in rows)

    # 依 decimal 分組（round 到 9 位避免浮點誤差把相同值拆開）。
    grouped: dict[float, list[ShareRow]] = {}
    for row in rows:
        grouped.setdefault(round(row.decimal, 9), []).append(row)

    groups: list[tuple[list[ShareRow], int, int]] = []
    for members in grouped.values():
        weight_sum = sum(member.weight for member in members)
        rounded_value = members[0].rounded
        groups.append((members, weight_sum, rounded_value))

    required = target - base_total

    # DP 狀態：累計加總增減量 -> (代價=已調整列數, 各組增減量 (group_index, shift) 序列)
    states: dict[int, tuple[int, tuple[tuple[int, int], ...]]] = {0: (0, ())}
    for index, (members, weight_sum, rounded_value) in enumerate(groups):
        row_count = len(members)
        options = [(0, 0, 0), (1, weight_sum, row_count)]
        if rounded_value >= 1:  # 不允許把持分調成負數
            options.append((-1, -weight_sum, row_count))

        next_states: dict[int, tuple[int, tuple[tuple[int, int], ...]]] = {}
        for sum_delta, (cost, assign) in states.items():
            for shift, delta_sum, delta_cost in options:
                new_sum = sum_delta + delta_sum
                new_cost = cost + delta_cost
                new_assign = assign if shift == 0 else assign + ((index, shift),)
                prev = next_states.get(new_sum)
                if prev is None or new_cost < prev[0]:
                    next_states[new_sum] = (new_cost, new_assign)
        states = next_states

    # 只接受能精準命中目標（sum_delta == required）的解；找不到就不調整。
    shifts: dict[int, int] = {}
    if required in states:
        _, best_assign = states[required]
        for index, shift in best_assign:
            for member in groups[index][0]:
                shifts[member.offset] = shift
        adjusted_total = target
    else:
        adjusted_total = base_total

    return ShareTweakResult(
        target=target,
        base_total=base_total,
        adjusted_total=adjusted_total,
        shifts=shifts,
    )
