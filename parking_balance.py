from __future__ import annotations

"""車位持分 ±1 微調計算。

試算表中（原始列號）：

- 車位總持分概算（框3）：D74 = 100000 - D73，其中
  D73 = round(C75 / D58 * 100000, 0)，
  C75 = round(E54 * C73 / C74, 2)，C74 = 100 - C73（公設比），
  E54 = 各戶（主+附）面積加權合計，D58 = 公設面積合計。
- 各尺寸車位持分（框1）：第77列 = round($H$76 * 尺寸比例, 0)，其中
  H76（基準車位持分）= round(D74 / H75, 0)，
  H75（小車位數倍數）= Σ(車位數 × 尺寸比例)。
- 車位總持分（框2）：H78 = Σ(車位數 × 各尺寸持分)。

openpyxl 不會計算公式，因此這裡以 Python 重現整條計算鏈，求出讓框2
剛好等於框3的微調組合（每種尺寸最多加減 MAX_DELTA）。若找不到能精準命中
框3的組合，則完全不調整，交由人工處理。
"""

from dataclasses import dataclass
from decimal import Decimal, ROUND_HALF_UP
from itertools import product
from typing import TYPE_CHECKING, Sequence

from csv_import import PublicAreaData, UnitData

if TYPE_CHECKING:  # 避免與 template_builder 形成循環匯入
    from template_builder import ParkingSpace

# 每種車位持分允許的最大微調幅度（±MAX_DELTA）
MAX_DELTA = 10

# 微調後車位總持分（框2）與概算（框3）可接受的最大差距
MAX_TOLERANCE = 20


def excel_round(value: float, digits: int = 0) -> float:
    """以 Excel ROUND 的「四捨五入、逢五進位」規則取整（非銀行家捨入）。"""
    quantum = Decimal(1).scaleb(-digits)
    return float(Decimal(repr(value)).quantize(quantum, rounding=ROUND_HALF_UP))


def total_unit_area(units: Sequence[UnitData]) -> float:
    """各戶（主+附）面積加權合計，對應試算表 E54。"""
    return sum(u.unit_count * (u.main_area + u.balcony_area) for u in units)


def total_public_area(public_areas: Sequence[PublicAreaData]) -> float:
    """公設面積合計，對應試算表 D58。

    左欄多層（如 4~9F*6）以單層面積 × 層數計入；其餘以面積本身計入。
    """
    total = 0.0
    for item in public_areas:
        multiplier = item.floor_multiplier
        if item.side != "right" and item.is_per_floor and multiplier:
            total += item.area * multiplier
        else:
            total += item.area
    return total


@dataclass(frozen=True)
class ShareAdjustment:
    """車位持分微調結果。"""

    target_total: int  # 框3：車位總持分概算（D74）
    base_total: int  # 框2：未微調前的車位總持分
    adjusted_total: int  # 框2：微調後的車位總持分
    base_shares: tuple[int, ...]  # 各尺寸未微調持分（依使用順序）
    deltas: tuple[int, ...]  # 各尺寸微調量，介於 -MAX_DELTA ~ +MAX_DELTA

    @property
    def has_adjustment(self) -> bool:
        return any(self.deltas)


def compute_share_adjustments(
    units: Sequence[UnitData] | None,
    public_areas: Sequence[PublicAreaData] | None,
    public_ratio: float | None,
    used_spaces: Sequence["ParkingSpace"],
    max_delta: int = MAX_DELTA,
    tolerance: int = MAX_TOLERANCE,
) -> ShareAdjustment | None:
    """計算讓框2盡量接近框3的每尺寸微調（每種最多 ±max_delta）。

    只接受微調後框2與框3差距在 tolerance 以內的組合；若無任何組合符合，
    回傳的 deltas 全為 0（不調整，交由人工處理）。
    資料不足或無法計算（如除以零）時回傳 None。
    """
    if not used_spaces or public_ratio is None or not units or not public_areas:
        return None

    e54 = total_unit_area(units)
    d58 = total_public_area(public_areas)
    c74 = 100 - public_ratio
    if c74 == 0 or d58 == 0:
        return None

    c75 = excel_round(e54 * public_ratio / c74, 2)
    d73 = excel_round(c75 / d58 * 100000, 0)
    target = int(round(100000 - d73))

    area0 = used_spaces[0].actual_area
    if area0 == 0:
        return None

    ratios: list[float] = []
    for idx, space in enumerate(used_spaces):
        if idx == 0:
            ratios.append(1.0)
        else:
            ratios.append(excel_round(space.actual_area / area0, 2))

    counts = [space.count for space in used_spaces]
    h75 = sum(count * ratio for count, ratio in zip(counts, ratios))
    if h75 == 0:
        return None

    h76 = excel_round(target / h75, 0)
    base_shares = [int(excel_round(h76 * ratio, 0)) for ratio in ratios]
    base_total = sum(count * share for count, share in zip(counts, base_shares))

    # 只接受微調後框2與框3差距在 tolerance 以內的組合；都湊不到就不調整（deltas 全 0）。
    # 排序：先選與目標差距最小，再選調整種類最少，最後選總調整幅度最小者。
    zero_deltas = tuple(0 for _ in used_spaces)
    best_key: tuple[int, int, int] | None = None
    best_deltas: tuple[int, ...] = zero_deltas
    best_total = base_total
    delta_range = range(-max_delta, max_delta + 1)
    for deltas in product(delta_range, repeat=len(used_spaces)):
        total = base_total + sum(count * d for count, d in zip(counts, deltas))
        diff = abs(target - total)
        if diff > tolerance:
            continue
        key = (diff, sum(1 for d in deltas if d), sum(abs(d) for d in deltas))
        if best_key is None or key < best_key:
            best_key = key
            best_deltas = deltas
            best_total = total

    return ShareAdjustment(
        target_total=target,
        base_total=base_total,
        adjusted_total=best_total,
        base_shares=tuple(base_shares),
        deltas=best_deltas,
    )
