"""모니터링 대상 판정.

대상 = (제네시스 전 차종) ∪ (BOF 플랫폼 이면서 SUV/픽업 이면서 전장 4,800~5,800mm)
"""
from __future__ import annotations

LENGTH_MIN_MM = 4800
LENGTH_MAX_MM = 5800
BODY_TYPES = {"SUV", "Pickup"}


def in_scope(vehicle: dict, length_mm: float | None) -> tuple[bool, str]:
    maker = (vehicle.get("maker") or "").strip().lower()
    if maker == "genesis":
        return True, "제네시스 전 차종"
    if (vehicle.get("platform") or "").upper() != "BOF":
        return False, "BOF 아님"
    if vehicle.get("body") not in BODY_TYPES:
        return False, "SUV/픽업 아님"
    if length_mm is None:
        return False, "전장 미확인"
    if not (LENGTH_MIN_MM <= length_mm <= LENGTH_MAX_MM):
        return False, f"전장 {length_mm:.0f}mm 범위 밖"
    return True, "BOF SUV/픽업, 전장 범위 내"


def comparison_allowed(vehicle: dict) -> bool:
    """비교 카드에는 범위 밖이라도 '현대차그룹 동급' 슬롯은 허용(라벨로 구분)."""
    return True
