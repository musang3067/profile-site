"""해외 단위(미국 기준) → SI 단위 환산.

원본 값과 단위는 그대로 보존하고, SI 환산값을 함께 계산한다.
가격(USD)은 통화이므로 환산하지 않는다.
"""
from __future__ import annotations

# unit -> (SI unit, factor)  ; si = value * factor
_LINEAR = {
    "in": ("mm", 25.4),
    "ft": ("mm", 304.8),
    "mm": ("mm", 1.0),
    "cm": ("mm", 10.0),
    "m": ("mm", 1000.0),
    "lb": ("kg", 0.45359237),
    "kg": ("kg", 1.0),
    "hp": ("kW", 0.745699872),
    "ps": ("kW", 0.73549875),
    "kw": ("kW", 1.0),
    "lb-ft": ("N·m", 1.3558179483),
    "nm": ("N·m", 1.0),
    "n·m": ("N·m", 1.0),
    "kgf·m": ("N·m", 9.80665),
    "mph": ("km/h", 1.609344),
    "km/h": ("km/h", 1.0),
    "mi": ("km", 1.609344),
    "km": ("km", 1.0),
    "gal": ("L", 3.785411784),
    "l": ("L", 1.0),
    "cu ft": ("L", 28.316846592),
    "s": ("s", 1.0),
    "deg": ("°", 1.0),
    "kwh": ("kWh", 1.0),
    "usd": ("USD", 1.0),
    "krw": ("KRW", 1.0),
    "cad": ("CAD", 1.0),
    "eur": ("EUR", 1.0),
    "seats": ("seats", 1.0),
}


def to_si(value, unit: str | None):
    """(si_value, si_unit) 반환. 값이 없거나 숫자가 아니면 (None, None)."""
    if value is None or value == "" or unit is None:
        return None, None
    if not isinstance(value, (int, float)):
        return None, None
    u = unit.strip().lower()
    if u in ("mpg", "mpg(us)", "mpg-us"):
        if value <= 0:
            return None, None
        return 235.214583 / value, "L/100km"
    if u in ("mpge",):
        # 33.705 kWh/gal 기준: kWh/100km
        return 3370.5 / (value * 1.609344), "kWh/100km"
    if u in _LINEAR:
        si_unit, f = _LINEAR[u]
        return value * f, si_unit
    return None, None


_DECIMALS = {"mm": 0, "kg": 0, "kW": 0, "N·m": 0, "km/h": 0, "km": 0, "L": 0,
             "L/100km": 1, "kWh/100km": 1, "s": 1, "°": 1, "kWh": 1, "USD": 0, "KRW": 0, "CAD": 0, "EUR": 0, "seats": 0}


def fmt(value, unit: str | None) -> str:
    if value is None:
        return "–"
    d = _DECIMALS.get(unit or "", 1)
    s = f"{value:,.{d}f}"
    if unit == "USD":
        return f"${s}"
    if unit == "KRW":
        return f"₩{s}"
    if unit == "CAD":
        return f"C${s}"
    if unit == "EUR":
        return f"€{s}"
    if unit == "seats":
        return f"{s}인승"
    return f"{s} {unit}" if unit else s


def fmt_original(value, unit: str | None) -> str:
    if value is None or value == "":
        return ""
    if isinstance(value, (int, float)):
        v = f"{value:,.1f}".rstrip("0").rstrip(".")
        return f"{v} {unit}" if unit else v
    return str(value)
