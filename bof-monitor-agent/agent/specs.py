"""제원 데이터 검증 · 정규화.

규칙
- 모든 값은 source_url 이 있어야 한다. 없으면 blank.
- tier: 1=제조사 공식 홈페이지/카탈로그, 2=공식 보도자료, 3=정부 DB(EPA/NHTSA), 4=언론/리뷰 매체
- tier 4 값은 서로 다른 출처 2곳 이상(source_urls)에서 일치해야 채택. 아니면 blank.
- 단위는 원본(해외) 단위를 보존하고 SI 환산값을 추가한다.
"""
from __future__ import annotations

from .units import to_si, fmt, fmt_original

# 카드에 쓰는 필드 정의: key -> (라벨, 그룹)
FIELDS = {
    "length":          ("전장", "dims"),
    "width":           ("전폭", "dims"),
    "height":          ("전고", "dims"),
    "wheelbase":       ("휠베이스", "dims"),
    "curb_weight":     ("공차중량", "dims"),
    "seats":           ("승차정원", "dims"),
    "ground_clearance": ("최저지상고", "dims"),
    "engine":          ("엔진/모터", "power"),
    "power":           ("최고출력", "power"),
    "torque":          ("최대토크", "power"),
    "transmission":    ("변속기", "power"),
    "drivetrain":      ("구동방식", "power"),
    "accel_0_60mph":   ("0→60mph(96.6km/h)", "perf"),
    "top_speed":       ("최고속도", "perf"),
    "towing":          ("최대 견인", "perf"),
    "fuel_economy":    ("복합연비", "perf"),
    "ev_range":        ("1회충전 주행거리", "perf"),
    "price_base":      ("시작가(US MSRP)", "price"),
    "price_top":       ("최상위 트림가", "price"),
}

TEXT_FIELDS = {"engine", "transmission", "drivetrain"}


def _blank(key, reason):
    return {"key": key, "label": FIELDS[key][0], "value": None, "unit": None,
            "si_value": None, "si_unit": None, "display": "–", "original": "",
            "sources": [], "tier": None, "blank_reason": reason}


def normalize_field(key: str, raw: dict | None) -> dict:
    if not raw or raw.get("value") in (None, ""):
        return _blank(key, "자료 미확인")
    sources = list(raw.get("source_urls") or ([raw["source_url"]] if raw.get("source_url") else []))
    sources = [s for s in sources if s]
    tier = raw.get("tier")
    if not sources:
        return _blank(key, "출처 없음 → 미채택")
    if tier == 4 and len(set(sources)) < 2:
        return _blank(key, "매체 단일 출처 → 미채택")
    value, unit = raw["value"], raw.get("unit")
    out = {"key": key, "label": FIELDS[key][0], "value": value, "unit": unit,
           "sources": sources, "tier": tier, "blank_reason": None,
           "original": fmt_original(value, unit)}
    if key in TEXT_FIELDS or not isinstance(value, (int, float)):
        out.update(si_value=None, si_unit=None, display=str(value))
        return out
    si_v, si_u = to_si(value, unit)
    if si_v is None:
        return _blank(key, f"단위 '{unit}' 환산 불가")
    out.update(si_value=si_v, si_unit=si_u, display=fmt(si_v, si_u))
    return out


def normalize_vehicle(vid: str, v: dict) -> dict:
    specs = v.get("specs") or {}
    fields = {k: normalize_field(k, specs.get(k)) for k in FIELDS}
    return {**{k: v.get(k) for k in ("name", "maker", "group", "platform", "body",
                                       "market", "trim", "model_year", "image", "image_credit")},
            "id": vid, "fields": fields}
