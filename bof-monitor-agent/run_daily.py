"""일일 브리프 빌드 엔트리포인트.

사용: python run_daily.py data/briefs/2026-11-10.json [--specs data/specs.json] [--out output]
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from agent.pptx_builder import build
from agent.scope import in_scope
from agent.specs import normalize_vehicle


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("brief")
    ap.add_argument("--specs", default="data/specs.json")
    ap.add_argument("--out", default="output")
    a = ap.parse_args(argv)

    brief = json.loads(Path(a.brief).read_text(encoding="utf-8"))
    raw = json.loads(Path(a.specs).read_text(encoding="utf-8")) if Path(a.specs).exists() else {}
    raw.update(brief.get("vehicles") or {})  # 당일 갱신분 우선
    vehicles = {vid: normalize_vehicle(vid, v) for vid, v in raw.items()}

    report = []
    for vid, v in vehicles.items():
        ok, why = in_scope(v, v["fields"]["length"]["si_value"])
        blanks = sum(1 for f in v["fields"].values() if f["value"] is None)
        report.append(f"{'IN ' if ok else 'OUT'} {vid:28s} {why} · blank {blanks}/{len(v['fields'])}")
    for comp in brief.get("comparisons") or []:
        t = vehicles.get(comp["target"])
        if t is None:
            sys.exit(f"비교 대상 '{comp['target']}' 제원 없음")
        kept = []
        for r in comp.get("rivals", []):
            v = vehicles.get(r)
            if v is None:
                print(f"경고: 경쟁차 '{r}' 제원 없음 → 제외")
            elif not in_scope(v, v["fields"]["length"]["si_value"])[0]:
                print(f"경고: 경쟁차 '{r}' 대상 범위 밖({in_scope(v, v['fields']['length']['si_value'])[1]}) → 제외")
            else:
                kept.append(r)
        comp["rivals"] = kept
        if len(kept) > 5:
            print(f"경고: '{comp['target']}' 경쟁차 5대 초과 → 앞 5대만 사용")

    d = brief["date"]
    out = Path(a.out) / d / f"BOF_Genesis_Brief_{d.replace('-', '')}.pptx"
    build(brief, vehicles, out)
    (out.parent / "scope_report.txt").write_text("\n".join(report) + "\n", encoding="utf-8")
    print("\n".join(report))
    print(f"saved: {out}")


if __name__ == "__main__":
    main()
