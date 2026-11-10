# BOF · 제네시스 일일 카드뉴스 에이전트

- 대상: 제네시스 전 차종 + BOF SUV/픽업(전장 4,800~5,800mm)
- 매일 Claude 예약 작업이 `prompts/daily_run.md` 지침대로 수집 → `data/specs.json`, `data/briefs/DATE.json` 갱신 → PPTX 생성
- 가격: 해외(US MSRP) · 단위: 해외 원 단위 기록 → SI 환산(원 단위 병기) · 미확인: `–`
- 결과물: `output/DATE/BOF_Genesis_Brief_YYYYMMDD.pptx`

```bash
pip install python-pptx Pillow pytest
python3 run_daily.py data/briefs/2026-11-10.json   # 빌드
python3 -m pytest -q tests                          # 테스트
```
구성: `agent/units.py`(SI 환산) · `agent/specs.py`(출처 규칙·blank 처리) · `agent/scope.py`(대상 판정) · `agent/pptx_builder.py`(4:5 카드뉴스)
