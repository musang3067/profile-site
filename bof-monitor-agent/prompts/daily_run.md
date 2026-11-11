# 일일 실행 지침 (Claude 예약 작업이 매일 이 문서를 따른다)

저장소: `musang3067/profile-site`, 브랜치: `claude/genesis-car-monitoring-agent-lwabcl`, 작업 폴더: `bof-monitor-agent/`
오늘 날짜(KST)를 `DATE`(YYYY-MM-DD)로 사용한다.

## 0. 준비
공식 사이트 렌더링 도구(출력은 스크래치 폴더에 저장):
`node tools/fetch.js <출력폴더> <URL...>` · `node tools/links.js <정규식> <URL...>` · `node tools/click.js <URL> <버튼텍스트> <출력파일>`
```bash
git fetch origin claude/genesis-car-monitoring-agent-lwabcl
git checkout claude/genesis-car-monitoring-agent-lwabcl && git pull origin claude/genesis-car-monitoring-agent-lwabcl
cd bof-monitor-agent && python3 -c "import pptx" || pip install python-pptx Pillow
```

## 1. 모니터링 (최근 24~36시간)
대상: `config/vehicles.yaml` 의 모든 차종 + 새로 발견된 BOF SUV/픽업.
- WebSearch(standard)로 영어·한국어 쿼리를 병렬 실행: 차종명 + `new / reveal / launch / pricing / recall / facelift / 2027`,
  `body-on-frame SUV`, `body-on-frame pickup`, `Genesis`, `제네시스 신차`, `프레임 바디 SUV`, `기아 타스만`, `현대차 픽업`.
- 제조사 뉴스룸(Genesis/Hyundai/Kia, GM, Ford, Stellantis, Toyota/Lexus, Nissan/Infiniti, Mercedes, INEOS)을 web-fetch로 확인.
- 공공 데이터: NHTSA 리콜(api.nhtsa.gov), fueleconomy.gov.
- 새 차종 발견 시 BOF 여부를 공식 자료로 확인하고 `config/vehicles.yaml` 에 추가(확인 못 하면 status: verify).

## 2. 분류 · 중요도
라벨: 신차공개, 출시/판매개시, 가격변경, 신기술, 페이스리프트/연식변경, 리콜/품질, 생산/단종, 루머/스파이샷
점수 0~100: 공식 발표 +30, 신차/출시 +25, 신기술 +20, 현대차그룹 관련 +15, 복수 매체 +10, 루머 −15.
70점 이상만 `issues` 에 넣는다(최대 6건, 점수순). 루머는 `"official": false`.
같은 이슈가 이전 브리프(`data/briefs/*.json`)에 이미 실렸고 새 사실이 없으면 제외.

## 3. 비교군 (최대 3세트/일)
상위 이슈 차종마다 `comparisons` 항목 1개:
- `target`: 이슈 차종
- `hmg`: 현대차그룹 동급 1대 (기아 타스만/모하비 등 BOF 우선, 없으면 가장 가까운 제네시스 SUV)
- `rivals`: 동급 경쟁 BOF 최대 5대 (가격 ±25%, 전장 ±250mm, 같은 용도). 범위 밖 차량은 빌드 시 자동 제외된다.
- `reason`: 선정 기준 한 줄
이슈가 없으면 비교 카드는 생략해도 된다.

## 4. 제원 수집 — `data/specs.json` 갱신
형식(차종 id → 객체):
```json
"lexus_gx": {"name": "Lexus GX 550", "maker": "Lexus", "group": "Toyota", "platform": "BOF", "body": "SUV",
  "market": "US", "trim": "Premium", "model_year": 2026, "image": "data/images/lexus_gx.jpg", "image_credit": "Lexus",
  "specs": {"length": {"value": 0, "unit": "in", "source_url": "https://...", "tier": 1, "fetched_at": "DATE",
                       "quote": "페이지에 적힌 원문 그대로"}}}
```
필드: length, width, height, wheelbase, curb_weight, seats, ground_clearance, engine, power, torque, transmission,
drivetrain, accel_0_60mph, top_speed, towing, fuel_economy, ev_range, price_base, price_top

규칙 (반드시 지킨다)
1. **해외(미국) 기준**: 가격은 US MSRP(USD, 배송비 제외). 미국 미판매 차는 본국/주요 해외시장 공식 사이트 가격을 쓰고 `market` 에 표시.
2. **단위는 공식 사이트의 원 단위 그대로** 기록(in, lb, hp, lb-ft, mph, mpg, mi 등). SI 환산은 빌더가 한다. 계산하지 말 것.
3. 출처 순위: tier1 제조사 공식 홈페이지/제원표/카탈로그 → tier2 공식 보도자료·미디어 사이트 → tier3 EPA/NHTSA → tier4 매체.
   tier4는 서로 다른 매체 2곳 이상이 일치할 때만 `source_urls` 배열로 기록.
4. **못 찾은 값은 필드를 빼거나 value를 null로 둔다(blank). 추정·기억·계산·다른 트림 값 대입 금지.**
5. `quote` 에는 페이지 원문 문자열을 그대로 넣는다(값 검증용).
6. 이미 있는 값은 `fetched_at` 이 30일 이내면 재사용, 신차/가격변경 이슈 차종은 무조건 재확인.
7. 사진: 제조사 공식 미디어/프레스 사이트 이미지만, `data/images/<id>.jpg` 로 저장(가로 1600px 이하), `image_credit` 기록. 없으면 비워 둔다.

## 5. 브리프 작성 · 빌드
`data/briefs/DATE.json`:
```json
{"date": "DATE", "headlines": ["..."], "issues": [{"title": "", "label": "", "score": 0, "official": true,
  "summary": ["3줄"], "tech": ["신기술/특이사항"], "image": "data/images/..", "image_credit": "", "sources": ["url"]}],
 "comparisons": [...], "insights": ["현대차그룹 관점 시사점 2~4개, 근거가 되는 사실만"]}
```
```bash
python3 run_daily.py data/briefs/DATE.json
```
출력: `output/DATE/BOF_Genesis_Brief_YYYYMMDD.pptx`, `output/DATE/scope_report.txt`
`python3 -m pytest -q tests` 통과 확인.

## 6. 저장
```bash
git add -A && git commit -m "Daily brief DATE" && git push -u origin claude/genesis-car-monitoring-agent-lwabcl
```
푸시 실패 시 네트워크 오류만 2/4/8/16초 간격으로 재시도. 마지막에 오늘의 이슈 제목, blank 항목 수, 파일 경로를 한 단락으로 보고한다.

## 7. 메일 발송
푸시가 끝나면 Gmail 커넥터(`mcp__Gmail__send_message`)로 오늘의 PPTX를 보낸다.
- 받는 사람: 예약 작업 프롬프트에 적힌 주소 (저장소에 주소를 적지 않는다)
- 제목: `[BOF·제네시스 데일리] DATE 카드뉴스 – <대표 헤드라인>`
- 본문(htmlBody): 오늘의 이슈 제목·라벨·점수, 3줄 요약, 신기술, 비교 차종 목록, blank 항목 수, 출처 링크
- 첨부: `output/DATE/BOF_Genesis_Brief_YYYYMMDD.pptx`
  (content 는 `sha256:<sha256sum 값>:<파일 절대경로>` 파일 참조 형식, mimeType `application/vnd.openxmlformats-officedocument.presentationml.presentation`)
- 이슈가 없는 날에도 "신규 주요 이슈 없음" 브리프를 보낸다.
- Gmail 도구가 없으면 메일은 건너뛰고 최종 보고에 "Gmail 커넥터 미연결"을 적는다.
