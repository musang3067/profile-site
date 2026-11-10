# 제네시스급 BOF(Body-on-Frame) 신차 모니터링 & 카드뉴스 자동 생성 에이전트 — 구축 계획서

> 작성일: 2026-11-10 · 상태: MVP 구현 · 산출물: 매일 카드뉴스 PPTX 1건

## 0. 확정 사항 (2026-11-10)
| 항목 | 결정 |
|---|---|
| 대상 | **제네시스 전 차종** + **BOF SUV·픽업 중 전장 4,800~5,800mm** (`agent/scope.py`) |
| 가격 | 해외 기준 (US MSRP, 미국 미판매 시 주요 해외시장 공식가) |
| 단위 | 해외 원 단위로 수집 → **SI 단위 환산 표시**, 원 단위 병기 (`agent/units.py`) |
| 배포 | 파일 저장 (`output/DATE/*.pptx`, Git 커밋) |
| 실행 | **Claude 예약 작업** 매일 06:43 KST, 지침 `prompts/daily_run.md` |

> 아래 1~8장은 초기 계획이며, 0장과 충돌하면 0장이 우선합니다.

---

## 1. 목표와 범위

| 항목 | 내용 |
|---|---|
| 모니터링 대상 | 제네시스급(프리미엄/럭셔리) **프레임 바디(BOF)** 플랫폼 SUV·픽업의 신차, 출시 정보, 신기술, 특이사항(리콜·가격 변경·단종·생산지 변경 등) |
| 주기 | ① 실시간에 가까운 폴링(1시간 간격, 속보 감지) ② **1일 1회 정기 수집 + 카드뉴스 생성(매일 07:00 KST)** |
| 산출물 | 매일 카드뉴스 형식의 `.pptx` (중요 뉴스 N건 + 차종별 비교 카드) |
| 비교 카드 | 이슈 차종 1대 + **현대차그룹 동급 차종** + **동급 경쟁차 최대 5대** → 기본 제원 · 가격 · 주요 성능 · 대표 사진 |
| 데이터 원칙 | **각사 공식 홈페이지 우선 → 없으면 신뢰 가능한 인터넷 출처 → 그래도 없으면 빈칸(blank)**. 추정치·생성값 금지, 모든 수치에 출처 URL 기록 |

### 1-1. 모니터링 차종 유니버스 (초기 시드, 운영 중 자동 확장)

> ※ 플랫폼 구분(BOF/모노코크)은 수집 단계에서 매번 재검증합니다. 아래는 초기 시드이며 확인 필요 항목은 표기했습니다.

| 구분 | 차종 (예시) | 비고 |
|---|---|---|
| 럭셔리 풀사이즈 SUV | Cadillac Escalade / Escalade ESV, Lincoln Navigator, GMC Yukon Denali, Jeep Grand Wagoneer, Infiniti QX80, Lexus LX | BOF |
| 프리미엄 중형/오프로드 SUV | Lexus GX, Toyota Land Cruiser, Mercedes-Benz G-Class, INEOS Grenadier, Toyota 4Runner, Nissan Patrol | BOF |
| 프리미엄 픽업 | Ford F-150 Platinum/Raptor, Ram 1500 Limited/TRX, GMC Sierra Denali, Toyota Tundra Capstone | BOF, 필요 시 포함 |
| 현대차그룹 동급 | 기아 타스만(BOF 픽업), 기아 모하비(BOF, 판매 상태 확인 필요), 현대차그룹 차세대 BOF 플랫폼(중형 픽업/대형 SUV 개발 보도 추적), 제네시스 GV80·GV90(모노코크 — BOF 동급 부재 시 **대체 비교군**으로 표기) | 동급 BOF가 없으면 "동급 BOF 없음 – 최근접 모델"로 명시 |
| 제외 | Land Rover Defender/Range Rover (모노코크) 등 | 비교 참고용으로만 사용 가능 |

---

## 2. 전체 아키텍처

```
 ┌──────────────┐   ┌──────────────┐   ┌──────────────┐   ┌──────────────┐   ┌──────────────┐
 │ ① 수집기      │→ │ ② 정규화/중복 │→ │ ③ 중요도 판정 │→ │ ④ 제원 수집기 │→ │ ⑤ 카드뉴스    │
 │ Collector    │   │ Normalizer   │   │ Classifier   │   │ Spec Fetcher │   │ PPTX Builder │
 └──────────────┘   └──────────────┘   └──────────────┘   └──────────────┘   └──────┬───────┘
   RSS·뉴스룸·       URL/제목 해시      LLM 분류+점수       공식HP→웹→blank          │
   공식HP·정부DB      SQLite 저장         (0~100)            출처 URL 필수          ▼
                                                                         ⑥ 배포(메일/Slack/
                                                                            Notion/Git 저장)
```

- 오케스트레이션: **Claude Agent SDK(Python)** 기반 에이전트 + 결정적(deterministic) 파이프라인 혼합
  - 수집·파싱·PPTX 렌더링 → 일반 Python 코드 (재현성·비용)
  - 뉴스 요약, 중요도 판정, 비정형 페이지에서 제원 추출, 경쟁차 선정 → LLM(Claude) 도구 호출
- 스케줄링: (택1) GitHub Actions cron / 서버 cron + systemd / Claude Code 예약 작업(scheduled task)
- 저장소: SQLite(초기) → PostgreSQL(확장), 이미지·PPTX는 날짜별 폴더 / 오브젝트 스토리지

---

## 3. 모듈별 상세 설계

### ① 수집기 (Collector)

| 소스 유형 | 예시 | 방법 | 주기 |
|---|---|---|---|
| 제조사 뉴스룸/프레스 | Genesis·Hyundai·Kia 뉴스룸, GM/Cadillac Pressroom, Ford Media, Toyota/Lexus Newsroom, Mercedes Media, Stellantis Media, Nissan/Infiniti Newsroom, INEOS | RSS 있으면 RSS, 없으면 Playwright 스크래핑 | 1시간 |
| 뉴스 집계 | Google News RSS(키워드 쿼리), 네이버 뉴스 검색 API | 키워드: `body-on-frame`, `프레임 바디`, 차종명, `신형`, `출시`, `페이스리프트`, `recall` | 1시간 |
| 전문 매체 | Motor1, Car and Driver, MotorTrend, Autocar, Automotive News, 오토뷰, 모터그래프, 다나와자동차 | RSS/스크래핑 | 1일 |
| 공공 데이터 | NHTSA 리콜 API, fueleconomy.gov(EPA) API, 국토부 자동차리콜센터, 환경부 인증(출시 전 신호) | REST API | 1일 |
| 보조(선택) | YouTube 공식 채널, 특허·상표 출원(신기술 선행 신호) | API | 1주 |

- 키워드/차종 사전은 `config/vehicles.yaml`로 관리, 새 차종이 감지되면 LLM이 BOF 여부 검증 후 사전에 추가 제안.

### ② 정규화 · 중복 제거
- URL 정규화 + 제목 유사도(임베딩/SimHash)로 동일 이슈 클러스터링
- 스키마: `article(id, source, url, title, published_at, lang, body, vehicle_ids[], cluster_id)`

### ③ 중요도 판정 (LLM)
- 분류 라벨: `신차공개` `출시/판매개시` `가격변경` `신기술` `페이스리프트/연식변경` `리콜/품질` `생산/단종` `루머/스파이샷`
- 점수(0~100) 가중치 예: 공식 발표 +30, 신차/출시 +25, 신기술 +20, 현대차그룹 관련 +15, 복수 매체 보도 +10, 루머 −15
- **70점 이상 → 당일 카드뉴스 수록**, 90점 이상 → 실시간 속보 알림(메일/Slack)

### ④ 제원 수집기 (Spec Fetcher) — 핵심 원칙: "못 찾으면 비운다"

**우선순위 체인**
1. 각사 공식 홈페이지(국내 판매 차종은 한국 홈페이지, 미판매 차종은 본국/북미 홈페이지) 제원표·가격표·카탈로그 PDF
2. 공식 프레스 릴리스 / 미디어 사이트
3. 공공 DB (EPA 연비, NHTSA 등)
4. 신뢰 매체 (Car and Driver, Edmunds, MotorTrend, 다나와자동차 등) — **2곳 이상 일치 시 채택**
5. 위 모두 실패 → **blank** (`""`)로 남기고 카드에는 `–` 표기

**필드 단위 출처 기록 (Provenance)**
```json
{
  "vehicle": "Lexus GX 550",
  "market": "US",
  "trim": "Overtrail+",
  "fields": {
    "price":       {"value": 0, "unit": "USD", "source_url": "", "source_tier": 1, "fetched_at": ""},
    "length_mm":   {"value": null, "source_url": "", "source_tier": null},
    "power_ps":    {"value": null, "source_url": "", "source_tier": null}
  }
}
```
- LLM은 페이지 텍스트에서 **인용 구간(quote)과 함께만** 값을 추출하도록 강제 → 인용문에 숫자가 실제로 존재하는지 코드로 재검증(환각 차단)
- 단위 통일: mm, kg, PS/hp 병기, kg·m/Nm 병기, km/L·MPG 병기, 가격은 현지 통화 원표기 + (선택) KRW 환산(환율 출처·기준일 명기)

**비교 제원 항목 (카드 표준)**

| 분류 | 항목 |
|---|---|
| 기본 제원 | 전장 / 전폭 / 전고 / 휠베이스 / 공차중량 / 승차정원 / 플랫폼(BOF 여부) |
| 파워트레인 | 엔진(배기량·형식) 또는 모터 / 최고출력 / 최대토크 / 변속기 / 구동방식 |
| 성능 | 0→100km/h(또는 0-60mph) / 최고속도 / 견인력 / 접근·이탈각, 도강 깊이(오프로드) / 연비·전비 |
| 가격 | 시작가 ~ 최상위 트림가 (시장·기준일 명기) |
| 신기술/특이사항 | 핵심 신기술 3줄 요약 |
| 사진 | 대표 외관 1장 (공식 미디어 사이트 이미지, 출처 표기) |

### ⑤ 비교군 선정 로직
1. 이슈 차종의 세그먼트(크기·가격대·용도)를 판정
2. **현대차그룹 동급 1대**: 동급 BOF 우선 → 없으면 최근접 모노코크 모델을 "대체 비교"로 라벨링
3. **경쟁차 최대 5대**: 가격대 ±25%, 전장 ±250mm, 같은 용도(SUV/픽업) 기준으로 점수화하여 상위 5대
4. 선정 이유를 카드 하단 각주로 1줄 표기

### ⑥ 카드뉴스 PPTX 빌더
- 라이브러리: `python-pptx` (템플릿 `.potx` 기반), 차트는 python-pptx 네이티브 차트(편집 가능)
- 판형: 카드뉴스용 **1080×1350(4:5)** 기본, 옵션으로 16:9 보고용 동시 출력
- 슬라이드 구성(1일분)

| # | 슬라이드 | 내용 |
|---|---|---|
| 1 | 표지 | 날짜, "오늘의 BOF 프리미엄 신차 브리핑", 헤드라인 3개 |
| 2 | 오늘의 요약 | 중요 이슈 Top 3~5 (라벨·점수·한 줄 요약) |
| 3~N | 이슈 카드 | 이슈 1건당 1장: 대표 사진, 핵심 내용 3줄, 신기술/특이사항, 출처 |
| N+1 | 비교 카드 ① 기본 제원 | 대상차 + 현대차그룹 동급 + 경쟁차 ≤5 — 사진 썸네일 + 제원 표 |
| N+2 | 비교 카드 ② 가격·성능 | 가격대 막대 차트, 출력/토크/0-100 비교 차트 |
| N+3 | 시사점 | 현대차그룹 관점 인사이트(전문가 코멘트, LLM 초안 + 근거 링크) |
| 마지막 | 출처 · 데이터 기준 | 필드별 출처 URL, 수집 시각, blank 항목 목록 |

- 빈칸 처리: 값 없음 → `–` + 회색 처리, 슬라이드 하단에 "공식/공개 자료 미확인 항목" 표기
- 사진: 공식 미디어(프레스킷) 이미지 우선, 캡션에 저작권자 표기. 이미지 없으면 회색 플레이스홀더

### ⑦ 배포
- Git 저장(`output/YYYY-MM-DD/briefing.pptx`), 이메일 첨부(Gmail), Slack/Notion 업로드(선택)
- 실패 시 재시도 3회 + 관리자 알림

---

## 4. 디렉터리 구조(안)

```
bof-monitor-agent/
├── config/
│   ├── vehicles.yaml        # 차종 유니버스, BOF 여부, 공식 URL, 세그먼트
│   ├── sources.yaml         # RSS/뉴스룸/API 목록
│   └── scoring.yaml         # 중요도 가중치
├── agent/
│   ├── collector/           # rss.py, newsroom_scraper.py, nhtsa.py, epa.py
│   ├── pipeline/            # normalize.py, dedup.py, classify.py
│   ├── specs/               # fetch_official.py, fetch_web.py, extract_llm.py, validate.py
│   ├── compare/             # select_competitors.py
│   ├── render/              # pptx_builder.py, charts.py, templates/card.potx
│   └── deliver/             # mailer.py, slack.py
├── data/                    # sqlite, 이미지 캐시
├── output/YYYY-MM-DD/       # 일자별 PPTX
├── tests/
└── run_daily.py / run_poll.py
```

---

## 5. 기술 스택

| 영역 | 선택 |
|---|---|
| 언어 | Python 3.12 |
| 에이전트 | Claude Agent SDK / Anthropic API (요약·분류·추출·인사이트) |
| 수집 | feedparser, httpx, Playwright(JS 렌더링 공식 HP), pdfplumber(카탈로그 PDF) |
| 저장 | SQLite → PostgreSQL |
| 렌더링 | python-pptx, Pillow(이미지 리사이즈/크롭) |
| 스케줄 | GitHub Actions cron (`0 22 * * *` UTC = 07:00 KST) + 1시간 폴링 |
| 품질 | pytest, 제원 검증 규칙(범위 체크: 전장 4,000~6,000mm 등) |

---

## 6. 단계별 일정 (제안)

| 단계 | 기간 | 산출물 |
|---|---|---|
| P0 준비 | 3일 | 차종 유니버스 확정, 소스 목록, 카드 템플릿 디자인 시안 |
| P1 MVP | 1주 | RSS·뉴스룸 수집 → 분류 → 요약 카드 PPTX (비교표 없이) |
| P2 제원 비교 | 1~2주 | 공식 HP 제원/가격 파서(우선 10개 차종), 출처 기록, blank 처리, 비교 카드 2장 |
| P3 고도화 | 1주 | 경쟁차 자동 선정, 차트, 사진 자동 수집, 속보 알림 |
| P4 운영 | 상시 | 파서 깨짐 모니터링, 차종 사전 자동 확장, 주간 품질 리뷰 |

---

## 7. 품질 · 리스크 관리

| 리스크 | 대응 |
|---|---|
| LLM 수치 환각 | 인용문 기반 추출 + 코드 재검증, 출처 없는 값은 저장 불가, 불확실 → blank |
| 공식 HP 구조 변경/차단 | 사이트별 파서 + 실패 감지 알림, 캐시된 직전 값 사용 시 "기준일" 표기 |
| 시장별 제원·가격 상이 | 시장(KR/US/EU/ME)과 트림을 필드에 명시, 카드에 기준 시장 표기 |
| BOF 여부 오분류 | 차종 사전에서 사람이 최종 승인(human-in-the-loop) |
| 이미지 저작권 | 공식 프레스 이미지만 사용, 캡션에 출처, 내부 보고용 원칙 |
| 루머성 기사 | `루머` 라벨 + 감점, 카드에 "미확인" 배지 |
| 비용 | 1차 키워드 필터 후 LLM 호출, 변경 없는 제원은 캐시(주 1회 재검증) |

**KPI**: 중요 이슈 누락률 < 5%, 제원 필드 출처 기록률 100%, 오기재 0건, 매일 07:00 생성 성공률 ≥ 99%

---

## 8. 결정이 필요한 사항

1. 비교 기준 시장: **한국 가격 우선**(국내 미출시 차는 미국 MSRP)으로 할지, 미국 기준 통일할지
2. 픽업트럭(F-150, 타스만 등)을 정식 모니터링 대상에 포함할지
3. 배포 채널: 이메일 / Slack / Notion / 단순 파일 저장 중 선택
4. 실행 환경: GitHub Actions(무료·간편) vs 자체 서버 vs Claude Code 예약 작업
5. 카드 판형: 4:5 카드뉴스 단일 vs 16:9 보고서 동시 생성
