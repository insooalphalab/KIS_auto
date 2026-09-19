# KIS_auto — 한투 Open API 수급·기술분석 리포트 자동화

한국투자증권 Open API로 보유·관심 종목의 시세·수급 데이터를 자동 수집하고,
**Claude Code CLI**가 그 데이터를 읽어 분석한 리포트를 **텔레그램**으로 발송하는 파이프라인입니다.
Windows 작업 스케줄러가 정해진 시각에 실행합니다.

> ⚠️ **면책** — 이 프로젝트는 개인 학습·기록 목적의 코드입니다. 생성되는 리포트는 정보 제공용이며
> 투자 권유가 아닙니다. 투자 판단과 그 결과에 대한 책임은 전적으로 사용자 본인에게 있습니다.
> 제공된 그대로 동작함을 보증하지 않으며, 실제 자금 운용에 사용하기 전에 결과를 직접 검증하세요.
> (이 코드는 주문을 내지 않습니다. 조회 API와 텔레그램 발송만 사용합니다.)

## 동작 방식

```
Windows 작업 스케줄러 ─▶ 한투API_스케줄러.py ─┬─ ① 수집기 (한투 Open API → xlsx)
                                              └─ ② 한투API_텔레그램분석.py
                                                    └─ claude -p 로 xlsx 분석 → 텔레그램 발송
```

| 구분 | 실행 시각 | 명령 | 수집 데이터 |
|---|---|---|---|
| 시간별 | 평일 10:00 / 13:00 / 15:20 | `--once` | 분봉, 체결강도, 외인·기관 가집계, 프로그램 매매 |
| 일별 | 평일 16:05 | `--daily` | 일봉, 수급, 뉴스, 대차거래, 신용잔고, 심화지표, 시장 비중, 회원사 동향, 매물대 |
| 주간 | 토요일 09:00 | `--weekly` | 주봉·40주선 이격도, 주간 수급 누적, 신용잔고 증감, 거래대금, 외인소진율 |

리포트는 **수급과 기술분석 중심**입니다. 금리·유가·환율 같은 매크로 판단은 이 프로젝트의 범위 밖이며,
프롬프트에서도 매크로를 판단하지 않도록 지시합니다. (`월간_매크로섹터분석.py`는 수동 실행용으로만
남아 있고 스케줄에는 등록하지 않습니다.)

## 파일 구성

| 파일 | 역할 |
|---|---|
| `한투API_스케줄러.py` | 진입점. `--once / --daily / --weekly`로 수집 → 분석을 순서대로 실행 |
| `한투API_시간별데이터.py` · `한투API_일별데이터.py` · `한투API_주간데이터.py` | 한투 Open API 수집기 (xlsx 저장) |
| `한투API_텔레그램분석.py` | 프롬프트 정의, Claude Code CLI 호출, 텔레그램 발송, 사용량 로깅 |
| `kis_config.py` | `.env` 로더 (시크릿 분리) |
| `setup_tasks.ps1` / `remove_tasks.ps1` | 작업 스케줄러 등록 / 제거 |
| `.env.example` · `tickers.example.json` | 설정 템플릿 |
| `월간_매크로섹터분석.py` | (수동 실행용) 매크로·섹터 데이터 수집 |

실행 중 생성되는 폴더/파일(모두 git 제외): `MarketData/`(시간별·일별), `WeeklyData/`, `_cache/`(일별 심화지표 캐시),
`claude_usage.jsonl`(Claude 호출별 토큰·비용 로그).

## 준비물

- Windows 10/11, Python 3.9 이상 (3.12에서 테스트)
- 한국투자증권 Open API 앱키/시크릿 (조회 권한)
- 텔레그램 봇 토큰과 채팅 ID (`@BotFather`, `@userinfobot`)
- Node.js와 Claude Code CLI — 설치 후 **사람이 직접 한 번 로그인**해 둘 것
  (`npm install -g @anthropic-ai/claude-code`, 이후 터미널에서 `claude` 실행)

## 설치와 설정

```
git clone https://github.com/insooalphalab/KIS_auto.git
cd KIS_auto
pip install -r requirements.txt

copy .env.example .env                    # 열어서 실제 값 입력
copy tickers.example.json tickers.json    # 열어서 본인 종목/비중으로 수정
```

**`.env`** — API 키·계좌번호·토큰은 코드가 아니라 이 파일에서 읽습니다. `.env`와 `tickers.json`은
`.gitignore`로 제외되어 저장소에 올라가지 않습니다. 같은 이름의 OS 환경변수가 있으면 그쪽이 우선합니다.

| 항목 | 설명 |
|---|---|
| `KIS_APP_KEY`, `KIS_APP_SECRET`, `KIS_CANO` | 한투 Open API 앱키/시크릿, 계좌번호 앞 8자리 |
| `TELEGRAM_BOT_TOKEN`, `TELEGRAM_CHAT_ID` | 텔레그램 봇 |
| `CLAUDE_CLI` | `claude.cmd` 전체 경로 (`where claude`로 확인). 작업 스케줄러 자동 실행에 필요 |
| `FRED_API_KEY`, `CUSTOMS_API_KEY`, `NAVER_CLIENT_ID`, `NAVER_CLIENT_SECRET` | 월간 스크립트를 수동으로 돌릴 때만 필요 |

**`tickers.json`** — 한 줄이 한 종목이며 형식은 `"종목코드;종목명;보유|관심;비중%"` 입니다.

```json
["000660;SK하이닉스;보유;30%", "005930;삼성전자;관심;0%"]
```

## 실행

먼저 한 번씩 수동으로 확인하세요.

```
python 한투API_스케줄러.py --once      # 시간별 수집 + 분석 + 텔레그램 발송
python 한투API_스케줄러.py --daily     # 일별
python 한투API_스케줄러.py --weekly    # 주간
```

텔레그램 발송 없이 콘솔에서 결과만 보려면 분석기를 직접 `--preview`로 실행합니다.

```
python 한투API_텔레그램분석.py --weekly --preview --date 20260919
python 한투API_텔레그램분석.py --daily  --preview
```

## 자동 실행 (Windows 작업 스케줄러)

1. PowerShell을 **관리자 권한**으로 실행 (창 제목이 `관리자: Windows PowerShell`)
2. 프로젝트 폴더에서:
   ```
   Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
   .\setup_tasks.ps1
   ```
3. `KIS_시간별_1000/1300/1520`, `KIS_일별_1605`, `KIS_주간_0900` 5개 작업이 등록됩니다.
4. `taskschd.msc`에서 작업 하나를 우클릭 → **실행**으로 시험하고, 텔레그램 도착과 마지막 실행 결과 `0x0`을 확인합니다.
5. 제거: 관리자 PowerShell에서 `.\remove_tasks.ps1`

자동 실행이 되려면 스케줄 시각에 PC가 켜져 있고 절전 상태가 아니어야 합니다. 놓친 회차는 켜질 때 늦게 실행됩니다.
작업은 "로그인 여부와 무관하게 실행"(S4U)으로 등록되는데, 이 방식은 **사용자 PATH를 읽지 않습니다**.
그래서 `setup_tasks.ps1`은 `python`을 전체 경로로 고정하고, `claude`는 `.env`의 `CLAUDE_CLI`로 찾습니다.

## 분석 프롬프트 수정

리포트 형식과 지침은 `한투API_텔레그램분석.py`의 프롬프트 상수에 있습니다.

- `HOURLY_SYSTEM_PROMPT`, `DAILY_SYSTEM_PROMPT`, `build_weekly_portfolio_prompt()`
- `COMPACT_STYLE_RULES` — 모든 리포트에 붙는 "문장 압축" 공통 규칙 (항목당 한 줄, 완곡 표현 제거)

Claude 호출 옵션은 `CLAUDE_EXTRA_ARGS`에 있습니다. `--disable-slash-commands`, `--strict-mcp-config`로
스킬·MCP 설명이 시스템 프롬프트에 실리는 것을 막아 호출당 약 6,000 토큰을 줄이고,
`--output-format json`으로 받은 토큰·비용을 `claude_usage.jsonl`에 누적합니다.
무인 실행이라 `--dangerously-skip-permissions`를 쓰는 대신, 작업 디렉터리를 그 회차의 데이터 폴더 하나로 한정합니다.
이 설정이 부담되면 실행 환경을 격리하거나 옵션을 조정하세요.

## 문제 해결

| 증상 | 확인 |
|---|---|
| `setup_tasks.ps1`에서 "액세스가 거부되었습니다" | 관리자 권한 PowerShell로 실행했는지 확인 |
| `.ps1` 실행 시 한글이 깨져 파일을 못 찾음 | 파일이 **UTF-8 BOM**으로 저장돼 있어야 함 (Windows PowerShell 5.1) |
| 작업 마지막 실행 결과 `0x80070002` | 작업 스케줄러가 `python`/`claude`를 못 찾음 → `setup_tasks.ps1` 재실행, `.env`의 `CLAUDE_CLI` 확인 |
| 텔레그램에 "Claude Code CLI를 찾을 수 없습니다" | `claude` 설치·로그인 여부, `CLAUDE_CLI` 경로 |
| "Claude Code 분석이 N초 안에 끝나지 않아 중단" | `CLAUDE_TIMEOUT_SEC` 값을 늘림 |
| 토큰 발급 실패 `EGW00133` | 한투 토큰은 1분에 1회만 발급 가능. 잠시 후 재시도 |
| 일별 심화지표 숫자가 이상함 | `_cache/advanced_intraday_{종목코드}.json`을 지우면 다음 실행 때 다시 계산 |

## 라이선스

[MIT](LICENSE)
