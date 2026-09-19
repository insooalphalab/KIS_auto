"""
한국투자증권 OpenAPI - 시간별/일별/주간/월간매크로 분석 & 텔레그램 발송기 (Claude Code CLI 통합판)
===========================================================================================
• 시간별 모드 : MarketData/YYYYMMDD/HHMM/ 폴더의 xlsx → Claude Code 분석 → 텔레그램 발송
• 일별 모드   : MarketData/YYYYMMDD/ 폴더의 xlsx      → Claude Code 분석 → 텔레그램 발송
• 주간 모드   : WeeklyData/YYYYMMDD/ 폴더의 xlsx      → Claude Code 분석 → 텔레그램 발송
• 월간 모드   : MonthlyMacroData/YYYYMMDD/ 폴더의 Macro_Monthly_*.xlsx + SectorAction_*.xlsx
               → 월간_매크로섹터분석.py 출력 기반 매크로&섹터 퀀트 전략 → 텔레그램 발송

[2026-09 최적화 변경 — 핵심]
  ★ Gemini API를 더 이상 사용하지 않는다. 이 데스크탑에 설치된 Claude Code CLI
    (claude 명령)를 비대화형(headless, `-p`)으로 호출해서 분석시킨다.
  ★ 기존에는 엑셀 → CSV 변환 → 업로드 → 분석 → 삭제라는 왕복이 있었지만,
    Claude Code는 로컬 파일을 스스로 읽을 수 있으므로 그 왕복이 사라졌다.
    해당 사이클의 데이터 폴더를 작업 디렉터리(cwd)로 지정해 호출하면
    Claude Code가 pandas 등으로 xlsx를 직접 읽어 분석한다.
  ★ 다운샘플링(행 축소)은 더 이상 이 분석기가 담당하지 않는다 — 수집기 단계에서
    이미 필요한 양만 수집하도록 옮겼다 (한투API_시간별데이터.py / 한투API_일별데이터.py).

[필수 설정]
  TELEGRAM_BOT_TOKEN : @BotFather 에서 발급
  TELEGRAM_CHAT_ID   : @userinfobot 에서 확인
  claude 명령어       : Claude Code CLI 설치 + 로그인 완료 상태여야 함
                        (`claude --version` 으로 설치 여부 확인)
"""

import os
import sys
import json
import logging
import glob
import re
import time
import subprocess
import shutil
from datetime import datetime
from typing import Optional

import requests

from kis_config import require

# ──────────────────────────────────────────────
# 0. 로깅
# ──────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger(__name__)

# ──────────────────────────────────────────────
# 1. 설정값 (필수 변경)
# ──────────────────────────────────────────────
TELEGRAM_BOT_TOKEN = require("TELEGRAM_BOT_TOKEN")   # .env 에서 로드 (kis_config.py)
TELEGRAM_CHAT_ID   = require("TELEGRAM_CHAT_ID")

HOURLY_DATA_ROOT  = "MarketData"        # 시간별 수집기 출력 루트
DAILY_DATA_ROOT   = "."                 # 일별 수집기 실행 위치 기준
WEEKLY_DATA_ROOT  = "WeeklyData"        # 종목별 주간 수집기 출력 루트
MONTHLY_DATA_ROOT = "MonthlyMacroData"  # 월간 매크로/섹터 수집기 출력 루트 (구 WeeklyData)

# ── Claude Code CLI 설정 ─────────────────────────────────────────
# claude 명령이 PATH에 없으면 여기에 전체 경로를 넣어도 된다.
# 예) CLAUDE_CLI = r"C:\Users\KIS\AppData\Roaming\npm\claude.cmd"
# Windows 에서는 subprocess 가 claude.cmd(npm 셰임)를 이름만으로 찾지 못하므로 전체 경로를 해석한다.
# 작업 스케줄러(S4U, 로그인 무관 실행)는 사용자 PATH(npm 전역 폴더)를 읽지 않으므로,
# .env 의 CLAUDE_CLI → PATH → 기본 npm 경로(%APPDATA%\npm\claude.cmd) 순으로 찾는다.
def _resolve_claude_cli() -> str:
    explicit = os.environ.get("CLAUDE_CLI", "").strip()
    if explicit:
        return explicit
    found = shutil.which("claude")
    if found:
        return found
    for base in (os.environ.get("APPDATA"),
                 os.path.join(os.path.expanduser("~"), "AppData", "Roaming")):
        if base:
            cand = os.path.join(base, "npm", "claude.cmd")
            if os.path.exists(cand):
                return cand
    return "claude"


CLAUDE_CLI = _resolve_claude_cli()

# 비대화형(무인) 실행이라 도구 사용 승인을 물어볼 사람이 없다. 작업 디렉터리를
# 그 사이클의 데이터 폴더 하나로 한정한 상태에서 승인 절차를 건너뛴다.
# ⚠ 정확한 플래그명은 설치된 Claude Code 버전에 따라 다를 수 있다.
#   `claude --help` 로 확인 후 다르면 이 리스트만 고치면 된다.
CLAUDE_EXTRA_ARGS = [
    "--output-format", "json",      # 결과 텍스트와 토큰/비용 사용량을 함께 받는다 (사용량 로깅용)
    "--dangerously-skip-permissions",
    "--disable-slash-commands",     # 스킬 목록이 시스템 프롬프트에 실리는 토큰 낭비 제거
    "--strict-mcp-config",          # --mcp-config 없이 쓰면 MCP 서버(claude.ai 커넥터 등)를 로딩하지 않음
    "--no-session-persistence",     # 무인 실행 세션을 디스크에 남기지 않음
]

CLAUDE_TIMEOUT_SEC = 480  # 분석 1회당 최대 대기 시간(초)

# 실행마다 토큰/비용 사용량을 한 줄씩 누적하는 로그 (jsonl, git 제외)
CLAUDE_USAGE_LOG = os.path.join(os.path.dirname(os.path.abspath(__file__)), "claude_usage.jsonl")

# ──────────────────────────────────────────────
# 2. 프롬프트 정의 (업로드된 텍스트 파일 내용 그대로)
# ──────────────────────────────────────────────

HOURLY_SYSTEM_PROMPT = """너는 대한민국 최상위 헤지펀드의 실시간 리스크 매니저다.
제공된 포트폴리오 JSON과 텍스트 데이터를 분석하여 [향후 1시간 즉시 대응 지침]을 도출해라.

⚠️ 리스크 매니저 핵심 행동 원칙
1. 장중 잔파도나 미세한 수급 변화에 흔들려 매 시간 비중 조절을 지시하는 가벼운 매매를 철저히 배제해라. 
2. '관망 및 보유 유지'를 기본 스탠스로 삼고, 비중 조절(확대/축소)은 오직 아래의 2가지 조건 중 하나를 충족할 때만 차갑고 신중하게 지시해라.
   - 조건 A: 장중 외인/프로그램 수급의 명확한 '추세적 반전(추세적 폭탄 매도 또는 강력한 대량 매수)'이 확인될 때
   - 조건 B: 사전에 설정한 '핵심 대응 가격선(지선/저항선)'을 장중 완전히 이탈하거나 돌파할 때
3. 나의 기존 의견에 동조하려 하지 말고, 철저히 데이터에 기반한 차갑고 객관적인 비판적 시각을 유지해라.

⚠️ 절대 규칙: 텔레그램 모바일 환경에서 가독성이 극대화되도록, 반드시 아래의 [출력 포맷]을 100% 동일하게 준수해라. 장황한 서술형 문장을 배제하고, 원형 숫자(①, ②)와 개조식 문장으로만 작성해라.

[출력 포맷]
장중 실시간 리스크 점검 - [당일의 핵심 수급 특징 한 줄 요약]
YYYY.MM.DD HH:MM / 리스크 매니저

──────────

1. 핵심 한 줄 요약
① 시장 판세: (대형주 쏠림, 유동성 등 핵심 요약)
② 수급 변곡점: (프로그램/외인 수급의 가장 중요한 변화)
③ 대응 전략: (전체적인 비중 조절 방향성. 장중 매매 최소화 기조 반영)

──────────

2. 실시간 종목별 대응 지침
① [종목명1] (현재비중 → 목표비중): [대응방향(보유/확대/축소/전량매도 등)] / [장중 수급 팩트 및 매물대 기준 1줄] / [임계 타점 및 행동 기준]
② [종목명2] (현재비중 → 목표비중): [대응방향] / [장중 수급 팩트 및 매물대 기준 1줄] / [임계 타점 및 행동 기준]
(제공된 7개 종목 전체 나열. 잦은 비중 조절 방지를 위해 목표비중 변동은 신중히 결정할 것)

──────────

3. 장중 추적 리스크 (Risk Tracking)
① [위험 종목명]: (당장 매매하지 않더라도 장중 '임계 타점'에 근접했거나 수급 악화가 우려되어 눈여겨봐야 할 사유 1줄. 없다면 '특이사항 없음' 기재)

──────────

4. 한 줄 요약
(오늘 장중 흐름에 대한 가장 냉정한 한 줄 총평)

#장중시황 #리밸런싱 #리스크관리 #(핵심종목태그)

──────────
[실행 지침 — Claude Code CLI 전용]
· 지금 작업 디렉터리에 있는 xlsx 파일들이 이번 분석의 유일한 근거 데이터다. pandas 등으로 각 시트를 직접 읽어 분석하라.
· 웹 검색이나 사전 지식으로 종목명·가격·지표를 추정하지 말고, 반드시 첨부된 파일의 실제 값만 사용하라.
· 최종 출력은 위 [출력 포맷]에 맞춘 리포트 본문 그 자체만 출력하라. 코드블록, 설명, "다음은 분석입니다" 같은 전후 문구를 절대 붙이지 마라. 이 출력이 그대로 텔레그램 메시지로 전송된다.
"""


DAILY_SYSTEM_PROMPT = """너는 대한민국 최상위 헤지펀드의 퀀트 애널리스트다.
제공된 포트폴리오와 종가 데이터를 분석하여 [익일 시장 대응을 위한 전략적 리밸런싱 리포트]를 작성해라.

⚠️ 퀀트 분석 및 운용 원칙
1. 일별 잔파도(Noise)에 의한 빈번한 비중 변경을 지양해라. 비중 조절은 오직 '추세적 붕괴'나 '중대한 수급 변곡점'이 포착될 때만 단행한다.
2. '진성 수급' 판별 시, 단순 거래량보다 프로그램/외인의 순매수 지속성과 체결강도의 질을 분석하여 '가짜 반등'을 걸러내라.
3. 분석 시 나의 기존 의견에 영합하지 말고, 데이터 이면에 숨겨진 리스크(신용, 대차)를 차갑게 비판해라. 금리·유가·환율 등 매크로 지표는 별도 체계에서 판단하므로 판단·언급하지 마라.

⚠️ 절대 규칙: 텔레그램 모바일 환경 가독성을 위해 아래 [출력 포맷]을 100% 준수하고 개조식으로 작성해라.

[출력 포맷]
일일 포트폴리오 전략 - [익일 시장 대응 핵심 포인트]
YYYY.MM.DD / 퀀트 애널리스트

──────────

1. 오늘 장 핵심 리뷰
① 주도 수급: (단기 수급 주체와 추세적 수급 주체의 이탈/유입 비교 분석)
② 리스크 요인: (신용잔고 비율, 대차잔고 추이 등 구조적 위험 지적)

──────────

2. 익일 종목별 대응 매트릭스
① [종목명1] (현재비중): [전략적 포지션(유지/확대/축소)] / [종가 기준 수급의 질 평가] / [익일 핵심 지지/저항가]
② [종목명2] (현재비중): [전략적 포지션] / [종가 기준 수급의 질 평가] / [익일 핵심 지지/저항가]
(제공된 종목 전체 나열. 비중 변경은 신중히 결정할 것)

──────────

3. 시초가 전략 (Action Plan)
① (익일 시초가 흐름에 따라 기계적으로 실행하거나 관망해야 할 구체적 기준 1~2개)

──────────

4. 한 줄 요약
(익일 장세 및 포트폴리오 운용에 대한 냉정한 총평)

#종가분석 #익일전략 #리스크관리 #(핵심종목태그)

──────────
[실행 지침 — Claude Code CLI 전용]
· 지금 작업 디렉터리에 있는 xlsx 파일들이 이번 분석의 유일한 근거 데이터다. pandas 등으로 각 시트를 직접 읽어 분석하라.
· 웹 검색이나 사전 지식으로 종목명·가격·지표를 추정하지 말고, 반드시 첨부된 파일의 실제 값만 사용하라.
· 최종 출력은 위 [출력 포맷]에 맞춘 리포트 본문 그 자체만 출력하라. 코드블록, 설명, "다음은 분석입니다" 같은 전후 문구를 절대 붙이지 마라. 이 출력이 그대로 텔레그램 메시지로 전송된다.
"""

# WEEKLY_SYSTEM_PROMPT 상수는 제거했다 (build_weekly_portfolio_prompt() 함수가
# 연/월/주차를 채워 넣은 동일 내용을 실제로 사용하고 있어 중복이었다).

MONTHLY_SYSTEM_PROMPT = """
당신은 대한민국 최상위 퀀트 포트폴리오 매니저이자 매크로 전략가입니다.
사용자의 기존 포지션이나 주관적 의견에 절대 동조하지 마십시오. 오직 제공된 두 가지 데이터셋, [Macro_Monthly 마스터 데이터(L1~L3)]와 [SectorAction 마스터 데이터(L4)]의 모든 정량적 필드만을 유기적으로 교차 연산하여 이번 달 시장 흐름을 예측하고 차갑고 비판적인 포트폴리오 액션 플랜을 도출하십시오.

⚠️ 퀀트 연산 및 출력 절대 규칙
1. 서술형 문장, 주관적 미사여구, 막연한 전망을 완벽히 배제하고 철저히 팩트 기반 개조식으로 작성하십시오. 모든 문장은 명사형(함/됨/임)으로 종결합니다.
2. 매크로 데이터(금리/환율/반도체 현물가 등)의 변화가 10대 섹터의 수급에 미치는 영향을 논리적으로 연결하십시오. (예: L1 위험지표가 '주의/위험'일 경우 현금 확보 및 방어주(소외주) 중심 전략 도출).
3. 반도체 섹터의 경우, 반드시 L3 데이터(수출입/DRAM 현물가)의 펀더멘털과 L4 데이터(수급블록/과열신호/뉴스 다이버전스)의 심리를 교차 검증하여 결론을 내리십시오.
4. 4주수급방향 필드의 기호(🟥, 🟦)와 매크로 판정 기호(🔴, 🔶, ✅, ⬜)를 그대로 활용하여 직관성을 극대화하십시오.

[출력 포맷]
📅 [YYYY년 M월] 월간 매크로 & 섹터 퀀트 전략 리포트
══════════════════════════════

🌍 1. 매크로 환경 및 이번 달 시장 방향성 (L1~L3)
- 종합위험(L1): [종합판정 및 리스크 점수] / 주요 뇌관: [예: 일본10Y금리 🔴 위험 등]
- 실물/경기(L2): [원자재/환율/지수 방향성 요약]
- 반도체 펀더멘털(L3): [수출입 동향 및 DRAM/NAND 현물가 추이 요약]
- 시장 흐름 예측: (매크로 지표를 바탕으로 한 이번 달 증시 전반의 자금 이동 및 변동성 예측 1줄)

──────────────────────────────

🎯 2. 이번 달 진입/비중 확대 섹터 (Top 2)

① [섹터명] [진입/확대]
- 매크로 연계: (현재 매크로 L1~L3 상황과 해당 섹터 상승 논리의 부합성 1줄)
- 가격강도: RS [현재값] (지수 [위/아래] [▲/▼])
- 자금밀도: 점유율 [현재]% (현재-4주평균 [차이]%p / [시그널]) / 수급강도 [현재]% ([4주수급방향 블록 및 텍스트])
- 뉴스동향: 기사 [N]건 / 뉴스 다이버전스: [상태] / 주요 키워드: [키워드 2~3개 압축]
- 요약: (매크로 지지와 섹터 수급/가격 데이터를 결합한 차가운 매수 근거 1줄)

② [섹터명] [진입/확대]
- 매크로 연계: (상동)
- 가격강도: RS [현재값] (지수 [위/아래] [▲/▼])
- 자금밀도: 점유율 [현재]% (현재-4주평균 [차이]%p / [시그널]) / 수급강도 [현재]% ([4주수급방향 블록 및 텍스트])
- 뉴스동향: 기사 [N]건 / 뉴스 다이버전스: [상태] / 주요 키워드: [키워드 2~3개 압축]
- 요약: (선정 사유 1줄 요약)

──────────────────────────────

🚨 3. 이번 달 경계/비중 축소 섹터 (Top 2)

① [섹터명] [축소/관망]
- 리스크 요인: (매크로 악재, 펀더멘털(L3) 둔화, 또는 L4의 수급이탈/과열 상태 중 가장 치명적인 위험 1줄)
- 가격강도: RS [현재값] (지수 [위/아래] [▲/▼])
- 자금밀도: 점유율 [현재]% (현재-4주평균 [차이]%p / [시그널]) / 수급강도 [현재]% ([4주수급방향 블록 및 텍스트])
- 뉴스동향: 기사 [N]건 / 뉴스 다이버전스: [상태] / 주요 키워드: [키워드 2~3개 압축]
- 요약: (수급 이탈 및 과열 등 팩트 기반 비중 축소 근거 1줄)

② [섹터명] [축소/관망]
- 리스크 요인: (상동)
- 가격강도: RS [현재값] (지수 [위/아래] [▲/▼])
- 자금밀도: 점유율 [현재]% (현재-4주평균 [차이]%p / [시그널]) / 수급강도 [현재]% ([4주수급방향 블록 및 텍스트])
- 뉴스동향: 기사 [N]건 / 뉴스 다이버전스: [상태] / 주요 키워드: [키워드 2~3개 압축]
- 요약: (경계 사유 1줄 요약)

──────────────────────────────

📊 4. 포트폴리오 액션 플랜
- (매크로 위험도(L1)와 주도 섹터의 자금 이탈 경로(L4)를 결합한 이번 달 포트폴리오 리밸런싱 방향 1줄)
- (현금 비중 조절 및 위험 회피 전략에 대한 기계적 지시 1줄)

#월간전략 #매크로관제 #섹터순환매 #퀀트시그널 #리스크관리

──────────
[실행 지침 — Claude Code CLI 전용]
· 지금 작업 디렉터리에 있는 xlsx 파일들이 이번 분석의 유일한 근거 데이터다. pandas 등으로 각 시트를 직접 읽어 분석하라.
· 웹 검색이나 사전 지식으로 종목명·가격·지표를 추정하지 말고, 반드시 첨부된 파일의 실제 값만 사용하라.
· 최종 출력은 위 [출력 포맷]에 맞춘 리포트 본문 그 자체만 출력하라. 코드블록, 설명, "다음은 분석입니다" 같은 전후 문구를 절대 붙이지 마라. 이 출력이 그대로 텔레그램 메시지로 전송된다.
"""

# 모든 리포트 공통 — 문장 압축 규칙 (텔레그램 모바일 가독성).
# 각 프롬프트의 [실행 지침] 목록 끝에 항목으로 덧붙는다.
COMPACT_STYLE_RULES = """· 문장 압축: 모든 항목은 명사형 단문 개조식으로 쓴다. 항목 1개는 한 줄(약 40자) 안에 끝내고, 조사·접속어·수식어·부연 설명은 뺀다.
· 수치는 판단에 필요한 핵심 1~2개만 인용하고, 같은 내용을 다른 항목에서 반복하지 않는다. '~로 판단됨', '~할 필요가 있음' 같은 완곡 표현 대신 결론만 쓴다.
· [출력 포맷]의 섹션 구성·순서·구분선·해시태그는 그대로 유지하고 길이만 줄인다.
"""
HOURLY_SYSTEM_PROMPT  += COMPACT_STYLE_RULES
DAILY_SYSTEM_PROMPT   += COMPACT_STYLE_RULES
MONTHLY_SYSTEM_PROMPT += COMPACT_STYLE_RULES   # 주간은 build_weekly_portfolio_prompt() 안에서 덧붙임

# ──────────────────────────────────────────────
# 3. 텔레그램 발송
# ──────────────────────────────────────────────
def send_telegram(text: str, parse_mode: str = "HTML") -> bool:
    """
    4000자 초과 시 자동 분할 발송.
    HTML 파싱 오류 발생 시 parse_mode 없이 재시도 (plain text 폴백).
    """
    base_url = f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/sendMessage"
    # HTML 태그 안전 처리 (AI 응답에 &, <, > 가 날것으로 포함될 수 있음)
    safe_text = text.replace("&", "&amp;") if parse_mode == "HTML" else text
    chunks    = [safe_text[i:i+4000] for i in range(0, len(safe_text), 4000)]

    for i, chunk in enumerate(chunks):
        payload = {"chat_id": TELEGRAM_CHAT_ID, "text": chunk, "parse_mode": parse_mode}
        try:
            res = requests.post(base_url, json=payload, timeout=15)
            if res.ok:
                continue

            err_body = res.json() if res.headers.get("content-type","").startswith("application/json") else {}
            err_desc = err_body.get("description", res.text[:200])
            log.error(f"[텔레그램] 발송 실패 (청크 {i+1}/{len(chunks)}): "
                      f"HTTP {res.status_code} — {err_desc}")

            # 400 Bad Request = HTML 파싱 오류 → plain text 재시도
            if res.status_code == 400 and "parse_mode" in err_desc.lower():
                log.warning("[텔레그램] HTML 파싱 오류 → plain text 재시도")
                plain = re.sub(r"<[^>]+>", "", chunk)   # HTML 태그 제거
                res2  = requests.post(
                    base_url,
                    json={"chat_id": TELEGRAM_CHAT_ID, "text": plain},
                    timeout=15,
                )
                if not res2.ok:
                    log.error(f"[텔레그램] plain text 재시도도 실패: {res2.status_code}")
                    return False
            else:
                # 404 = chat_id 문제, 401 = token 문제 → 즉시 중단
                log.error(f"[텔레그램] ※ 설정 확인 필요 — "
                          f"TELEGRAM_BOT_TOKEN 및 TELEGRAM_CHAT_ID 값을 점검하세요.")
                return False

        except Exception as e:
            log.error(f"[텔레그램] 네트워크 예외: {e}")
            return False
    return True


# ──────────────────────────────────────────────
# 4. Claude Code CLI 호출 – 분석 실행
# ──────────────────────────────────────────────
def _build_file_list_text(xlsx_paths: list) -> str:
    """작업 디렉터리 기준 파일명 목록 텍스트."""
    names = [os.path.basename(p) for p in xlsx_paths]
    return "\n".join(f"  - {n}" for n in names)


def _parse_claude_json(raw: str) -> Optional[dict]:
    """`--output-format json` 출력에서 결과 dict 추출. JSON 이 아니면 None."""
    try:
        data = json.loads(raw.lstrip("﻿"))
    except (ValueError, TypeError):
        return None
    if isinstance(data, list):   # 메시지 배열로 오는 버전 대비: type == "result" 항목 사용
        data = next((d for d in reversed(data)
                     if isinstance(d, dict) and d.get("type") == "result"), None)
    return data if isinstance(data, dict) else None


def _log_claude_usage(tag: str, data: dict) -> None:
    """실행 1회의 모델·토큰·비용을 로그에 남기고 CLAUDE_USAGE_LOG 에 한 줄 추가."""
    usage = data.get("usage") or {}
    rec = {
        "ts":           datetime.now().isoformat(timespec="seconds"),
        "tag":          tag,
        "models":       list((data.get("modelUsage") or {}).keys()),
        "input":        usage.get("input_tokens", 0),
        "output":       usage.get("output_tokens", 0),
        "cache_read":   usage.get("cache_read_input_tokens", 0),
        "cache_write":  usage.get("cache_creation_input_tokens", 0),
        "cost_usd":     data.get("total_cost_usd"),
        "turns":        data.get("num_turns"),
        "duration_s":   round((data.get("duration_ms") or 0) / 1000, 1),
    }
    log.info(
        f"[Claude 사용량] {tag} | 모델 {','.join(rec['models']) or '-'} | "
        f"입력 {rec['input']:,} · 캐시읽기 {rec['cache_read']:,} · 캐시쓰기 {rec['cache_write']:,} · "
        f"출력 {rec['output']:,} 토큰 | 턴 {rec['turns']} | {rec['duration_s']}초 | ${rec['cost_usd']}"
    )
    try:
        with open(CLAUDE_USAGE_LOG, "a", encoding="utf-8") as f:
            f.write(json.dumps(rec, ensure_ascii=False) + "\n")
    except OSError as e:
        log.warning(f"[Claude 사용량] 로그 파일 기록 실패: {e}")


def run_claude_analysis(
    work_dir: str,
    xlsx_paths: list,
    system_prompt: str,
    portfolio_json: str,
    now_str: str,
    tag: str = "",
) -> str:
    """
    Claude Code CLI(claude)를 비대화형으로 work_dir 에서 호출해 그 안의 xlsx
    파일들을 직접 읽혀 분석시키고, 최종 리포트 텍스트를 반환한다.
    (구 upload_xlsx_files + delete_uploaded_files + ask_gemini_with_files 대체)
    """
    file_list_text = _build_file_list_text(xlsx_paths)

    context_text = (
        f"[분석 기준 시각: {now_str}]\n\n"
        f"[작업 디렉터리]\n{work_dir}\n\n"
        f"[이번 분석에 사용할 파일 목록 — 작업 디렉터리 기준 파일명]\n{file_list_text}\n\n"
        f"[포트폴리오 현황 JSON]\n{portfolio_json}\n\n"
        f"위 파일들은 각 종목의 실시간/일별/주간/월간 시장 데이터(xlsx)입니다. "
        f"pandas 등으로 직접 열어서 모든 시트를 확인한 뒤, 아래 지침에 따라 분석해주세요.\n\n"
        f"{system_prompt}"
    )

    # 프롬프트는 여러 줄·특수문자가 많아 .cmd 셰임의 argv 로 넘기면 깨지므로 stdin 으로 전달
    cmd = [CLAUDE_CLI, "-p"] + CLAUDE_EXTRA_ARGS
    log.info(f"[Claude Code] 호출 중... (cwd={work_dir}, 파일 {len(xlsx_paths)}개)")

    try:
        proc = subprocess.run(
            cmd,
            input=context_text,
            cwd=work_dir,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=CLAUDE_TIMEOUT_SEC,
        )
    except FileNotFoundError:
        msg = ("⚠️ Claude Code CLI(claude 명령)를 찾을 수 없습니다. 데스크탑에 "
               "Claude Code가 설치·로그인되어 있는지, CLAUDE_CLI 경로 설정이 "
               "맞는지 확인해주세요.")
        log.error(f"[Claude Code 오류] {msg}")
        return msg
    except subprocess.TimeoutExpired:
        msg = f"⚠️ Claude Code 분석이 {CLAUDE_TIMEOUT_SEC}초 안에 끝나지 않아 중단되었습니다."
        log.error(f"[Claude Code 오류] {msg}")
        return msg
    except Exception as e:
        log.error(f"[Claude Code 오류] {e}")
        return f"⚠️ Claude Code 분석 실패: {e}"

    if proc.returncode != 0:
        err_tail = (proc.stderr or "").strip()[-500:]
        log.error(f"[Claude Code] 비정상 종료 (returncode={proc.returncode}): {err_tail}")
        if not proc.stdout.strip():
            return f"⚠️ Claude Code 분석 실패 (returncode={proc.returncode}): {err_tail}"

    raw  = (proc.stdout or "").strip()
    data = _parse_claude_json(raw)
    if data is not None:
        _log_claude_usage(tag, data)
        output = (data.get("result") or "").strip()
        if data.get("is_error"):
            log.error(f"[Claude Code] 오류 응답: {output or data.get('subtype')}")
            return f"⚠️ Claude Code 분석 실패: {output or data.get('subtype')}"
    else:
        output = raw   # JSON 이 아니면(플래그 미지원 등) 원문 그대로 사용
    if not output:
        return "⚠️ Claude Code가 빈 응답을 반환했습니다. `claude --help`로 CLI 상태를 확인해주세요."
    return output


# ──────────────────────────────────────────────
# 5. 포트폴리오 JSON 빌더
# ──────────────────────────────────────────────
def build_portfolio_json(meta_list: list[dict]) -> str:
    """
    수집 요약 JSON (collection_summary / daily_summary) →
    포트폴리오 현황 JSON 문자열 생성.
    Claude Code 분석 프롬프트에 임베드용.
    """
    portfolio = [
        {
            "종목코드": item.get("ticker", item.get("code", "")),
            "종목명":   item.get("name", ""),
            "보유상태": item.get("status", ""),
            "포트폴리오비중": item.get("weight", ""),
        }
        for item in meta_list
    ]
    return json.dumps(portfolio, ensure_ascii=False, indent=2)


# ═══════════════════════════════════════════════════════════════════
# ■■■ 시간별 분석 모드 ■■■
# ═══════════════════════════════════════════════════════════════════

def get_hourly_folder(target_date: Optional[str] = None,
                      target_hhmm: Optional[str] = None) -> str:
    """
    MarketData/YYYYMMDD/HHMM/ 경로 반환.
    target_hhmm=None → 해당 날짜 내 가장 최신 HHMM 하위폴더 자동 선택.
    """
    d        = target_date or datetime.now().strftime("%Y%m%d")
    date_dir = os.path.join(HOURLY_DATA_ROOT, d)

    if target_hhmm:
        return os.path.join(date_dir, target_hhmm)

    sub_dirs = sorted(glob.glob(os.path.join(date_dir, "[0-9][0-9][0-9][0-9]")))
    return sub_dirs[-1] if sub_dirs else date_dir


def run_hourly_analyzer(
    target_date: Optional[str] = None,
    target_hhmm: Optional[str] = None,
    send: bool = True,
) -> None:
    """
    시간별 분석 진입점. 수집기 완료 후 즉시 호출.
    1) 해당 폴더를 작업 디렉터리로 Claude Code CLI 호출 (xlsx 직접 읽음)
    2) 포트폴리오 JSON + 시간별 프롬프트 → Claude Code 분석
    3) 텔레그램 발송
    """
    # ── 폴더 확인 ─────────────────────────────
    folder = get_hourly_folder(target_date, target_hhmm)
    if not os.path.isdir(folder):
        log.error(f"[시간별] 폴더 없음: {folder}")
        return

    hhmm_key = os.path.basename(folder)
    date_key = os.path.basename(os.path.dirname(folder))
    date_fmt = f"{date_key[:4]}-{date_key[4:6]}-{date_key[6:]}" if len(date_key) == 8 else date_key
    time_fmt = f"{hhmm_key[:2]}:{hhmm_key[2:]}"                  if len(hhmm_key) == 4  else hhmm_key
    now_str  = f"{date_fmt} {time_fmt}"

    # ── xlsx 목록 ─────────────────────────────
    xlsx_files = sorted(glob.glob(os.path.join(folder, "MarketData_*.xlsx")))
    log.info(f"[시간별] {len(xlsx_files)}개 xlsx 발견 → Claude Code 분석 준비")
    if not xlsx_files:
        log.warning("[시간별] 분석할 xlsx 없음")
        return

    # ── 포트폴리오 JSON ───────────────────────
    summary_files = sorted(glob.glob(os.path.join(folder, "collection_summary_*.json")))
    meta_list: list[dict] = []
    if summary_files:
        with open(summary_files[-1], "r", encoding="utf-8") as f:
            meta_list = json.load(f)
    portfolio_json = build_portfolio_json(meta_list)

    # ── 헤더 메시지 즉시 발송 (분석 중 알림) ──
    n = len(xlsx_files)
    header_msg = (
        f"📡 <b>장중 시황 분석 시작</b>\n"
        f"📅 {date_fmt}  ⏰ {time_fmt}  종목 {n}개\n"
        f"🤖 Claude Code 분석 중... 잠시 후 리포트가 발송됩니다."
    )
    if send:
        send_telegram(header_msg)

    # ── Claude Code 분석 ──────────────────────
    log.info(f"[Claude Code] 시간별 분석 요청 ({len(xlsx_files)}개 파일)...")
    ai_text = run_claude_analysis(
        folder, xlsx_files, HOURLY_SYSTEM_PROMPT, portfolio_json, now_str, tag="시간별"
    )

    # ── 최종 메시지 발송 ──────────────────────
    full_msg = (
        f"📊 <b>장중 시황 리포트</b>\n"
        f"📅 {date_fmt}  ⏰ {time_fmt}\n"
        f"{'─' * 30}\n\n"
        f"{ai_text}"
    )

    if send:
        ok = send_telegram(full_msg, parse_mode="")
        log.info(f"[텔레그램] {'✅ 발송 완료' if ok else '❌ 발송 실패'}")
    else:
        print("\n" + "=" * 60 + "\n[미리보기]\n" + "=" * 60)
        print(re.sub(r"<[^>]+>", "", full_msg))


# ═══════════════════════════════════════════════════════════════════
# ■■■ 일별 분석 모드 ■■■
# ═══════════════════════════════════════════════════════════════════

def get_daily_folder(target_date: Optional[str] = None) -> str:
    """output_YYYYMMDD 폴더 경로 반환."""
    d = target_date or datetime.now().strftime("%Y%m%d")
    return os.path.join(DAILY_DATA_ROOT, "MarketData", d)


def run_daily_analyzer(
    target_date: Optional[str] = None,
    send: bool = True,
) -> None:
    """
    일별 분석 진입점. 수집기 완료 후 즉시 호출.
    1) 해당 폴더를 작업 디렉터리로 Claude Code CLI 호출 (xlsx 직접 읽음)
    2) 포트폴리오 JSON + 일별 프롬프트 → Claude Code 분석
    3) 텔레그램 발송
    """
    # ── 폴더 확인 ─────────────────────────────
    folder = get_daily_folder(target_date)
    if not os.path.isdir(folder):
        log.error(f"[일별] 폴더 없음: {folder}")
        return

    date_key = target_date or datetime.now().strftime("%Y%m%d")
    date_fmt = f"{date_key[:4]}-{date_key[4:6]}-{date_key[6:]}"
    now_str  = f"{date_fmt} 장 마감"

    # ── xlsx 목록 ─────────────────────────────
    xlsx_files = sorted(glob.glob(os.path.join(folder, "*.xlsx")))
    log.info(f"[일별] {len(xlsx_files)}개 xlsx 발견")
    if not xlsx_files:
        log.warning("[일별] 분석할 xlsx 없음")
        return

    # ── 포트폴리오 JSON ───────────────────────
    summary_files = sorted(glob.glob(os.path.join(folder, "daily_summary_*.json")))
    meta_list: list[dict] = []
    if summary_files:
        with open(summary_files[-1], "r", encoding="utf-8") as f:
            meta_list = json.load(f)
    portfolio_json = build_portfolio_json(meta_list)

    # ── 헤더 메시지 즉시 발송 ─────────────────
    n = len(xlsx_files)
    header_msg = (
        f"🌙 <b>장 마감 종합 리포트 생성 중</b>\n"
        f"📅 {date_fmt}  종목 {n}개\n"
        f"{'═' * 30}\n"
        f"🤖 Claude Code 심층 분석 중... 잠시 후 전체 리포트가 발송됩니다."
    )
    if send:
        send_telegram(header_msg)

    # ── Claude Code 분석 (폴더 내 전체 xlsx를 한 번에 통합 분석) ──
    log.info(f"[Claude Code] 일별 통합 분석 요청 ({len(xlsx_files)}개 파일)...")
    ai_text = run_claude_analysis(
        folder, xlsx_files, DAILY_SYSTEM_PROMPT, portfolio_json, now_str, tag="일별"
    )

    # ── 최종 메시지 발송 ──────────────────────
    full_msg = (
        f"🌙 <b>장 마감 종합 리밸런싱 전략 리포트</b>\n"
        f"📅 {date_fmt}\n"
        f"{'═' * 30}\n\n"
        f"{ai_text}"
    )

    if send:
        ok = send_telegram(full_msg, parse_mode="")
        log.info(f"[텔레그램] {'✅ 발송 완료' if ok else '❌ 발송 실패'}")
    else:
        print("\n" + "=" * 60 + "\n[미리보기]\n" + "=" * 60)
        print(re.sub(r"<[^>]+>", "", full_msg))


# ──────────────────────────────────────────────
# [추가] 주간 포트폴리오 리스크 관제용 프롬프트
# ──────────────────────────────────────────────
def build_weekly_portfolio_prompt(target_date):
    year = target_date[:4]
    month = target_date[4:6]
    week = str((int(target_date[6:8]) - 1) // 7 + 1)
    
    prompt = f"""
    당신은 전체 포트폴리오의 리스크를 총괄하는 냉정하고 기계적인 퀀트 시스템입니다. 주관적 서술과 미사여구를 철저히 배제하고, 포트폴리오 레벨의 구조적 리스크와 개별 종목의 핵심 이탈 시그널만을 극도로 간결하게 출력하십시오.

    [입력 데이터]
    - 첨부된 엑셀 파일들은 이번 주의 종목별 주간 수급, 40주선 이격도, 신용잔고 증감 데이터를 포함하고 있습니다.

    [출력 조건 및 분석 지침]
    1. 모든 문장은 명사형(음/함/됨/임)으로 종결하며, 수치와 팩트 위주로 건조하게 작성할 것.
    2. 개별 종목 나열을 지양하고, '주간 수급 쏠림·이탈이 현재 포트폴리오 섹터 비중에 미치는 구조적 리스크'를 우선적으로 관조할 것.
    3. 개별 종목은 VWAP 하회, 신용잔고 급증, 40주선 이격도 과다 등 '치명적 다이버전스'가 발생한 경우에만 1줄로 요약하여 경고할 것.
    4. 알고리즘이 즉각적으로 반영해야 할 포트폴리오 비중 조절(현금 확보, 섹터 축소 등) 지침을 명확히 제시할 것.
    5. 금리·유가·환율 등 매크로 지표는 별도 체계에서 판단하므로 이 리포트에서 판단·언급·요구하지 말 것. 데이터에 없는 항목의 부재도 언급하지 말고, 첨부 파일의 수급·기술 지표(주간 수급, 이격도, 신용잔고, 거래대금, 외인소진율, VWAP)만으로 판단할 것.

    [출력 포맷]
    📅 [{year}년 {month}월 {week}주차] 포트폴리오 리스크 관제 리포트
    ══════════════════════════════
    🌍 [수급 & 포트폴리오 총평]
    - (외인·기관 주간 수급 방향과 신용잔고 추이가 포트폴리오 전반에 미치는 영향)
    - (반도체 등 특정 섹터의 과열이나 메이저 수급 이탈 현상 요약)

    🚨 [핵심 리스크 종목] (※ 시스템 경고가 발생한 종목만 압축 출력)
    - [종목명]: (이격도/VWAP/신용잔고 기준 이탈 사유 1줄 요약)

    🎯 [시스템 실행 전략]
    - (현금 비중 N% 확대, 특정 섹터 비중 축소 등 기계적 액션 지시)

──────────
[실행 지침 — Claude Code CLI 전용]
· 지금 작업 디렉터리에 있는 xlsx 파일들이 이번 분석의 유일한 근거 데이터다. pandas 등으로 각 시트를 직접 읽어 분석하라.
· 웹 검색이나 사전 지식으로 종목명·가격·지표를 추정하지 말고, 반드시 첨부된 파일의 실제 값만 사용하라.
· 최종 출력은 위 [출력 포맷]에 맞춘 리포트 본문 그 자체만 출력하라. 코드블록, 설명, "다음은 분석입니다" 같은 전후 문구를 절대 붙이지 마라. 이 출력이 그대로 텔레그램 메시지로 전송된다.
""" + COMPACT_STYLE_RULES
    return prompt

def run_weekly_analyzer(target_date: Optional[str] = None, send: bool = True):
    """
    주간 포트폴리오 리스크 분석기 (종목별 수급/이격도/신용잔고).
    WeeklyData/YYYYMMDD/ 폴더의 xlsx → Claude Code 분석 → 텔레그램 발송
    """
    if target_date is None:
        target_date = datetime.now().strftime("%Y%m%d")

    base_dir   = os.path.dirname(os.path.abspath(__file__))
    target_dir = os.path.join(base_dir, WEEKLY_DATA_ROOT, target_date)
    files      = glob.glob(os.path.join(target_dir, "*.xlsx"))

    # fallback: 날짜 하위폴더가 없으면 WeeklyData 루트에서 탐색
    # (★ 버그 수정: 기존 코드는 fallback으로 찾은 파일 경로와 달리 target_dir을
    #  갱신하지 않아, 분석 작업 디렉터리가 실제 파일 위치와 어긋날 수 있었다.)
    if not files:
        fallback_dir = os.path.join(base_dir, WEEKLY_DATA_ROOT)
        files = glob.glob(os.path.join(fallback_dir, "*.xlsx"))
        if files:
            target_dir = fallback_dir

    if not files:
        log.error(f"[주간] 분석할 xlsx 없음: {target_dir}")
        return

    date_fmt = f"{target_date[:4]}-{target_date[4:6]}-{target_date[6:]}"
    log.info(f"[주간] {len(files)}개 xlsx 발견")

    # 헤더 메시지 즉시 발송
    if send:
        send_telegram(
            f"📅 <b>주간 리스크 관제 리포트 생성 중</b>\n"
            f"🗓 {date_fmt}  종목 {len(files)}개\n"
            f"🤖 Claude Code 분석 중... 잠시 후 리포트가 발송됩니다."
        )

    system_prompt = build_weekly_portfolio_prompt(target_date)
    log.info(f"[Claude Code] 주간 분석 요청 ({len(files)}개 파일)...")
    ai_text = run_claude_analysis(
        target_dir, files, system_prompt, "{}", f"{date_fmt} 주간", tag="주간"
    )

    full_msg = (
        f"📅 <b>주간 포트폴리오 리스크 관제 리포트</b>\n"
        f"🗓 {date_fmt}\n"
        f"{'═' * 30}\n\n"
        f"{ai_text}"
    )

    if send:
        ok = send_telegram(full_msg, parse_mode="")
        log.info(f"[텔레그램] {'✅ 발송 완료' if ok else '❌ 발송 실패'}")
    else:
        print("\n" + "=" * 60 + "\n[미리보기]\n" + "=" * 60)
        print(re.sub(r"<[^>]+>", "", full_msg))

    log.info(f"[Claude Code] 주간 분석 완료")

# ──────────────────────────────────────────────
# [추가] 월간매크로 분석기
# ──────────────────────────────────────────────
def run_monthly_analyzer(target_date: Optional[str] = None, send: bool = True):
    """
    월간 매크로 & 섹터 퀀트 전략 분석기.
    월간_매크로섹터분석.py(구 주간_섹터_동향분석.py) 가 생성한
    MonthlyMacroData/YYYYMMDD/ 폴더의
      Macro_Monthly_YYYYMMDD.xlsx  (L1~L3 통합 매크로)
      SectorAction_YYYYMMDD.xlsx   (L4 섹터 대응)
    두 파일을 Claude Code CLI로 읽혀 MONTHLY_SYSTEM_PROMPT(섹터 퀀트 전략)로
    분석한 뒤 텔레그램 발송.
    """
    if target_date is None:
        target_date = datetime.now().strftime("%Y%m%d")

    base_dir   = os.path.dirname(os.path.abspath(__file__))
    target_dir = os.path.join(base_dir, MONTHLY_DATA_ROOT, target_date)

    # ── Macro_Monthly + SectorAction 두 파일 우선 탐색 ───────────
    macro_files  = sorted(glob.glob(os.path.join(target_dir, "Macro_Monthly_*.xlsx")))
    sector_files = sorted(glob.glob(os.path.join(target_dir, "SectorAction_*.xlsx")))

    # fallback: 날짜 하위폴더가 없으면 MonthlyMacroData 루트에서 최신 파일 탐색
    if not macro_files or not sector_files:
        fallback_dir = os.path.join(base_dir, MONTHLY_DATA_ROOT)
        date_dirs = sorted(glob.glob(os.path.join(fallback_dir, "[0-9]" * 8)), reverse=True)
        for d in date_dirs:
            macro_files  = sorted(glob.glob(os.path.join(d, "Macro_Monthly_*.xlsx")))
            sector_files = sorted(glob.glob(os.path.join(d, "SectorAction_*.xlsx")))
            if macro_files and sector_files:
                target_dir = d
                log.warning(f"[월간] 날짜 폴더 없음 → 최신 폴더 사용: {d}")
                break

    if not macro_files:
        log.error(f"[월간] Macro_Monthly_*.xlsx 없음: {target_dir}")
        return
    if not sector_files:
        log.error(f"[월간] SectorAction_*.xlsx 없음: {target_dir}")
        return

    # 각각 가장 최신 파일 1개씩 사용
    files = [macro_files[-1], sector_files[-1]]

    ym_fmt   = f"{target_date[:4]}년 {int(target_date[4:6])}월"
    date_fmt = f"{target_date[:4]}-{target_date[4:6]}-{target_date[6:]}"

    log.info(f"[월간] 분석 대상:")
    log.info(f"  Macro  : {os.path.basename(files[0])}")
    log.info(f"  Sector : {os.path.basename(files[1])}")

    # 헤더 메시지 즉시 발송
    if send:
        send_telegram(
            f"📅 <b>월간 매크로 &amp; 섹터 퀀트 전략 리포트 생성 중</b>\n"
            f"🗓 {ym_fmt}\n"
            f"📊 Macro_Monthly + SectorAction 2파일 교차 분석\n"
            f"🤖 Claude Code 심층 분석 중... 잠시 후 리포트가 발송됩니다."
        )

    log.info(f"[Claude Code] 월간(매크로/섹터) 분석 요청 ({len(files)}개 파일)...")
    ai_text = run_claude_analysis(
        target_dir, files, MONTHLY_SYSTEM_PROMPT, "{}", f"{ym_fmt}", tag="월간"
    )

    full_msg = (
        f"📅 <b>월간 매크로 &amp; 섹터 퀀트 전략 리포트</b>\n"
        f"🗓 {ym_fmt}\n"
        f"{'═' * 30}\n\n"
        f"{ai_text}"
    )

    if send:
        ok = send_telegram(full_msg, parse_mode="")
        log.info(f"[텔레그램] {'✅ 발송 완료' if ok else '❌ 발송 실패'}")
    else:
        print("\n" + "=" * 60 + "\n[미리보기]\n" + "=" * 60)
        print(re.sub(r"<[^>]+>", "", full_msg))


# ═══════════════════════════════════════════════════════════════════
# ■■■ 하위 호환 래퍼 ■■■
# ═══════════════════════════════════════════════════════════════════
def run_analyzer(
    target_date: Optional[str] = None,
    target_hhmm: Optional[str] = None,
    send: bool = True,
    mode: str = "hourly",
) -> None:
    if mode == "daily":
        run_daily_analyzer(target_date=target_date, send=send)
    elif mode == "weekly":
        run_weekly_analyzer(target_date=target_date, send=send)
    elif mode == "monthly":
        run_monthly_analyzer(target_date=target_date, send=send)
    else:
        run_hourly_analyzer(target_date=target_date, target_hhmm=target_hhmm, send=send)

# ──────────────────────────────────────────────
# CLI 진입점
# ──────────────────────────────────────────────
if __name__ == "__main__":
    """
    사용법 (직접 실행):
      py 한투API_텔레그램분석.py                           # 오늘 최신 시간별 분석 & 즉시 발송
      py 한투API_텔레그램분석.py --daily                   # 오늘 일별 분석 & 즉시 발송
      py 한투API_텔레그램분석.py --weekly                  # 금주 주간 분석 & 즉시 발송
      py 한투API_텔레그램분석.py --monthly                 # 이번 달 월간 분석 & 즉시 발송
      py 한투API_텔레그램분석.py --preview                 # 발송 없이 콘솔 미리보기
      py 한투API_텔레그램분석.py --daily   --preview       # 일별 콘솔 미리보기
      py 한투API_텔레그램분석.py --weekly  --preview       # 주간 콘솔 미리보기
      py 한투API_텔레그램분석.py --monthly --preview       # 월간 콘솔 미리보기

    스케줄러 subprocess 호출 방식:
      py 한투API_텔레그램분석.py --date 20260604 --hhmm 0939
      py 한투API_텔레그램분석.py --daily   --date 20260604
      py 한투API_텔레그램분석.py --weekly  --date 20260607
      py 한투API_텔레그램분석.py --monthly --date 20260608
    """
    args    = sys.argv[1:]
    preview = "--preview" in args
    daily   = "--daily"   in args
    weekly  = "--weekly"  in args
    monthly = "--monthly" in args

    # --date YYYYMMDD
    t_date = None
    if "--date" in args:
        idx = args.index("--date")
        if idx + 1 < len(args) and re.match(r"^\d{8}$", args[idx + 1]):
            t_date = args[idx + 1]
    if t_date is None:
        pos_dates = [a for a in args if re.match(r"^\d{8}$", a)]
        t_date = pos_dates[0] if pos_dates else None

    # --hhmm HHMM
    t_hhmm = None
    if "--hhmm" in args:
        idx = args.index("--hhmm")
        if idx + 1 < len(args) and re.match(r"^\d{4}$", args[idx + 1]):
            t_hhmm = args[idx + 1]

    # 실행 분기
    if monthly:
        run_monthly_analyzer(target_date=t_date, send=not preview)
    elif weekly:
        run_weekly_analyzer(target_date=t_date, send=not preview)
    elif daily:
        run_daily_analyzer(target_date=t_date, send=not preview)
    else:
        run_hourly_analyzer(target_date=t_date, target_hhmm=t_hhmm, send=not preview)
