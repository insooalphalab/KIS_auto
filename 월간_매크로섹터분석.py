"""
월간 매크로/섹터 통합 파이프라인 v3.0  (구 주간_섹터_동향분석.py)
===========================================
★ 매월 첫째주 일요일 1회 실행 (Windows 작업 스케줄러: setup_tasks.ps1)

[2026-09 최적화 변경]
  · 파일명을 주간_섹터_동향분석.py → 월간_매크로섹터분석.py 로 변경했다.
    금리·신용스프레드·반도체 수출입 등 L1~L3 지표는 월 단위로 갱신되는
    성격이 강해, 기존 "매주 실행"에서 "매월 1회 실행"으로 주기를 늦췄다.
    (종목별 주간 수급/이격도는 별도로 한투API_주간데이터.py가 토요일마다
    계속 담당하므로 데이터 손실은 없다.)
  · 저장 폴더를 WeeklyData/ → MonthlyMacroData/ 로 분리해, 종목별 주간
    데이터(WeeklyData/)와 매크로/섹터 월간 데이터가 더 이상 섞이지 않는다.
  · 산출 파일명도 Macro_Weekly_*.xlsx → Macro_Monthly_*.xlsx 로 변경했다.

[3파일 구조]
  Macro_Monthly_YYYYMMDD.xlsx  ← L1(위험관리) + L2(경기방향) + L3(펀더멘털) 통합
  SectorAction_YYYYMMDD.xlsx   ← 섹터 대응 (Bottom-up)

[사용 API]
  · FRED API   (fredapi)  : HY스프레드, 금리 역전, BBB스프레드, 일본10Y
  · yfinance              : 미국 금리, VIX, 구리/WTI, 달러인덱스, 환율
  · 한국투자증권 오픈API  : ETF 일봉, 수급, KOSPI 지수
  · 네이버 검색 API       : 섹터별 뉴스
  · 관세청 공공데이터 API : 반도체(HS8542) 수출입
  · DRAMeXchange 스크래핑: DRAM/NAND 현물가

[저장 경로]
  (스크립트 위치)\\MonthlyMacroData\\YYYYMMDD\\   ← 프로젝트 폴더 어디로 옮겨도 동작하도록
                                                 절대경로 하드코딩을 상대경로로 변경했다.

[필요 패키지]
  pip install requests pandas openpyxl yfinance fredapi beautifulsoup4 lxml
"""

import os, re, sys, json, time, io
import urllib.request, urllib.parse
import requests
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta
from collections import Counter
from concurrent.futures import ThreadPoolExecutor
import traceback

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8')

try:
    import pandas as pd
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError as e:
    print(f"❌ 필수 패키지 없음: {e}\npip install pandas openpyxl 실행 후 재시도")
    sys.exit(1)

try:
    import yfinance as yf
except ImportError:
    print("❌ yfinance 없음: pip install yfinance")
    sys.exit(1)

try:
    from fredapi import Fred
    FRED_AVAILABLE = True
except ImportError:
    print("⚠️  fredapi 없음. L1 FRED 데이터 건너뜀. (pip install fredapi)")
    FRED_AVAILABLE = False

try:
    from bs4 import BeautifulSoup
    BS4_AVAILABLE = True
except ImportError:
    print("⚠️  beautifulsoup4 없음. L3 DRAM 스크래핑 건너뜀. (pip install beautifulsoup4 lxml)")
    BS4_AVAILABLE = False

# ============================================================
# ★ API 인증키 (기존 파일에서 그대로 유지)
# ============================================================
from urllib.parse import quote as _urlquote
from kis_config import require   # 시크릿은 .env 에서 로드

FRED_API_KEY   = require("FRED_API_KEY")

CUSTOMS_API_KEY      = require("CUSTOMS_API_KEY")
CUSTOMS_ENCODING_KEY = _urlquote(CUSTOMS_API_KEY, safe="")   # 디코딩 키 → URL 인코딩 키

APP_KEY    = require("KIS_APP_KEY")
APP_SECRET = require("KIS_APP_SECRET")
BASE_URL   = "https://openapi.koreainvestment.com:9443"

NAVER_CLIENT_ID     = require("NAVER_CLIENT_ID")
NAVER_CLIENT_SECRET = require("NAVER_CLIENT_SECRET")

# ============================================================
# 저장 경로
# ============================================================
NOW      = datetime.now()
RUN_DATE = NOW.strftime("%Y%m%d")
# ★ 2026-09 변경: 절대경로 하드코딩 제거 → 스크립트 위치 기준 상대경로.
#   폴더명도 WeeklyData(종목별 주간데이터와 혼동)에서 MonthlyMacroData로 분리.
OUTPUT_DIR = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "MonthlyMacroData", RUN_DATE
)

# ============================================================
# 리스크 임계치
# ============================================================
THRESHOLD = {
    "JP_10Y_RISK":       1.5,
    "US_10Y_HIGH":       5.0,
    "HY_SPREAD_CAUTION": 4.0,
    "HY_SPREAD_RISKOFF": 6.0,
    "HY_MOM_ALERT":      0.5,
    "VIX_CAUTION":       20.0,
    "VIX_RISKOFF":       30.0,
    "WTI_MOM_SPIKE":     15.0,
    "SEMI_MOM_DECLINE":  -10.0,
}

# ============================================================
# 섹터 정의 (주간 파이프라인과 동일)
# ============================================================
SECTORS = [
    {"id": 1,  "name": "반도체",  "sheet_name": "01_반도체",
     "etfs": [("396500", "TIGER반도체TOP10"), ("395160", "KODEX AI반도체TOP2플러스")],
     "news_query": "반도체 | HBM | 메모리 | 파운드리 | 소부장 | 팹 | 출하량 | 가동률"},
    {"id": 2,  "name": "자동차",  "sheet_name": "02_자동차",
     "etfs": [("091170", "KODEX자동차"), ("466930", "SOL자동차TOP3플러스")],
     "news_query": "자동차 | 완성차 | 전기차 | 자동차부품 | 북미판매 | 내수판매"},
    {"id": 3,  "name": "조선",    "sheet_name": "03_조선",
     "etfs": [("449020", "KODEX조선TOP10"), ("494670", "TIGER조선TOP10")],
     "news_query": "조선 | 선박 | 수주 | LNG선 | 신조선가 | 조선업"},
    {"id": 4,  "name": "정유화학","sheet_name": "04_정유화학",
     "etfs": [("117460", "KODEX에너지화학"), ("139250", "TIGER200에너지화학")],
     "news_query": "정유 | 석유화학 | 정제마진 | 에틸렌 | 나프타 | 석화"},
    {"id": 5,  "name": "2차전지", "sheet_name": "05_2차전지",
     "etfs": [("305540", "KODEX2차전지산업"), ("364980", "TIGER2차전지TOP10")],
     "news_query": "배터리 | 양극재 | 전기차배터리 | 셀업체 | LFP | 2차전지"},
    {"id": 6,  "name": "화장품",  "sheet_name": "06_화장품",
     "etfs": [("228790", "TIGER화장품"), ("479840", "SOL화장품TOP3플러스")],
     "news_query": "화장품 | K뷰티 | ODM | 인디브랜드 | 화장품수출"},
    {"id": 7,  "name": "전력기기","sheet_name": "07_전력기기",
     "etfs": [("487240", "KODEX AI전력핵심설비"), ("488000", "TIGER코리아AI전력기기TOP3플러스")],
     "news_query": "전력기기 | 변압기 | 송배전 | 데이터센터전력 | 전력망"},
    {"id": 8,  "name": "음식료",  "sheet_name": "08_음식료",
     "etfs": [("102970", "KODEX필수소비재"), ("453630", "KODEX미국S&P500필수소비재")],
     "news_query": "라면 | K푸드 | 가공식품 | 음식료 | 식품수출"},
    {"id": 9,  "name": "인바운드","sheet_name": "09_인바운드",
     "etfs": [("228800", "TIGER여행레저"), ("388280", "RISE K엔터&여행레저")],
     "news_query": "면세점 | 외국인관광객 | 카지노 | 방한관광 | 인바운드"},
    {"id": 10, "name": "바이오",  "sheet_name": "10_바이오",
     "etfs": [("364970", "TIGER바이오TOP10"), ("244580", "KODEX바이오")],
     "news_query": "바이오 | 위탁생산 | CMO | CDMO | 바이오시밀러 | FDA"},
]

KOSPI_ETF_CODE = "069500"

# 퀀트 파라미터
RS_MA_PERIOD   = 20
SUPPLY_WEEKS   = 4
OVERHEAT_PCT   = 35.0
NEGLECT_PCT    = 5.0
TURNOVER_MULT  = 2.0
SUPPLY_NORM_WINDOW = 20
NEWS_SURGE_MULT    = 1.5
NEWS_MIN_COUNT     = 5

ONE_MONTH_AGO = NOW - timedelta(days=30)

# 뉴스 필터
EXCLUDE_KEYWORDS = ["[특징주]", "특징주", "장마감", "상한가", "하한가"]
STOP_WORDS = [
    "기자","뉴스","올해","지난해","최근","위해","대비","가장","통해","경우",
    "있다","있는","있습니다","있어","있고","없는","없다",
    "등을","등이","등은","등으로","따라","대한","대해","중심으로",
    "위한","하는","하며","하고","하다","합니다",
    "크게","많이","매우","특히","현재","다양한","관련","이번",
    "이후","전년","동기","분기","만원","억원","달러","수준","기록",
]

# ============================================================
# 공통 스타일
# ============================================================
THIN   = Side(border_style="thin", color="000000")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
C_CTR  = Alignment(horizontal="center", vertical="center", wrap_text=True)
C_LFT  = Alignment(horizontal="left",   vertical="center", wrap_text=True)
B_FONT = Font(name="맑은 고딕", size=10)

# 파일별 헤더 색상
HC = {
    "L1": "8B0000",   # 진적색 - 위험관리
    "L2": "1A3A5C",   # 진남색 - 매크로 방향
    "L3": "006400",   # 진녹색 - 펀더멘털
    "L4": "4A235A",   # 보라   - 시장 대응
    "summary": "2E4057",
    "signal":  "4A235A",
}

POS      = PatternFill("solid", start_color="E8F4EA")
NEG      = PatternFill("solid", start_color="FDE8E8")
WARN     = PatternFill("solid", start_color="FFE699")
GREEN    = PatternFill("solid", start_color="C6EFCE")
BUY      = PatternFill("solid", start_color="C6EFCE")
SELL     = PatternFill("solid", start_color="FFC7CE")
OHT      = PatternFill("solid", start_color="FF9999")
TRN      = PatternFill("solid", start_color="99CCFF")
RISKOFF  = PatternFill("solid", start_color="8B0000")
ALERT    = PatternFill("solid", start_color="FF4500")


def hdr(ws, headers, widths, level_key="L1"):
    fill = PatternFill("solid", start_color=HC.get(level_key, "1F497D"))
    ws.append(headers)
    r = ws.max_row
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for cell in ws[r]:
        cell.font      = Font(bold=True, color="FFFFFF", name="맑은 고딕", size=10)
        cell.fill      = fill
        cell.alignment = C_CTR
        cell.border    = BORDER


def srow(ws, ri, n, num_fmts=None):
    for c in range(1, n + 1):
        cell = ws.cell(ri, c)
        cell.font      = B_FONT
        cell.border    = BORDER
        cell.alignment = C_CTR
        if num_fmts and c in num_fmts:
            cell.number_format = num_fmts[c]


def color_mom(cell):
    if isinstance(cell.value, (int, float)):
        cell.fill = POS if cell.value >= 0 else NEG


def to_float(v, default=0.0):
    try:
        return float(str(v).replace(",", "").replace("+", "") or default)
    except:
        return default


def safe_int(v):
    try:
        return int(str(v).replace(",", "").replace("+", "") or 0)
    except:
        return 0


def clean_html(t):
    t = re.sub(r'<[^>]+>', '', t)
    for o, n in [('&quot;', '"'), ('&apos;', "'"), ('&amp;', '&'), ('&lt;', '<'), ('&gt;', '>')]:
        t = t.replace(o, n)
    return t.strip()


def iso_week(date_str):
    d = datetime.strptime(date_str, "%Y%m%d")
    y, w, _ = d.isocalendar()
    return f"{y}-W{w:02d}"


def fmt_won(n):
    if n is None:
        return "N/A"
    return f"{n / 100_000_000:+.2f}억"


def add_title_row(ws, title_text, n_cols, color_key="L1"):
    ws.append([title_text])
    r = ws.max_row
    ws[f"A{r}"].font = Font(bold=True, size=12, name="맑은 고딕", color="FFFFFF")
    ws[f"A{r}"].fill = PatternFill("solid", start_color=HC.get(color_key, "2E4057"))
    ws[f"A{r}"].alignment = C_CTR
    if n_cols > 1:
        ws.merge_cells(f"A{r}:{get_column_letter(n_cols)}{r}")
    ws.row_dimensions[r].height = 22
    ws.append([])


# ============================================================
# ========== LEVEL 1: 위험 관리 (Risk-Off 판독기) =============
# ============================================================

def collect_l1_data():
    """
    L1 수집: HY스프레드(FRED), 미국/일본 장기금리(yfinance+FRED), VIX(yfinance)
    반환: dict with keys 'rates', 'credit'
    """
    print("\n" + "=" * 60)
    print("  [L1] 위험 관리 데이터 수집")
    print("=" * 60)

    end   = datetime.today()
    start = end - timedelta(days=365 * 2)
    s_str = start.strftime("%Y-%m-%d")

    # ── yfinance: 미국 금리, VIX ────────────────────────────────
    yf_tickers = {
        "^TNX":   ("미국10Y금리(%)",  "US_10Y"),
        "^TYX":   ("미국30Y금리(%)",  "US_30Y"),
        "^VIX":   ("VIX공포지수",     "VIX"),
        "KRW=X":  ("원/달러환율",     "USDKRW"),
    }
    rates = {}
    for ticker, (kor, key) in yf_tickers.items():
        try:
            df = yf.download(ticker, start=s_str, end=end.strftime("%Y-%m-%d"),
                             progress=False, auto_adjust=True)
            if df.empty:
                continue
            if isinstance(df.columns, pd.MultiIndex):
                col = [c for c in df.columns if c[0] == 'Close']
                series = df[col[0]] if col else None
            else:
                series = df['Close']
            if series is None:
                continue
            series.index = pd.to_datetime(series.index)
            # 주간(W-FRI) 마지막값
            weekly = series.resample('W-FRI').last().dropna()
            rates[key] = {"kor": kor, "weekly": weekly}
            print(f"  ✅ {ticker} ({kor}): {len(weekly)}주")
        except Exception as e:
            print(f"  ❌ {ticker}: {e}")
        time.sleep(0.3)

    # ── FRED: 일본 10Y, HY스프레드, 금리역전 ────────────────────
    credit = {}
    if FRED_AVAILABLE:
        fred_series = {
            "IRLTLT01JPM156N": ("일본10Y금리(%)",    "JP_10Y"),
            "BAMLH0A0HYM2":    ("HY스프레드(%)",     "HY"),
            "T10Y2Y":          ("10Y-2Y역전(%)",     "INVERT_10Y2Y"),
            "T10Y3M":          ("10Y-3M역전(%)",     "INVERT_10Y3M"),
            "BAMLC0A4CBBB":    ("BBB스프레드(%)",    "BBB"),
        }
        try:
            fred = Fred(api_key=FRED_API_KEY)
            for sid, (kor, key) in fred_series.items():
                try:
                    s = fred.get_series(sid, observation_start=s_str)
                    s.index = pd.to_datetime(s.index)
                    weekly = s.resample("W-FRI").last().dropna()
                    credit[key] = {"kor": kor, "weekly": weekly}
                    print(f"  ✅ FRED {sid} ({kor}): {len(weekly)}주")
                    time.sleep(0.4)
                except Exception as e:
                    print(f"  ❌ FRED {sid}: {e}")
        except Exception as e:
            print(f"  ❌ FRED 전체 오류: {e}")

    return {"rates": rates, "credit": credit}


def _latest(d, n=0):
    """weekly Series에서 최신 n번째 값을 안전하게 반환 (내림차순 정렬 보장)"""
    try:
        s = d["weekly"].sort_index(ascending=False)
        if len(s) <= n:
            return None
        return round(float(s.iloc[n]), 4)
    except:
        return None


def build_macro_excel(l1_data, l2_data, l3_customs_df, l3_dram_data, output_dir):
    """
    L1+L2+L3 통합 엑셀
    시트 구성:
      00_종합판정     : Risk-Off 게이트키퍼
      01_금리_신용    : 미국/일본 금리, VIX, HY스프레드 (L1)
      02_경기방향     : 구리/WTI/DXY/환율 방향 신호 (L2)
      03_원자재추이   : 26주 원자재·지수 추이 (L2)
      04_반도체수출입 : 관세청 HS8542 (L3)
      05_DRAM현물가   : DRAMeXchange 스크래핑 (L3)
    """
    wb = openpyxl.Workbook()

    rates  = l1_data.get("rates", {})
    credit = l1_data.get("credit", {})

    # ── 공통: 주간 날짜 인덱스 (최신→과거, 26주) ─────────────────
    all_idx = set()
    for d in list(rates.values()) + list(credit.values()) + list(l2_data.values()):
        for dt in d["weekly"].index:
            all_idx.add(dt)
    idx = sorted(all_idx, reverse=True)[:26]

    # ── 시트1: 종합 판정 (Risk Gauge) ────────────────────────────
    ws1 = wb.active
    ws1.title = "00_종합판정"
    add_title_row(ws1, f"매크로 종합 판정 (L1 위험관리)   기준: {NOW.strftime('%Y-%m-%d')}", 5, "L1")

    hdr(ws1, ["기준주", "지표", "현재값", "임계치", "판정"], [14, 22, 14, 22, 30], "L1")

    # ★ 수정: _latest()로 정렬된 최신값 사용
    us10y  = _latest(rates["US_10Y"])  if "US_10Y"  in rates  else None
    us30y  = _latest(rates["US_30Y"])  if "US_30Y"  in rates  else None
    jp10y  = _latest(credit["JP_10Y"]) if "JP_10Y"  in credit else None
    vix    = _latest(rates["VIX"])     if "VIX"     in rates  else None
    hy     = _latest(credit["HY"])     if "HY"      in credit else None
    week_str = idx[0].strftime("%Y-%m-%d") if idx else NOW.strftime("%Y-%m-%d")

    def judge_cell(ws, ri, val, thr_warn, thr_risk, higher_is_bad=True):
        cell = ws.cell(ri, 5)
        if val is None:
            cell.value = "-"
            return "-", 0
        if higher_is_bad:
            if val >= thr_risk:
                cell.value = "🔴 위험"
                cell.fill  = RISKOFF
                cell.font  = Font(bold=True, color="FFFFFF", name="맑은 고딕", size=10)
                return "🔴", 2
            elif val >= thr_warn:
                cell.value = "🔶 주의"
                cell.fill  = WARN
                return "🔶", 1
            else:
                cell.value = "✅ 정상"
                cell.fill  = GREEN
                return "✅", 0
        else:
            if val <= thr_risk:
                cell.value = "🔴 위험"
                cell.fill  = RISKOFF
                cell.font  = Font(bold=True, color="FFFFFF", name="맑은 고딕", size=10)
                return "🔴", 2
            elif val <= thr_warn:
                cell.value = "🔶 주의"
                cell.fill  = WARN
                return "🔶", 1
            else:
                cell.value = "✅ 정상"
                cell.fill  = GREEN
                return "✅", 0

    score = 0
    check_items = [
        (week_str, "미국10Y금리(%)", us10y,  4.5, THRESHOLD["US_10Y_HIGH"], True),
        (week_str, "미국30Y금리(%)", us30y,  4.8, 5.2,                      True),
        (week_str, "일본10Y금리(%)", jp10y,  1.0, THRESHOLD["JP_10Y_RISK"], True),
        (week_str, "VIX공포지수",    vix,    THRESHOLD["VIX_CAUTION"], THRESHOLD["VIX_RISKOFF"], True),
        (week_str, "HY스프레드(%)",  hy,     THRESHOLD["HY_SPREAD_CAUTION"], THRESHOLD["HY_SPREAD_RISKOFF"], True),
    ]

    for c_week, c_name, c_val, c_warn, c_risk, c_hib in check_items:
        thr_str = f"주의>{c_warn} / 위험>{c_risk}" if c_hib else f"주의<{c_warn} / 위험<{c_risk}"
        ws1.append([c_week, c_name, c_val if c_val is not None else "N/A", thr_str, ""])
        ri = ws1.max_row
        srow(ws1, ri, 5)
        _, s = judge_cell(ws1, ri, c_val, c_warn, c_risk, c_hib)
        score += s

    ws1.append([])
    if score >= 4:
        verdict = "🔴 RISK-OFF  → 현금 비중 확대, 방어주 전환"
        v_fill, v_color = RISKOFF, "FFFFFF"
    elif score >= 2:
        verdict = "🔶 CAUTION   → 방어적 운용, 신규 매수 축소"
        v_fill, v_color = WARN, "000000"
    else:
        verdict = "✅ RISK-ON   → 정상 운용"
        v_fill, v_color = GREEN, "000000"

    ws1.append(["", "【 종합 판정 】", verdict, f"리스크점수: {score}점", ""])
    ri = ws1.max_row
    for c in range(1, 6):
        cell = ws1.cell(ri, c)
        cell.fill = v_fill
        cell.font = Font(bold=True, color=v_color, name="맑은 고딕", size=11)
        cell.border = BORDER
        cell.alignment = C_CTR

    ws1.freeze_panes = "A4"
    print("  ✅ 시트 00_종합판정 완료")

    # ── 시트2: L1 금리·VIX·신용 통합 추이 ───────────────────────
    ws2 = wb.create_sheet("01_금리_신용")
    add_title_row(ws2, "미국/일본 금리 · VIX · HY스프레드 주간 추이 (26주)", 11, "L1")

    # ★ 수정: 미국30Y 컬럼 추가
    cols_2 = ["기준주",
              "미국10Y(%)", "미국30Y(%)", "일본10Y(%)", "VIX",
              "HY스프레드(%)", "10Y-2Y역전(%)", "BBB스프레드(%)",
              "미국10Y WoW(%p)", "미국30Y WoW(%p)", "HY WoW(%p)"]
    hdr(ws2, cols_2, [14, 12, 12, 12, 10, 14, 14, 14, 16, 16, 12], "L1")

    # 룩업 테이블 구성 (dt → 값)
    def make_lookup(d_dict, keys):
        lk = {k: {} for k in keys}
        for k in keys:
            if k in d_dict:
                s = d_dict[k]["weekly"].sort_index(ascending=False)
                for dt in idx:
                    # idx는 내림차순이므로 s.get()으로 조회
                    try:
                        lk[k][dt] = float(s.get(dt, float('nan')))
                        if pd.isna(lk[k][dt]):
                            lk[k][dt] = None
                    except:
                        lk[k][dt] = None
        return lk

    rate_lk  = make_lookup(rates,  ["US_10Y", "US_30Y", "VIX", "USDKRW"])
    cred_lk  = make_lookup(credit, ["JP_10Y", "HY", "INVERT_10Y2Y", "BBB"])

    def gv_r(k, dt): return rate_lk.get(k, {}).get(dt)
    def gv_c(k, dt): return cred_lk.get(k, {}).get(dt)
    def rv(v, d=3): return round(float(v), d) if v is not None else None

    for i, dt in enumerate(idx):
        us10 = rv(gv_r("US_10Y", dt))
        us30 = rv(gv_r("US_30Y", dt))
        jp10 = rv(gv_c("JP_10Y", dt))
        vx   = rv(gv_r("VIX",   dt), 2)
        hy_v = rv(gv_c("HY",    dt), 2)
        i2   = rv(gv_c("INVERT_10Y2Y", dt), 2)
        bbb  = rv(gv_c("BBB",   dt), 2)

        # WoW: 이번주 - 전주 (idx[i+1]이 전주)
        us10_wow = us30_wow = hy_wow = None
        if i < len(idx) - 1:
            dt_p = idx[i + 1]
            us10_p = rv(gv_r("US_10Y", dt_p))
            us30_p = rv(gv_r("US_30Y", dt_p))
            hy_p   = rv(gv_c("HY",    dt_p), 2)
            if us10 is not None and us10_p is not None:
                us10_wow = round(us10 - us10_p, 3)
            if us30 is not None and us30_p is not None:
                us30_wow = round(us30 - us30_p, 3)
            if hy_v is not None and hy_p is not None:
                hy_wow = round(hy_v - hy_p, 3)

        ws2.append([dt.strftime("%Y-%m-%d"),
                    us10, us30, jp10, vx,
                    hy_v, i2, bbb,
                    us10_wow, us30_wow, hy_wow])
        ri = ws2.max_row
        srow(ws2, ri, 11, {2:"0.000",3:"0.000",4:"0.000",5:"0.00",
                           6:"0.00",7:"0.00",8:"0.00",
                           9:"+0.000;-0.000;0.000",
                           10:"+0.000;-0.000;0.000",
                           11:"+0.000;-0.000;0.000"})
        # 임계치 색상
        for ci, val, tw, tr in [(2,us10,4.5,THRESHOLD["US_10Y_HIGH"]),
                                 (3,us30,4.8,5.2),
                                 (4,jp10,1.0,THRESHOLD["JP_10Y_RISK"]),
                                 (5,vx,  THRESHOLD["VIX_CAUTION"],THRESHOLD["VIX_RISKOFF"]),
                                 (6,hy_v,THRESHOLD["HY_SPREAD_CAUTION"],THRESHOLD["HY_SPREAD_RISKOFF"])]:
            c = ws2.cell(ri, ci)
            if isinstance(val, float):
                if val >= tr:
                    c.fill = RISKOFF
                    c.font = Font(bold=True, color="FFFFFF", name="맑은 고딕", size=10)
                elif val >= tw:
                    c.fill = WARN
        for ci in [9, 10, 11]:
            color_mom(ws2.cell(ri, ci))

    ws2.freeze_panes = "A4"
    print("  ✅ 시트 01_금리_신용 완료")

    # ── 시트3: L2 경기 방향 신호 ─────────────────────────────────
    ws3 = wb.create_sheet("02_경기방향")
    add_title_row(ws3, f"L2 경기 방향성 종합   기준: {NOW.strftime('%Y-%m-%d')}", 5, "L2")
    hdr(ws3, ["지표", "현재값", "전주대비(%)", "4주전대비(%)", "방향성 판단"], [20, 14, 14, 14, 30], "L2")

    l2_lk = make_lookup(l2_data, ["COPPER", "WTI", "GOLD", "DXY", "USDKRW", "KOSPI", "SPX"])

    macro_checks = [
        ("구리($/파운드)",  "COPPER"),
        ("WTI원유($/배럴)", "WTI"),
        ("금($/온스)",      "GOLD"),
        ("달러인덱스(DXY)", "DXY"),
        ("원/달러환율",     "USDKRW"),
        ("KOSPI",           "KOSPI"),
        ("S&P500",          "SPX"),
    ]

    for kor, key in macro_checks:
        # ★ Fix 3: idx(내림차순 정렬된 주간 날짜)로 n번째 주 값 조회
        def get_nth(k, n, _idx=idx, _lk=l2_lk):
            if n >= len(_idx):
                return None
            v = _lk.get(k, {}).get(_idx[n])
            return round(float(v), 4) if v is not None else None

        v0 = get_nth(key, 0)   # 이번주 (최신)
        v1 = get_nth(key, 1)   # 전주
        v4 = get_nth(key, 4)   # 4주전

        # ★ Fix 3: 전주대비/4주전대비 모두 % 변화율로 통일
        wow_pct = round((v0 - v1) / abs(v1) * 100, 2) if (v0 is not None and v1) else None
        m4w_pct = round((v0 - v4) / abs(v4) * 100, 2) if (v0 is not None and v4) else None

        if key == "COPPER":
            if m4w_pct is not None:
                verdict = "🟢 경기확장" if m4w_pct > 5 else ("🔴 경기둔화" if m4w_pct < -5 else "⬜ 중립")
            else:
                verdict = "-"
        elif key == "WTI":
            if m4w_pct is not None:
                verdict = "🔶 원가압박" if m4w_pct > THRESHOLD["WTI_MOM_SPIKE"] else ("🟡 원유상승" if m4w_pct > 5 else "✅ 안정")
            else:
                verdict = "-"
        elif key in ("DXY", "USDKRW"):
            if wow_pct is not None:
                verdict = "🔴 달러강세(신흥국압박)" if wow_pct > 1.0 else ("🟢 달러약세" if wow_pct < -1.0 else "⬜ 안정")
            else:
                verdict = "-"
        else:
            verdict = "-"

        ws3.append([kor, v0, wow_pct, m4w_pct, verdict])
        ri = ws3.max_row
        srow(ws3, ri, 5, {2: "#,##0.000", 3: "+0.00;-0.00;0.00", 4: "+0.00;-0.00;0.00"})
        color_mom(ws3.cell(ri, 3))
        color_mom(ws3.cell(ri, 4))

    ws3.freeze_panes = "A4"
    print("  ✅ 시트 02_경기방향 완료")

    # ── 시트4: L2 원자재·지수 26주 추이 ──────────────────────────
    ws4 = wb.create_sheet("03_원자재추이")
    add_title_row(ws4, "원자재 & 지수 주간 추이 (26주)", 10, "L2")
    cols_4 = ["기준주", "구리($/lb)", "WTI($/배럴)", "금($/온스)",
              "DXY", "원/달러", "KOSPI", "S&P500",
              "구리 WoW%", "WTI WoW%"]
    hdr(ws4, cols_4, [14, 12, 14, 12, 10, 12, 10, 10, 12, 12], "L2")

    for i, dt in enumerate(idx):
        def g2(k, _dt=dt, _lk=l2_lk):
            v = _lk.get(k, {}).get(_dt)
            return round(float(v), 3) if v is not None else None

        cu = g2("COPPER"); wt = g2("WTI"); go = g2("GOLD")
        dx = g2("DXY");    kr = g2("USDKRW")
        ks = g2("KOSPI");  sp = g2("SPX")

        cu_wow = wt_wow = None
        if i < len(idx) - 1:
            dt_p = idx[i + 1]
            cu_p = l2_lk.get("COPPER", {}).get(dt_p)
            wt_p = l2_lk.get("WTI",    {}).get(dt_p)
            if cu and cu_p:
                cu_wow = round((cu - float(cu_p)) / float(cu_p) * 100, 2)
            if wt and wt_p:
                wt_wow = round((wt - float(wt_p)) / float(wt_p) * 100, 2)

        ws4.append([dt.strftime("%Y-%m-%d"),
                    cu, wt, go, dx, kr, ks, sp, cu_wow, wt_wow])
        ri = ws4.max_row
        srow(ws4, ri, 10, {2:"0.000",3:"0.00",4:"#,##0.0",5:"0.00",
                           6:"#,##0.0",7:"#,##0",8:"#,##0.0",
                           9:"+0.00;-0.00;0.00",10:"+0.00;-0.00;0.00"})
        color_mom(ws4.cell(ri, 9))
        color_mom(ws4.cell(ri, 10))

    ws4.freeze_panes = "A4"
    print("  ✅ 시트 03_원자재추이 완료")

    # ── 시트5: L3 반도체 수출입 (관세청) ─────────────────────────
    ws5 = wb.create_sheet("04_반도체수출입")
    add_title_row(ws5, f"L3 반도체 수출입 동향 (관세청 HS8542/8541)   기준: {NOW.strftime('%Y-%m-%d')}", 8, "L3")
    hdrs5 = ["기준월", "HS코드", "품목명", "수출(달러)", "수입(달러)", "무역수지(달러)", "수출MoM(%)", "펀더멘털신호"]
    hdr(ws5, hdrs5, [12, 10, 22, 18, 18, 18, 14, 24], "L3")

    if l3_customs_df is not None and not l3_customs_df.empty:
        for _, row in l3_customs_df.iterrows():
            mom = row.get("수출MoM(%)")
            if mom is None or (isinstance(mom, float) and pd.isna(mom)):
                sig = "-"
            elif mom <= THRESHOLD["SEMI_MOM_DECLINE"]:
                sig = "🔴 수출 급감"
            elif mom < 0:
                sig = "🔶 수출 감소"
            elif mom >= 20:
                sig = "🚀 수출 급증"
            else:
                sig = "✅ 수출 증가"
            ws5.append([row.get("기준월"), row.get("HS코드"), row.get("품목명"),
                        row.get("수출(달러)"), row.get("수입(달러)"), row.get("무역수지(달러)"),
                        mom, sig])
            ri = ws5.max_row
            srow(ws5, ri, 8, {4:"#,##0",5:"#,##0",6:"#,##0;[Red]-#,##0",7:"0.00"})
            color_mom(ws5.cell(ri, 7))
            sig_c = ws5.cell(ri, 8)
            if "급감" in sig:
                sig_c.fill = RISKOFF
                sig_c.font = Font(bold=True, color="FFFFFF", name="맑은 고딕", size=10)
            elif "감소" in sig: sig_c.fill = WARN
            elif "급증" in sig or "증가" in sig: sig_c.fill = GREEN
    else:
        ws5.append(["관세청 데이터 없음 - API 재확인 필요"])

    ws5.freeze_panes = "A4"
    print("  ✅ 시트 04_반도체수출입 완료")

    # ── 시트6: L3 DRAM/NAND 현물가 ───────────────────────────────
    ws6 = wb.create_sheet("05_DRAM현물가")
    add_title_row(ws6, f"L3 DRAM/NAND/GDDR 현물가 (DRAMeXchange)   수집: {NOW.strftime('%Y-%m-%d %H:%M')}", 8, "L3")

    dx_title_font = Font(size=12, bold=True, color="003399")
    dx_time_font  = Font(size=10, italic=True, color="CC0000")
    dx_hdr_fill   = PatternFill("solid", start_color="D9E1F2")
    dx_hdr_font   = Font(bold=True)

    if l3_dram_data:
        for sec_title, sec_data in l3_dram_data.items():
            last_upd = sec_data.get("last_update", "")
            df_d     = sec_data.get("df", pd.DataFrame())
            if df_d.empty:
                continue
            ws6.append([sec_title, last_upd])
            tr = ws6.max_row
            ws6.cell(tr, 1).font      = dx_title_font
            ws6.cell(tr, 2).font      = dx_time_font
            ws6.cell(tr, 2).alignment = C_LFT
            cols_d = df_d.columns.tolist()
            ws6.append(cols_d)
            hr = ws6.max_row
            for ci in range(len(cols_d)):
                c = ws6.cell(hr, ci + 1)
                c.fill = dx_hdr_fill; c.font = dx_hdr_font; c.alignment = C_CTR
            for _, drow in df_d.iterrows():
                ws6.append(drow.values.tolist())
                ri = ws6.max_row
                for ci in range(1, len(cols_d) + 1):
                    ws6.cell(ri, ci).font      = B_FONT
                    ws6.cell(ri, ci).alignment = C_CTR
                for ci, col_name in enumerate(cols_d, 1):
                    if "Change" in str(col_name):
                        cv = str(ws6.cell(ri, ci).value or "")
                        if cv.startswith("-"):
                            ws6.cell(ri, ci).fill = NEG
                        elif cv and cv != "0" and not cv.startswith("-"):
                            ws6.cell(ri, ci).fill = POS
            ws6.append([]); ws6.append([])
        ws6.column_dimensions['A'].width = 42
        for col in ['B','C','D','E','F','G']:
            ws6.column_dimensions[col].width = 16
    else:
        ws6.append(["DRAMeXchange 데이터 없음 - 네트워크 또는 사이트 구조 변경 확인"])

    print("  ✅ 시트 05_DRAM현물가 완료")

    # ── 시트 정리 + 저장 ─────────────────────────────────────────
    if "Sheet" in wb.sheetnames and len(wb.sheetnames) > 1:
        wb.remove(wb["Sheet"])

    path = os.path.join(output_dir, f"Macro_Monthly_{RUN_DATE}.xlsx")
    wb.save(path)
    print(f"  💾 Macro 통합 저장: {path}")
    return path


# ============================================================
# ========== LEVEL 2 데이터 수집 ==============================
# ============================================================

def collect_l2_data():
    """L2 수집: 구리/WTI/금/DXY/환율/KOSPI/S&P500 (yfinance 주간)"""
    print("\n" + "=" * 60)
    print("  [L2] 경기 방향성 데이터 수집")
    print("=" * 60)

    end   = datetime.today()
    start = end - timedelta(days=365 * 2)
    s_str = start.strftime("%Y-%m-%d")

    yf_tickers = {
        "HG=F":     ("구리($/파운드)",   "COPPER"),
        "CL=F":     ("WTI원유($/배럴)",  "WTI"),
        "GC=F":     ("금($/온스)",       "GOLD"),
        "DX-Y.NYB": ("달러인덱스(DXY)", "DXY"),
        "KRW=X":    ("원/달러환율",      "USDKRW"),
        "^KS11":    ("KOSPI",            "KOSPI"),
        "SPY":      ("S&P500",           "SPX"),
    }

    results = {}
    for ticker, (kor, key) in yf_tickers.items():
        try:
            df = yf.download(ticker, start=s_str, end=end.strftime("%Y-%m-%d"),
                             progress=False, auto_adjust=True)
            if df.empty:
                continue
            if isinstance(df.columns, pd.MultiIndex):
                col = [c for c in df.columns if c[0] == 'Close']
                series = df[col[0]] if col else None
            else:
                series = df['Close']
            if series is None:
                continue
            series.index = pd.to_datetime(series.index)
            weekly = series.resample('W-FRI').last().dropna()
            results[key] = {"kor": kor, "weekly": weekly}
            print(f"  ✅ {ticker} ({kor}): {len(weekly)}주")
        except Exception as e:
            print(f"  ❌ {ticker}: {e}")
        time.sleep(0.3)

    return results


# ============================================================
# ========== LEVEL 3 데이터 수집 ==============================
# ============================================================

def collect_l3_customs():
    """관세청 API: 반도체(HS8542/8541) - 1년 이내 범위 조회 (API 제한)"""
    print("\n  [L3] 관세청 API - 반도체 수출입(HS8542) 수집...")

    BASE = "https://apis.data.go.kr/1220000/Itemtrade"
    hs_targets = {"8542": "전자집적회로(IC반도체)", "8541": "개별반도체소자"}

    today = datetime.today().replace(day=1)
    # ★ Fix 4: API 1년 제한 대응 → 11개월치만 요청 (안전 마진 1개월 확보)
    end_month   = (today - timedelta(days=30)).strftime("%Y%m")    # 전월 (확정치)
    start_month = (today - timedelta(days=335)).strftime("%Y%m")   # 약 11개월 전

    print(f"  수집 범위: {start_month} ~ {end_month}")

    def _fetch(hs):
        url = (f"{BASE}/getItemtradeList"
               f"?serviceKey={CUSTOMS_ENCODING_KEY}"
               f"&strtYymm={start_month}&endYymm={end_month}"
               f"&hsSgn={hs}&numOfRows=500&pageNo=1&type=xml")
        try:
            res = requests.get(url, timeout=30)
            if res.text.lstrip().startswith("{"):
                return [], f"JSON 오류: {res.text[:80]}"
            root = ET.fromstring(res.content)
            rc   = root.findtext(".//resultCode") or "00"
            if rc not in ("00", "0000", ""):
                return [], f"API 오류 {rc}: {root.findtext('.//resultMsg')}"
            return root.findall(".//item"), None
        except Exception as e:
            return [], str(e)

    all_rows = []
    for hs, hs_name in hs_targets.items():
        items, err = _fetch(hs)
        if err:
            print(f"  ❌ HS{hs}: {err}")
            continue
        bucket = {}
        for item in items:
            yr = item.findtext("year") or ""
            if yr == "총계" or not yr:
                continue
            ym = yr.replace(".", "")
            if ym not in bucket:
                bucket[ym] = {"exp": 0.0, "imp": 0.0, "bal": 0.0}
            bucket[ym]["exp"] += to_float(item.findtext("expDlr"))
            bucket[ym]["imp"] += to_float(item.findtext("impDlr"))
            bucket[ym]["bal"] += to_float(item.findtext("balPayments"))
        for ym, v in sorted(bucket.items(), reverse=True):
            all_rows.append({"기준월": ym, "HS코드": hs, "품목명": hs_name,
                             "수출(달러)": v["exp"], "수입(달러)": v["imp"],
                             "무역수지(달러)": v["bal"]})
        print(f"  ✅ HS{hs} ({hs_name}): {len(bucket)}개월")
        time.sleep(0.5)

    df = pd.DataFrame(all_rows)
    if df.empty:
        return df

    df["수출MoM(%)"] = None
    for hs in hs_targets:
        mask = df["HS코드"] == hs
        sub  = df[mask].sort_values("기준월", ascending=False)
        vals = sub["수출(달러)"].values
        moms = [None] + [
            round((vals[i - 1] - vals[i]) / vals[i] * 100, 2) if vals[i] > 0 else None
            for i in range(1, len(vals))
        ]
        for j, idx2 in enumerate(sub.index):
            df.at[idx2, "수출MoM(%)"] = moms[j]

    return df.reset_index(drop=True)


def collect_l3_dram():
    """DRAMeXchange 스크래핑: DRAM/NAND/GDDR 현물가"""
    if not BS4_AVAILABLE:
        print("  ⚠️  bs4 없음 → DRAM 스크래핑 건너뜀")
        return {}

    print("\n  [L3] DRAMeXchange 현물가 스크래핑...")
    url = "https://www.dramexchange.com/"
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                      "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
    }
    try:
        res  = requests.get(url, headers=headers, timeout=15)
        res.raise_for_status()
        soup = BeautifulSoup(res.text, 'lxml')
    except Exception as e:
        print(f"  ❌ DRAMeXchange 접속 실패: {e}")
        return {}

    sections = soup.find_all('div', class_='left_tab')
    result   = {}
    for sec in sections:
        title_node = sec.find(class_='title_left')
        if not title_node:
            continue
        title    = title_node.get_text(strip=True)
        time_node= sec.find(class_='tab_time')
        last_upd = ""
        if time_node:
            last_upd = time_node.get_text(strip=True).replace('Price Notice', '').strip()
        target_tbl = None
        for tbl in sec.find_all('table'):
            text = tbl.get_text(separator=' ')
            if "Item" in text and ("High" in text or "Average" in text):
                target_tbl = tbl
                break
        if not target_tbl:
            continue
        try:
            df = pd.read_html(io.StringIO(str(target_tbl)))[0]
            if df.iloc[0].astype(str).str.contains("Item", case=False).any():
                df.columns = df.iloc[0]
                df = df[1:].reset_index(drop=True)
            df.dropna(how='all', inplace=True)
            df = df.loc[:, df.notna().any()]
            if 'History' in df.columns:
                df.drop(columns=['History'], inplace=True)
            if df.empty or 'Item' not in df.columns:
                continue
            result[title] = {"last_update": last_upd, "df": df}
            print(f"  ✅ {title}: {len(df)}개 품목")
        except Exception as e:
            print(f"  ❌ {title} 파싱 오류: {e}")

    return result





# ============================================================
# ========== LEVEL 4: 시장 대응 (Bottom-up Action) ============
# ============================================================

# ── 한투 API 공통 ────────────────────────────────────────────
def get_token():
    res  = requests.post(f"{BASE_URL}/oauth2/tokenP",
                         headers={"content-type": "application/json"},
                         json={"grant_type": "client_credentials",
                               "appkey": APP_KEY.strip(),
                               "appsecret": APP_SECRET.strip()},
                         timeout=15)
    data = res.json()
    if "access_token" not in data:
        raise Exception(f"토큰 오류: {data}")
    print("  ✅ 한투 토큰 발급")
    return data["access_token"]


def make_hantoo_hdr(token, tr_id):
    return {
        "content-type":  "application/json; charset=utf-8",
        "authorization": f"Bearer {token}",
        "appkey":        APP_KEY.strip(),
        "appsecret":     APP_SECRET.strip(),
        "tr_id":         tr_id,
        "custtype":      "P",
    }


def safe_get(url, hdrs, params, retries=3, timeout=12):
    for i in range(retries):
        try:
            return requests.get(url, headers=hdrs, params=params, timeout=timeout)
        except Exception as e:
            if i < retries - 1:
                time.sleep(1.5 * (i + 1))
            else:
                raise


# ── ETF 일봉 수집 ────────────────────────────────────────────
def get_ohlcv(token, code):
    """
    ETF 일봉 수집.
    ★ 수정: 2구간 병렬 → 단일 구간 직렬 수집으로 날짜 연속성 보장.
      - RS MA20 계산 + 표시 30행 = 최소 50 영업일 필요
      - 안전 마진 포함해 최근 65 영업일(약 90 캘린더일) 단일 요청
      - 한투 API 1회 최대 100건 반환 → 65 영업일은 단일 호출로 충분
    """
    url   = f"{BASE_URL}/uapi/domestic-stock/v1/quotations/inquire-daily-itemchartprice"
    hdrs  = make_hantoo_hdr(token, "FHKST03010100")
    today = datetime.today()
    # 약 90 캘린더일 = 65 영업일 (주말+공휴일 여유분 포함)
    s_date = (today - timedelta(days=90)).strftime("%Y%m%d")
    e_date = today.strftime("%Y%m%d")

    try:
        r = safe_get(url, hdrs, {
            "FID_COND_MRKT_DIV_CODE": "J",
            "FID_INPUT_ISCD":         code,
            "FID_INPUT_DATE_1":       s_date,
            "FID_INPUT_DATE_2":       e_date,
            "FID_PERIOD_DIV_CODE":    "D",
            "FID_ORG_ADJ_PRC":        "0",
        })
        items = r.json().get("output2", [])
    except Exception as e:
        print(f"    ❌ get_ohlcv({code}) 오류: {e}")
        return []

    seen, rows = set(), []
    for item in items:
        d = item.get("stck_bsop_date")
        if not d or d in seen:
            continue
        seen.add(d)
        rows.append({
            "날짜":     d,
            "종가":     safe_int(item.get("stck_clpr", 0)),
            "거래대금": safe_int(item.get("acml_tr_pbmn", 0)),
        })

    rows.sort(key=lambda x: x["날짜"])
    return rows  # 최대 ~65행, 한투 100건 제한 안에서 연속 보장


# ── KOSPI 지수 수집 ──────────────────────────────────────────
def get_kospi_index(token):
    """
    KOSPI 지수 일봉 수집.
    ★ 수정: 단일 구간(90 캘린더일) 수집으로 날짜 연속성 보장.
    """
    url   = f"{BASE_URL}/uapi/domestic-stock/v1/quotations/inquire-daily-indexchartprice"
    hdrs  = make_hantoo_hdr(token, "FHKUP03500100")
    today = datetime.today()
    s_date = (today - timedelta(days=90)).strftime("%Y%m%d")
    e_date = today.strftime("%Y%m%d")

    result = {}
    try:
        r = safe_get(url, hdrs, {
            "FID_COND_MRKT_DIV_CODE": "U",
            "FID_INPUT_ISCD":         "0001",
            "FID_INPUT_DATE_1":       s_date,
            "FID_INPUT_DATE_2":       e_date,
            "FID_PERIOD_DIV_CODE":    "D",
        })
        for item in r.json().get("output2", []):
            d = item.get("stck_bsop_date", "")
            if not d:
                continue
            v = item.get("bstp_nmix_prpr") or item.get("stck_clpr", "0")
            try:
                result[d] = float(str(v).replace(",", ""))
            except:
                pass
    except Exception as e:
        print(f"    ❌ get_kospi_index 오류: {e}")

    # fallback: KOSPI ETF로 대체
    if len(result) < 10:
        rows = get_ohlcv(token, KOSPI_ETF_CODE)
        result = {r["날짜"]: float(r["종가"]) for r in rows if r["종가"] > 0}

    return result


# ── 수급 수집 ────────────────────────────────────────────────
def get_investor(token, code, ohlcv_rows):
    if not ohlcv_rows:
        return []
    url   = f"{BASE_URL}/uapi/domestic-stock/v1/quotations/investor-trade-by-stock-daily"
    hdrs  = make_hantoo_hdr(token, "FHPTJ04160001")
    dates = sorted(r["날짜"] for r in ohlcv_rows)
    valid, m_start, m_end = set(dates), dates[0], dates[-1]
    all_rows, seen, cursor = [], set(), m_end
    while True:
        try:
            res  = safe_get(url, hdrs, {
                "FID_COND_MRKT_DIV_CODE": "J", "FID_INPUT_ISCD": code,
                "FID_INPUT_DATE_1": cursor,    "FID_INPUT_DATE_2": m_start,
                "FID_ORG_ADJ_PRC": "0",        "FID_ETC_CLS_CODE": "1",
            })
            data = res.json()
        except:
            break
        if data.get("rt_cd") != "0":
            msg = data.get("msg1", "")
            if ("TIME" in msg.upper() or "15:40" in msg) and cursor == m_end:
                cursor = (datetime.strptime(cursor, "%Y%m%d") - timedelta(days=1)).strftime("%Y%m%d")
                continue
            break
        output2 = data.get("output2", [])
        if not output2:
            break
        oldest = None
        for item in output2:
            d = item.get("stck_bsop_date", "")
            if not d:
                continue
            oldest = d
            if d in seen or d not in valid:
                continue
            seen.add(d)
            all_rows.append({
                "날짜": d,
                "외국인": safe_int(item.get("frgn_ntby_tr_pbmn", 0)),
                "기관합계": safe_int(item.get("orgn_ntby_tr_pbmn", 0)),
                "연기금": safe_int(item.get("ivtr_ntby_tr_pbmn", 0)),
                "금융투자": safe_int(item.get("scrt_ntby_tr_pbmn", 0)),
            })
        if not oldest or oldest <= m_start:
            break
        cursor = (datetime.strptime(oldest, "%Y%m%d") - timedelta(days=1)).strftime("%Y%m%d")
        if cursor < m_start:
            break
        time.sleep(0.2)
    all_rows.sort(key=lambda x: x["날짜"], reverse=True)
    return all_rows


# ── 퀀트 연산 ────────────────────────────────────────────────
def merge_etf(etf_data_list):
    dp = {}
    for ed in etf_data_list:
        for r in ed["ohlcv"]:
            d = r["날짜"]
            if d not in dp:
                dp[d] = {"종가_sum": 0, "거래대금": 0, "cnt": 0}
            dp[d]["종가_sum"]  += r["종가"]
            dp[d]["거래대금"]  += r.get("거래대금", 0)
            dp[d]["cnt"]       += 1
    m_ohlcv = [{"날짜": d, "종가": round(v["종가_sum"] / v["cnt"]),
                "거래대금": v["거래대금"]}
               for d, v in sorted(dp.items()) if v["cnt"] > 0]
    di = {}
    for ed in etf_data_list:
        for r in ed["inv"]:
            d = r["날짜"]
            if d not in di:
                di[d] = {"외국인": 0, "기관합계": 0, "연기금": 0, "금융투자": 0}
            for k in ("외국인", "기관합계", "연기금", "금융투자"):
                di[d][k] += r.get(k, 0)
    m_inv = [{"날짜": d, **v} for d, v in sorted(di.items(), reverse=True)]
    return m_ohlcv, m_inv


def compute_rs(merged_ohlcv, kospi_dict):
    if not merged_ohlcv or not kospi_dict:
        return [], "none", None, None
    series = []
    for r in sorted(merged_ohlcv, key=lambda x: x["날짜"]):
        d = r["날짜"]
        if d not in kospi_dict or kospi_dict[d] <= 0 or r["종가"] <= 0:
            continue
        series.append({"날짜": d, "RS": r["종가"] / kospi_dict[d]})
    if len(series) < RS_MA_PERIOD:
        return series, "none", None, None
    vals = [r["RS"] for r in series]
    for i, row in enumerate(series):
        win = vals[max(0, i - RS_MA_PERIOD + 1): i + 1]
        row["MA20"] = sum(win) / len(win)
        row["위치"] = "above" if row["RS"] >= row["MA20"] else "below"
    crossover = "none"
    if len(series) >= 2:
        prev, curr = series[-2], series[-1]
        if prev["위치"] == "below" and curr["위치"] == "above":
            crossover = "golden"
        elif prev["위치"] == "above" and curr["위치"] == "below":
            crossover = "dead"
    return series, crossover, series[-1]["RS"], series[-1]["MA20"]


def compute_turnover_ratio(etf_ohlcv_list, market_total_dict):
    date_sector = {}
    for ohlcv in etf_ohlcv_list:
        for r in ohlcv:
            d = r["날짜"]
            date_sector[d] = date_sector.get(d, 0) + r.get("거래대금", 0)
    if not date_sector or not market_total_dict:
        return [], None, None, "normal"
    daily_ratio = []
    for d in sorted(date_sector.keys()):
        mkt = market_total_dict.get(d, 0)
        sec = date_sector[d]
        pct = (sec / mkt * 100) if mkt > 0 else 0.0
        daily_ratio.append({"날짜": d, "섹터대금(억)": round(sec / 1e8, 2),
                            "10대섹터합산(억)": round(mkt / 1e8, 2), "점유율(%)": round(pct, 4)})
    if not daily_ratio:
        return [], None, None, "normal"
    latest_pct = daily_ratio[-1]["점유율(%)"]
    recent20   = [r["점유율(%)"] for r in daily_ratio[-20:]]
    avg4w      = round(sum(recent20) / len(recent20), 4) if recent20 else 0.0
    prev20     = [r["점유율(%)"] for r in daily_ratio[-40:-20]]
    avg_prev   = round(sum(prev20) / len(prev20), 4) if prev20 else avg4w
    if latest_pct >= OVERHEAT_PCT:
        t_signal = "overheat"
    elif latest_pct < NEGLECT_PCT and avg_prev < NEGLECT_PCT and avg4w >= avg_prev * TURNOVER_MULT:
        t_signal = "turnover"
    elif latest_pct < NEGLECT_PCT:
        t_signal = "neglect"
    else:
        t_signal = "normal"
    return daily_ratio, latest_pct, avg4w, t_signal


def compute_weekly_supply(inv_rows):
    if not inv_rows:
        return None, None, []
    week_map = {}
    for r in inv_rows:
        wk = iso_week(r["날짜"])
        if wk not in week_map:
            week_map[wk] = {"외국인": 0, "기관합계": 0, "연기금": 0, "금융투자": 0}
        for k in ("외국인", "기관합계", "연기금", "금융투자"):
            week_map[wk][k] += r.get(k, 0)
    recent = sorted(week_map.keys(), reverse=True)[:SUPPLY_WEEKS]
    weekly = []
    for wk in sorted(recent):
        d = week_map[wk]
        total = d["외국인"] + d["기관합계"]
        weekly.append({"주": wk, "외국인": d["외국인"], "기관합계": d["기관합계"],
                       "연기금": d["연기금"], "금융투자": d["금융투자"],
                       "합계(외+기관)": total, "방향": "▲" if total > 0 else "▼"})
    net4w = sum(w["합계(외+기관)"] for w in weekly)
    daily_abs = [abs(r.get("외국인", 0) + r.get("기관합계", 0)) for r in inv_rows[:SUPPLY_NORM_WINDOW]]
    if daily_abs:
        avg_d = sum(daily_abs) / len(daily_abs)
        total_vol = avg_d * 20
        intensity = round((net4w / total_vol) * 100, 2) if total_vol > 0 else 0.0
    else:
        intensity = None
    return net4w, intensity, weekly


def build_market_turnover(all_etf_ohlcv_list):
    total = {}
    for rows in all_etf_ohlcv_list:
        for r in rows:
            d = r["날짜"]
            total[d] = total.get(d, 0) + r.get("거래대금", 0)
    return total


# ── 네이버 뉴스 ──────────────────────────────────────────────
NAVER_RETRY = 3
NAVER_DELAY = 2.0


def _naver_page(enc_query, start_idx):
    url = (f"https://openapi.naver.com/v1/search/news.json"
           f"?query={enc_query}&display=100&start={start_idx}&sort=sim")
    for attempt in range(1, NAVER_RETRY + 1):
        try:
            req = urllib.request.Request(url)
            req.add_header("X-Naver-Client-Id",     NAVER_CLIENT_ID)
            req.add_header("X-Naver-Client-Secret", NAVER_CLIENT_SECRET)
            res = urllib.request.urlopen(req, timeout=15)
            return json.loads(res.read().decode("utf-8")).get("items", [])
        except Exception as e:
            if attempt < NAVER_RETRY:
                time.sleep(NAVER_DELAY * attempt)
            else:
                print(f"      ❌ 네이버 {start_idx}p 실패: {e}")
    return []


def fetch_news(sector):
    enc = urllib.parse.quote(sector["news_query"])
    raw = []
    for si in [1, 101]:
        for item in _naver_page(enc, si):
            try:
                pub = pd.to_datetime(item["pubDate"]).replace(tzinfo=None)
            except:
                continue
            if pub < ONE_MONTH_AGO:
                continue
            title = clean_html(item["title"])
            desc  = clean_html(item["description"])
            if len(desc) > 150:
                desc = desc[:147] + "..."
            if any(n in title or n in desc for n in EXCLUDE_KEYWORDS):
                continue
            raw.append({"발행일": pub, "기사 제목": title, "기사 요약": desc})
        time.sleep(0.5)
    if not raw:
        return pd.DataFrame(columns=["발행일", "기사 제목", "기사 요약"])
    df = (pd.DataFrame(raw)
          .drop_duplicates(subset=["기사 제목"])
          .sort_values("발행일", ascending=False)
          .reset_index(drop=True))
    df["발행일"] = df["발행일"].dt.strftime("%Y-%m-%d %H:%M")
    return df


def keywords(df, n=5):
    if df.empty:
        return "기사 없음"
    text  = " ".join(df["기사 제목"].tolist() + df["기사 요약"].tolist())
    words = re.findall(r"[가-힣A-Za-z]{2,}", text)
    clean = []
    for w in words:
        if w in STOP_WORDS:
            continue
        if w[-1] in "을를이가은는":
            w = w[:-1]
        if len(w) >= 2 and w not in STOP_WORDS:
            clean.append(w)
    return ", ".join(f"{w}({c})" for w, c in Counter(clean).most_common(n))


# ── 종합 시그널 ──────────────────────────────────────────────
def build_signal(crossover, t_signal, net4w, intensity,
                 news_surge=False, news_count=0):
    if intensity is not None:
        pos_flow = intensity > 0
        neg_flow = intensity < 0
    else:
        pos_flow = net4w is not None and net4w > 0
        neg_flow = net4w is not None and net4w < 0
    rs_down = crossover in ("dead", "none_below")
    if news_surge and (crossover == "dead" or rs_down) and neg_flow:
        int_str = f"{intensity:+.2f}%" if intensity is not None else ""
        return f"☠ 강제매도 (기사급증{news_count}건+RS하향+수급이탈{int_str})"
    if t_signal == "overheat":
        return "🔴🔴 과열+수급이탈 (매도 우선)" if neg_flow else "🔴 과열주의 (추격매수금지)"
    if t_signal == "turnover":
        return "💙 턴어라운드 (소외→폭증)"
    if crossover == "golden" and pos_flow:
        return f"★ BUY  (RS돌파+수급강도{intensity:+.2f}%)" if intensity is not None else "★ BUY  (RS돌파+수급유입)"
    if crossover == "dead" and neg_flow:
        return f"▼ SELL (RS이탈+수급강도{intensity:+.2f}%)" if intensity is not None else "▼ SELL (RS이탈+수급이탈)"
    if crossover == "golden" and not pos_flow:
        return "△ 관망 (RS돌파/수급미확인)"
    if crossover == "dead" and pos_flow:
        return "◇ 관망 (RS이탈/수급유입)"
    if t_signal == "neglect":
        return "⬜ 소외 (거래대금 극소)"
    if pos_flow:
        int_str = f" 강도{intensity:+.2f}%" if intensity is not None else ""
        return f"◈ 수급유입{int_str} (RS중립)"
    return "- 중립"


# ── L4 엑셀 저장 ─────────────────────────────────────────────
def save_l4_summary(wb, results):
    ws = wb.create_sheet("0_주간요약") if "0_주간요약" not in wb.sheetnames else wb["0_주간요약"]
    add_title_row(ws, f"L4 10대 섹터 주간 퀀트 시그널   기준: {NOW.strftime('%Y-%m-%d')}", 15, "L4")

    legend = ("★BUY: RS 20일선 상향돌파+수급유입 │ ▼SELL: RS하향이탈+수급이탈 │ "
              "🔴과열: 거래대금점유율≥35% │ 💙턴어라운드: 소외구간 폭증")
    ws.append([legend])
    ws[f"A{ws.max_row}"].font = Font(italic=True, name="맑은 고딕", size=9, color="555555")
    ws.merge_cells(f"A{ws.max_row}:O{ws.max_row}")
    ws.append([])

    cols = ["섹터", "대표 ETF",
            "RS(현재)", "RS MA20", "RS위치", "RS크로스",
            "거래대금점유율\n(10대섹터,%)", "현재-4주평균\n차이(%p)", "거래대금시그널",
            "수급강도(%)\n4주수급/일평균", "4주수급방향",
            "뉴스기사수\n(이번달)", "뉴스다이버전스",
            "뉴스키워드Top5", "★ 종합시그널"]
    widths = [12, 30, 12, 12, 10, 14, 16, 14, 14, 18, 16, 12, 14, 40, 34]
    hdr(ws, cols, widths, "L4")

    for r in results:
        rs_pos    = "▲ above" if (r["rs_now"] and r["ma20"] and r["rs_now"] >= r["ma20"]) else "▼ below"
        cross_lbl = {"golden": "🟢 상향돌파", "dead": "🔴 하향이탈", "none": "-"}[r["crossover"]]
        t_lbl     = {"overheat": "🔴 과열", "turnover": "💙 턴어라운드",
                     "neglect": "⬜ 소외", "normal": "-"}[r["t_signal"]]
        dirs      = [w["방향"] for w in sorted(r["weekly"], key=lambda x: x["주"])]
        emoji_str = "".join("🟥" if d == "▲" else "🟦" for d in dirs)
        if emoji_str == "🟥🟥🟥🟥":     trend = " (연속매수)"
        elif emoji_str == "🟦🟦🟦🟦":   trend = " (연속매도)"
        elif emoji_str.count("🟥") >= 3: trend = " (매수우위)"
        elif emoji_str.count("🟦") >= 3: trend = " (매도우위)"
        else:                             trend = " (혼조세)"
        intensity = r.get("intensity")
        news_cnt  = r.get("news_count", 0)
        nd_lbl    = "⚠ 급증" if r.get("news_surge") else "-"

        # ★ Fix 5: 현재점유율 - 4주평균 차이(%p)
        latest_pct = r["latest_pct"]
        avg4w_pct  = r["avg4w_pct"]
        pct_delta  = round(latest_pct - avg4w_pct, 2) if (latest_pct is not None and avg4w_pct is not None) else None

        row_data = [
            r["name"], " / ".join(n for _, n in r["etfs"]),
            round(r["rs_now"],  2) if r["rs_now"]  else "N/A",
            round(r["ma20"],    2) if r["ma20"]     else "N/A",
            rs_pos, cross_lbl,
            round(latest_pct, 2) if latest_pct is not None else "N/A",
            pct_delta if pct_delta is not None else "N/A",   # ← 차이(%p)
            t_lbl,
            round(intensity, 2) if intensity is not None else "N/A",
            emoji_str + trend,
            news_cnt, nd_lbl,
            r["news_kw"], r["signal"],
        ]
        ws.append(row_data)
        ri = ws.max_row
        srow(ws, ri, 15)

        # 색상 처리
        c5 = ws.cell(ri, 5)
        c5.fill = POS if "above" in str(c5.value) else NEG
        c5.alignment = C_CTR

        c6 = ws.cell(ri, 6)
        if "상향" in str(c6.value): c6.fill = BUY
        elif "하향" in str(c6.value): c6.fill = SELL
        c6.alignment = C_CTR

        pct_c = ws.cell(ri, 7)
        if isinstance(pct_c.value, (int, float)):
            pct_c.number_format = "0.00"
            if pct_c.value >= OVERHEAT_PCT:
                pct_c.fill = OHT
            elif pct_c.value < NEGLECT_PCT:
                pct_c.fill = WARN

        # ★ Fix 5: 차이(%p) 컬럼 색상 — 양수(현재>평균)=과열쪽, 음수(현재<평균)=위축
        delta_c = ws.cell(ri, 8)
        delta_c.number_format = "+0.00;-0.00;0.00"
        if isinstance(delta_c.value, (int, float)):
            if delta_c.value >= 5:
                delta_c.fill = OHT          # 평균 대비 5%p 이상 급등 → 과열 경고
            elif delta_c.value > 0:
                delta_c.fill = POS          # 평균 상회
            elif delta_c.value <= -5:
                delta_c.fill = WARN         # 평균 대비 급감 → 주의
            else:
                delta_c.fill = NEG          # 평균 하회
        delta_c.alignment = C_CTR

        c9 = ws.cell(ri, 9)
        if "과열" in str(c9.value):      c9.fill = OHT
        elif "턴어라운드" in str(c9.value): c9.fill = TRN
        elif "소외" in str(c9.value):    c9.fill = WARN
        c9.alignment = C_CTR

        sv = ws.cell(ri, 10)
        if isinstance(sv.value, (int, float)):
            sv.number_format = "+0.00;-0.00;0.00"
            if sv.value > 5:     sv.fill = BUY
            elif sv.value > 0:   sv.fill = POS
            elif sv.value < -5:  sv.fill = SELL
            else:                sv.fill = NEG
        sv.alignment = C_CTR

        nd_c = ws.cell(ri, 13)
        if "급증" in str(nd_c.value):
            nd_c.fill = WARN
            nd_c.font = Font(bold=True, name="맑은 고딕", size=10)
        nd_c.alignment = C_CTR

        sc = ws.cell(ri, 15)
        sc.font = Font(bold=True, name="맑은 고딕", size=10)
        sc.alignment = C_CTR
        sig_v = str(sc.value)
        if "강제매도" in sig_v:
            sc.fill = PatternFill("solid", start_color="8B0000")
            sc.font = Font(bold=True, color="FFFFFF", name="맑은 고딕", size=10)
        elif "BUY"       in sig_v: sc.fill = BUY
        elif "SELL"      in sig_v: sc.fill = SELL
        elif "과열"      in sig_v: sc.fill = OHT
        elif "턴어라운드" in sig_v: sc.fill = TRN
        elif "관망"      in sig_v: sc.fill = WARN

    ws.freeze_panes = "A5"


def save_l4_sector_sheet(wb, sector, etf_data_list,
                         rs_series, crossover, rs_now, ma20,
                         t_ratio, latest_pct, avg4w_pct, t_signal,
                         weekly, net4w, intensity,
                         news_df, news_count=0, news_surge=False):
    ws = wb.create_sheet(sector["sheet_name"])
    etf_str = " / ".join(f"{c}({n})" for c, n in sector["etfs"])
    ws.append([f"【 {sector['name']} 】  ETF: {etf_str}"])
    ws["A1"].font = Font(bold=True, size=11, name="맑은 고딕", color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", start_color="1A3A5C")
    ws["A1"].alignment = C_LFT
    ws.merge_cells("A1:J1")
    ws.row_dimensions[1].height = 20
    ws.append([])

    # RS 추이 (최근 30행 표시 — MA20 계산은 전체 사용)
    ws.append([f"▶ RS 상대강도   현재={round(rs_now, 4) if rs_now else 'N/A'}  "
               f"MA20={round(ma20, 4) if ma20 else 'N/A'}  "
               f"크로스={'🟢상향' if crossover == 'golden' else '🔴하향' if crossover == 'dead' else '-'}"])
    ws[f"A{ws.max_row}"].font = Font(bold=True, name="맑은 고딕", size=10)
    ws.merge_cells(f"A{ws.max_row}:J{ws.max_row}")
    hdr(ws, ["날짜", "RS값", "MA20", "RS위치", "vs MA20"], [14, 16, 16, 12, 16], "L4")
    for row in sorted(rs_series[-30:], key=lambda x: x["날짜"], reverse=True):
        ma   = row.get("MA20")
        diff = round(row["RS"] - ma, 6) if ma else None
        ws.append([row["날짜"], round(row["RS"], 6),
                   round(ma, 6) if ma else "", row.get("위치", ""), diff])
        ri = ws.max_row
        srow(ws, ri, 5, {2: "0.000000", 3: "0.000000", 5: "0.000000"})
        if diff is not None:
            ws.cell(ri, 5).fill = POS if diff >= 0 else NEG
        pos_c = ws.cell(ri, 4)
        pos_c.fill = POS if row.get("위치") == "above" else NEG
        pos_c.alignment = C_CTR
    ws.append([])

    # 거래대금 점유율 (최근 30행)
    ws.append([f"▶ 거래대금 점유율   최신={round(latest_pct, 2) if latest_pct else 'N/A'}%  "
               f"4주평균={round(avg4w_pct, 2) if avg4w_pct else 'N/A'}%  시그널={t_signal}"])
    ws[f"A{ws.max_row}"].font = Font(bold=True, name="맑은 고딕", size=10)
    ws.merge_cells(f"A{ws.max_row}:J{ws.max_row}")
    hdr(ws, ["날짜", "섹터거래대금(억)", "10대섹터합산(억)", "점유율(%)"], [14, 18, 20, 14], "L4")
    for row in sorted(t_ratio[-30:], key=lambda x: x["날짜"], reverse=True):
        ws.append([row["날짜"], row["섹터대금(억)"], row["10대섹터합산(억)"], row["점유율(%)"]])
        ri = ws.max_row
        srow(ws, ri, 4, {2: "#,##0.00", 3: "#,##0.00", 4: "0.0000"})
        pct = row["점유율(%)"]
        p_c = ws.cell(ri, 4)
        if pct >= OVERHEAT_PCT:  p_c.fill = OHT
        elif pct < NEGLECT_PCT:  p_c.fill = WARN
        else:                    p_c.fill = POS
    ws.append([])

    # 주간 수급
    int_str = f"{intensity:+.2f}%" if intensity is not None else "N/A"
    ws.append([f"▶ 주간 수급 (최근 4주)   4주합계={fmt_won(net4w)}   수급강도={int_str}"])
    ws[f"A{ws.max_row}"].font = Font(bold=True, name="맑은 고딕", size=10)
    ws.merge_cells(f"A{ws.max_row}:J{ws.max_row}")
    hdr(ws, ["주(ISO)", "외국인(천원)", "기관합계(천원)", "연기금(천원)", "금융투자(천원)", "합계(외+기관)", "방향"],
        [14, 18, 18, 15, 15, 20, 8], "L4")
    for wr in sorted(weekly, key=lambda x: x["주"], reverse=True):
        ws.append([wr["주"], wr["외국인"], wr["기관합계"], wr["연기금"], wr["금융투자"],
                   wr["합계(외+기관)"], wr["방향"]])
        ri = ws.max_row
        srow(ws, ri, 7, {c: "#,##0;[Red]-#,##0" for c in range(2, 7)})
        for ci in range(2, 7):
            c = ws.cell(ri, ci)
            if isinstance(c.value, (int, float)):
                c.fill = POS if c.value > 0 else NEG
        ws.cell(ri, 7).alignment = C_CTR
    ws.append([])

    # 뉴스 — 건수·키워드는 전체 기준, 엑셀 저장은 상위 20건만
    nd_mark = "  ⚠ 기사 급증!" if news_surge else ""
    ws.append([f"▶ 최근 1개월 뉴스   전체 {news_count}건{nd_mark}  (상위 20건 표시)"])
    ws[f"A{ws.max_row}"].font = Font(bold=True, name="맑은 고딕", size=10)
    hdr(ws, ["발행일", "기사 제목", "기사 요약"], [20, 70, 100], "L4")
    if not news_df.empty:
        for _, nr in news_df.head(20).iterrows():   # ★ 상위 20건만 저장
            ws.append([nr.get(h, "") for h in ["발행일", "기사 제목", "기사 요약"]])
            srow(ws, ws.max_row, 3)
    else:
        ws.append(["해당 기간 뉴스 없음", "", ""])

    ws.freeze_panes = "A2"


def collect_and_build_l4(output_dir):
    """L4 수집 + 저장"""
    print("\n" + "=" * 60)
    print("  [L4] 시장 대응 데이터 수집 (한투 + 네이버)")
    print("=" * 60)

    try:
        token = get_token()
    except Exception as e:
        print(f"  ❌ 한투 토큰 오류: {e}")
        return None

    print("\n  [벤치마크] KOSPI 지수 수집...")
    kospi_dict = get_kospi_index(token)
    time.sleep(0.5)

    wb      = openpyxl.Workbook()
    results = []

    # 1패스: ETF 데이터 수집
    print("\n  [1패스] 전체 섹터 ETF 수집...")
    all_sector_data = []
    all_etf_ohlcv   = []

    for sector in SECTORS:
        print(f"\n  [{sector['id']:02d}] {sector['name']}")
        etf_data_list = []
        for etf_code, etf_name in sector["etfs"]:
            ohlcv = get_ohlcv(token, etf_code)
            time.sleep(0.5)
            inv = get_investor(token, etf_code, ohlcv)
            time.sleep(0.5)
            etf_data_list.append({"code": etf_code, "name": etf_name,
                                  "ohlcv": ohlcv, "inv": inv})
            all_etf_ohlcv.append(ohlcv)
        all_sector_data.append({"sector": sector, "etf_data_list": etf_data_list})

    print("\n  [거래대금 분모] 전체 ETF 합산...")
    market_turnover = build_market_turnover(all_etf_ohlcv)

    # 2패스: 연산 + 저장
    print("\n  [2패스] 퀀트 연산 + 시트 저장...")
    for item in all_sector_data:
        sector        = item["sector"]
        etf_data_list = item["etf_data_list"]

        m_ohlcv, m_inv = merge_etf(etf_data_list)
        rs_series, crossover, rs_now, ma20 = compute_rs(m_ohlcv, kospi_dict)
        t_ratio, latest_pct, avg4w_pct, t_signal = compute_turnover_ratio(
            [ed["ohlcv"] for ed in etf_data_list], market_turnover)
        net4w, intensity, weekly = compute_weekly_supply(m_inv)

        print(f"  [뉴스] {sector['name']} 수집...")
        news_df   = fetch_news(sector)
        news_kw   = keywords(news_df)
        news_count = len(news_df) if not news_df.empty else 0
        news_surge = (news_count >= NEWS_MIN_COUNT and
                      news_count >= 10 * NEWS_SURGE_MULT)

        sig = build_signal(crossover, t_signal, net4w, intensity,
                           news_surge, news_count)

        results.append({
            "name": sector["name"], "etfs": sector["etfs"],
            "rs_now": rs_now, "ma20": ma20, "crossover": crossover,
            "latest_pct": latest_pct, "avg4w_pct": avg4w_pct, "t_signal": t_signal,
            "net4w": net4w, "intensity": intensity, "weekly": weekly,
            "news_kw": news_kw, "news_count": news_count,
            "news_surge": news_surge, "signal": sig,
        })

        save_l4_sector_sheet(wb, sector, etf_data_list,
                             rs_series, crossover, rs_now, ma20,
                             t_ratio, latest_pct, avg4w_pct, t_signal,
                             weekly, net4w, intensity,
                             news_df, news_count, news_surge)
        time.sleep(0.5)

    # 요약 시트 (맨 앞으로)
    save_l4_summary(wb, results)
    wb.move_sheet(wb["0_주간요약"], offset=-len(wb.sheetnames) + 1)

    if "Sheet" in wb.sheetnames and len(wb.sheetnames) > 1:
        wb.remove(wb["Sheet"])

    path = os.path.join(output_dir, f"SectorAction_{RUN_DATE}.xlsx")
    wb.save(path)
    print(f"  💾 L4 저장: {path}")

    # 시그널 요약 출력
    print("\n  [L4 시그널 요약]")
    for r in results:
        print(f"    {r['name']:10s}  {r['signal']}")

    return path


# ============================================================
# 메인
# ============================================================
def main():
    print("=" * 65)
    print("  월간 매크로/섹터 통합 파이프라인 v3.0")
    print("  Macro_Monthly (L1+L2+L3 통합)  |  SectorAction")
    print("=" * 65)
    print(f"  실행: {NOW.strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"  출력: {OUTPUT_DIR}")

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    saved = {}

    # ── 데이터 수집 ─────────────────────────────────────────────
    l1_data = {}
    l2_data = {}
    customs_df = pd.DataFrame()
    dram_data  = {}

    try:
        l1_data = collect_l1_data()
    except Exception as e:
        print(f"\n❌ L1 수집 오류: {e}"); traceback.print_exc()

    try:
        l2_data = collect_l2_data()
    except Exception as e:
        print(f"\n❌ L2 수집 오류: {e}"); traceback.print_exc()

    try:
        customs_df = collect_l3_customs()
    except Exception as e:
        print(f"\n❌ L3 관세청 수집 오류: {e}"); traceback.print_exc()

    try:
        dram_data = collect_l3_dram()
    except Exception as e:
        print(f"\n❌ L3 DRAM 수집 오류: {e}"); traceback.print_exc()

    # ── Macro 통합 엑셀 (L1+L2+L3) ──────────────────────────────
    try:
        saved["Macro"] = build_macro_excel(
            l1_data, l2_data, customs_df, dram_data, OUTPUT_DIR)
    except Exception as e:
        print(f"\n❌ Macro 통합 저장 오류: {e}"); traceback.print_exc()

    # ── L4: 시장 대응 ─────────────────────────────────────────
    try:
        saved["L4"] = collect_and_build_l4(OUTPUT_DIR)
    except Exception as e:
        print(f"\n❌ L4 오류: {e}"); traceback.print_exc()

    # ── 완료 요약 ─────────────────────────────────────────────
    print("\n" + "=" * 65)
    print("  ✅ 월간 매크로/섹터 파이프라인 완료")
    print(f"  📁 저장 위치: {OUTPUT_DIR}")
    print()
    for level, path in saved.items():
        if path:
            print(f"  {level}: {os.path.basename(path)}")
        else:
            print(f"  {level}: ❌ 생성 실패")
    print("=" * 65)


if __name__ == "__main__":
    main()
