"""
한국투자증권 오픈API - 종목 데이터 수집 및 대차 병합 프로그램 (경량화판 v2)
=============================================
기능:
  1. 일봉 데이터 6개월치 → 엑셀 탭 1
  2. 수급 데이터 (최신 120영업일) → 엑셀 탭 2
  3. 최신 뉴스 제목 100개 → 엑셀 탭 3
  4. HTS 대차거래 CSV 자동 탐색 및 수급 데이터와 병합 → 엑셀 탭 4
  5. 신용잔고 일별추이 → 엑셀 탭 5
  6. 일별 심화 지표 → 엑셀 탭 6
  7. 시장 비중 분석 → 엑셀 탭 7
  8. 회원사동향 (★ 이관: 시간별 수집기에서 이동, 일별 1회 스냅샷) → 엑셀 탭 8
  9. 매물대   (★ 이관: 시간별 수집기에서 이동, 일별 1회 스냅샷) → 엑셀 탭 9
공통: 모든 탭 날짜 내림차순 정렬

[2026-09 최적화 변경]
  · 탭6(일별심화지표): 최대낙폭/최대수익은 일봉만으로 계산 가능한데 기존 코드가
    분봉 API를 불필요하게 호출하고 있었다 → API 호출 제거, 즉시 계산으로 전환.
    저가형성시간/초반거래대금비중은 한번 확정되면 바뀌지 않는 과거값이므로
    _cache/advanced_intraday_{종목코드}.json 에 누적 저장하고, 캐시에 없는
    "신규 거래일"만 분봉 API를 호출한다 (최초 1회 이후 매일 최대 120배 API 절감).
  · 회원사동향/매물대: 장중 시간 단위로 크게 변하지 않는 스냅샷 데이터라
    기존 시간별(1일 7회) 수집기에서 이 일별 수집기(1일 1회)로 이관했다.
"""

import requests
import json
import time
import os
import glob
import sys      
import io       
import traceback
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

# 1. 터미널 인코딩 에러(이모지 출력 등) 방지
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8')

# 2. 전체 코드 들여쓰기 없이 에러를 완벽하게 잡아내는 Global Hook
def global_exception_handler(exc_type, exc_value, exc_traceback):
    print("\n==========================================")
    print("🚨 [치명적 오류 발생] 에러 상세 내역:")
    print("==========================================")
    traceback.print_exception(exc_type, exc_value, exc_traceback)
    print("==========================================")
    sys.exit(1)

sys.excepthook = global_exception_handler


# 병합을 위한 pandas 라이브러리 체크
try:
    import pandas as pd
except ImportError:
    print("❌ pandas 라이브러리가 설치되어 있지 않습니다.")
    print("   터미널에서 'pip install pandas'를 실행한 후 다시 시도해주세요.")
    exit()

# =============================================
# ★ 본인의 API 정보를 입력하세요 ★
# =============================================
from kis_config import require

APP_KEY    = require("KIS_APP_KEY")      # .env 에서 로드 (kis_config.py)
APP_SECRET = require("KIS_APP_SECRET")
CANO       = require("KIS_CANO")
ACNT_PRDT_CD = "01"
BASE_URL   = "https://openapi.koreainvestment.com:9443"


# 엑셀 스타일 설정 (공통)
HEADER_FONT  = Font(bold=True, color="FFFFFF")
HEADER_FILL  = PatternFill("solid", start_color="4F81BD")
BODY_FONT    = Font(name="맑은 고딕", size=10)
CENTER_ALIGN = Alignment(horizontal="center", vertical="center")
THIN_SIDE    = Side(border_style="thin", color="000000")
BORDER       = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)

# ──────────────────────────────────────────────
# 0-1. 공용 숫자 변환 헬퍼
# ──────────────────────────────────────────────
def to_int(v):
    try:
        return int(str(v).replace(",", "").replace("+", "") or 0)
    except Exception:
        return 0

def to_float(v):
    try:
        return float(str(v).replace(",", "").replace("+", "") or 0)
    except Exception:
        return 0.0

# ──────────────────────────────────────────────
# 0-2. 일별 심화지표 분봉 캐시 (신규 거래일만 API 호출하기 위함)
# ──────────────────────────────────────────────
def _cache_dir():
    d = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_cache")
    os.makedirs(d, exist_ok=True)
    return d

def _load_intraday_cache(stock_code):
    path = os.path.join(_cache_dir(), f"advanced_intraday_{stock_code}.json")
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def _save_intraday_cache(stock_code, cache):
    path = os.path.join(_cache_dir(), f"advanced_intraday_{stock_code}.json")
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(cache, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"  \u26a0\ufe0f  \uce90\uc2dc \uc800\uc7a5 \uc2e4\ud328({stock_code}): {e}")

# ──────────────────────────────────────────────
# 0-3. 회원사동향 / 매물대  (★ 이관: 기존 시간별 수집기에서 이동)
#      장중 시간 단위로 크게 변하지 않는 스냅샷 데이터라 하루 1회
#      (장마감 후 일별 수집 시점)에만 수집한다.
# ──────────────────────────────────────────────
_SNAPSHOT_COL_MAP = {
    'stck_bsop_date': '일자',
    'stck_prpr': '현재가', 'stck_oprc': '시가', 'stck_hgpr': '최고가', 'stck_lwpr': '최저가', 'stck_clpr': '종가',
    'acml_vol': '누적거래량', 'acml_tr_pbmn': '누적거래대금',
    'prdy_vrss': '전일대비', 'prdy_vrss_sign': '대비부호', 'prdy_ctrt': '등락율',
    'data_rank': '순위', 'acml_vol_rlim': '매물대비중(%)',
}

def _snapshot_process_dataframe(df):
    if df.empty:
        return pd.DataFrame(["데이터 없음"], columns=["상태"])
    for col in df.columns:
        if not str(col).endswith('name') and not str(col).endswith('isnm'):
            try:
                temp = df[col].replace(r'^\s*$', float('NaN'), regex=True)
                df[col] = pd.to_numeric(temp)
            except Exception:
                pass
    return df.rename(columns=_SNAPSHOT_COL_MAP)

def get_member_trend(token, stock_code):
    """회원사동향 창구 상위 5개 (일별 1회 스냅샷)."""
    url = f"{BASE_URL}/uapi/domestic-stock/v1/quotations/inquire-member"
    headers = make_headers(token, "FHKST01010600")
    try:
        res = safe_get(url, headers, {"FID_COND_MRKT_DIV_CODE": "J", "FID_INPUT_ISCD": stock_code})
        output = res.json().get("output", {})
        if isinstance(output, list):   # API \uac00 \uc694\uc18c 1\uac1c\uc9dc\ub9ac list \ub85c \uc751\ub2f5\ud568
            output = output[0] if output else {}
    except Exception as e:
        print(f"  \u274c \ud68c\uc6d0\uc0ac\ub3d9\ud5a5 \uc624\ub958: {e}")
        return pd.DataFrame(["데이터 없음"], columns=["상태"])
    if not output:
        return pd.DataFrame(["데이터 없음"], columns=["상태"])

    today_str = datetime.today().strftime("%Y-%m-%d")
    rows = []
    for i in range(1, 6):
        rows.append({
            "일자": today_str, "순위": i,
            "매도회원사": output.get(f"seln_mbcr_name{i}", ""),
            "매도수량":   to_int(output.get(f"total_seln_qty{i}", 0)),
            "매수회원사": output.get(f"shnu_mbcr_name{i}", ""),
            "매수수량":   to_int(output.get(f"total_shnu_qty{i}", 0)),
        })
    return pd.DataFrame(rows)

def get_price_volume_bar(token, stock_code):
    """누적 거래 매물대 (일별 1회 스냅샷)."""
    url = f"{BASE_URL}/uapi/domestic-stock/v1/quotations/pbar-tratio"
    headers = make_headers(token, "FHPST01130000")
    params = {"FID_COND_MRKT_DIV_CODE": "J", "FID_INPUT_ISCD": stock_code,
              "FID_COND_SCR_DIV_CODE": "20171", "FID_SQC_DATA_YN": "N", "FID_INPUT_HOUR_1": ""}
    try:
        res = safe_get(url, headers, params)
        data = res.json()
        target = data.get("output2") or data.get("output1") or data.get("output") or []
        if isinstance(target, dict):
            target = [target]
        df = pd.DataFrame(target)
    except Exception as e:
        print(f"  \u274c \ub9e4\ubb3c\ub300 \uc624\ub958: {e}")
        return pd.DataFrame(["데이터 없음"], columns=["상태"])
    return _snapshot_process_dataframe(df)

def save_member_trend_excel(df, wb):
    ws = wb.create_sheet("회원사동향")
    if df.empty or "상태" in df.columns:
        ws.append(["데이터 없음"])
        print("  \u26a0\ufe0f  [탭8] 회원사동향 데이터 없음")
        return
    headers = list(df.columns)
    widths  = [14] + [12] * (len(headers) - 1)
    set_header(ws, headers, widths)
    for i, row in df.iterrows():
        ws.append(list(row))
        ri = i + 2
        for col in range(1, len(headers) + 1):
            cell = ws.cell(ri, col)
            cell.font, cell.border, cell.alignment = BODY_FONT, BORDER, CENTER_ALIGN
    ws.freeze_panes = "A2"
    print("  \u2705 [탭8] 회원사동향 완료 (일별 1회 스냅샷)")

def save_price_volume_bar_excel(df, wb):
    ws = wb.create_sheet("매물대")
    if df.empty or "상태" in df.columns:
        ws.append(["데이터 없음"])
        print("  \u26a0\ufe0f  [탭9] 매물대 데이터 없음")
        return
    headers = list(df.columns)
    widths  = [12] * len(headers)
    set_header(ws, headers, widths)
    for i, row in df.iterrows():
        ws.append(list(row))
        ri = i + 2
        for col in range(1, len(headers) + 1):
            cell = ws.cell(ri, col)
            cell.font, cell.border, cell.alignment = BODY_FONT, BORDER, CENTER_ALIGN
    ws.freeze_panes = "A2"
    print("  \u2705 [탭9] 매물대 완료 (일별 1회 스냅샷)")

def set_header(ws, headers, col_widths):
    ws.append(headers)
    for col, width in enumerate(col_widths, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = width
    for cell in ws[1]:
        cell.font      = HEADER_FONT
        cell.fill      = HEADER_FILL
        cell.alignment = CENTER_ALIGN
        cell.border    = BORDER

def make_headers(token, tr_id):
    return {
        "content-type": "application/json; charset=utf-8",
        "authorization": f"Bearer {token}",
        "appkey": APP_KEY.strip().replace('\n', '').replace('\r', ''),
        "appsecret": APP_SECRET.strip().replace('\n', '').replace('\r', ''),
        "tr_id": tr_id,
        "custtype": "P"
    }

def safe_get(url, headers, params, retries=3, timeout=10):
    """
    ConnectionResetError 등 네트워크 단절에 대비한 재시도 래퍼.
    재시도 간격: 1초, 2초 (지수 백오프)
    """
    for attempt in range(retries):
        try:
            return requests.get(url, headers=headers, params=params, timeout=timeout)
        except Exception as e:
            if attempt < retries - 1:
                time.sleep(1.0 * (attempt + 1))
            else:
                raise

def get_access_token():
    url = f"{BASE_URL}/oauth2/tokenP"
    headers = {"content-type": "application/json"}
    
    clean_app_key = APP_KEY.strip().replace('\n', '').replace('\r', '')
    clean_app_secret = APP_SECRET.strip().replace('\n', '').replace('\r', '')
    
    body = {
        "grant_type": "client_credentials",
        "appkey": clean_app_key,
        "appsecret": clean_app_secret
    }
    res = requests.post(url, headers=headers, json=body)
    res_data = res.json()
    if "access_token" in res_data:
        print("✅ 액세스 토큰 발급 완료")
        return res_data["access_token"]
    else:
        raise Exception(f"토큰 발급 실패: {res_data}")

# =============================================
# 1. 일봉 데이터 수집
# =============================================
def get_ohlcv(token, stock_code):
    """
    일봉 데이터 수집 (마스터 캘린더용 120영업일 확보)
    ─ 한국투자증권 일봉 API는 1회 호출 시 최대 100영업일만 반환한다.
    ─ 120영업일 ≒ 달력 170일이므로, 아래 2구간으로 나눠 호출 후 합산한다.
      · 구간A: (today - 365일) ~ (today - 181일)   ← 과거 원거리 구간
      · 구간B: (today - 180일) ~ today              ← 최근 구간
    ─ 두 구간 합산 후 중복 제거·오름차순 정렬 → 마지막 120건만 사용
    """
    url = f"{BASE_URL}/uapi/domestic-stock/v1/quotations/inquire-daily-itemchartprice"
    headers = make_headers(token, "FHKST03010100")

    today = datetime.today()

    # 120영업일을 확보하기 위한 2구간 설정
    # (공휴일·연휴 여유분 포함, 각 구간이 100영업일 이내에 들어오도록 분할)
    fetch_ranges = [
        (
            (today - timedelta(days=365)).strftime("%Y%m%d"),
            (today - timedelta(days=181)).strftime("%Y%m%d"),
            "구간A(원거리)"
        ),
        (
            (today - timedelta(days=180)).strftime("%Y%m%d"),
            today.strftime("%Y%m%d"),
            "구간B(최근)"
        ),
    ]

    print(f"\n📊 [1단계] 일봉 데이터 수집 중 (종목: {stock_code}) ...")

    def _fetch_range(start_date, end_date, range_name):
        print(f"   [{range_name}] {start_date} ~ {end_date}", end=" ")
        params = {
            "FID_COND_MRKT_DIV_CODE": "J",
            "FID_INPUT_ISCD":         stock_code,
            "FID_INPUT_DATE_1":       start_date,
            "FID_INPUT_DATE_2":       end_date,
            "FID_PERIOD_DIV_CODE":    "D",
            "FID_ORG_ADJ_PRC":        "0"
        }
        try:
            res  = safe_get(url, headers=headers, params=params)
            data = res.json()
            items = data.get("output2", [])
            print(f"→ {len(items)}건 ✅")
            return items
        except Exception as e:
            print(f"❌ 통신 오류: {e}")
            return []

    seen_dates = set()
    rows = []

    # 두 구간 병렬 호출
    with ThreadPoolExecutor(max_workers=2) as ex:
        f_a = ex.submit(_fetch_range, *fetch_ranges[0])
        f_b = ex.submit(_fetch_range, *fetch_ranges[1])
        results_ab = [f_a.result(), f_b.result()]

    for output2 in results_ab:
        for item in output2:
            try:
                d = item.get("stck_bsop_date")
                if not d or d in seen_dates:
                    continue
                seen_dates.add(d)
                rows.append({
                    "날짜":   d,
                    "저가":   int(item.get("stck_lwpr", 0)),
                    "시가":   int(item.get("stck_oprc", 0)),
                    "고가":   int(item.get("stck_hgpr", 0)),
                    "종가":   int(item.get("stck_clpr", 0)),
                    "거래량": int(item.get("acml_vol", 0))
                })
            except:
                continue

    # 오름차순 정렬 후 가장 최근 120영업일만 유지 (마스터 캘린더 기준)
    rows.sort(key=lambda x: x["날짜"])
    TARGET_DAYS = 120
    if len(rows) > TARGET_DAYS:
        rows = rows[-TARGET_DAYS:]

    print(f"  → 마스터 캘린더 확정: {len(rows)}영업일 "
          f"({rows[0]['날짜']} ~ {rows[-1]['날짜']})")
    return rows

def save_ohlcv_excel(rows, wb):
    ws = wb.create_sheet("일봉데이터")

    headers = ["날짜", "저가", "시가", "고가", "종가", "거래량"]
    col_widths = [14, 12, 12, 12, 12, 15]
    set_header(ws, headers, col_widths)

    num_fmt = "#,##0"
    for i, row in enumerate(sorted(rows, key=lambda x: x["날짜"], reverse=True), 2):
        ws.append([row["날짜"], row["저가"], row["시가"], row["고가"], row["종가"], row["거래량"]])
        for col in range(1, 7):
            cell = ws.cell(i, col)
            cell.font      = BODY_FONT
            cell.border    = BORDER
            cell.alignment = CENTER_ALIGN
            if col > 1:
                cell.number_format = num_fmt

    ws.freeze_panes = "A2"
    print("  ✅ [탭1] 일봉데이터 완료")

# =============================================
# 2. 수급 데이터 수집 (일봉 마스터 캘린더 연동 및 20일 분할 방식)
# =============================================
# =============================================
# 2. 수급 데이터 수집 (cursor 페이지네이션 방식)
# =============================================
def get_investor_data(token, stock_code, ohlcv_rows):
    """
    ★ 수급 API(FHPTJ04160001) 페이지네이션 분석 결과:
      - FID_INPUT_DATE_1 / DATE_2 파라미터는 무시됨
      - 항상 FID_INPUT_DATE_1 날짜 기준 최신 30건 고정 반환
      - 다음 페이지: FID_INPUT_DATE_1을 이전 페이지 마지막 날짜 -1일로 설정

    해결책: 신용잔고 수집과 동일한 cursor 방식으로 전환
      → master_start 날짜에 도달할 때까지 반복 호출
      → 청크 분할 불필요, 누락 원천 차단
    """
    url = f"{BASE_URL}/uapi/domestic-stock/v1/quotations/investor-trade-by-stock-daily"
    headers = make_headers(token, "FHPTJ04160001")

    if not ohlcv_rows:
        print("  ⚠️ 일봉 데이터가 없어 수급 데이터를 수집할 기준 날짜가 없습니다.")
        return []

    dates        = sorted([row["날짜"] for row in ohlcv_rows])
    valid_dates  = set(dates)
    total_days   = len(dates)
    master_start = dates[0]
    master_end   = dates[-1]

    print(f"\n📊 [2단계] 수급 데이터 수집 중 (종목: {stock_code})")
    print(f"   마스터 캘린더 범위: {master_start} ~ {master_end} ({total_days}영업일)")
    print(f"   [cursor 페이지네이션] {master_end}부터 역방향 수집...")

    def to_int(v):
        try:
            return int(str(v).replace(",", "").replace("+", "") or 0)
        except:
            return 0

    all_rows   = []
    seen_dates = set()
    cursor     = master_end   # 가장 최신 날짜부터 시작
    page       = 0

    while True:
        params = {
            "FID_COND_MRKT_DIV_CODE": "J",
            "FID_INPUT_ISCD":         stock_code,
            "FID_INPUT_DATE_1":       cursor,
            "FID_INPUT_DATE_2":       master_start,
            "FID_ORG_ADJ_PRC":        "0",
            "FID_ETC_CLS_CODE":       "0"
        }

        try:
            res  = safe_get(url, headers=headers, params=params)
            data = res.json()
        except Exception as e:
            print(f"   ❌ 통신 오류: {e}")
            break

        # 당일 15:40 이전 에러는 cursor를 전날로 당겨 재시도
        if data.get("rt_cd") != "0":
            msg = data.get("msg1", "")
            if ("TIME" in msg.upper() or "15:40" in msg) and cursor == master_end:
                prev = datetime.strptime(cursor, "%Y%m%d") - timedelta(days=1)
                cursor = prev.strftime("%Y%m%d")
                print(f"   💡 당일 데이터 미집계, {cursor}부터 재시도...")
                continue
            else:
                print(f"   ❌ API 거부: {msg}")
                break

        output2 = data.get("output2", [])
        if not output2:
            break

        oldest_date = None
        added_page  = 0
        for item in output2:
            d = item.get("stck_bsop_date", "")
            if not d:
                continue
            oldest_date = d
            if d in seen_dates or d not in valid_dates:
                continue
            seen_dates.add(d)
            all_rows.append({
                "날짜":     d,
                "개인":     to_int(item.get("prsn_ntby_tr_pbmn",    0)),
                "외국인":   to_int(item.get("frgn_ntby_tr_pbmn",    0)),
                "기관합계": to_int(item.get("orgn_ntby_tr_pbmn",    0)),
                "금융투자": to_int(item.get("scrt_ntby_tr_pbmn",    0)),
                "보험":     to_int(item.get("insu_ntby_tr_pbmn",    0)),
                "투신":     to_int(item.get("fund_ntby_tr_pbmn",    0)),
                "은행":     to_int(item.get("bank_ntby_tr_pbmn",    0)),
                "연기금":   to_int(item.get("ivtr_ntby_tr_pbmn",    0)),
                "사모펀드": to_int(item.get("pe_fund_ntby_tr_pbmn", 0)),
                "기타법인": to_int(item.get("etc_corp_ntby_tr_pbmn",0)),
                "프로그램": to_int(item.get("mrbn_ntby_tr_pbmn",    0)),
            })
            added_page += 1

        page += 1
        print(f"   [페이지{page}] cursor={cursor} → {added_page}건 추가 (누적 {len(all_rows)}건)")

        # 종료 조건: 마지막 날짜가 master_start 이전이면 수집 완료
        if oldest_date and oldest_date <= master_start:
            break
        if oldest_date is None:
            break

        # cursor를 이전 페이지 마지막 날짜 -1일로 이동
        prev   = datetime.strptime(oldest_date, "%Y%m%d") - timedelta(days=1)
        cursor = prev.strftime("%Y%m%d")
        if cursor < master_start:
            break
        time.sleep(0.2)

    all_rows.sort(key=lambda x: x["날짜"], reverse=True)  # 내림차순

    collected = {r["날짜"] for r in all_rows}
    missing   = sorted(valid_dates - collected, reverse=True)
    print(f"  → 수급 수집 완료: {len(all_rows)}건 / 목표 {total_days}건")
    if missing:
        print(f"  ⚠️ 누락 날짜 {len(missing)}건: {missing[:5]}{'...' if len(missing)>5 else ''}")
    else:
        print("  ✅ 마스터 캘린더 전체 날짜 수급 완료 — 누락 없음")

    return all_rows

def save_investor_excel(rows, wb):
    ws = wb.create_sheet("수급데이터")

    headers = ["날짜","개인","외국인","기관합계","금융투자","보험","투신","은행","연기금","사모펀드","기타법인","프로그램"]
    col_widths = [14,12,12,12,12,10,10,10,10,12,12,12]
    set_header(ws, headers, col_widths)

    pos_fill = PatternFill("solid", start_color="E8F4EA")  
    neg_fill = PatternFill("solid", start_color="FDE8E8")  
    num_fmt  = "#,##0;[Red]-#,##0"

    for i, row in enumerate(rows, 2):
        ws.append([
            row["날짜"], row["개인"], row["외국인"], row["기관합계"],
            row["금융투자"], row["보험"], row["투신"], row["은행"],
            row["연기금"], row["사모펀드"], row["기타법인"], row["프로그램"],
        ])
        for col in range(1, 13):
            cell = ws.cell(i, col)
            cell.font      = BODY_FONT
            cell.border    = BORDER
            cell.alignment = CENTER_ALIGN
            if col > 1:
                cell.number_format = num_fmt
                if isinstance(cell.value, (int, float)):
                    cell.fill = pos_fill if cell.value >= 0 else neg_fill

    ws.freeze_panes = "A2"
    print("  ✅ [탭2] 수급데이터 완료")

# =============================================
# 3. 뉴스 수집
# =============================================
def get_news(token, stock_code):
    url = f"{BASE_URL}/uapi/domestic-stock/v1/quotations/news-title"
    headers = make_headers(token, "FHKST01011800")

    print(f"\n📊 [3단계] 뉴스 수집 중 (종목: {stock_code}) ...")

    TARGET = 100
    all_news   = []
    seen_srno  = set()
    params = {
        "FID_NEWS_OFER_ENTP_CODE": "",
        "FID_COND_MRKT_DIV_CODE":  "J",
        "FID_COND_MRKT_CLS_CODE":  "",
        "FID_INPUT_ISCD":          stock_code,
        "FID_TITL_CNTT":           "",
        "FID_DATA_RANK":           "",
        "FID_INPUT_DATE_1":        "",
        "FID_INPUT_HOUR_1":        "",
        "FID_RANK_SORT_CLS_CODE":  "",
        "FID_INPUT_SRNO":          ""
    }

    for _ in range(10):   # 최대 10페이지(×약 20건 = 최대 200건 시도)
        try:
            res    = safe_get(url, headers=headers, params=params)
            output = res.json().get("output", [])
        except Exception as e:
            print(f"   ❌ 뉴스 오류: {e}")
            break

        if not output:
            break

        added = 0
        for item in output:
            srno = item.get("news_srno", "")
            if srno and srno in seen_srno:
                continue
            if srno:
                seen_srno.add(srno)
            all_news.append(item)
            added += 1

        if len(all_news) >= TARGET:
            break
        if added == 0:   # 중복만 왔으면 페이지 소진
            break

        # 다음 페이지: 마지막 뉴스의 날짜·시간으로 커서 이동
        last = output[-1]
        params["FID_INPUT_DATE_1"] = last.get("data_dt", "")
        params["FID_INPUT_HOUR_1"] = last.get("data_tm", "")
        time.sleep(0.15)

    result = all_news[:TARGET]
    print(f"  → {len(result)}건 수집 완료")
    return result

def save_news_excel(news_list, wb):
    ws = wb.create_sheet("최신뉴스")

    headers = ["날짜", "시간", "뉴스제목"]
    col_widths = [14, 12, 80]
    set_header(ws, headers, col_widths)

    for i, item in enumerate(news_list, 2):
        dt = item.get("data_dt", "")
        tm = item.get("data_tm", "")
        raw_title = str(item.get("hts_pbnt_titl_cntt", ""))
        clean_title = ILLEGAL_CHARACTERS_RE.sub('', raw_title)
        ws.append([dt, tm, clean_title])
        for col in range(1, 4):
            cell = ws.cell(i, col)
            cell.font   = BODY_FONT
            cell.border = BORDER
            cell.alignment = CENTER_ALIGN if col < 3 else Alignment(vertical="center", wrap_text=True)

    ws.freeze_panes = "A2"
    print("  ✅ [탭3] 최신뉴스 완료")

# =============================================
# 4. 일별 대차거래 추이 수집 (API)
# =============================================
def get_short_selling(token, stock_code, ohlcv_rows=None):
    """
    [버그2 수정] start_date / end_date를 timedelta 하드코딩이 아닌
    마스터 캘린더(ohlcv_rows)의 실제 첫날·마지막날로 고정.
    → 모든 탭의 날짜 범위가 일봉 기준으로 일치하게 됨.
    """
    url = f"{BASE_URL}/uapi/domestic-stock/v1/quotations/daily-loan-trans"
    headers = make_headers(token, "HHPST074500C0")

    if ohlcv_rows:
        dates      = sorted([r["날짜"] for r in ohlcv_rows])
        start_date = dates[0]
        end_date   = dates[-1]
    else:
        print("  ❌ 마스터 캘린더(ohlcv_rows)가 없어 대차거래 수집을 건너뜁니다.")
        return []

    print(f"\n📊 [4단계] 대차거래 데이터 수집 중 (종목: {stock_code}) ...")

    def to_int(v):
        try:
            return int(str(v).replace(",", "").replace("+", "") or 0)
        except:
            return 0

    all_rows = []
    cts = ""

    while True:
        params = {
            "MRKT_DIV_CLS_CODE": "3",       
            "MKSC_SHRN_ISCD":    stock_code,
            "START_DATE":        start_date,
            "END_DATE":          end_date,
            "CTS":               cts,
        }

        try:
            res  = safe_get(url, headers=headers, params=params)
            data = res.json()
        except Exception as e:
            print(f"   ❌ 응답 오류: {e}")
            break

        if data.get("rt_cd") != "0":
            break

        output = data.get("output1", [])
        if not output:
            break

        for item in output:
            d = item.get("bsop_date", "")
            if not d:
                continue
            all_rows.append({
                "날짜":       d,
                "대차증가":   to_int(item.get("new_stcn",       0)),
                "대차상환":   to_int(item.get("rdmp_stcn",      0)),
                "전일대비":   to_int(item.get("prdy_rmnd_vrss", 0)),
                "잔고수량":   to_int(item.get("rmnd_stcn",      0)),
                "잔고금액(백만)": to_int(item.get("rmnd_amt",   0)),
            })

        next_cts = data.get("cts", "").strip()
        if not next_cts or len(output) < 100:
            break
        cts = next_cts
        time.sleep(0.2)

    all_rows.sort(key=lambda x: x["날짜"], reverse=True)
    print(f"  → {len(all_rows)}건 수집 완료")
    return all_rows


def save_short_selling_excel(rows, wb):
    ws = wb.create_sheet("대차거래추이")

    headers    = ["날짜", "대차증가", "대차상환", "전일대비", "잔고수량", "잔고금액(백만)"]
    col_widths = [14, 14, 14, 14, 16, 18]
    set_header(ws, headers, col_widths)

    num_fmt = "#,##0"
    pos_fill = PatternFill("solid", start_color="E8F4EA")
    neg_fill = PatternFill("solid", start_color="FDE8E8")

    for i, row in enumerate(rows, 2):
        ws.append([
            row["날짜"], row["대차증가"], row["대차상환"],
            row["전일대비"], row["잔고수량"], row["잔고금액(백만)"],
        ])
        for col in range(1, len(headers) + 1):
            cell = ws.cell(i, col)
            cell.font      = BODY_FONT
            cell.border    = BORDER
            cell.alignment = CENTER_ALIGN
            if col > 1:
                cell.number_format = num_fmt
            if col == 4 and isinstance(cell.value, (int, float)):
                cell.fill = pos_fill if cell.value >= 0 else neg_fill

    ws.freeze_panes = "A2"
    print("  ✅ [탭4] 대차거래추이 완료")


# =============================================
# 5. 신용잔고 일별추이 수집 (API)
# =============================================
def get_credit_balance(token, stock_code, ohlcv_rows=None):
    """
    [버그3 수정] start_date / end_date를 timedelta 하드코딩이 아닌
    마스터 캘린더(ohlcv_rows)의 실제 첫날·마지막날로 고정.
    → 탭5 신용잔고 날짜 범위가 탭1 일봉과 완전히 일치함.
    """
    url = f"{BASE_URL}/uapi/domestic-stock/v1/quotations/daily-credit-balance"
    headers = make_headers(token, "FHPST04760000")

    if ohlcv_rows:
        dates      = sorted([r["날짜"] for r in ohlcv_rows])
        start_date = dates[0]
        end_date   = dates[-1]
    else:
        print("  ❌ 마스터 캘린더(ohlcv_rows)가 없어 신용잔고 수집을 건너뜁니다.")
        return []

    print(f"\n📊 [5단계] 신용잔고 데이터 수집 중 (종목: {stock_code}) ...")

    def to_int(v):
        try:
            return int(str(v).replace(",", "").replace("+", "") or 0)
        except:
            return 0

    def to_float(v):
        try:
            return float(str(v).replace(",", "") or 0)
        except:
            return 0.0

    all_rows  = []
    seen_dates = set()
    cursor    = end_date

    while True:
        params = {
            "fid_cond_mrkt_div_code": "J",
            "fid_cond_scr_div_code":  "20476",
            "fid_input_iscd":         stock_code,
            "fid_input_date_1":       cursor,
        }

        try:
            res  = safe_get(url, headers=headers, params=params)
            data = res.json()
        except Exception as e:
            print(f"   ❌ 응답 오류: {e}")
            break

        if data.get("rt_cd") != "0":
            break

        output = data.get("output", [])
        if not output:
            break

        oldest_date = None
        for item in output:
            d = item.get("deal_date", "")
            if not d:
                continue
            oldest_date = d
            if d in seen_dates:
                continue
            seen_dates.add(d)
            if d < start_date:
                continue
            try:
                all_rows.append({
                    "날짜":              d,
                    "융자신규(주)":      to_int(item.get("whol_loan_new_stcn",   0)),
                    "융자상환(주)":      to_int(item.get("whol_loan_rdmp_stcn",  0)),
                    "융자잔고(주)":      to_int(item.get("whol_loan_rmnd_stcn",  0)),
                    "융자잔고금액(백만)": to_int(item.get("whol_loan_rmnd_amt",   0)),
                    "융자잔고비율(%)":   to_float(item.get("whol_loan_rmnd_rate", 0)),
                    "대주신규(주)":      to_int(item.get("whol_stln_new_stcn",   0)),
                    "대주상환(주)":      to_int(item.get("whol_stln_rdmp_stcn",  0)),
                    "대주잔고(주)":      to_int(item.get("whol_stln_rmnd_stcn",  0)),
                    "대주잔고금액(백만)": to_int(item.get("whol_stln_rmnd_amt",   0)),
                    "대주잔고비율(%)":   to_float(item.get("whol_stln_rmnd_rate", 0)),
                })
            except Exception as e:
                pass

        if len(output) < 30:
            break
        if oldest_date and oldest_date < start_date:
            break

        prev   = datetime.strptime(oldest_date, "%Y%m%d") - timedelta(days=1)
        cursor = prev.strftime("%Y%m%d")
        if cursor < start_date:
            break
        time.sleep(0.2)

    all_rows.sort(key=lambda x: x["날짜"], reverse=True)
    print(f"  → {len(all_rows)}건 수집 완료")
    return all_rows


def save_credit_balance_excel(rows, wb):
    ws = wb.create_sheet("신용잔고추이")

    headers = [
        "날짜",
        "융자신규(주)", "융자상환(주)", "융자잔고(주)", "융자잔고금액(백만)", "융자잔고비율(%)",
        "대주신규(주)", "대주상환(주)", "대주잔고(주)", "대주잔고금액(백만)", "대주잔고비율(%)",
    ]
    col_widths = [14, 14, 14, 14, 18, 16, 14, 14, 14, 18, 16]
    set_header(ws, headers, col_widths)

    int_fmt   = "#,##0"
    float_fmt = "0.00"
    pos_fill  = PatternFill("solid", start_color="E8F4EA")
    neg_fill  = PatternFill("solid", start_color="FDE8E8")

    for i, row in enumerate(rows, 2):
        ws.append([
            row["날짜"],
            row["융자신규(주)"],      row["융자상환(주)"],      row["융자잔고(주)"],
            row["융자잔고금액(백만)"], row["융자잔고비율(%)"],
            row["대주신규(주)"],      row["대주상환(주)"],      row["대주잔고(주)"],
            row["대주잔고금액(백만)"], row["대주잔고비율(%)"],
        ])
        for col in range(1, len(headers) + 1):
            cell = ws.cell(i, col)
            cell.font      = BODY_FONT
            cell.border    = BORDER
            cell.alignment = CENTER_ALIGN
            if col in (6, 11):                    
                cell.number_format = float_fmt
            elif col > 1:
                cell.number_format = int_fmt
            if col in (4, 9) and isinstance(cell.value, (int, float)):
                cell.fill = pos_fill if cell.value > 0 else neg_fill

    ws.freeze_panes = "A2"
    print("  ✅ [탭5] 신용잔고추이 완료")


# =============================================
# 6. 일별 심화 지표 (120영업일, 분봉 하이브리드)
# =============================================

def _floor_to_30min(raw_time):
    """
    1분봉 시각(HHMMSS)을 30분 단위로 버림하여 HH:MM 문자열 반환.
    예) '132600' → '13:00', '093500' → '09:30', '153000' → '15:30'
    """
    if len(raw_time) < 4:
        return "N/A"
    hh = int(raw_time[:2])
    mm = int(raw_time[2:4])
    mm_floored = (mm // 30) * 30
    return f"{hh:02d}:{mm_floored:02d}"


def _call_minute_candles(token, stock_code, date_str, hour_str):
    """
    분봉 API 단일 호출. 1회 호출 = 최대 120개 1분봉 캔들 반환.
    TR: FHKST03010230 (주식일별분봉조회)
    date_str : YYYYMMDD
    hour_str : HHMMSS (예: '153000')
               해당 시각을 기준으로 역순(최신→과거) 120개 반환
    ConnectionResetError 등 네트워크 오류 시 최대 2회 재시도.
    """
    url = f"{BASE_URL}/uapi/domestic-stock/v1/quotations/inquire-time-dailychartprice"
    headers = make_headers(token, "FHKST03010230")
    params = {
        "FID_COND_MRKT_DIV_CODE": "J",
        "FID_INPUT_ISCD":         stock_code,
        "FID_INPUT_DATE_1":       date_str,
        "FID_INPUT_HOUR_1":       hour_str,
        "FID_PW_DATA_INCU_YN":   "Y",
        "FID_FAKE_TICK_INCU_YN": "",
    }
    for attempt in range(3):  # 최초 1회 + 재시도 2회
        try:
            res  = requests.get(url, headers=headers, params=params, timeout=10)
            data = res.json()
            if data.get("rt_cd") == "0":
                return data.get("output2", [])
            return []
        except Exception as e:
            if attempt < 2:
                time.sleep(1.0 * (attempt + 1))  # 1초, 2초 간격 재시도
            else:
                print(f"   ❌ 분봉 API 오류 ({date_str} {hour_str}): {e}")
    return []


def _build_intraday_map(token, stock_code, target_dates):
    """
    target_dates: YYYYMMDD 문자열 리스트 (최대 120일)

    ★ 개선 v2: 호출 횟수 절반 + 올바른 병렬 처리
    ─────────────────────────────────────────────────────────────
    [핵심 변경] 날짜당 2회 → 1회로 축소
      기존: 호출A(093000) + 호출B(153000) = 날짜당 2회 = 240회 총 호출
      개선: 호출1(153000) 단 1회만 호출 = 날짜당 1회 = 120회 총 호출

    왜 153000 단일 호출로 충분한가?
      - API는 FID_INPUT_HOUR_1 기준으로 역순(최신→과거) 최대 120캔들 반환
      - 153000 기준이면 09:14~15:30 구간 캔들 약 120개를 커버
        (1분봉 120개 = 120분 = 2시간 분량 → 장 시작 09:00부터 대부분 포함)
      - 저가 형성시간: 전 구간 캔들에서 탐색 가능
      - 초반 거래대금(09:01~09:30): acml_tr_pbmn은 누적값이므로
        09:30에 해당하는 캔들 1개만 있으면 계산 가능

    [올바른 병렬화] Lock은 시간 기록용, sleep은 Lock 바깥
      - 이전 버전의 버그: with _lock 블록 안에서 sleep → Lock을 쥔 채로 대기
        → 다른 스레드가 Lock 획득 불가 → 사실상 순차 실행과 동일
      - 수정: sleep을 Lock 바깥으로 이동 → 대기 중에도 다른 스레드 진입 가능

    결과: 120회 호출 × 0.22초(네트워크) / 5병렬 ≈ 약 25~35초 예상
    ─────────────────────────────────────────────────────────────
    반환값: {
        'early':    {date_str: [candle, ...]},   # 09:30 이전 캔들 (필터링)
        'intraday': {date_str: [candle, ...]},   # 전체 캔들 (153000 기준)
    }
    """
    # ── Rate Limit 제어 (초당 최대 4 req, 여유 있게) ───────────────
    # 한투 API 실측 기준: 연속 호출 시 약 0.25초 이상 간격 권장
    # Lock은 순서 보장용, sleep은 Lock 바깥 → 병렬 실효성 확보
    _MIN_INTERVAL   = 0.26          # 초당 ~4 req (0.21은 현실 네트워크에서 차단 유발)
    _lock           = threading.Lock()
    _last_call_time = [0.0]

    def _throttled_call(date_str, hour_str):
        """Rate Limit 준수 단일 분봉 API 호출."""
        # ① Lock에서 대기 시간 계산 후 즉시 해제 (sleep은 바깥에서)
        with _lock:
            now  = time.monotonic()
            wait = _MIN_INTERVAL - (now - _last_call_time[0])
            # 다음 스레드 진입 가능 시각을 미리 예약 (겹침 방지)
            _last_call_time[0] = now + max(wait, 0)
        if wait > 0:
            time.sleep(wait)
        result = _call_minute_candles(token, stock_code, date_str, hour_str)
        return result

    def _fetch_one_date(date_str):
        """
        153000 호출로 전체 캔들 취득. 09:30 이전 캔들이 없으면 093000 보완 호출.

        왜 보완 호출이 필요한가?
          거래량이 많은 날 153000 기준 120캔들은 오후에 집중되어
          09:14 이전 캔들이 잘릴 수 있음 → 초반 거래대금 계산 불가.
          이 경우에만 093000 추가 호출로 오전 구간을 보완.
          → 대부분 날은 1회, 오전 데이터 없는 날만 2회 호출.
        """
        raw = _throttled_call(date_str, "153000")
        day_candles = [c for c in raw if c.get("stck_bsop_date") == date_str]

        # 09:01~09:30 구간 필터링
        early = [
            c for c in day_candles
            if "090100" <= c.get("stck_cntg_hour", "") <= "093000"
        ]

        # 오전 캔들이 없으면 093000 보완 호출
        if not early:
            raw2 = _throttled_call(date_str, "093000")
            morning = [c for c in raw2 if c.get("stck_bsop_date") == date_str]
            early = [
                c for c in morning
                if "090100" <= c.get("stck_cntg_hour", "") <= "093000"
            ]
            # 전체 캔들에도 합산 (저가 탐색 범위 보완)
            existing_hours = {c.get("stck_cntg_hour") for c in day_candles}
            for c in morning:
                if c.get("stck_cntg_hour") not in existing_hours:
                    day_candles.append(c)

        return date_str, early, day_candles

    # ── 병렬 실행 ────────────────────────────────────────────────
    early_map    = {}
    intraday_map = {}

    dates_sorted = sorted(set(target_dates), reverse=True)
    total        = len(dates_sorted)

    # 5 스레드 × 0.26초 간격 = 스레드당 실효 간격 1.3초
    # 이론 처리량: 120일 / (0.26초 × 120 / 5스레드) ≈ 32초 (보완 호출 없을 경우)
    MAX_WORKERS = 5
    print(f"   분봉 병렬 수집: {total}일 × 최소1회 (병렬 {MAX_WORKERS}, "
          f"간격 {_MIN_INTERVAL}s, 예상 {int(total*_MIN_INTERVAL/MAX_WORKERS)+10}~"
          f"{int(total*_MIN_INTERVAL/MAX_WORKERS)+30}초)")

    t0        = time.monotonic()
    completed = 0
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = {executor.submit(_fetch_one_date, d): d for d in dates_sorted}
        for future in as_completed(futures):
            try:
                date_str, early, all_day = future.result()
                if early:
                    early_map[date_str]    = early
                if all_day:
                    intraday_map[date_str] = all_day
            except Exception as e:
                print(f"   ⚠️ {futures[future]} 처리 오류: {e}")
            completed += 1
            if completed % 20 == 0 or completed == total:
                elapsed = time.monotonic() - t0
                rate    = completed / elapsed if elapsed > 0 else 0
                eta     = (total - completed) / rate if rate > 0 else 0
                print(f"   [{completed}/{total}] {elapsed:.0f}s 경과, 잔여 ~{eta:.0f}s")

    print(f"   → 분봉 수집 완료: early={len(early_map)}일, intraday={len(intraday_map)}일")
    return {"early": early_map, "intraday": intraday_map}


def get_daily_advanced_metrics(token, stock_code, ohlcv_rows):
    """
    일별 심화 지표 4가지 산출 (2026-09 최적화판):
    ① 시가대비 최대낙폭(%)  : 일봉 재활용 — API 호출 없이 즉시 계산
    ② 시가대비 최대수익(%)  : 일봉 재활용 — API 호출 없이 즉시 계산
    ③ 당일 저가 형성시간    : 1분봉 → 30분 단위 버림 (신규 거래일만 API 호출)
    ④ 초반 거래대금 비중(%) : 09:01~09:30 누적 거래대금 / 일봉 총거래대금 (신규 거래일만 API 호출)

    ★ 최적화 포인트: ①②는 원래 분봉이 전혀 필요 없는 값인데 기존 코드는 매일
      120거래일치를 분봉 API로 재계산하고 있었다 → 제거.
      ③④는 한번 확정되면 바뀌지 않는 과거값이므로 _cache/advanced_intraday_
      {종목코드}.json 에 거래일별로 누적 저장해두고, 캐시에 없는 "신규 거래일"만
      분봉 API를 호출한다. 최초 1회(캐시 없음)는 기존과 동일하게 최대 120일치를
      수집하지만, 이후에는 매일 신규 1일치만 호출한다.
    """
    print(f"\n📊 [6단계] 일별 심화 지표 수집 중 (종목: {stock_code}, 최근 120영업일) ...")
    t_start = time.monotonic()

    sorted_ohlcv = sorted(ohlcv_rows, key=lambda x: x["날짜"], reverse=True)
    target_rows  = sorted_ohlcv[:120]
    target_dates = [r["날짜"] for r in target_rows]
    ohlcv_map    = {r["날짜"]: r for r in target_rows}

    cache = _load_intraday_cache(stock_code)
    uncached_dates = [d for d in target_dates if d not in cache]

    if uncached_dates:
        print(f"   캐시 미보유 거래일 {len(uncached_dates)}건만 분봉 API 호출 "
              f"(캐시 보유: {len(target_dates) - len(uncached_dates)}건)")
        candle_maps  = _build_intraday_map(token, stock_code, uncached_dates)
        early_map    = candle_maps["early"]
        intraday_map = candle_maps["intraday"]

        for date_str in uncached_dates:
            low_time    = "N/A"
            early_ratio = 0.0
            early_candles    = early_map.get(date_str, [])
            intraday_candles = intraday_map.get(date_str, [])

            if intraday_candles:
                min_candle = min(
                    intraday_candles,
                    key=lambda c: int(c.get("stck_lwpr") or 999999999)
                )
                raw_time = min_candle.get("stck_cntg_hour", "")
                low_time = _floor_to_30min(raw_time)

            if early_candles and intraday_candles:
                latest_all = max(intraday_candles, key=lambda c: c.get("stck_cntg_hour", "000000"))
                total_vol  = int(latest_all.get("acml_tr_pbmn", 0) or 0)
                if total_vol > 0:
                    latest_morning = max(early_candles, key=lambda c: c.get("stck_cntg_hour", "000000"))
                    early_vol      = int(latest_morning.get("acml_tr_pbmn", 0) or 0)
                    early_ratio    = round(early_vol / total_vol * 100, 2)

            cache[date_str] = {"저가형성시간": low_time, "초반거래대금비중(%)": early_ratio}

        _save_intraday_cache(stock_code, cache)
    else:
        print(f"   전체 {len(target_dates)}거래일 캐시 적중 — 분봉 API 호출 0건")

    results = []
    for date_str in sorted(target_dates):
        row    = ohlcv_map.get(date_str, {})
        open_p = row.get("시가", 0)
        high_p = row.get("고가", 0)
        low_p  = row.get("저가", 0)

        # ①② 일봉만으로 즉시 계산 (분봉 API 불필요)
        max_drop   = round((low_p  - open_p) / open_p * 100, 2) if open_p > 0 else 0.0
        max_profit = round((high_p - open_p) / open_p * 100, 2) if open_p > 0 else 0.0

        cached = cache.get(date_str, {"저가형성시간": "N/A", "초반거래대금비중(%)": 0.0})

        results.append({
            "날짜":               date_str,
            "시가":               open_p,
            "고가":               high_p,
            "저가":               low_p,
            "시가대비최대낙폭(%)": max_drop,
            "시가대비최대수익(%)": max_profit,
            "저가형성시간":        cached["저가형성시간"],
            "초반거래대금비중(%)": cached["초반거래대금비중(%)"],
        })

    results.sort(key=lambda x: x["날짜"], reverse=True)
    elapsed = time.monotonic() - t_start
    print(f"  → {len(results)}건 지표 산출 완료 "
          f"(소요시간: {elapsed:.1f}초, 신규 분봉 호출 {len(uncached_dates)}거래일)")
    return results

def save_daily_advanced_excel(rows, wb):
    ws = wb.create_sheet("일별심화지표")

    headers = [
        "날짜", "시가", "고가", "저가",
        "시가대비 최대낙폭(%)", "시가대비 최대수익(%)",
        "당일 저가 형성시간", "초반 거래대금 비중(%)"
    ]
    col_widths = [14, 12, 12, 12, 20, 20, 18, 22]
    set_header(ws, headers, col_widths)

    pct_fmt     = "0.00"
    price_fmt   = "#,##0"
    drop_fill   = PatternFill("solid", start_color="FDE8E8")
    profit_fill = PatternFill("solid", start_color="E8F4EA")
    early_fill  = PatternFill("solid", start_color="FFF9E6")

    for i, row in enumerate(rows, 2):
        ws.append([
            row["날짜"],
            row["시가"],
            row["고가"],
            row["저가"],
            row["시가대비최대낙폭(%)"],
            row["시가대비최대수익(%)"],
            row["저가형성시간"],
            row["초반거래대금비중(%)"],
        ])
        for col in range(1, len(headers) + 1):
            cell = ws.cell(i, col)
            cell.font      = BODY_FONT
            cell.border    = BORDER
            cell.alignment = CENTER_ALIGN
            if col in (2, 3, 4):
                cell.number_format = price_fmt
            elif col == 5:
                cell.number_format = pct_fmt
                cell.fill = drop_fill
            elif col == 6:
                cell.number_format = pct_fmt
                cell.fill = profit_fill
            elif col == 8:
                cell.number_format = pct_fmt
                cell.fill = early_fill

    ws.freeze_panes = "A2"
    print("  ✅ [탭6] 일별심화지표 완료")


# =============================================
# 7. 시장 비중 분석 (삼성전자+SK하이닉스 / 코스피)
# =============================================

def _get_index_daily(token, iscd, start_date, end_date):
    """지수 일별 종가·거래대금 조회 (TR: FHKUP03500100)"""
    url = f"{BASE_URL}/uapi/domestic-stock/v1/quotations/inquire-daily-indexchartprice"
    h = {
        "content-type":  "application/json; charset=utf-8",
        "authorization": f"Bearer {token}",
        "appkey":        APP_KEY,
        "appsecret":     APP_SECRET,
        "tr_id":         "FHKUP03500100",
        "custtype":      "P",
    }
    params = {
        "FID_COND_MRKT_DIV_CODE": "U",
        "FID_INPUT_ISCD":         iscd,
        "FID_INPUT_DATE_1":       start_date,
        "FID_INPUT_DATE_2":       end_date,
        "FID_PERIOD_DIV_CODE":    "D",
    }
    data = safe_get(url, headers=h, params=params).json()
    result = {}
    if data.get("rt_cd") == "0":
        for item in data.get("output2", []):
            d = item.get("stck_bsop_date")
            if d:
                result[d] = {
                    "종가":     float(item.get("bstp_nmix_prpr", 0)),
                    "거래대금": int(item.get("acml_tr_pbmn", 0)) * 1_000_000,
                }
    else:
        print(f"   ⚠️ 지수 조회 오류 ({iscd}): {data.get('msg1')}")
    return result


def _get_stock_daily(token, iscd, start_date, end_date):
    """종목 일별 종가·거래대금 조회 (TR: FHKST03010100)"""
    url = f"{BASE_URL}/uapi/domestic-stock/v1/quotations/inquire-daily-itemchartprice"
    h = {
        "content-type":  "application/json; charset=utf-8",
        "authorization": f"Bearer {token}",
        "appkey":        APP_KEY,
        "appsecret":     APP_SECRET,
        "tr_id":         "FHKST03010100",
        "custtype":      "P",
    }
    params = {
        "FID_COND_MRKT_DIV_CODE": "J",
        "FID_INPUT_ISCD":         iscd,
        "FID_INPUT_DATE_1":       start_date,
        "FID_INPUT_DATE_2":       end_date,
        "FID_PERIOD_DIV_CODE":    "D",
        "FID_ORG_ADJ_PRC":        "0",
    }
    data = safe_get(url, headers=h, params=params).json()
    result = {}
    if data.get("rt_cd") == "0":
        for item in data.get("output2", []):
            d = item.get("stck_bsop_date")
            if d:
                result[d] = {
                    "종가":     int(item.get("stck_clpr", 0)),
                    "거래대금": int(item.get("acml_tr_pbmn", 0)),
                }
    else:
        print(f"   ⚠️ 종목 조회 오류 ({iscd}): {data.get('msg1')}")
    return result


def write_market_ratio_sheet(wb, token, ohlcv_rows=None):
    """코스피·코스피200·코스닥150·삼성전자·SK하이닉스 일별 비중 분석 탭 추가.
    ohlcv_rows가 있으면 마스터 캘린더 날짜 범위를 기준으로 조회한다."""
    print("\n📊 [7단계] 시장 비중 분석 데이터 수집 중 ...")

    if ohlcv_rows:
        dates      = sorted([r["날짜"] for r in ohlcv_rows])
        start_date = dates[0]
        end_date   = dates[-1]
    else:
        today      = datetime.today()
        end_date   = today.strftime("%Y%m%d")
        start_date = (today - timedelta(days=120)).strftime("%Y%m%d")

    print(f"   기간: {start_date} ~ {end_date}")

    # 5개 시장 데이터 병렬 호출
    with ThreadPoolExecutor(max_workers=5) as ex:
        f_kospi     = ex.submit(_get_index_daily, token, "0001", start_date, end_date)
        f_kospi200  = ex.submit(_get_index_daily, token, "2001", start_date, end_date)
        f_kosdaq150 = ex.submit(_get_index_daily, token, "2203", start_date, end_date)
        f_sec       = ex.submit(_get_stock_daily, token, "005930", start_date, end_date)
        f_skh       = ex.submit(_get_stock_daily, token, "000660", start_date, end_date)
        kospi     = f_kospi.result()
        kospi200  = f_kospi200.result()
        kosdaq150 = f_kosdaq150.result()
        sec       = f_sec.result()
        skh       = f_skh.result()

    all_dates = sorted(kospi.keys(), reverse=True)
    print(f"   수집 영업일: {len(all_dates)}일")

    ws = wb.create_sheet("시장비중분석")

    hdr_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    hdr_font = Font(color="FFFFFF", bold=True)
    c_align  = Alignment(horizontal="center", vertical="center")
    r_align  = Alignment(horizontal="right",  vertical="center")

    headers    = [
        "날짜", "코스피 종가", "코스피 거래대금(원)", "코스피200 종가", "코스닥150 종가",
        "삼성전자 종가", "삼성전자 거래대금", "SK하이닉스 종가", "SK하이닉스 거래대금",
        "삼전+닉스 합계금액", "코스피내 쏠림비중(%)",
    ]
    col_widths = [14, 14, 22, 14, 14, 14, 18, 16, 18, 22, 18]

    ws.append(headers)
    for col, (cell, w) in enumerate(zip(ws[1], col_widths), 1):
        cell.fill = hdr_fill; cell.font = hdr_font
        cell.alignment = c_align; cell.border = BORDER
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = w

    for d in all_dates:
        k_close     = kospi.get(d,     {}).get("종가",     0)
        k_vol       = kospi.get(d,     {}).get("거래대금", 0)
        k200_close  = kospi200.get(d,  {}).get("종가",     0)
        kq150_close = kosdaq150.get(d, {}).get("종가",     0)
        sec_close   = sec.get(d,       {}).get("종가",     0)
        sec_vol     = sec.get(d,       {}).get("거래대금", 0)
        skh_close   = skh.get(d,       {}).get("종가",     0)
        skh_vol     = skh.get(d,       {}).get("거래대금", 0)
        sum_vol     = sec_vol + skh_vol
        ratio       = (sum_vol / k_vol * 100) if k_vol > 0 else 0
        fmt_date    = f"{d[:4]}-{d[4:6]}-{d[6:]}"

        ws.append([fmt_date, k_close, k_vol, k200_close, kq150_close,
                   sec_close, sec_vol, skh_close, skh_vol, sum_vol, ratio])

        ri = ws.max_row
        for col in range(1, len(headers) + 1):
            cell = ws.cell(ri, col)
            cell.border    = BORDER
            cell.alignment = c_align if col == 1 else r_align
            if col in (2, 4, 5):
                cell.number_format = "#,##0.00"
            elif col == 11:
                cell.number_format = '0.00"%"'
                if ratio > 35:
                    cell.font = Font(color="FF0000", bold=True)
            elif col > 1:
                cell.number_format = "#,##0"

    ws.freeze_panes = "A2"
    print(f"  ✅ [탭7] 시장비중분석 ({len(all_dates)}행)")


# =============================================
# 단일 종목 처리 함수
# =============================================
def process_stock(token, stock_code, stock_name, output_dir):
    """단일 종목의 전체 데이터를 수집하고 엑셀 파일로 저장한다."""

    print("\n" + "=" * 60)
    print(f"  처리 중: [{stock_code}] {stock_name}")
    print("=" * 60)

    wb = openpyxl.Workbook()

    # ─────────────────────────────────────────────────────────────
    # 1. 일봉 (★ 마스터 캘린더 생성 - 가장 먼저 실행, 실패 시 전체 중단)
    # ─────────────────────────────────────────────────────────────
    ohlcv_rows = []
    try:
        ohlcv_rows = get_ohlcv(token, stock_code)
        if ohlcv_rows:
            save_ohlcv_excel(ohlcv_rows, wb)
        else:
            print("  ❌ 일봉 데이터가 없어 이 종목을 건너뜁니다.")
            return None
    except Exception as e:
        print(f"  ❌ 일봉 오류: {e}")
        return None

    # ─────────────────────────────────────────────────────────────
    # 2. 수급 (마스터 캘린더 기준 날짜 범위 사용)
    # ─────────────────────────────────────────────────────────────
    try:
        inv_rows = get_investor_data(token, stock_code, ohlcv_rows)
        if inv_rows:
            save_investor_excel(inv_rows, wb)
    except Exception as e: print(f"  ❌ 수급 오류: {e}")

    # ─────────────────────────────────────────────────────────────
    # 3. 뉴스 (날짜 범위 무관 - 최신 50건)
    # ─────────────────────────────────────────────────────────────
    try:
        news = get_news(token, stock_code)
        if news:
            save_news_excel(news, wb)
    except Exception as e: print(f"  ❌ 뉴스 오류: {e}")

    # ─────────────────────────────────────────────────────────────
    # 4. 대차거래 (★ ohlcv_rows 전달 → 마스터 캘린더 날짜 범위 사용)
    # ─────────────────────────────────────────────────────────────
    try:
        short_rows = get_short_selling(token, stock_code, ohlcv_rows)
        if short_rows:
            save_short_selling_excel(short_rows, wb)
    except Exception as e: print(f"  ❌ 대차거래 오류: {e}")

    # ─────────────────────────────────────────────────────────────
    # 5. 신용잔고 (★ ohlcv_rows 전달 → 마스터 캘린더 날짜 범위 사용)
    # ─────────────────────────────────────────────────────────────
    try:
        credit_rows = get_credit_balance(token, stock_code, ohlcv_rows)
        if credit_rows:
            save_credit_balance_excel(credit_rows, wb)
    except Exception as e: print(f"  ❌ 신용잔고 오류: {e}")

    # ─────────────────────────────────────────────────────────────
    # 6. 일별 심화 지표 (최대낙폭·최대수익·저가형성시간·초반거래대금비중)
    # ─────────────────────────────────────────────────────────────
    try:
        adv_rows = get_daily_advanced_metrics(token, stock_code, ohlcv_rows)
        if adv_rows:
            save_daily_advanced_excel(adv_rows, wb)
    except Exception as e: print(f"  ❌ 일별심화지표 오류: {e}")

    # ─────────────────────────────────────────────────────────────
    # 7. 시장 비중 분석 (삼성전자+SK하이닉스 / 코스피)
    #    ohlcv_rows 전달 → 마스터 캘린더 날짜 범위 기준으로 조회
    # ─────────────────────────────────────────────────────────────
    try:
        write_market_ratio_sheet(wb, token, ohlcv_rows)
    except Exception as e: print(f"  ❌ 시장비중분석 오류: {e}")

    # ─────────────────────────────────────────────────────────────
    # 8. 회원사동향 (★ 이관: 기존 시간별 수집기에서 이동, 일별 1회 스냅샷)
    # ─────────────────────────────────────────────────────────────
    try:
        mem_df = get_member_trend(token, stock_code)
        save_member_trend_excel(mem_df, wb)
    except Exception as e: print(f"  ❌ 회원사동향 오류: {e}")

    # ─────────────────────────────────────────────────────────────
    # 9. 매물대 (★ 이관: 기존 시간별 수집기에서 이동, 일별 1회 스냅샷)
    # ─────────────────────────────────────────────────────────────
    try:
        pbar_df = get_price_volume_bar(token, stock_code)
        save_price_volume_bar_excel(pbar_df, wb)
    except Exception as e: print(f"  ❌ 매물대 오류: {e}")

    # openpyxl이 워크북 생성 시 기본으로 만드는 "Sheet" 탭 제거
    if "Sheet" in wb.sheetnames and len(wb.sheetnames) > 1:
        wb.remove(wb["Sheet"])

    # ★ 종목별 파일로 저장 (output_dir 하위)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_name = stock_name.replace("/", "_").replace("\\", "_")
    final_filename = os.path.join(output_dir, f"{stock_code}_{safe_name}_{timestamp}.xlsx")

    try:
        wb.save(final_filename)
        print(f"\n  ✅ 저장 완료: {os.path.abspath(final_filename)}")
        return final_filename
    except Exception as e:
        print(f"\n  ❌ 파일 저장 실패: {e}")
        return None


# =============================================
# 메인 실행부 (JSON 파일에서 티커 순차 처리)
# =============================================
if __name__ == "__main__":
    print("=" * 60)
    print("  한국투자증권 오픈API 데이터 수집 (JSON 티커 일괄 처리)")
    print("=" * 60)

    # ─────────────────────────────────────────────────────────────
    # tickers.json 로드
    # ─────────────────────────────────────────────────────────────
    TICKERS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "tickers.json")
    if not os.path.exists(TICKERS_FILE):
        print(f"❌ 티커 파일을 찾을 수 없습니다: {TICKERS_FILE}")
        exit()

    with open(TICKERS_FILE, "r", encoding="utf-8") as f:
        raw_tickers = json.load(f)

    # "종목코드;종목명;상태;비중" 형식에서 코드·이름·상태·비중 파싱
    tickers = []
    for entry in raw_tickers:
        parts = entry.split(";")
        if len(parts) >= 2:
            tickers.append({
                "code":   parts[0].strip(),
                "name":   parts[1].strip(),
                "status": parts[2].strip() if len(parts) > 2 else "",
                "weight": parts[3].strip() if len(parts) > 3 else "",
            })
        else:
            print(f"  ⚠️ 파싱 불가 항목 건너뜀: {entry}")

    if not tickers:
        print("❌ 처리할 종목이 없습니다.")
        exit()

    print(f"\n📋 총 {len(tickers)}개 종목 로드 완료:")
    for i, t in enumerate(tickers, 1):
        print(f"   {i:2d}. [{t['code']}] {t['name']}")

# ─────────────────────────────────────────────────────────────
    # 출력 폴더 생성 (MarketData/YYYYMMDD)
    # ─────────────────────────────────────────────────────────────
    run_date   = datetime.now().strftime("%Y%m%d")
    output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "MarketData", run_date)
    os.makedirs(output_dir, exist_ok=True)
    print(f"\n📁 출력 폴더: {os.path.abspath(output_dir)}")

    # ─────────────────────────────────────────────────────────────
    # 토큰 발급 (1회)
    # ─────────────────────────────────────────────────────────────
    try:
        token = get_access_token()
    except Exception as e:
        print(f"❌ 토큰 발급 오류: {e}")
        exit()

    # ─────────────────────────────────────────────────────────────
    # 종목별 순차 처리
    # ─────────────────────────────────────────────────────────────
    results        = {"성공": [], "실패": []}
    summary_list   = []
    collected_time = datetime.now()

    for idx, ticker in enumerate(tickers, 1):
        code   = ticker["code"]
        name   = ticker["name"]
        status = ticker.get("status", "")
        weight = ticker.get("weight", "")
        print(f"\n\n{'━' * 60}")
        print(f"  [{idx}/{len(tickers)}] {name} ({code}) 처리 시작")
        print(f"{'━' * 60}")

        saved_path = process_stock(token, code, name, output_dir)
        if saved_path:
            results["성공"].append(f"[{code}] {name}")
            summary_list.append({
                "ticker":       code,
                "name":         name,
                "status":       status,
                "weight":       weight,
                "file":         saved_path,
                "collected_at": collected_time.isoformat(),
            })
        else:
            results["실패"].append(f"[{code}] {name}")

        # 마지막 종목이 아니면 API 과부하 방지를 위해 잠시 대기
        if idx < len(tickers):
            print(f"\n  ⏳ 다음 종목 처리 전 3초 대기...")
            time.sleep(3)

    # ─────────────────────────────────────────────────────────────
    # 수집 요약 JSON 저장 (시간별 데이터와 동일한 포맷)
    # ─────────────────────────────────────────────────────────────
    summary_filename = f"daily_summary_{run_date}.json"
    summary_path     = os.path.join(output_dir, summary_filename)
    try:
        with open(summary_path, "w", encoding="utf-8") as f:
            json.dump(summary_list, f, ensure_ascii=False, indent=4)
        print(f"\n📄 요약 JSON 저장 완료: {os.path.abspath(summary_path)}")
    except Exception as e:
        print(f"\n❌ 요약 JSON 저장 실패: {e}")

    # ─────────────────────────────────────────────────────────────
    # 최종 결과 요약
    # ─────────────────────────────────────────────────────────────
    print(f"\n\n{'=' * 60}")
    print("  📊 전체 처리 결과 요약")
    print(f"{'=' * 60}")
    print(f"  ✅ 성공 ({len(results['성공'])}건): {', '.join(results['성공']) or '없음'}")
    print(f"  ❌ 실패 ({len(results['실패'])}건): {', '.join(results['실패']) or '없음'}")
    print(f"\n  📁 저장 위치: {os.path.abspath(output_dir)}")
    print("=" * 60)