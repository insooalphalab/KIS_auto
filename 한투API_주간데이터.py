"""
한국투자증권 오픈API - 주간 데이터 수집 프로그램
================================================
★ API 스펙 확인 기반으로 작성 (2026-06-06)

수집 탭 구성:
  탭1. 주봉_OHLCV_이격도   : FHKST03010100 (FID_PERIOD_DIV_CODE='W')
                             10주/40주 MA + 이격도, 외인소진율
  탭2. 주간_수급_누적       : FHPTJ04160001 (일봉 수급 → 주차별 집계)
                             외인/기관/연기금/금투/사모 5영업일 누적 순매수
                             주간 VWAP + 현재가 위/아래 판별
  탭3. 신용잔고_주간증감    : FHPST04760000 (일별 → 금요일 샘플링)
                             신용다이버전스 자동 판별
  탭4. 주간_거래대금_분석   : 탭1 재활용 (추가 API 없음)
                             4주 평균 대비 배율, 52주 백분위, 폭증 신호
  탭5. 외인소진율_추이      : 탭1 재활용 (추가 API 없음)
                             4주 추세, 소진율 구간 분류

API 주요 수정사항 (스펙 확인 후):
  - FHPTJ04160001: FID_ETC_CLS_CODE="1" (공란→"1" 수정)
  - 수급 단위: 백만원 (내부적으로 ×1_000_000 하지 않고 백만원 표기)
  - 신용잔고 단위: 융자금액 만원 단위 (whol_loan_rmnd_amt)
  - 주봉 API: FID_PERIOD_DIV_CODE='W', 1회 최대 100건
  - output2 필드명: stck_oprc(시가), stck_hgpr(고가), stck_lwpr(저가),
                    stck_clpr(종가), acml_vol(거래량), acml_tr_pbmn(거래대금)
                    hts_frgn_ehrt(외인소진율), prdy_vrss/prdy_ctrt(전주대비)

수집 시점: 매주 금요일 장 마감 후 실행
저장 위치: WeeklyData/YYYYMMDD/
"""

import requests
import json
import time
import os
import sys
import traceback
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor, as_completed

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE

# ── 터미널 인코딩 에러 방지 ─────────────────────────────────────────────
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8')

def global_exception_handler(exc_type, exc_value, exc_traceback):
    print("\n==========================================")
    print("🚨 [치명적 오류 발생] 에러 상세 내역:")
    print("==========================================")
    traceback.print_exception(exc_type, exc_value, exc_traceback)
    print("==========================================")
    sys.exit(1)

sys.excepthook = global_exception_handler

try:
    import pandas as pd
except ImportError:
    print("❌ pandas 라이브러리가 필요합니다. 'pip install pandas' 실행 후 재시도하세요.")
    exit()

# =============================================
# ★ API 정보
# =============================================
from kis_config import require

APP_KEY      = require("KIS_APP_KEY")    # .env 에서 로드 (kis_config.py)
APP_SECRET   = require("KIS_APP_SECRET")
CANO         = require("KIS_CANO")
ACNT_PRDT_CD = "01"
BASE_URL     = "https://openapi.koreainvestment.com:9443"

# ── 엑셀 공통 스타일 ────────────────────────────────────────────────────
HEADER_FONT  = Font(bold=True, color="FFFFFF")
# 주간 탭 헤더: 진청색 (일봉 4F81BD와 구분)
HEADER_FILL  = PatternFill("solid", start_color="1F497D")
BODY_FONT    = Font(name="맑은 고딕", size=10)
CENTER_ALIGN = Alignment(horizontal="center", vertical="center")
THIN_SIDE    = Side(border_style="thin", color="000000")
BORDER       = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)
POS_FILL     = PatternFill("solid", start_color="E8F4EA")   # 연초록 (양수)
NEG_FILL     = PatternFill("solid", start_color="FDE8E8")   # 연빨강 (음수)

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
        "content-type":  "application/json; charset=utf-8",
        "authorization": f"Bearer {token}",
        "appkey":        APP_KEY.strip(),
        "appsecret":     APP_SECRET.strip(),
        "tr_id":         tr_id,
        "custtype":      "P"
    }

def safe_get(url, headers, params, retries=3, timeout=10):
    for attempt in range(retries):
        try:
            return requests.get(url, headers=headers, params=params, timeout=timeout)
        except Exception as e:
            if attempt < retries - 1:
                time.sleep(1.0 * (attempt + 1))
            else:
                raise

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

# =============================================
# 토큰 발급
# =============================================
def get_access_token():
    url = f"{BASE_URL}/oauth2/tokenP"
    body = {
        "grant_type": "client_credentials",
        "appkey":     APP_KEY.strip(),
        "appsecret":  APP_SECRET.strip()
    }
    res = requests.post(url, headers={"content-type": "application/json"}, json=body)
    data = res.json()
    if "access_token" in data:
        print("✅ 액세스 토큰 발급 완료")
        return data["access_token"]
    raise Exception(f"토큰 발급 실패: {data}")


# =============================================
# 탭1. 주봉 OHLCV + 이평선 이격도
# =============================================
def get_weekly_ohlcv(token, stock_code):
    """
    TR: FHKST03010100  FID_PERIOD_DIV_CODE='W'
    ─ 1회 최대 100건 반환
    ─ 52주 확보를 위해 2구간 분할 호출
    ─ 이격도 계산용 버퍼: 40주 추가 (총 최대 92주 수집 시도)
    ─ 이격도(%) = (종가 - MAn) / MAn × 100
      · MA10이격도 > +10%: 단기 과매수 경고
      · MA10이격도 < -10%: 단기 과매도 (반등 구간)
      · MA40이격도는 중장기 추세 위치 판별
    ─ output2 응답 필드:
      stck_bsop_date(날짜), stck_oprc(시가), stck_hgpr(고가),
      stck_lwpr(저가), stck_clpr(종가), acml_vol(거래량),
      acml_tr_pbmn(거래대금), prdy_vrss(전주대비), prdy_ctrt(전주대비율),
      hts_frgn_ehrt(외인소진율)
    """
    url     = f"{BASE_URL}/uapi/domestic-stock/v1/quotations/inquire-daily-itemchartprice"
    headers = make_headers(token, "FHKST03010100")
    today   = datetime.today()

    # 92주 버퍼 포함 2구간 (각 최대 100건이므로 46주씩 분할)
    fetch_ranges = [
        (
            (today - timedelta(weeks=92)).strftime("%Y%m%d"),
            (today - timedelta(weeks=47)).strftime("%Y%m%d"),
            "구간A(원거리92~47주)"
        ),
        (
            (today - timedelta(weeks=46)).strftime("%Y%m%d"),
            today.strftime("%Y%m%d"),
            "구간B(최근46주~이번주)"
        ),
    ]

    print(f"\n📊 [탭1] 주봉 OHLCV 수집 (종목: {stock_code}) ...")

    seen_dates = set()
    raw_rows   = []

    def _fetch(start_date, end_date, label):
        print(f"   [{label}] {start_date} ~ {end_date}", end=" ")
        params = {
            "FID_COND_MRKT_DIV_CODE": "J",
            "FID_INPUT_ISCD":         stock_code,
            "FID_INPUT_DATE_1":       start_date,
            "FID_INPUT_DATE_2":       end_date,
            "FID_PERIOD_DIV_CODE":    "W",   # ★ 주봉
            "FID_ORG_ADJ_PRC":        "0"
        }
        try:
            res   = safe_get(url, headers=headers, params=params)
            data  = res.json()
            items = data.get("output2", [])
            print(f"→ {len(items)}건 ✅")
            return items
        except Exception as e:
            print(f"❌ 오류: {e}")
            return []

    with ThreadPoolExecutor(max_workers=2) as ex:
        f_a = ex.submit(_fetch, *fetch_ranges[0])
        f_b = ex.submit(_fetch, *fetch_ranges[1])
        for items in [f_a.result(), f_b.result()]:
            for item in items:
                d = item.get("stck_bsop_date")
                if not d or d in seen_dates:
                    continue
                seen_dates.add(d)
                raw_rows.append({
                    "날짜":       d,
                    "시가":       to_int(item.get("stck_oprc",     0)),
                    "고가":       to_int(item.get("stck_hgpr",     0)),
                    "저가":       to_int(item.get("stck_lwpr",     0)),
                    "종가":       to_int(item.get("stck_clpr",     0)),
                    "거래량":     to_int(item.get("acml_vol",      0)),
                    # acml_tr_pbmn 단위: 원 (일봉/주봉 공통)
                    "거래대금":   to_int(item.get("acml_tr_pbmn",  0)),
                    "전주대비":   to_int(item.get("prdy_vrss",     0)),
                    "전주대비율": to_float(item.get("prdy_ctrt",   0)),
                    "외인소진율": to_float(item.get("hts_frgn_ehrt", 0)),
                })

    raw_rows.sort(key=lambda x: x["날짜"])

    # ── MA 및 이격도 계산 (전체 버퍼 포함 순서로 수행) ───────────────
    closes = [r["종가"] for r in raw_rows]

    def _sma(idx, window):
        if idx < window - 1:
            return None
        return sum(closes[idx - window + 1: idx + 1]) / window

    def _divergence(close, ma):
        if ma is None or ma == 0:
            return None
        return round((close - ma) / ma * 100, 2)

    for i, row in enumerate(raw_rows):
        ma10 = _sma(i, 10)
        ma40 = _sma(i, 40)
        row["MA10"]       = round(ma10, 0) if ma10 else None
        row["MA40"]       = round(ma40, 0) if ma40 else None
        row["MA10이격도"] = _divergence(row["종가"], ma10)
        row["MA40이격도"] = _divergence(row["종가"], ma40)

    # 최근 52주만 유지
    TARGET = 52
    if len(raw_rows) > TARGET:
        raw_rows = raw_rows[-TARGET:]

    raw_rows.sort(key=lambda x: x["날짜"], reverse=True)  # 내림차순 (최신 상단)

    if raw_rows:
        print(f"  → 주봉 마스터: {len(raw_rows)}주 "
              f"({raw_rows[-1]['날짜']} ~ {raw_rows[0]['날짜']})")
    return raw_rows


def save_weekly_ohlcv_excel(rows, wb):
    ws = wb.create_sheet("주봉_OHLCV_이격도")

    headers = [
        "주차(마지막영업일)", "시가", "고가", "저가", "종가",
        "거래량", "거래대금(원)", "전주대비(원)", "전주대비율(%)",
        "MA10(10주)", "MA40(40주)", "MA10이격도(%)", "MA40이격도(%)", "외인소진율(%)"
    ]
    col_widths = [20, 12, 12, 12, 12, 14, 18, 14, 16, 14, 14, 14, 14, 14]
    set_header(ws, headers, col_widths)

    # 이격도 색상 구간
    c_strong_pos = PatternFill("solid", start_color="375623")  # 진초록 (>+10%)
    c_pos        = PatternFill("solid", start_color="C6EFCE")  # 연초록 (0~+10%)
    c_neg        = PatternFill("solid", start_color="FFC7CE")  # 연빨강 (-10~0%)
    c_strong_neg = PatternFill("solid", start_color="9C0006")  # 진빨강 (<-10%)

    for i, row in enumerate(rows, 2):
        ws.append([
            row["날짜"],
            row["시가"], row["고가"], row["저가"], row["종가"],
            row["거래량"], row["거래대금"],
            row["전주대비"], row["전주대비율"],
            row["MA10"], row["MA40"],
            row["MA10이격도"], row["MA40이격도"],
            row["외인소진율"],
        ])
        for col in range(1, len(headers) + 1):
            cell = ws.cell(i, col)
            cell.font      = BODY_FONT
            cell.border    = BORDER
            cell.alignment = CENTER_ALIGN

            if col in (2, 3, 4, 5, 10, 11):
                cell.number_format = "#,##0"
            elif col == 6:
                cell.number_format = "#,##0"
            elif col == 7:
                cell.number_format = "#,##0"
            elif col == 8:
                cell.number_format = "#,##0;[Red]-#,##0"
            elif col in (9, 14):
                cell.number_format = "0.00"
            elif col in (12, 13):
                cell.number_format = "0.00"
                val = cell.value
                if val is not None:
                    if val > 10:
                        cell.fill = c_strong_pos
                        cell.font = Font(name="맑은 고딕", size=10, bold=True, color="FFFFFF")
                    elif val > 0:
                        cell.fill = c_pos
                    elif val < -10:
                        cell.fill = c_strong_neg
                        cell.font = Font(name="맑은 고딕", size=10, bold=True, color="FFFFFF")
                    elif val < 0:
                        cell.fill = c_neg

        # 전주대비 색상
        cell_vrss = ws.cell(i, 8)
        if isinstance(cell_vrss.value, (int, float)) and col != 12 and col != 13:
            cell_vrss.fill = POS_FILL if cell_vrss.value >= 0 else NEG_FILL

    ws.freeze_panes = "A2"
    print("  ✅ [탭1] 주봉_OHLCV_이격도 저장 완료")


# =============================================
# 탭2. 주간 수급 누적
# =============================================
def get_weekly_investor(token, stock_code, weekly_rows):
    """
    TR: FHPTJ04160001 (종목별 투자자매매동향 일별)
    ─ 파라미터:
        FID_INPUT_DATE_1: 조회기준일 (해당일 기준 역순 약 30건)
        FID_ETC_CLS_CODE: "1" (★ API 스펙상 "1" 입력 필수)
        FID_ORG_ADJ_PRC:  공란
    ─ 단위: 순매수 대금 = 백만원
    ─ 수급 cursor 방식으로 52주 × 5영업일 ≈ 260일 수집

    ─ 주간 집계 방법:
        주봉 마지막 영업일(금요일)을 기준으로
        해당 주의 월~금(최대 5영업일) 일봉 수급을 합산

    ─ 주간 VWAP(거래량 가중 평균 체결가):
        = Σ(일별 종가 × 일별 거래량) / Σ(일별 거래량)
        일봉 OHLCV API(D)로 별도 수집 후 계산

    ─ 외인/기관 주간 평균 매수단가 추정:
        = |주간 순매수대금(백만원)| × 1_000_000 / 추정 거래량
        추정 거래량 = |순매수대금| / 주봉 VWAP
        → 단순 추정이므로 참고 수치로 활용
    """
    url_inv   = f"{BASE_URL}/uapi/domestic-stock/v1/quotations/investor-trade-by-stock-daily"
    url_ohlcv = f"{BASE_URL}/uapi/domestic-stock/v1/quotations/inquire-daily-itemchartprice"
    h_inv     = make_headers(token, "FHPTJ04160001")
    h_ohlcv   = make_headers(token, "FHKST03010100")

    if not weekly_rows:
        return []

    wdates       = sorted([r["날짜"] for r in weekly_rows])
    master_start = (datetime.strptime(wdates[0], "%Y%m%d") - timedelta(days=7)).strftime("%Y%m%d")
    master_end   = wdates[-1]

    print(f"\n📊 [탭2] 주간 수급 누적 수집 (종목: {stock_code})")
    print(f"   수집 범위: {master_start} ~ {master_end}")

    # ── 2-1. 일봉 수급 cursor 방식 수집 ──────────────────────────────
    def _collect_investor_daily():
        all_rows   = []
        seen_dates = set()
        cursor     = master_end
        page       = 0

        while True:
            params = {
                "FID_COND_MRKT_DIV_CODE": "J",
                "FID_INPUT_ISCD":         stock_code,
                "FID_INPUT_DATE_1":       cursor,
                "FID_ORG_ADJ_PRC":        "",       # 공란
                "FID_ETC_CLS_CODE":       "1",      # ★ 스펙: "1" 입력
            }
            try:
                res  = safe_get(url_inv, headers=h_inv, params=params)
                data = res.json()
            except Exception as e:
                print(f"   ❌ 수급 API 오류: {e}")
                break

            if data.get("rt_cd") != "0":
                msg = data.get("msg1", "")
                # 당일 15:40 이전 에러: cursor 하루 당겨서 재시도
                if ("TIME" in msg.upper() or "15:40" in msg) and cursor == master_end:
                    prev   = datetime.strptime(cursor, "%Y%m%d") - timedelta(days=1)
                    cursor = prev.strftime("%Y%m%d")
                    print(f"   💡 당일 미집계, {cursor}부터 재시도...")
                    continue
                print(f"   ❌ API 거부: {msg}")
                break

            output2 = data.get("output2", [])
            if not output2:
                break

            oldest_date = None
            added = 0
            for item in output2:
                d = item.get("stck_bsop_date", "")
                if not d:
                    continue
                oldest_date = d
                if d in seen_dates or d < master_start:
                    continue
                seen_dates.add(d)
                all_rows.append({
                    "날짜":     d,
                    # 단위: 백만원 (API 스펙 확인)
                    "외국인":   to_int(item.get("frgn_ntby_tr_pbmn",    0)),
                    "기관합계": to_int(item.get("orgn_ntby_tr_pbmn",    0)),
                    "연기금":   to_int(item.get("ivtr_ntby_tr_pbmn",    0)),
                    "금융투자": to_int(item.get("scrt_ntby_tr_pbmn",    0)),
                    "보험":     to_int(item.get("insu_ntby_tr_pbmn",    0)),
                    "투신":     to_int(item.get("fund_ntby_tr_pbmn",    0)),
                    "사모펀드": to_int(item.get("pe_fund_ntby_tr_pbmn", 0)),
                    "기타법인": to_int(item.get("etc_corp_ntby_tr_pbmn",0)),
                    # 일봉 OHLCV (VWAP 계산용)
                    "종가":     to_int(item.get("stck_clpr",  0)),
                    "거래량":   to_int(item.get("acml_vol",   0)),
                })
                added += 1

            page += 1
            print(f"   [페이지{page}] cursor={cursor} → {added}건 추가 (누적 {len(all_rows)}건)")

            if oldest_date and oldest_date <= master_start:
                break
            if oldest_date is None:
                break

            prev   = datetime.strptime(oldest_date, "%Y%m%d") - timedelta(days=1)
            cursor = prev.strftime("%Y%m%d")
            if cursor < master_start:
                break
            time.sleep(0.2)

        print(f"   일봉 수급 수집 완료: {len(all_rows)}건")
        return {r["날짜"]: r for r in all_rows}

    inv_map = _collect_investor_daily()

    # ── 2-2. 주차별 집계 ──────────────────────────────────────────────
    result = []
    for wrow in sorted(weekly_rows, key=lambda x: x["날짜"], reverse=True):
        week_end_date = wrow["날짜"]
        week_end_dt   = datetime.strptime(week_end_date, "%Y%m%d")
        week_start_dt = week_end_dt - timedelta(days=6)
        wstart_str    = week_start_dt.strftime("%Y%m%d")

        # 해당 주차 일봉 날짜 목록
        week_dates = sorted([
            d for d in inv_map.keys()
            if wstart_str <= d <= week_end_date
        ])

        # 수급 누적 (단위: 백만원)
        acc = {k: 0 for k in
               ("외국인", "기관합계", "연기금", "금융투자", "보험", "투신", "사모펀드", "기타법인")}
        for d in week_dates:
            for k in acc:
                acc[k] += inv_map[d].get(k, 0)

        # 주간 VWAP (종가×거래량 가중평균)
        # output2에 stck_clpr, acml_vol 포함되어 있어 추가 API 불필요
        vol_sum   = sum(inv_map[d]["거래량"] for d in week_dates if d in inv_map)
        price_vol = sum(
            inv_map[d]["종가"] * inv_map[d]["거래량"]
            for d in week_dates if d in inv_map
        )
        vwap = round(price_vol / vol_sum, 0) if vol_sum > 0 else 0

        # 현재가(주봉 종가) vs 주간 VWAP
        close = wrow["종가"]
        if vwap > 0:
            vwap_vs = "▲위" if close >= vwap else "▽아래"
        else:
            vwap_vs = "-"

        # 외인 주간 평균 매수단가 추정
        # 단위 환산: 백만원 × 1,000,000 = 원
        frgn_won   = abs(acc["외국인"]) * 1_000_000
        frgn_vol_e = frgn_won / vwap if vwap > 0 and frgn_won > 0 else 0
        frgn_avg   = int(frgn_won / frgn_vol_e) if frgn_vol_e > 0 else int(vwap)

        # 기관 주간 평균 매수단가 추정
        orgn_won   = abs(acc["기관합계"]) * 1_000_000
        orgn_vol_e = orgn_won / vwap if vwap > 0 and orgn_won > 0 else 0
        orgn_avg   = int(orgn_won / orgn_vol_e) if orgn_vol_e > 0 else int(vwap)

        result.append({
            "주차(마지막영업일)":         week_end_date,
            "영업일수":                  len(week_dates),
            "외인누적순매수(백만원)":     acc["외국인"],
            "기관누적순매수(백만원)":     acc["기관합계"],
            "연기금누적순매수(백만원)":   acc["연기금"],
            "금융투자누적순매수(백만원)": acc["금융투자"],
            "보험누적순매수(백만원)":     acc["보험"],
            "투신누적순매수(백만원)":     acc["투신"],
            "사모펀드누적순매수(백만원)": acc["사모펀드"],
            "기타법인누적순매수(백만원)": acc["기타법인"],
            "주간VWAP(원)":              int(vwap),
            "주봉종가(원)":              close,
            "현재가vs주간VWAP":          vwap_vs,
            "외인추정평균단가(원)":       frgn_avg,
            "기관추정평균단가(원)":       orgn_avg,
        })

    print(f"  → 주간 수급 집계 완료: {len(result)}주")
    return result


def save_weekly_investor_excel(rows, wb):
    ws = wb.create_sheet("주간_수급_누적")

    headers = [
        "주차(마지막영업일)", "영업일수",
        "외인순매수(백만원)", "기관순매수(백만원)", "연기금순매수(백만원)",
        "금투순매수(백만원)", "보험순매수(백만원)", "투신순매수(백만원)",
        "사모펀드순매수(백만원)", "기타법인순매수(백만원)",
        "주간VWAP(원)", "주봉종가(원)", "현재가vs주간VWAP",
        "외인추정평균단가(원)", "기관추정평균단가(원)"
    ]
    col_widths = [20, 10, 20, 20, 20, 18, 14, 14, 20, 20, 16, 16, 16, 20, 20]
    set_header(ws, headers, col_widths)

    num_fmt = "#,##0;[Red]-#,##0"
    for i, row in enumerate(rows, 2):
        ws.append([
            row["주차(마지막영업일)"], row["영업일수"],
            row["외인누적순매수(백만원)"], row["기관누적순매수(백만원)"],
            row["연기금누적순매수(백만원)"], row["금융투자누적순매수(백만원)"],
            row["보험누적순매수(백만원)"], row["투신누적순매수(백만원)"],
            row["사모펀드누적순매수(백만원)"], row["기타법인누적순매수(백만원)"],
            row["주간VWAP(원)"], row["주봉종가(원)"], row["현재가vs주간VWAP"],
            row["외인추정평균단가(원)"], row["기관추정평균단가(원)"],
        ])
        for col in range(1, len(headers) + 1):
            cell = ws.cell(i, col)
            cell.font      = BODY_FONT
            cell.border    = BORDER
            cell.alignment = CENTER_ALIGN

            if col in range(3, 11):
                cell.number_format = num_fmt
                if isinstance(cell.value, (int, float)):
                    cell.fill = POS_FILL if cell.value >= 0 else NEG_FILL
            elif col in (11, 12, 14, 15):
                cell.number_format = "#,##0"
            elif col == 13:
                if cell.value == "▲위":
                    cell.fill = POS_FILL
                elif cell.value == "▽아래":
                    cell.fill = NEG_FILL

    ws.freeze_panes = "A2"
    print("  ✅ [탭2] 주간_수급_누적 저장 완료")


# =============================================
# 탭3. 신용잔고 주간 증감률
# =============================================
def get_weekly_credit(token, stock_code, weekly_rows):
    """
    TR: FHPST04760000 (국내주식 신용잔고 일별추이)
    ─ 파라미터:
        fid_cond_mrkt_div_code: "J"
        fid_cond_scr_div_code:  "20476"  (Unique key, 스펙 필수)
        fid_input_iscd:         종목코드
        fid_input_date_1:       결제일자 (cursor 방식)
    ─ 1회 최대 30건
    ─ 단위: 금액 만원 (whol_loan_rmnd_amt)
    ─ 주봉 금요일에 가장 가까운 일봉값으로 샘플링

    신용 다이버전스 판별 기준:
      · 융자잔고비율 주간 증가 > +0.1%p + 주봉 등락 ≤ 0% → ⚠️ 경고
      · 융자잔고비율 주간 급증 > +0.3%p                    → 📌 급증 주의
      · 융자잔고비율 주간 감소 < -0.2%p                    → ✅ 청산 진행
    """
    url = f"{BASE_URL}/uapi/domestic-stock/v1/quotations/daily-credit-balance"
    h   = make_headers(token, "FHPST04760000")

    if not weekly_rows:
        return []

    wdates     = sorted([r["날짜"] for r in weekly_rows])
    start_date = (datetime.strptime(wdates[0], "%Y%m%d") - timedelta(days=7)).strftime("%Y%m%d")
    end_date   = wdates[-1]

    print(f"\n📊 [탭3] 신용잔고 주간 추이 수집 (종목: {stock_code}) ...")

    # cursor 방식 수집 (최대 30건/회, 소문자 파라미터 사용)
    all_credit = []
    seen_dates = set()
    cursor     = end_date

    while True:
        params = {
            "fid_cond_mrkt_div_code": "J",
            "fid_cond_scr_div_code":  "20476",   # ★ 스펙 Unique key
            "fid_input_iscd":         stock_code,
            "fid_input_date_1":       cursor,
        }
        try:
            res  = safe_get(url, headers=h, params=params)
            data = res.json()
        except Exception as e:
            print(f"   ❌ 신용잔고 오류: {e}")
            break

        if data.get("rt_cd") != "0":
            break

        output = data.get("output", [])
        if not output:
            break

        oldest_dt = None
        for item in output:
            d = item.get("deal_date", "")
            if not d:
                continue
            oldest_dt = d
            if d in seen_dates or d < start_date:
                continue
            seen_dates.add(d)
            all_credit.append({
                "날짜":              d,
                "융자잔고(주)":      to_int(item.get("whol_loan_rmnd_stcn",  0)),
                # 단위: 만원 (스펙 확인)
                "융자잔고금액(만원)": to_int(item.get("whol_loan_rmnd_amt",   0)),
                "융자잔고비율(%)":   to_float(item.get("whol_loan_rmnd_rate", 0)),
                "대주잔고(주)":      to_int(item.get("whol_stln_rmnd_stcn",  0)),
                "대주잔고비율(%)":   to_float(item.get("whol_stln_rmnd_rate", 0)),
            })

        if len(output) < 30 or (oldest_dt and oldest_dt <= start_date):
            break
        if oldest_dt is None:
            break

        prev   = datetime.strptime(oldest_dt, "%Y%m%d") - timedelta(days=1)
        cursor = prev.strftime("%Y%m%d")
        if cursor < start_date:
            break
        time.sleep(0.2)

    credit_map  = {r["날짜"]: r for r in all_credit}
    weekly_map  = {r["날짜"]: r for r in weekly_rows}
    print(f"   일별 신용잔고 수집: {len(credit_map)}건")

    # ── 주봉 샘플링 + 다이버전스 판별 ─────────────────────────────────
    sorted_wdates = sorted(wdates, reverse=True)
    result = []

    for i, wdate in enumerate(sorted_wdates):
        wend_dt   = datetime.strptime(wdate, "%Y%m%d")
        wstart_dt = wend_dt - timedelta(days=6)
        wstart_str = wstart_dt.strftime("%Y%m%d")

        # 해당 주 내 가장 최근 신용잔고 데이터 탐색
        credit_val = None
        for delta in range(7):
            cand = (wend_dt - timedelta(days=delta)).strftime("%Y%m%d")
            if cand in credit_map and cand >= wstart_str:
                credit_val = credit_map[cand]
                break

        if credit_val is None:
            credit_val = {
                "융자잔고(주)": None,
                "융자잔고금액(만원)": None,
                "융자잔고비율(%)": None,
                "대주잔고(주)": None,
                "대주잔고비율(%)": None,
            }

        # 전주 데이터
        prev_rate = None
        if i + 1 < len(sorted_wdates):
            prev_wdate  = sorted_wdates[i + 1]
            prev_wend   = datetime.strptime(prev_wdate, "%Y%m%d")
            prev_wstart = (prev_wend - timedelta(days=6)).strftime("%Y%m%d")
            for delta in range(7):
                cand = (prev_wend - timedelta(days=delta)).strftime("%Y%m%d")
                if cand in credit_map and cand >= prev_wstart:
                    prev_rate = credit_map[cand]["융자잔고비율(%)"]
                    break

        curr_rate = credit_val["융자잔고비율(%)"]
        if curr_rate is not None and prev_rate is not None:
            rate_chg = round(curr_rate - prev_rate, 3)
        else:
            rate_chg = None

        # 다이버전스 판별
        week_change = weekly_map.get(wdate, {}).get("전주대비율", None)
        divergence  = "-"
        if rate_chg is not None and week_change is not None:
            if rate_chg > 0.1 and week_change <= 0:
                divergence = "⚠️ 신용↑ 주가↓ 경고"
            elif rate_chg > 0.3:
                divergence = "📌 신용 급증"
            elif rate_chg < -0.2:
                divergence = "✅ 신용 감소"

        result.append({
            "주차(마지막영업일)":    wdate,
            "융자잔고(주)":         credit_val["융자잔고(주)"],
            "융자잔고금액(만원)":    credit_val["융자잔고금액(만원)"],
            "융자잔고비율(%)":       curr_rate,
            "전주대비잔고비율변화":   rate_chg,
            "주봉종가등락율(%)":     week_change,
            "신용다이버전스":        divergence,
            "대주잔고(주)":         credit_val["대주잔고(주)"],
            "대주잔고비율(%)":       credit_val["대주잔고비율(%)"],
        })

    print(f"  → 주간 신용잔고 집계 완료: {len(result)}주")
    return result


def save_weekly_credit_excel(rows, wb):
    ws = wb.create_sheet("신용잔고_주간증감")

    headers = [
        "주차(마지막영업일)", "융자잔고(주)", "융자잔고금액(만원)",
        "융자잔고비율(%)", "전주대비잔고비율변화(%p)", "주봉종가등락율(%)",
        "신용다이버전스", "대주잔고(주)", "대주잔고비율(%)"
    ]
    col_widths = [20, 16, 18, 16, 22, 18, 24, 14, 16]
    set_header(ws, headers, col_widths)

    warn_fill  = PatternFill("solid", start_color="FFE699")   # 노랑 (급증)
    alert_fill = PatternFill("solid", start_color="FF7043")   # 주황 (경고)
    good_fill  = PatternFill("solid", start_color="C6EFCE")   # 초록 (감소)

    for i, row in enumerate(rows, 2):
        ws.append([
            row["주차(마지막영업일)"],
            row["융자잔고(주)"],
            row["융자잔고금액(만원)"],
            row["융자잔고비율(%)"],
            row["전주대비잔고비율변화"],
            row["주봉종가등락율(%)"],
            row["신용다이버전스"],
            row["대주잔고(주)"],
            row["대주잔고비율(%)"],
        ])
        for col in range(1, len(headers) + 1):
            cell = ws.cell(i, col)
            cell.font      = BODY_FONT
            cell.border    = BORDER
            cell.alignment = CENTER_ALIGN

            if col in (2, 3, 8):
                cell.number_format = "#,##0"
            elif col in (4, 5, 6, 9):
                cell.number_format = "0.000"

            if col == 5 and isinstance(cell.value, (int, float)):
                cell.fill = POS_FILL if cell.value >= 0 else NEG_FILL
            if col == 6 and isinstance(cell.value, (int, float)):
                cell.fill = POS_FILL if cell.value >= 0 else NEG_FILL

        div_cell = ws.cell(i, 7)
        div_str  = str(div_cell.value) if div_cell.value else ""
        if "경고" in div_str:
            div_cell.fill = alert_fill
            div_cell.font = Font(name="맑은 고딕", size=10, bold=True, color="FF0000")
        elif "급증" in div_str:
            div_cell.fill = warn_fill
        elif "감소" in div_str:
            div_cell.fill = good_fill

    ws.freeze_panes = "A2"
    print("  ✅ [탭3] 신용잔고_주간증감 저장 완료")


# =============================================
# 탭4. 주간 거래대금 분석 (추가 API 없음)
# =============================================
def build_weekly_volume_analysis(weekly_rows):
    """
    탭1 주봉 OHLCV 데이터 재활용 (추가 API 호출 없음)

    분석 항목:
      · 전주 대비 거래대금 변화율(%)
      · 최근 4주 평균 거래대금 대비 이번 주 배율
      · 52주 내 거래대금 백분위 (하위 몇 % 위치)
      · 폭증/감소 신호 (4주 평균의 2배 이상 = 🔥 폭증)
    """
    print(f"\n📊 [탭4] 주간 거래대금 분석 ...")

    sorted_rows = sorted(weekly_rows, key=lambda x: x["날짜"], reverse=True)

    # 52주 내 거래대금 백분위 계산용
    all_vols_sorted = sorted([r["거래대금"] for r in sorted_rows if r["거래대금"] > 0])
    total_count     = len(all_vols_sorted)

    result = []
    for i, row in enumerate(sorted_rows):
        vol  = row["거래대금"]

        # 전주 대비 증감율
        if i + 1 < len(sorted_rows):
            prev_vol    = sorted_rows[i + 1]["거래대금"]
            vol_chg_pct = round((vol - prev_vol) / prev_vol * 100, 1) if prev_vol > 0 else None
        else:
            vol_chg_pct = None

        # 4주 이동평균 (직전 4주 = i+1 ~ i+4)
        prev_vols = [
            sorted_rows[j]["거래대금"]
            for j in range(i + 1, min(i + 5, len(sorted_rows)))
            if sorted_rows[j]["거래대금"] > 0
        ]
        ma4_vol   = round(sum(prev_vols) / len(prev_vols), 0) if prev_vols else None
        vol_ratio = round(vol / ma4_vol, 2) if ma4_vol and ma4_vol > 0 else None

        # 52주 백분위 (낮을수록 거래 적음, 높을수록 많음)
        if vol > 0 and all_vols_sorted:
            rank = all_vols_sorted.index(min(all_vols_sorted, key=lambda x: abs(x - vol)))
            pct  = round(rank / total_count * 100, 1)
        else:
            pct = 0.0

        # 폭증/감소 신호
        if vol_ratio and vol_ratio >= 2.0:
            surge = f"🔥 {vol_ratio:.1f}배 폭증"
        elif vol_ratio and vol_ratio >= 1.5:
            surge = f"📈 {vol_ratio:.1f}배 증가"
        elif vol_ratio and vol_ratio <= 0.5:
            surge = f"📉 {vol_ratio:.1f}배 감소"
        else:
            surge = "-"

        result.append({
            "주차(마지막영업일)": row["날짜"],
            "주간거래대금(원)":   vol,
            "전주대비증감율(%)":  vol_chg_pct,
            "4주평균거래대금(원)": int(ma4_vol) if ma4_vol else None,
            "4주평균대비배율":    vol_ratio,
            "52주내백분위(%)":   pct,
            "거래대금신호":       surge,
            "주봉종가(원)":       row["종가"],
            "주봉등락율(%)":      row["전주대비율"],
        })

    print(f"  → 주간 거래대금 분석 완료: {len(result)}주")
    return result


def save_weekly_volume_excel(rows, wb):
    ws = wb.create_sheet("주간_거래대금_분석")

    headers = [
        "주차(마지막영업일)", "주간거래대금(원)", "전주대비증감율(%)",
        "4주평균거래대금(원)", "4주평균대비배율", "52주내백분위(%)",
        "거래대금신호", "주봉종가(원)", "주봉등락율(%)"
    ]
    col_widths = [20, 20, 18, 20, 16, 16, 20, 16, 14]
    set_header(ws, headers, col_widths)

    fire_fill = PatternFill("solid", start_color="FF4500")

    for i, row in enumerate(rows, 2):
        ws.append([
            row["주차(마지막영업일)"],
            row["주간거래대금(원)"],
            row["전주대비증감율(%)"],
            row["4주평균거래대금(원)"],
            row["4주평균대비배율"],
            row["52주내백분위(%)"],
            row["거래대금신호"],
            row["주봉종가(원)"],
            row["주봉등락율(%)"],
        ])
        for col in range(1, len(headers) + 1):
            cell = ws.cell(i, col)
            cell.font      = BODY_FONT
            cell.border    = BORDER
            cell.alignment = CENTER_ALIGN

            if col in (2, 4, 8):
                cell.number_format = "#,##0"
            elif col in (3, 6, 9):
                cell.number_format = "0.0"
            elif col == 5:
                cell.number_format = "0.00"

            if col == 3 and isinstance(cell.value, (int, float)):
                cell.fill = POS_FILL if cell.value >= 0 else NEG_FILL
            if col == 9 and isinstance(cell.value, (int, float)):
                cell.fill = POS_FILL if cell.value >= 0 else NEG_FILL

        sig_cell = ws.cell(i, 7)
        if "폭증" in str(sig_cell.value):
            sig_cell.fill = fire_fill
            sig_cell.font = Font(name="맑은 고딕", size=10, bold=True, color="FFFFFF")

    ws.freeze_panes = "A2"
    print("  ✅ [탭4] 주간_거래대금_분석 저장 완료")


# =============================================
# 탭5. 외국인 소진율 추이
# =============================================
def build_weekly_frgn_exhaustion(weekly_rows):
    """
    탭1 주봉 OHLCV의 hts_frgn_ehrt(외인소진율) 재활용

    분석 항목:
      · 주봉 기준 외인소진율(%)
      · 전주 대비 변화(%p)
      · 4주 연속 추세 (↑/↓ N주 연속)
      · 소진율 구간 분류:
          ≥80%: 🔴 극고소진 (외인 추가 매수 여력 제한)
        60~80%: 🟠 고소진
        40~60%: 🟡 중립
          <40%: 🟢 저소진 (매수 여력 충분)
    """
    print(f"\n📊 [탭5] 외인소진율 추이 분석 ...")

    sorted_rows = sorted(weekly_rows, key=lambda x: x["날짜"], reverse=True)
    result      = []

    for i, row in enumerate(sorted_rows):
        frgn_ex = row.get("외인소진율", 0) or 0
        prev_ex = sorted_rows[i + 1].get("외인소진율", 0) if i + 1 < len(sorted_rows) else None
        chg     = round(frgn_ex - prev_ex, 2) if prev_ex is not None else None

        # 4주 연속 추세 계산
        trend_label = "-"
        if prev_ex is not None:
            direction   = 1 if frgn_ex >= prev_ex else -1
            trend_count = 1
            for j in range(i + 1, min(i + 5, len(sorted_rows) - 1)):
                curr_v = sorted_rows[j].get("외인소진율", 0) or 0
                next_v = sorted_rows[j + 1].get("외인소진율", 0) or 0
                if direction == 1 and curr_v >= next_v:
                    trend_count += 1
                elif direction == -1 and curr_v <= next_v:
                    trend_count += 1
                else:
                    break
            trend_label = f"{'↑' if direction == 1 else '↓'} {trend_count}주 연속"

        # 소진율 구간
        if frgn_ex >= 80:
            zone = "🔴 극고소진(≥80%)"
        elif frgn_ex >= 60:
            zone = "🟠 고소진(60~80%)"
        elif frgn_ex >= 40:
            zone = "🟡 중립(40~60%)"
        else:
            zone = "🟢 저소진(<40%)"

        result.append({
            "주차(마지막영업일)": row["날짜"],
            "외인소진율(%)":      frgn_ex,
            "전주대비변화(%p)":   chg,
            "4주추세":            trend_label,
            "소진율구간":          zone,
            "주봉종가(원)":        row.get("종가", 0),
            "주봉등락율(%)":       row.get("전주대비율", 0),
            "MA10이격도(%)":       row.get("MA10이격도"),
            "MA40이격도(%)":       row.get("MA40이격도"),
        })

    print(f"  → 외인소진율 분석 완료: {len(result)}주")
    return result


def save_weekly_frgn_excel(rows, wb):
    ws = wb.create_sheet("외인소진율_추이")

    headers = [
        "주차(마지막영업일)", "외인소진율(%)", "전주대비변화(%p)",
        "4주추세", "소진율구간", "주봉종가(원)",
        "주봉등락율(%)", "MA10이격도(%)", "MA40이격도(%)"
    ]
    col_widths = [20, 14, 16, 16, 22, 16, 14, 14, 14]
    set_header(ws, headers, col_widths)

    red_fill    = PatternFill("solid", start_color="FFB3B3")
    orange_fill = PatternFill("solid", start_color="FFD9B3")
    yellow_fill = PatternFill("solid", start_color="FFF9C4")
    green_fill  = PatternFill("solid", start_color="C8E6C9")
    zone_map    = {"극고소진": red_fill, "고소진": orange_fill, "중립": yellow_fill, "저소진": green_fill}

    for i, row in enumerate(rows, 2):
        ws.append([
            row["주차(마지막영업일)"],
            row["외인소진율(%)"],
            row["전주대비변화(%p)"],
            row["4주추세"],
            row["소진율구간"],
            row["주봉종가(원)"],
            row["주봉등락율(%)"],
            row["MA10이격도(%)"],
            row["MA40이격도(%)"],
        ])
        for col in range(1, len(headers) + 1):
            cell = ws.cell(i, col)
            cell.font      = BODY_FONT
            cell.border    = BORDER
            cell.alignment = CENTER_ALIGN

            if col == 6:
                cell.number_format = "#,##0"
            elif col in (2, 3, 7, 8, 9):
                cell.number_format = "0.00"

            if col == 3 and isinstance(cell.value, (int, float)):
                cell.fill = POS_FILL if cell.value >= 0 else NEG_FILL
            if col == 7 and isinstance(cell.value, (int, float)):
                cell.fill = POS_FILL if cell.value >= 0 else NEG_FILL

        zone_cell = ws.cell(i, 5)
        zone_str  = str(zone_cell.value) if zone_cell.value else ""
        for key, fill in zone_map.items():
            if key in zone_str:
                zone_cell.fill = fill
                break

    ws.freeze_panes = "A2"
    print("  ✅ [탭5] 외인소진율_추이 저장 완료")


# =============================================
# 단일 종목 처리
# =============================================
def process_stock_weekly(token, stock_code, stock_name, output_dir):
    print("\n" + "=" * 60)
    print(f"  처리 중: [{stock_code}] {stock_name}")
    print("=" * 60)

    wb = openpyxl.Workbook()

    # ── 탭1: 주봉 OHLCV + 이격도 (마스터 캘린더) ─────────────────────
    weekly_rows = []
    try:
        weekly_rows = get_weekly_ohlcv(token, stock_code)
        if not weekly_rows:
            print("  ❌ 주봉 데이터가 없어 이 종목을 건너뜁니다.")
            return None
        save_weekly_ohlcv_excel(weekly_rows, wb)
    except Exception as e:
        print(f"  ❌ 주봉 OHLCV 오류: {e}")
        return None

    # ── 탭2: 주간 수급 누적 ──────────────────────────────────────────
    try:
        inv_rows = get_weekly_investor(token, stock_code, weekly_rows)
        if inv_rows:
            save_weekly_investor_excel(inv_rows, wb)
    except Exception as e:
        print(f"  ❌ 주간 수급 오류: {e}")

    # ── 탭3: 신용잔고 주간 증감 ─────────────────────────────────────
    try:
        credit_rows = get_weekly_credit(token, stock_code, weekly_rows)
        if credit_rows:
            save_weekly_credit_excel(credit_rows, wb)
    except Exception as e:
        print(f"  ❌ 신용잔고 오류: {e}")

    # ── 탭4: 주간 거래대금 분석 (API 없음) ──────────────────────────
    try:
        vol_rows = build_weekly_volume_analysis(weekly_rows)
        if vol_rows:
            save_weekly_volume_excel(vol_rows, wb)
    except Exception as e:
        print(f"  ❌ 거래대금 분석 오류: {e}")

    # ── 탭5: 외인 소진율 추이 (API 없음) ────────────────────────────
    try:
        frgn_rows = build_weekly_frgn_exhaustion(weekly_rows)
        if frgn_rows:
            save_weekly_frgn_excel(frgn_rows, wb)
    except Exception as e:
        print(f"  ❌ 외인소진율 오류: {e}")

    # 기본 Sheet 제거
    if "Sheet" in wb.sheetnames and len(wb.sheetnames) > 1:
        wb.remove(wb["Sheet"])

    timestamp  = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_name  = stock_name.replace("/", "_").replace("\\", "_")
    final_path = os.path.join(output_dir, f"{stock_code}_{safe_name}_주간_{timestamp}.xlsx")

    try:
        wb.save(final_path)
        print(f"\n  ✅ 저장 완료: {os.path.abspath(final_path)}")
        return final_path
    except Exception as e:
        print(f"\n  ❌ 파일 저장 실패: {e}")
        return None


# =============================================
# 메인 실행부
# =============================================
if __name__ == "__main__":
    print("=" * 60)
    print("  한투 오픈API - 주간 데이터 수집 (tickers.json 기반)")
    print("  수집 탭: 주봉OHLCV+이격도 / 주간수급 / 신용잔고 / 거래대금 / 외인소진율")
    print("=" * 60)

    TICKERS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "tickers.json")
    if not os.path.exists(TICKERS_FILE):
        print(f"❌ 티커 파일을 찾을 수 없습니다: {TICKERS_FILE}")
        exit()

    with open(TICKERS_FILE, "r", encoding="utf-8") as f:
        raw_tickers = json.load(f)

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

    if not tickers:
        print("❌ 처리할 종목이 없습니다.")
        exit()

    print(f"\n📋 총 {len(tickers)}개 종목:")
    for i, t in enumerate(tickers, 1):
        print(f"   {i:2d}. [{t['code']}] {t['name']}")

    # 저장 폴더: WeeklyData/YYYYMMDD
    run_date   = datetime.now().strftime("%Y%m%d")
    output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "WeeklyData", run_date)
    os.makedirs(output_dir, exist_ok=True)
    print(f"\n📁 출력 폴더: {os.path.abspath(output_dir)}")

    try:
        token = get_access_token()
    except Exception as e:
        print(f"❌ 토큰 오류: {e}")
        exit()

    results = {"성공": [], "실패": []}

    for idx, ticker in enumerate(tickers, 1):
        print(f"\n{'━' * 60}")
        print(f"  [{idx}/{len(tickers)}] {ticker['name']} ({ticker['code']}) 처리 시작")
        print(f"{'━' * 60}")

        saved = process_stock_weekly(token, ticker["code"], ticker["name"], output_dir)
        if saved:
            results["성공"].append(f"[{ticker['code']}] {ticker['name']}")
        else:
            results["실패"].append(f"[{ticker['code']}] {ticker['name']}")

        if idx < len(tickers):
            print(f"\n  ⏳ 다음 종목 전 3초 대기...")
            time.sleep(3)

    print(f"\n\n{'=' * 60}")
    print("  📊 전체 처리 결과 요약")
    print(f"{'=' * 60}")
    print(f"  ✅ 성공 ({len(results['성공'])}건): {', '.join(results['성공']) or '없음'}")
    print(f"  ❌ 실패 ({len(results['실패'])}건): {', '.join(results['실패']) or '없음'}")
    print(f"\n  📁 저장 위치: {os.path.abspath(output_dir)}")
    print("=" * 60)
