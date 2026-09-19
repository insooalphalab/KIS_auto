"""
한국투자증권 OpenAPI - 실시간 시장 데이터 수집기 (경량화판 v2)
==================================================================
[2026-09 최적화 변경]
  · 6개 모듈 → 4개 모듈로 축소 (분봉 / 체결강도 / 외인기관추정 / 프로그램매매)
    - 회원사동향, 매물대는 장중 시간 단위로 크게 변하지 않는 스냅샷 성격이라
      한투API_일별데이터.py 쪽으로 이관(하루 1회만 수집)했다.
  · 체결강도 수집 목표를 100건 → 40건으로 축소
    - 어차피 분석 단계에서 최근 일부만 사용했던 항목이라 원천에서부터 줄였다.
  · 하루 실행 횟수는 이 파일이 아니라 Windows 작업 스케줄러(원래 7회 →
    3회: 10:00 / 13:00 / 15:20)에서 결정한다. setup_tasks.ps1 참고.

• tickers.json 에서 종목 리스트 로드
• 날짜 및 실행 시간별 하위 폴더 자동 생성 (MarketData/YYYYMMDD/HHMM/)
• 파일명: MarketData_{티커}_{종목명}_{HHMM}.xlsx
• 수집 완료 후 해당 시간 폴더 내에 통합 collection_summary_YYYYMMDD_HHMM.json 생성
"""

import requests
import pandas as pd
from datetime import datetime, timedelta
import json
import time
import os
import sys
import logging

# ──────────────────────────────────────────────
# 0. 로깅 설정
# ──────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger(__name__)

# ──────────────────────────────────────────────
# 1. 기본 설정 (사용자님의 Key값과 토큰 유지)
# ──────────────────────────────────────────────
from kis_config import require

DOMAIN      = "https://openapi.koreainvestment.com:9443"
APP_KEY     = require("KIS_APP_KEY")      # .env 에서 로드 (kis_config.py)
APP_SECRET  = require("KIS_APP_SECRET")
ACCESS_TOKEN = "YOUR_ACCESS_TOKEN"  # 실제 발급받은 액세스 토큰 입력

# 체결강도 목표 건수 (기존 100 → 40, 원천에서부터 다운샘플링)
TICK_STRENGTH_TARGET = 40

def get_headers(tr_id):
    return {
        "content-type": "application/json; charset=utf-8",
        "authorization": f"Bearer {ACCESS_TOKEN}",
        "appkey": APP_KEY,
        "appsecret": APP_SECRET,
        "tr_id": tr_id
    }

# 공통 컬럼명 매핑 딕셔너리
COL_MAP = {
    'stck_bsop_date': '일자',
    'stck_prpr': '현재가', 'stck_oprc': '시가', 'stck_hgpr': '최고가', 'stck_lwpr': '최저가', 'stck_clpr': '종가',
    'cntg_vol': '체결거래량', 'acml_vol': '누적거래량', 'acml_tr_pbmn': '누적거래대금',
    'prdy_vrss': '전일대비', 'prdy_vrss_sign': '대비부호', 'prdy_ctrt': '등락율',
    'askp': '매도호가', 'bidp': '매수호가', 'tday_rltv': '체결강도', 'cnqn': '체결건수',
    'frgn_fake_ntby_qty': '외국인순매수(추정)', 'orgn_fake_ntby_qty': '기관순매수(추정)', 'sum_fake_ntby_qty': '외인기관합산순매수',
    'whol_smtn_ntby_qty': '전체프로그램순매수', 'whol_smtn_seln_vol': '전체프로그램매도수량', 'whol_smtn_shnu_vol': '전체프로그램매수수량',
    'whol_ntby_vol_icdc': '전체순매수수량증감', 'whol_smtn_seln_tr_pbmn': '프로그램매도대금', 'whol_smtn_shnu_tr_pbmn': '프로그램매수대금',
    'whol_smtn_ntby_tr_pbmn': '프로그램순매수대금', 'whol_ntby_tr_pbmn_icdc': '전체순매수대금증감',
    'data_rank': '순위', 'acml_vol_rlim': '매물대비중(%)'
}

# ──────────────────────────────────────────────
# 2. 공통 API 통신 및 데이터 처리 함수
# ──────────────────────────────────────────────
def call_api_with_log(api_name, url, headers, params):
    log.info(f"  [요청] {api_name}")
    try:
        res = requests.get(url, headers=headers, params=params)
        if res.status_code != 200:
            log.error(f"    → HTTP 에러 ({res.status_code}): {res.text}")
            return pd.DataFrame()

        data = res.json()
        target_data = []
        if 'output2' in data and isinstance(data['output2'], list): target_data = data['output2']
        elif 'output' in data and isinstance(data['output'], list): target_data = data['output']
        elif 'output1' in data and isinstance(data['output1'], list): target_data = data['output1']
        else:
            if 'output' in data and data['output']: target_data = [data['output']]
            elif 'output1' in data and data['output1']: target_data = [data['output1']]
            elif 'output2' in data and data['output2']: target_data = [data['output2']]

        return pd.DataFrame(target_data)
    except Exception as e:
        log.error(f"    → 시스템 예외 에러: {e}")
        return pd.DataFrame()

def process_dataframe(df):
    if df.empty:
        return pd.DataFrame(["데이터 없음"], columns=["상태"])

    for col in ['stck_cntg_hour', 'bsop_hour']:
        if col in df.columns:
            df[col] = df[col].apply(lambda x: f"{str(x).zfill(6)[:2]}:{str(x).zfill(6)[2:4]}:{str(x).zfill(6)[4:]}" if pd.notnull(x) and str(x).strip() != '' else x)
            time_col = df.pop(col)
            df.insert(0, '시간', time_col)

    if 'stck_bsop_date' in df.columns:
        df['stck_bsop_date'] = df['stck_bsop_date'].apply(
            lambda x: f"{str(x)[:4]}-{str(x)[4:6]}-{str(x)[6:]}" if pd.notnull(x) and len(str(x))==8 else x
        )
        date_col = df.pop('stck_bsop_date')
        insert_loc = 1 if '시간' in df.columns else 0
        df.insert(insert_loc, '일자', date_col)

    if 'bsop_hour_gb' in df.columns:
        time_mapping = {'1': '09:30', '2': '11:20', '3': '13:20', '4': '14:30', '5': '15:30'}
        df['bsop_hour_gb'] = df['bsop_hour_gb'].astype(str).map(time_mapping).fillna(df['bsop_hour_gb'])
        time_col = df.pop('bsop_hour_gb')
        df.insert(0, '집계시간', time_col)

    for col in df.columns:
        if col not in ['시간', '집계시간', '일자'] and not col.endswith('name') and not col.endswith('isnm'):
            try:
                temp = df[col].replace(r'^\s*$', float('NaN'), regex=True)
                df[col] = pd.to_numeric(temp)
            except Exception:
                pass

    df = df.rename(columns=COL_MAP)
    return df

# ──────────────────────────────────────────────
# 3. 모듈별 API 수집 함수 (4종)
#    회원사동향 / 매물대는 한투API_일별데이터.py 로 이관됨
# ──────────────────────────────────────────────
def fetch_1min_chart(ticker, target_time):
    url = f"{DOMAIN}/uapi/domestic-stock/v1/quotations/inquire-time-itemchartprice"
    params = {"FID_ETC_CLS_CODE": "", "FID_COND_MRKT_DIV_CODE": "J", "FID_INPUT_ISCD": ticker, "FID_INPUT_HOUR_1": target_time, "FID_PW_DATA_INCU_YN": "Y"}
    df = call_api_with_log("1. 분봉 차트", url, get_headers("FHKST03010200"), params)
    return process_dataframe(df)

def fetch_tick_strength_safe(ticker, target_time, target_count=TICK_STRENGTH_TARGET):
    """체결강도 수집. 목표건수를 40건으로 낮춰 호출 횟수(최대 4회→최대 2회)와
    다운스트림 데이터량을 원천에서부터 줄였다."""
    url = f"{DOMAIN}/uapi/domestic-stock/v1/quotations/inquire-time-itemconclusion"
    headers = get_headers("FHPST01060000")
    all_data, curr_time = [], target_time

    for i in range(4):
        params = {"FID_COND_MRKT_DIV_CODE": "J", "FID_INPUT_ISCD": ticker, "FID_INPUT_HOUR_1": curr_time}
        df_temp = call_api_with_log(f"2. 체결강도 조각 ({curr_time})", url, headers, params)
        if df_temp.empty or 'stck_cntg_hour' not in df_temp.columns: break
        all_data.append(df_temp)

        if sum(len(d) for d in all_data) >= target_count: break
        last_time_str = str(df_temp['stck_cntg_hour'].iloc[-1])
        try:
            dt_obj = datetime.strptime(last_time_str, "%H%M%S")
            curr_time = (dt_obj - timedelta(seconds=1)).strftime("%H%M%S")
        except:
            curr_time = last_time_str
        time.sleep(0.1)

    if not all_data: return pd.DataFrame(["데이터 없음"], columns=["상태"])
    return process_dataframe(pd.concat(all_data, ignore_index=True).head(target_count))

def fetch_foreign_inst_est(ticker):
    url = f"{DOMAIN}/uapi/domestic-stock/v1/quotations/investor-trend-estimate"
    return process_dataframe(call_api_with_log("3. 외인기관 가집계 추정", url, get_headers("HHPTJ04160200"), {"MKSC_SHRN_ISCD": ticker}))

def fetch_program_trade(ticker):
    url = f"{DOMAIN}/uapi/domestic-stock/v1/quotations/program-trade-by-stock"
    return process_dataframe(call_api_with_log("4. 프로그램 매매 동향", url, get_headers("FHPPG04650101"), {"FID_COND_MRKT_DIV_CODE": "J", "FID_INPUT_ISCD": ticker}))

# ──────────────────────────────────────────────
# 4. JSON 티커 파일 파서 및 폴더 트리 생성 함수
# ──────────────────────────────────────────────
def load_tickers_from_json(file_path="tickers.json"):
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            raw_data = json.load(f)
        parsed = []
        for item in raw_data:
            parts = str(item).split(';')
            if parts[0].strip():
                parsed.append({
                    "code": parts[0].strip(),
                    "name": parts[1].strip() if len(parts) > 1 else "",
                    "status": parts[2].strip() if len(parts) > 2 else "",
                    "weight": parts[3].strip() if len(parts) > 3 else ""
                })
        return parsed
    except FileNotFoundError:
        log.error(f"'{file_path}' 파일을 찾을 수 없습니다.")
        return []

def prepare_directories(base_root="MarketData"):
    """
    MarketData/YYYYMMDD/HHMM/ 형태로 시간 단위 세부 폴더까지 트리 생성
    """
    now = datetime.now()
    date_str = now.strftime("%Y%m%d")
    time_str = now.strftime("%H%M")

    target_folder = os.path.join(base_root, date_str, time_str)
    os.makedirs(target_folder, exist_ok=True)
    return target_folder, date_str, time_str

def build_excel_path(folder, info, target_time):
    filename = f"MarketData_{info['code']}_{info['name']}_{target_time[:4]}.xlsx"
    return os.path.join(folder, filename)

# ──────────────────────────────────────────────
# 5. 메인 런처 엔트리포인트
# ──────────────────────────────────────────────
def run_extractor(tickers_file: str = "tickers.json") -> str:
    now = datetime.now()
    target_time = now.strftime("%H%M%S")

    folder, date_str, time_str = prepare_directories()
    log.info(f"========== 실시간 시장 수집기 구동 [경량화 4모듈 모드] ==========")
    log.info(f"  → 저장 위치: {folder}")

    portfolio = load_tickers_from_json(tickers_file)
    if not portfolio:
        log.error("수집할 자산 티커 리스트가 비어있습니다.")
        return folder

    results_summary = []

    for info in portfolio:
        ticker = info["code"]
        log.info(f"● 종목 프로세싱: {info['name']} ({ticker}) [{info['status']}/비중:{info['weight']}]")

        df_chart = fetch_1min_chart(ticker, target_time)
        df_tick  = fetch_tick_strength_safe(ticker, target_time)
        df_est   = fetch_foreign_inst_est(ticker)
        df_prog  = fetch_program_trade(ticker)

        excel_path = build_excel_path(folder, info, target_time)
        with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
            df_chart.to_excel(writer, sheet_name='1.분봉',        index=False)
            df_tick.to_excel( writer, sheet_name='2.체결강도',    index=False)
            df_est.to_excel(  writer, sheet_name='3.외인기관추정', index=False)
            df_prog.to_excel( writer, sheet_name='4.프로그램매매', index=False)

        log.info(f"  ▶ 저장 완료: {os.path.basename(excel_path)}")
        results_summary.append({
            "ticker": ticker,
            "name": info["name"],
            "status": info["status"],
            "weight": info["weight"],
            "file": excel_path,
            "collected_at": now.isoformat(),
        })

        time.sleep(1)  # Rate limit 방지

    summary_path = os.path.join(folder, f"collection_summary_{date_str}_{time_str}.json")
    with open(summary_path, 'w', encoding='utf-8') as f:
        json.dump(results_summary, f, ensure_ascii=False, indent=4)
    log.info(f"========== 전체 작업 완료 및 요약 파일 패키징 완료: {os.path.basename(summary_path)} ==========")

    return folder

if __name__ == "__main__":
    run_extractor()
