"""
한투API 자동 스케줄러 (시간별 / 일별 / 주간 / 월간매크로 통합 마스터) — 경량화판 v2
=================================================================
[핵심 설계]
  모든 수집기·분석기를 subprocess(독립 프로세스)로 실행합니다.
  → Windows 환경에서 패키지 충돌 완전 차단
  → 수집기 완료 즉시 분석기 호출 (대기 없음)

[2026-09 최적화 변경 — 실행 방식]
  ★ 이 파일의 무한루프 run_scheduler()는 더 이상 상시 실행을 권장하지 않습니다.
    대신 Windows 작업 스케줄러(Task Scheduler)가 아래 스케줄대로 이 스크립트를
    "--once / --daily / --weekly / --monthly" 인자로 그때그때 호출합니다.
    (등록 스크립트: setup_tasks.ps1 참고)
    이렇게 하면 컴퓨터 재부팅·수면 이후에도 스케줄이 끊기지 않고,
    파이썬 프로세스가 24시간 상주할 필요가 없습니다.
    run_scheduler() 무한루프는 개발/테스트용으로만 남겨두었습니다.

[스케줄 요약 — 기존 대비 변경]
  ─ 평일(월~금) 시간별: 10:00 / 13:00 / 15:20                (기존 7회 → 3회로 축소)
  ─ 평일(월~금) 일별  : 16:05                                  (변경 없음)
  ─ 토요일       주간  : 09:00  (종목별 수급·이격도·신용잔고)   (변경 없음)
  ─ 매월 첫째주 일요일 월간매크로: 09:00 (매크로+섹터, 이전엔 매주 일요일이었음)

  시간별 모듈도 6개 → 4개로 축소(회원사동향/매물대는 일별로 이관)했습니다.
  자세한 내용은 각 수집기 파일 상단 docstring 참고.

[사이클별 흐름]
  시간별 ① 한투API_시간별데이터.py       → ② 한투API_텔레그램분석.py --date YYYYMMDD --hhmm HHMM
  일별   ① 한투API_일별데이터.py         → ② 한투API_텔레그램분석.py --daily --date YYYYMMDD
  주간   ① 한투API_주간데이터.py         → ② 한투API_텔레그램분석.py --weekly --date YYYYMMDD
  월간   ① 월간_매크로섹터분석.py        → ② 한투API_텔레그램분석.py --monthly --date YYYYMMDD
         (구 주간_섹터_동향분석.py를 리네임 + 월 1회 주기로 전환)

  ※ 분석 단계(②)는 더 이상 Gemini API를 쓰지 않습니다. 이 데스크탑에 설치된
    Claude Code CLI(claude 명령)를 비대화형으로 호출해 로컬 파일을 직접 읽고
    분석하게 한 뒤, 결과를 텔레그램으로 발송합니다.

실행 방법:
  python 한투API_스케줄러.py            # (테스트용) 자동 스케줄 무한루프
  python 한투API_스케줄러.py --once     # 시간별 1회 즉시 실행
  python 한투API_스케줄러.py --daily    # 일별  1회 즉시 실행
  python 한투API_스케줄러.py --weekly   # 주간  1회 즉시 실행
  python 한투API_스케줄러.py --monthly  # 월간매크로 1회 즉시 실행

  실제 운영은 setup_tasks.ps1 로 등록한 Windows 작업 스케줄러가
  위 4개 플래그를 정해진 시각에 자동으로 호출합니다.
"""

import sys
import time
import logging
import subprocess
import os
import re
from datetime import datetime, date

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
# 1. 스케줄 설정
#    ★ 실제 운영 트리거는 Windows 작업 스케줄러(setup_tasks.ps1)가 담당한다.
#      아래 상수는 (a) run_scheduler() 테스트 루프, (b) 로그/문서 참고용이다.
# ──────────────────────────────────────────────
HOURLY_SCHEDULE  = ["10:00", "13:00", "15:20"]   # 기존 7회 → 3회
DAILY_SCHEDULE   = ["16:05"]
WEEKLY_SCHEDULE  = ["09:00"]   # 토요일(weekday=5)에만 실행 — 종목별 주간데이터
MONTHLY_SCHEDULE = ["09:00"]   # 매월 "첫째주" 일요일(weekday=6)에만 실행 — 매크로/섹터

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
MONTHLY_MACRO_SCRIPT = "월간_매크로섹터분석.py"   # 구 주간_섹터_동향분석.py


# ──────────────────────────────────────────────
# 2. subprocess 실행 헬퍼
# ──────────────────────────────────────────────
def _decode_line(b: bytes) -> str:
    """bytes 한 줄을 utf-8 → cp949 → euc-kr 순으로 폴백 디코딩."""
    if not b:
        return ""
    for enc in ("utf-8", "cp949", "euc-kr", "utf-8-sig"):
        try:
            return b.decode(enc).rstrip("\r\n")
        except (UnicodeDecodeError, LookupError):
            continue
    return b.decode("utf-8", errors="replace").rstrip("\r\n")


def run_script(script_name: str, args: list, timeout: int = 900) -> tuple:
    """
    script_name 을 subprocess로 실행.
    stdout/stderr 를 한 줄씩 실시간으로 스케줄러 터미널에 출력.
    반환: (성공여부, stdout 전체 텍스트)
    """
    path = os.path.join(BASE_DIR, script_name)
    cmd  = [sys.executable, "-u", path] + args   # -u : 출력 버퍼링 비활성화

    log.info(f"  $ python {script_name} {' '.join(args)}")

    collected_lines = []

    try:
        proc = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,   # stderr 를 stdout 에 합쳐서 순서 유지
            bufsize=0,                  # 버퍼 없음
        )

        import threading, queue as _queue

        line_q = _queue.Queue()

        def _reader():
            """별도 스레드에서 바이트 줄 읽기 → 큐에 적재."""
            try:
                for raw in proc.stdout:
                    line_q.put(raw)
            finally:
                line_q.put(None)   # 종료 신호

        reader_thread = threading.Thread(target=_reader, daemon=True)
        reader_thread.start()

        deadline = time.time() + timeout

        while True:
            if time.time() > deadline:
                proc.kill()
                log.error(f"  → [{script_name}] 타임아웃 ({timeout}초 초과) — 프로세스 강제 종료")
                return False, "\n".join(collected_lines)

            try:
                raw = line_q.get(timeout=1)
            except _queue.Empty:
                if proc.poll() is not None:
                    break
                continue

            if raw is None:
                break

            line = _decode_line(raw)
            if line:
                log.info(f"    | {line}")
                collected_lines.append(line)

        reader_thread.join(timeout=5)
        proc.wait(timeout=10)

        stdout_str = "\n".join(collected_lines)

        if proc.returncode != 0:
            log.error(f"  → [{script_name}] 비정상 종료 (returncode={proc.returncode})")
            return False, stdout_str

        log.info(f"  → [{script_name}] 정상 완료")
        return True, stdout_str

    except Exception as e:
        log.error(f"  → [{script_name}] 실행 오류: {e}")
        return False, ""


def extract_folder_from_output(stdout: str) -> tuple:
    """
    수집기 stdout에서 저장 폴더 경로 파싱.
    예: 'MarketData\\20260604\\0939'
    반환: (date_str, hhmm_str)  예: ('20260604', '0939')
    stdout이 None이거나 빈 문자열이면 (None, None) 반환.
    """
    if not stdout:
        return None, None
    m = re.search(r"MarketData[/\\](\d{8})[/\\](\d{4})", stdout)
    if m:
        return m.group(1), m.group(2)
    return None, None


# ──────────────────────────────────────────────
# 3. 시간별 사이클 (하루 3회: 10:00 / 13:00 / 15:20)
# ──────────────────────────────────────────────
def run_hourly_cycle():
    start = datetime.now()
    log.info("━" * 55)
    log.info(f"▶ [시간별 사이클 시작] {start.strftime('%Y-%m-%d %H:%M:%S')}")
    log.info("━" * 55)

    # Step 1. 수집기 (4모듈: 분봉/체결강도/외인기관/프로그램)
    log.info("[Step 1] 시간별 데이터 수집...")
    ok, stdout = run_script("한투API_시간별데이터.py", [], timeout=600)

    date_str, hhmm_str = extract_folder_from_output(stdout)
    if date_str and hhmm_str:
        log.info(f"  → 저장 폴더: MarketData/{date_str}/{hhmm_str}")
    else:
        date_str = start.strftime("%Y%m%d")
        hhmm_str = start.strftime("%H%M")
        log.warning(f"  → 폴더 파싱 실패, 추정값 사용: {date_str}/{hhmm_str}")

    # Step 2. 분석기 (Claude Code CLI 기반, 수집 완료 즉시 실행)
    log.info("[Step 2] Claude Code 시간별 분석기 실행 (즉시)...")
    run_script(
        "한투API_텔레그램분석.py",
        ["--date", date_str, "--hhmm", hhmm_str],
        timeout=600,
    )

    elapsed = (datetime.now() - start).total_seconds()
    log.info(f"▶ [시간별 사이클 완료] 총 소요: {elapsed/60:.1f}분\n")


# ──────────────────────────────────────────────
# 4. 일별 사이클
# ──────────────────────────────────────────────
def run_daily_cycle():
    start    = datetime.now()
    date_str = start.strftime("%Y%m%d")
    log.info("━" * 55)
    log.info(f"▶ [일별 사이클 시작] {start.strftime('%Y-%m-%d %H:%M:%S')}")
    log.info("━" * 55)

    # Step 1. 일별 수집기 (일봉/수급/뉴스/대차/신용잔고/심화지표/시장비중/회원사동향/매물대)
    log.info("[Step 1] 일별 데이터 수집...")
    ok, _ = run_script("한투API_일별데이터.py", [], timeout=3600)
    if not ok:
        log.error("[Step 1 실패] 일별 수집 오류 — 분석 중단")
        return

    # Step 2. 분석기 (Claude Code CLI 기반, 수집 완료 즉시 실행)
    log.info("[Step 2] Claude Code 일별 분석기 실행 (즉시)...")
    run_script(
        "한투API_텔레그램분석.py",
        ["--daily", "--date", date_str],
        timeout=900,
    )

    elapsed = (datetime.now() - start).total_seconds()
    log.info(f"▶ [일별 사이클 완료] 총 소요: {elapsed/60:.1f}분\n")


# ──────────────────────────────────────────────
# 5. 주간 사이클 (매주 토요일 09:00) — 종목별 수급/이격도/신용잔고
# ──────────────────────────────────────────────
def run_weekly_cycle():
    start    = datetime.now()
    date_str = start.strftime("%Y%m%d")
    log.info("━" * 55)
    log.info(f"▶ [주간 사이클 시작] {start.strftime('%Y-%m-%d %H:%M:%S')}")
    log.info("━" * 55)

    # Step 1. 주간 수집기
    log.info("[Step 1] 주간 데이터 수집...")
    ok, _ = run_script("한투API_주간데이터.py", [], timeout=3600)
    if not ok:
        log.error("[Step 1 실패] 주간 수집 오류 — 분석 중단")
        return

    # Step 2. 주간 분석기 (Claude Code CLI 기반, 수집 완료 즉시 실행)
    log.info("[Step 2] Claude Code 주간 분석기 실행 (즉시)...")
    run_script(
        "한투API_텔레그램분석.py",
        ["--weekly", "--date", date_str],
        timeout=900,
    )

    elapsed = (datetime.now() - start).total_seconds()
    log.info(f"▶ [주간 사이클 완료] 총 소요: {elapsed/60:.1f}분\n")


# ──────────────────────────────────────────────
# 6. 월간매크로 사이클 (매월 첫째주 일요일 09:00)
#    수집기: 월간_매크로섹터분석.py (구 주간_섹터_동향분석.py, Macro_Weekly + SectorAction 생성)
#    분석기: 한투API_텔레그램분석.py --monthly
#    ★ 변경: 기존에는 매주 일요일 실행되었으나, 금리/신용스프레드/반도체 수출입 등
#      매크로 지표는 월 단위로 갱신되는 성격이 강해 월 1회로 주기를 늦췄다.
#      (실제 "월 1회"는 Windows 작업 스케줄러의 monthly 트리거가 보장한다.
#       아래 run_scheduler() 무한루프를 쓸 경우에만 "첫째주 일요일" 자체 판별이 필요하다.)
# ──────────────────────────────────────────────
def run_monthly_macro_cycle():
    start    = datetime.now()
    date_str = start.strftime("%Y%m%d")
    log.info("━" * 55)
    log.info(f"▶ [월간매크로 사이클 시작] {start.strftime('%Y-%m-%d %H:%M:%S')}")
    log.info("━" * 55)

    # Step 1. 매크로/섹터 데이터 수집
    log.info(f"[Step 1] 매크로/섹터 데이터 수집 ({MONTHLY_MACRO_SCRIPT})...")
    ok, _ = run_script(MONTHLY_MACRO_SCRIPT, [], timeout=3600)
    if not ok:
        log.error(f"[Step 1 실패] {MONTHLY_MACRO_SCRIPT} 오류 — 분석 중단")
        return

    # Step 2. 매크로/섹터 퀀트 전략 분석기 (Claude Code CLI 기반, 수집 완료 즉시 실행)
    log.info("[Step 2] Claude Code 월간매크로 분석기 실행 (즉시)...")
    run_script(
        "한투API_텔레그램분석.py",
        ["--monthly", "--date", date_str],
        timeout=900,
    )

    elapsed = (datetime.now() - start).total_seconds()
    log.info(f"▶ [월간매크로 사이클 완료] 총 소요: {elapsed/60:.1f}분\n")


# ──────────────────────────────────────────────
# 7. 메인 스케줄 루프 (★ 개발/테스트용 — 운영은 Windows 작업 스케줄러 권장)
# ──────────────────────────────────────────────
def run_scheduler():
    log.warning("⚠️  run_scheduler() 무한루프는 테스트용입니다. 실제 운영에는")
    log.warning("    setup_tasks.ps1 로 등록하는 Windows 작업 스케줄러를 권장합니다.")
    log.info("🚀 스케줄러 가동 (테스트 모드)")
    log.info(f"  시간별 (평일)         : {' / '.join(HOURLY_SCHEDULE)}")
    log.info(f"  일별   (평일)         : {' / '.join(DAILY_SCHEDULE)}")
    log.info(f"  주간   (토요일)       : {' / '.join(WEEKLY_SCHEDULE)}")
    log.info(f"  월간매크로(첫째주 일) : {' / '.join(MONTHLY_SCHEDULE)}")
    log.info(f"  방식                  : 수집 완료 즉시 Claude Code 분석 (subprocess 독립 실행)")
    log.info(f"  월간매크로 수집기     : {MONTHLY_MACRO_SCRIPT}")

    executed_hourly  = set()
    executed_daily   = set()
    executed_weekly  = set()
    executed_monthly = set()
    last_date = date.today()

    while True:
        now   = datetime.now()
        today = now.date()
        hhmm  = now.strftime("%H:%M")
        weekday = today.weekday()   # 0=월 ... 4=금 / 5=토 / 6=일
        is_first_week = today.day <= 7   # "첫째주" 판별 (테스트 모드 전용)

        if today != last_date:
            executed_hourly.clear()
            executed_daily.clear()
            executed_weekly.clear()
            executed_monthly.clear()
            last_date = today
            log.info(f"📅 [날짜 변경] {today} — 이력 초기화")

        # ── 평일(월~금) 전용 스케줄 ──────────────
        if weekday <= 4:   # 0~4: 월~금
            if hhmm in HOURLY_SCHEDULE and hhmm not in executed_hourly:
                executed_hourly.add(hhmm)
                run_hourly_cycle()

            if hhmm in DAILY_SCHEDULE and hhmm not in executed_daily:
                executed_daily.add(hhmm)
                run_daily_cycle()

        # ── 토요일 주간 스케줄 ────────────────────
        elif weekday == 5:
            if hhmm in WEEKLY_SCHEDULE and hhmm not in executed_weekly:
                executed_weekly.add(hhmm)
                run_weekly_cycle()

        # ── 일요일 월간매크로 스케줄 (첫째주만) ───
        elif weekday == 6 and is_first_week:
            if hhmm in MONTHLY_SCHEDULE and hhmm not in executed_monthly:
                executed_monthly.add(hhmm)
                run_monthly_macro_cycle()

        time.sleep(10)


# ──────────────────────────────────────────────
# 8. CLI 진입점
# ──────────────────────────────────────────────
if __name__ == "__main__":
    if "--once" in sys.argv:
        run_hourly_cycle()
    elif "--daily" in sys.argv:
        run_daily_cycle()
    elif "--weekly" in sys.argv:
        run_weekly_cycle()
    elif "--monthly" in sys.argv:
        run_monthly_macro_cycle()
    else:
        run_scheduler()
