<#
.SYNOPSIS
  한투API 자동화 파이프라인을 Windows 작업 스케줄러(Task Scheduler)에 등록한다.

.설명
  기존에는 한투API_스케줄러.py 가 while True 무한루프로 상주하며 스스로 시각을
  체크했다. 2026-09 개편으로 그 루프는 테스트용 폴백으로만 남기고, 실제 운영은
  Windows 작업 스케줄러가 정해진 시각에 --once/--daily/--weekly/--monthly 플래그로
  한투API_스케줄러.py 를 "1회성"으로 호출하는 방식으로 바꿨다.
  (컴퓨터가 꺼져있으면 그 실행만 건너뛰고, 다음 예약 시각에 정상적으로 다시 실행된다 —
   상주 프로세스가 없어 메모리 누수나 크래시 후 방치 위험이 없다.)

  등록하는 작업 (평일 = 월~금):
    KIS_시간별_1000   평일 10:00   python 한투API_스케줄러.py --once
    KIS_시간별_1300   평일 13:00   python 한투API_스케줄러.py --once
    KIS_시간별_1520   평일 15:20   python 한투API_스케줄러.py --once
    KIS_일별_1605     평일 16:05   python 한투API_스케줄러.py --daily
    KIS_주간_0900     토요일 09:00 python 한투API_스케줄러.py --weekly

  ※ 월간 매크로(--monthly)는 이 파이프라인에서 운영하지 않기로 해 등록하지 않는다
    (매크로 판단은 별도 도구에서 수행). 스크립트의 --monthly 플래그는 수동 실행용으로만 남아 있다.

.사용법
  1) 이 스크립트(setup_tasks.ps1)와 프로젝트 폴더(한투API_스케줄러.py 등)를
     같은 위치에 두거나, 아래 $ProjectDir 변수를 직접 프로젝트 경로로 수정한다.
  2) PowerShell을 "관리자 권한으로 실행"한 뒤 아래처럼 실행한다.
       cd "C:\Users\KIS\Desktop\Curser\project1"
       Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
       .\setup_tasks.ps1
  3) 등록 후 작업 스케줄러(taskschd.msc)에서 "KIS_" 로 시작하는 5개 작업이
     보이는지 확인한다.
  4) 제거하려면 remove_tasks.ps1 을 관리자 권한으로 실행한다.

.주의
  · python 이 PATH에 잡혀 있어야 한다 (cmd 에서 `python --version` 확인).
    안 잡혀 있으면 아래 $PythonExe 변수에 python.exe 전체 경로를 직접 넣는다.
  · claude(Claude Code CLI)도 이 컴퓨터의 이 사용자 계정에 설치·로그인되어
    있어야 한다 (`claude --version`). 작업 스케줄러는 로그인 세션이 아니어도
    실행되므로, "로그인 여부와 무관하게 실행" 옵션을 쓰려면 claude 인증이
    사용자 프로필 영역(토큰 캐시)에 저장되어 있어야 한다 — 최초 1회는 반드시
    사람이 직접 `claude` 를 실행해 로그인까지 마친 뒤 이 스크립트를 등록할 것.
#>

# ============================================================
# 0. 경로 설정 — 필요하면 이 두 줄만 직접 고치면 된다
# ============================================================
$ProjectDir = $PSScriptRoot
if (-not $ProjectDir) { $ProjectDir = (Get-Location).Path }

$PythonExe = "python"   # 그대로 두면 아래에서 전체 경로로 자동 해석. 직접 지정하려면 예) "C:\Users\KIS\AppData\Local\Programs\Python\Python311\python.exe"

# 작업 스케줄러(S4U, 로그인 무관 실행)는 '사용자 PATH'를 읽지 않아 이름만 쓰면 python 을 못 찾는다
# (LastTaskResult 0x80070002). 그래서 등록 시점에 전체 경로로 고정한다.
if ($PythonExe -eq "python") {
    $found = Get-Command python -ErrorAction SilentlyContinue
    if ($found) { $PythonExe = $found.Source }
}

$SchedulerScript = Join-Path $ProjectDir "한투API_스케줄러.py"

if (-not (Test-Path $SchedulerScript)) {
    Write-Host "❌ 한투API_스케줄러.py 를 찾을 수 없습니다: $SchedulerScript" -ForegroundColor Red
    Write-Host "   이 스크립트를 프로젝트 폴더 안에 두고 다시 실행하거나," -ForegroundColor Red
    Write-Host "   스크립트 상단의 `$ProjectDir 값을 직접 수정하세요." -ForegroundColor Red
    exit 1
}

# S4U(로그인 여부와 무관하게 실행) 작업 등록은 관리자 권한이 필요하다. 없으면 여기서 바로 중단.
$IsAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
           ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $IsAdmin) {
    Write-Host "❌ 관리자 권한으로 실행되지 않았습니다. 등록하지 않고 중단합니다." -ForegroundColor Red
    Write-Host "   시작 메뉴에서 PowerShell 우클릭 → '관리자 권한으로 실행' 후 다시 실행하세요." -ForegroundColor Red
    Write-Host "   (창 제목이 '관리자: Windows PowerShell' 이어야 합니다)" -ForegroundColor Red
    exit 1
}

Write-Host "프로젝트 경로: $ProjectDir"
Write-Host "python 실행:   $PythonExe"
Write-Host ""

# ============================================================
# 1. 공통 액션/조건 헬퍼
# ============================================================
function New-KisAction($Flag) {
    New-ScheduledTaskAction -Execute $PythonExe `
        -Argument "`"$SchedulerScript`" $Flag" `
        -WorkingDirectory $ProjectDir
}

# 배터리 노트북이어도 실행되도록, 전원 조건은 모두 끈다.
$Settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 20) `
    -MultipleInstances IgnoreNew

# 로그인 여부와 무관하게 현재 사용자 권한으로 실행 (화면에 창 안 띄움)
$Principal = New-ScheduledTaskPrincipal `
    -UserId "$env:USERDOMAIN\$env:USERNAME" `
    -LogonType S4U `
    -RunLevel Limited

# ============================================================
# 2. 작업 등록
# ============================================================
$tasks = @()

# ── 시간별 (평일 10:00 / 13:00 / 15:20, --once) ─────────────
foreach ($t in @(
        @{Name="KIS_시간별_1000"; Time="10:00"},
        @{Name="KIS_시간별_1300"; Time="13:00"},
        @{Name="KIS_시간별_1520"; Time="15:20"}
    )) {
    $trigger = New-ScheduledTaskTrigger -Weekly `
        -DaysOfWeek Monday,Tuesday,Wednesday,Thursday,Friday `
        -At $t.Time
    $tasks += [PSCustomObject]@{ Name=$t.Name; Trigger=$trigger; Action=(New-KisAction "--once") }
}

# ── 일별 (평일 16:05, --daily) ───────────────────────────────
$trigger = New-ScheduledTaskTrigger -Weekly `
    -DaysOfWeek Monday,Tuesday,Wednesday,Thursday,Friday -At "16:05"
$tasks += [PSCustomObject]@{ Name="KIS_일별_1605"; Trigger=$trigger; Action=(New-KisAction "--daily") }

# ── 주간 (토요일 09:00, --weekly) ────────────────────────────
$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Saturday -At "09:00"
$tasks += [PSCustomObject]@{ Name="KIS_주간_0900"; Trigger=$trigger; Action=(New-KisAction "--weekly") }

# ============================================================
# 3. 실제 등록 실행
# ============================================================
foreach ($t in $tasks) {
    try {
        Unregister-ScheduledTask -TaskName $t.Name -Confirm:$false -ErrorAction SilentlyContinue
        Register-ScheduledTask -TaskName $t.Name `
            -Action $t.Action -Trigger $t.Trigger `
            -Settings $Settings -Principal $Principal `
            -Description "한투API 자동화 파이프라인 ($($t.Name))" `
            -ErrorAction Stop | Out-Null   # 실패를 catch 로 보내야 거짓 "등록 완료"가 안 찍힌다
        Write-Host "✅ 등록 완료: $($t.Name)" -ForegroundColor Green
    } catch {
        Write-Host "❌ 등록 실패: $($t.Name) — $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "   (관리자 권한 PowerShell에서 실행했는지 확인하세요)" -ForegroundColor Yellow
    }
}

Write-Host ""
Write-Host "완료. 작업 스케줄러(taskschd.msc)에서 'KIS_' 로 시작하는 작업 5개를 확인하세요." -ForegroundColor Cyan
Write-Host "테스트하려면 작업을 마우스 우클릭 → '실행'을 누르거나, 아래처럼 직접 실행해도 됩니다:" -ForegroundColor Cyan
Write-Host "  python 한투API_스케줄러.py --once" -ForegroundColor Cyan
