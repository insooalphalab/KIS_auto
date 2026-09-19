<#
.SYNOPSIS
  setup_tasks.ps1 로 등록한 'KIS_' 로 시작하는 작업 스케줄러 작업들을 모두 제거한다.

.사용법
  PowerShell을 "관리자 권한으로 실행"한 뒤:
    Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
    .\remove_tasks.ps1
#>

$taskNames = @(
    "KIS_시간별_1000",
    "KIS_시간별_1300",
    "KIS_시간별_1520",
    "KIS_일별_1605",
    "KIS_주간_0900",
    "KIS_월간_0900"
)

foreach ($name in $taskNames) {
    $existing = Get-ScheduledTask -TaskName $name -ErrorAction SilentlyContinue
    if ($existing) {
        Unregister-ScheduledTask -TaskName $name -Confirm:$false
        Write-Host "🗑️  제거 완료: $name" -ForegroundColor Yellow
    } else {
        Write-Host "ℹ️  등록되어 있지 않음(건너뜀): $name" -ForegroundColor DarkGray
    }
}

Write-Host ""
Write-Host "완료." -ForegroundColor Cyan
