# ============================================================
# TRUXT Macro -- Setup do Task Scheduler para exportacao horaria
# Rodar UMA vez como usuario normal (NAO precisa de admin):
#   powershell -ExecutionPolicy Bypass -File "S:\Macro\Site\scripts\setup_market_task.ps1"
# ============================================================

$taskName   = "TRUXT Market Data"
$scriptPath = "S:\Macro\Site\scripts\update_market_hourly.ps1"

# Remove tarefa anterior se existir
Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue

$action = New-ScheduledTaskAction `
    -Execute  "powershell.exe" `
    -Argument "-WindowStyle Hidden -NonInteractive -ExecutionPolicy Bypass -File `"$scriptPath`""

# Seg-sex, 8h ate 19h, uma trigger por hora
$triggers = 8..19 | ForEach-Object {
    New-ScheduledTaskTrigger `
        -Weekly `
        -DaysOfWeek Monday, Tuesday, Wednesday, Thursday, Friday `
        -At ('{0:D2}:00' -f $_)
}

$settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 10) `
    -RestartCount       3 `
    -RestartInterval    (New-TimeSpan -Minutes 2) `
    -MultipleInstances  IgnoreNew

# Sem -RunLevel Highest: nao precisa de admin.
Register-ScheduledTask `
    -TaskName $taskName `
    -Action   $action `
    -Trigger  $triggers `
    -Settings $settings `
    -Force

if ($?) {
    Write-Host ""
    Write-Host "OK -- Tarefa '$taskName' registrada!" -ForegroundColor Green
    Write-Host "  - Seg-sex, 8h as 19h, a cada 1h"
    Write-Host "  - Abre Excel (ou usa instancia existente), roda ExportarMercado, fecha"
    Write-Host "  - Watcher detecta o market.json atualizado e publica no GitHub"
    Write-Host ""
    Write-Host "Testar agora? Execute:"
    Write-Host "  powershell -ExecutionPolicy Bypass -File `"$scriptPath`"" -ForegroundColor Cyan
} else {
    Write-Host "ERRO ao registrar tarefa." -ForegroundColor Red
}
