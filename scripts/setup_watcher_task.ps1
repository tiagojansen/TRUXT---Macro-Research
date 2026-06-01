# ============================================================
# TRUXT Macro -- Setup do Task Scheduler para o watcher
# Rodar UMA vez como Administrador no Main PC:
#   powershell -ExecutionPolicy Bypass -File "S:\Macro\Site\scripts\setup_watcher_task.ps1"
# ============================================================

$taskName   = "TRUXT Market Watcher"
$scriptPath = "S:\Macro\Site\scripts\watch_and_publish.ps1"

# Remove tarefa anterior se existir
Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue

$action = New-ScheduledTaskAction `
    -Execute  "powershell.exe" `
    -Argument "-WindowStyle Hidden -NonInteractive -ExecutionPolicy Bypass -File `"$scriptPath`""

# Dispara no logon E a cada 30 minutos (recupera apos hibernate/reinicio)
$triggerLogon  = New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME
$triggerRepeat = New-ScheduledTaskTrigger -RepetitionInterval (New-TimeSpan -Minutes 30) `
                                          -Once -At (Get-Date)

$settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Seconds 0) `
    -RestartCount       5 `
    -RestartInterval    (New-TimeSpan -Minutes 1) `
    -MultipleInstances  IgnoreNew

Register-ScheduledTask `
    -TaskName $taskName `
    -Action   $action `
    -Trigger  @($triggerLogon, $triggerRepeat) `
    -Settings $settings `
    -RunLevel Highest `
    -Force

if ($?) {
    Write-Host ""
    Write-Host "OK -- Tarefa registrada!" -ForegroundColor Green
    Write-Host "  - Inicia no logon"
    Write-Host "  - Reinicia automaticamente em caso de falha (ate 5x, intervalo 1min)"
    Write-Host "  - Trigger a cada 30min (recupera apos hibernate)"
    Write-Host "Iniciando agora..."
    Start-ScheduledTask -TaskName $taskName
    Write-Host "Watcher rodando em background." -ForegroundColor Green
} else {
    Write-Host "ERRO ao registrar tarefa." -ForegroundColor Red
}
