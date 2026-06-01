# ============================================================
# TRUXT Macro -- Setup do Task Scheduler para o watcher
# Rodar UMA vez como usuario normal (NAO precisa de admin):
#   powershell -ExecutionPolicy Bypass -File "S:\Macro\Site\scripts\setup_watcher_task.ps1"
# ============================================================

$taskName   = "TRUXT Market Watcher"
$scriptPath = "S:\Macro\Site\scripts\watch_and_publish.ps1"

# Remove tarefa anterior se existir
Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue

$action = New-ScheduledTaskAction `
    -Execute  "powershell.exe" `
    -Argument "-WindowStyle Hidden -NonInteractive -ExecutionPolicy Bypass -File `"$scriptPath`""

# Trigger 1: inicia ao fazer logon (qualquer horario)
$triggerLogon = New-ScheduledTaskTrigger -AtLogOn

# Triggers 2-13: seg-sex, 8h ate 19h, uma por hora
# (garante reinicio mesmo que o loop trave no meio do dia)
$triggerHourly = 8..19 | ForEach-Object {
    New-ScheduledTaskTrigger `
        -Weekly `
        -DaysOfWeek Monday, Tuesday, Wednesday, Thursday, Friday `
        -At ('{0:D2}:00' -f $_)
}

$settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Seconds 0) `
    -RestartCount       5 `
    -RestartInterval    (New-TimeSpan -Minutes 1) `
    -MultipleInstances  IgnoreNew

# Sem -RunLevel Highest: nao precisa de admin.
# Por padrao roda apenas quando o usuario estiver logado.
Register-ScheduledTask `
    -TaskName $taskName `
    -Action   $action `
    -Trigger  (@($triggerLogon) + $triggerHourly) `
    -Settings $settings `
    -Force

if ($?) {
    Write-Host ""
    Write-Host "OK -- Tarefa '$taskName' registrada!" -ForegroundColor Green
    Write-Host "  - Inicia no logon (qualquer horario)"
    Write-Host "  - Reinicia a cada 1h (seg-sex, 8h-19h) para recuperar de travamentos"
    Write-Host "  - Reinicia em falha ate 5x (intervalo 1min)"
    Write-Host "  - MultipleInstances: IgnoreNew (nao duplica)"
    Write-Host ""
    Write-Host "Iniciando agora..."
    Start-ScheduledTask -TaskName $taskName
    Write-Host "Watcher rodando em background." -ForegroundColor Green
} else {
    Write-Host "ERRO ao registrar tarefa." -ForegroundColor Red
}
