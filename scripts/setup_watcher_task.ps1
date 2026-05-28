# ============================================================
# TRUXT Macro — Setup do Task Scheduler para o watcher
# Rodar UMA vez como Administrador no Main PC:
#   powershell -ExecutionPolicy Bypass -File "S:\Macro\Site\scripts\setup_watcher_task.ps1"
# ============================================================

$taskName   = "TRUXT Market Watcher"
$scriptPath = "S:\Macro\Site\scripts\watch_and_publish.ps1"

# Remove tarefa anterior se existir
Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue

$action = New-ScheduledTaskAction `
    -Execute "powershell.exe" `
    -Argument "-WindowStyle Hidden -NonInteractive -ExecutionPolicy Bypass -File `"$scriptPath`""

# Dispara ao fazer logon do usuário atual
$trigger = New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME

$settings = New-ScheduledTaskSettingsSet `
    -ExecutionTimeLimit (New-TimeSpan -Seconds 0) `   # sem limite de tempo
    -RestartCount 3 `
    -RestartInterval (New-TimeSpan -Minutes 1) `       # reinicia se travar
    -StartWhenAvailable

Register-ScheduledTask `
    -TaskName $taskName `
    -Action   $action `
    -Trigger  $trigger `
    -Settings $settings `
    -RunLevel Highest `
    -Force

Write-Host ""
Write-Host "Tarefa '$taskName' registrada com sucesso!" -ForegroundColor Green
Write-Host "O watcher vai iniciar automaticamente ao fazer login no Windows."
Write-Host ""
Write-Host "Para iniciar agora sem reiniciar:" -ForegroundColor Yellow
Write-Host "  Start-ScheduledTask -TaskName '$taskName'"
