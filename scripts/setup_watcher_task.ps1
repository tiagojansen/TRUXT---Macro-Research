# ============================================================
# TRUXT Macro — Setup do Task Scheduler para o watcher
# Rodar UMA vez como Administrador no Main PC:
#   Clique direito no PowerShell → "Executar como administrador"
#   powershell -ExecutionPolicy Bypass -File "S:\Macro\Site\scripts\setup_watcher_task.ps1"
# ============================================================

$taskName   = "TRUXT Market Watcher"
$scriptPath = "S:\Macro\Site\scripts\watch_and_publish.ps1"

# Remove tarefa anterior se existir
Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue

$action = New-ScheduledTaskAction `
    -Execute "powershell.exe" `
    -Argument "-WindowStyle Hidden -NonInteractive -ExecutionPolicy Bypass -File `"$scriptPath`""

$trigger = New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME

$settings = New-ScheduledTaskSettingsSet -ExecutionTimeLimit (New-TimeSpan -Seconds 0)

Register-ScheduledTask `
    -TaskName $taskName `
    -Action   $action `
    -Trigger  $trigger `
    -Settings $settings `
    -RunLevel Highest `
    -Force

if ($?) {
    Write-Host ""
    Write-Host "OK — Tarefa '$taskName' registrada!" -ForegroundColor Green
    Write-Host "Iniciando agora..."
    Start-ScheduledTask -TaskName $taskName
    Write-Host "Watcher rodando em background. Site vai atualizar sozinho a partir de agora." -ForegroundColor Green
} else {
    Write-Host "ERRO ao registrar tarefa." -ForegroundColor Red
}
