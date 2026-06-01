# ============================================================
# TRUXT Macro -- Setup autostart do watcher (SEM admin)
# Usa HKCU\Run + VBScript wrapper com loop de auto-reinicio.
#
# Rodar UMA vez (usuario normal, sem admin):
#   powershell -ExecutionPolicy Bypass -File "S:\Macro\Site\scripts\setup_autostart.ps1"
# ============================================================

$scriptPath = "S:\Macro\Site\scripts\watch_and_publish.ps1"
$vbsPath    = "S:\Macro\Site\scripts\watcher_loop.vbs"

# --- Para watcher anterior se estiver rodando ---
Get-Process -Name "wscript" -ErrorAction SilentlyContinue | ForEach-Object {
    $cmdline = (Get-CimInstance Win32_Process -Filter "ProcessId=$($_.Id)" `
                -ErrorAction SilentlyContinue).CommandLine
    if ($cmdline -like "*watcher_loop*") {
        Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue
        Write-Host "  Parou instancia anterior (PID $($_.Id))" -ForegroundColor Yellow
    }
}

# --- 1. Cria VBScript com loop de auto-reinicio ---
# wscript //B roda sem janela. Se o PowerShell travar, reinicia em 30s.
$vbs = @"
' TRUXT Market Watcher -- loop de auto-reinicio (sem janela)
Set sh = CreateObject("WScript.Shell")
Do
    sh.Run "powershell -WindowStyle Hidden -NonInteractive -ExecutionPolicy Bypass -File """ & "$scriptPath" & """", 0, True
    WScript.Sleep 30000
Loop
"@
Set-Content -Path $vbsPath -Value $vbs -Encoding ASCII
Write-Host "VBScript criado: $vbsPath"

# --- 2. Registra no HKCU\Run (sem admin, persiste para todo logon) ---
$runKey = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
$runCmd = "wscript //B `"$vbsPath`""
Set-ItemProperty -Path $runKey -Name "TRUXTWatcher" -Value $runCmd
Write-Host "Registrado em HKCU\Run."

# --- 3. Inicia agora sem precisar fazer logoff ---
Write-Host ""
Write-Host "OK -- Autostart configurado!" -ForegroundColor Green
Write-Host "  - Inicia automaticamente a cada logon (sem admin)"
Write-Host "  - Se o watcher travar, reinicia sozinho em 30 segundos"
Write-Host "  - Sem janela visivel"
Write-Host ""
Write-Host "Iniciando agora..."
Start-Process "wscript.exe" -ArgumentList "//B `"$vbsPath`""
Start-Sleep -Seconds 2

# Verifica se iniciou
$proc = Get-Process -Name "wscript" -ErrorAction SilentlyContinue | Where-Object {
    (Get-CimInstance Win32_Process -Filter "ProcessId=$($_.Id)" `
     -ErrorAction SilentlyContinue).CommandLine -like "*watcher_loop*"
}
if ($proc) {
    Write-Host "Watcher rodando em background (PID $($proc.Id))." -ForegroundColor Green
} else {
    Write-Host "Aviso: nao foi possivel confirmar inicio. Verifique manualmente." -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Para parar o watcher:"
Write-Host "  Get-Process wscript | Stop-Process -Force" -ForegroundColor Cyan
Write-Host "Para remover o autostart:"
Write-Host "  Remove-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run' -Name 'TRUXTWatcher'" -ForegroundColor Cyan
