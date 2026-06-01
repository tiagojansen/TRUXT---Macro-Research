# ============================================================
# TRUXT Macro -- Setup do servico horario de mercado
# Rodar UMA vez (sem admin):
#   powershell -ExecutionPolicy Bypass -File "S:\Macro\Site\scripts\setup_hourly.ps1"
# ============================================================

$scriptPath = "S:\Macro\Site\scripts\hourly_market.ps1"
$vbsPath    = "S:\Macro\Site\scripts\hourly_market_loop.vbs"

# Para instancia anterior se existir
Get-Process -Name "wscript" -ErrorAction SilentlyContinue | ForEach-Object {
    $cmd = (Get-CimInstance Win32_Process -Filter "ProcessId=$($_.Id)" `
            -ErrorAction SilentlyContinue).CommandLine
    if ($cmd -like "*hourly_market_loop*") {
        Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue
        Write-Host "  Parou instancia anterior (PID $($_.Id))" -ForegroundColor Yellow
    }
}

# VBScript sem janela com loop de auto-reinicio
$vbs = @"
Set sh = CreateObject("WScript.Shell")
Do
    sh.Run "powershell -WindowStyle Hidden -NonInteractive -ExecutionPolicy Bypass -File """ & "$scriptPath" & """", 0, True
    WScript.Sleep 30000
Loop
"@
Set-Content -Path $vbsPath -Value $vbs -Encoding ASCII
Write-Host "VBScript criado: $vbsPath"

# Registra no HKCU\Run (persiste entre logons, sem admin)
$runKey = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
Set-ItemProperty -Path $runKey -Name "TRUXTHourly" -Value "wscript //B `"$vbsPath`""
Write-Host "Registrado em HKCU\Run."

# Inicia agora
Start-Process "wscript.exe" -ArgumentList "//B `"$vbsPath`""
Start-Sleep -Seconds 2

$proc = Get-Process -Name "wscript" -ErrorAction SilentlyContinue | Where-Object {
    (Get-CimInstance Win32_Process -Filter "ProcessId=$($_.Id)" `
     -ErrorAction SilentlyContinue).CommandLine -like "*hourly_market_loop*"
}

Write-Host ""
if ($proc) {
    Write-Host "OK -- Servico horario ativo!" -ForegroundColor Green
} else {
    Write-Host "OK -- Servico registrado (iniciara no proximo logon)." -ForegroundColor Green
}
Write-Host "  Seg-sex, 8:59 ate 18:59: Bloomberg refresh + save + ExportarMercado"
Write-Host "  Log: S:\Macro\Site\scripts\update_log.txt"
Write-Host ""
Write-Host "Para parar:"
Write-Host "  Get-Process wscript | Stop-Process -Force" -ForegroundColor Cyan
