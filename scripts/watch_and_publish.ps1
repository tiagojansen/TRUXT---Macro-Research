# ============================================================
# TRUXT Macro Research — Publicação automática (watcher)
#
# Roda CONTINUAMENTE no Main PC (loop infinito).
# Detecta quando market.json ou market_history.json mudam
# e faz git push automaticamente em segundos.
#
# Setup (rodar UMA vez como admin no Main PC):
#   powershell -File "S:\Macro\Site\scripts\setup_watcher_task.ps1"
# ============================================================

$SITE_DIR  = "S:\Macro\Site"
$JSON_FILE = "$SITE_DIR\data\market.json"
$HIST_FILE = "$SITE_DIR\data\market_history.json"
$LOG_FILE  = "$SITE_DIR\scripts\update_log.txt"

$POLL_SECS = 10    # verifica a cada 10 segundos
$GRACE_SEC = 5     # espera após detectar mudança (garante que ambos os arquivos fecharam)
$COOLDOWN  = 90    # segundos mínimos entre publicações consecutivas

function Log($msg) {
    $line = "$(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')  $msg"
    Write-Host $line
    Add-Content -Path $LOG_FILE -Value $line -Encoding UTF8
}

function Publish {
    Log "=== Auto-publicando ==="
    try {
        Set-Location $SITE_DIR

        $staged = $false
        if (Test-Path $JSON_FILE) { git add "data/market.json";         $staged = $true }
        if (Test-Path $HIST_FILE) { git add "data/market_history.json"; $staged = $true }

        if ($staged) {
            $commitMsg = "data: market $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
            $result = git commit -m $commitMsg 2>&1
            if ($result -match "nothing to commit") {
                Log "Sem alterações novas."
            } else {
                git push 2>&1 | ForEach-Object { Log $_ }
                Log "Publicado com sucesso!"
            }
        }
    } catch {
        Log "ERRO: $_"
    }
    Log "=== Concluído ===`n"
}

Log "=== Watcher iniciado (poll: ${POLL_SECS}s) ==="
$lastPublished = [datetime]::MinValue

while ($true) {
    Start-Sleep -Seconds $POLL_SECS

    $trigger = $false
    $window  = $POLL_SECS * 2   # janela de detecção = 2x o intervalo de polling

    if (Test-Path $JSON_FILE) {
        $age = ((Get-Date) - (Get-Item $JSON_FILE).LastWriteTime).TotalSeconds
        if ($age -le $window) { $trigger = $true }
    }

    if (-not $trigger -and (Test-Path $HIST_FILE)) {
        $age = ((Get-Date) - (Get-Item $HIST_FILE).LastWriteTime).TotalSeconds
        if ($age -le $window) { $trigger = $true }
    }

    $cooldownOk = ((Get-Date) - $lastPublished).TotalSeconds -ge $COOLDOWN

    if ($trigger -and $cooldownOk) {
        Start-Sleep -Seconds $GRACE_SEC   # aguarda VBA fechar todos os arquivos
        Publish
        $lastPublished = Get-Date
    }
}
