# ============================================================
# TRUXT Macro Research — Publicacao automatica (watcher)
#
# Roda CONTINUAMENTE no Main PC (loop infinito).
# Detecta mudancas em:
#   - data/market.json e data/market_history.json  → publica market data
#   - reports/**/*.html e presentations/**          → atualiza config.json e publica
#
# Setup (rodar UMA vez como admin no Main PC):
#   powershell -File "S:\Macro\Site\scripts\setup_watcher_task.ps1"
# ============================================================

$SITE_DIR     = "S:\Macro\Site"
$JSON_FILE    = "$SITE_DIR\data\market.json"
$HIST_FILE    = "$SITE_DIR\data\market_history.json"
$REPORTS_DIR  = "$SITE_DIR\reports"
$PRES_DIR     = "$SITE_DIR\presentations"
$CONFIG_FILE  = "$SITE_DIR\config.json"
$UPDATE_DATES = "$SITE_DIR\scripts\update_config_dates.ps1"
$LOG_FILE     = "$SITE_DIR\scripts\update_log.txt"

$POLL_SECS = 10    # verifica a cada 10 segundos
$GRACE_SEC = 5     # espera apos detectar mudanca
$COOLDOWN  = 90    # segundos minimos entre publicacoes consecutivas

function Log($msg) {
    $line = "$(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')  $msg"
    Write-Host $line
    Add-Content -Path $LOG_FILE -Value $line -Encoding UTF8
}

function PublishMarket {
    Log "=== Publicando market data ==="
    try {
        Set-Location $SITE_DIR
        $staged = $false
        if (Test-Path $JSON_FILE) { git add "data/market.json";         $staged = $true }
        if (Test-Path $HIST_FILE) { git add "data/market_history.json"; $staged = $true }
        if ($staged) {
            $msg    = "data: market $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
            $result = git commit -m $msg 2>&1
            if ($result -match "nothing to commit") {
                Log "Sem alteracoes novas."
            } else {
                git push 2>&1 | ForEach-Object { Log $_ }
                Log "Market data publicado!"
            }
        }
    } catch { Log "ERRO: $_" }
    Log "=== Concluido ===`n"
}

function PublishReports {
    Log "=== Publicando reports ==="
    try {
        Set-Location $SITE_DIR

        # Atualiza datas no config.json conforme data real dos arquivos
        if (Test-Path $UPDATE_DATES) {
            Log "  Atualizando datas do config.json..."
            powershell -ExecutionPolicy Bypass -File $UPDATE_DATES 2>&1 | ForEach-Object { Log "  $_" }
        }

        git add "config.json"
        git add "reports/"
        if (Test-Path $PRES_DIR) { git add "presentations/" }

        $msg    = "reports: update $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
        $result = git commit -m $msg 2>&1
        if ($result -match "nothing to commit") {
            Log "Sem alteracoes novas nos reports."
        } else {
            git push 2>&1 | ForEach-Object { Log $_ }
            Log "Reports publicados!"
        }
    } catch { Log "ERRO: $_" }
    Log "=== Concluido ===`n"
}

Log "=== Watcher iniciado (poll: ${POLL_SECS}s) ==="
$lastPublishedMarket  = [datetime]::MinValue
$lastPublishedReports = [datetime]::MinValue

while ($true) {
    Start-Sleep -Seconds $POLL_SECS
    $window = $POLL_SECS * 2   # janela de deteccao = 2x o intervalo de polling

    # ── Detecta mudancas em market data ──────────────────────────
    $triggerMarket = $false
    if (Test-Path $JSON_FILE) {
        $age = ((Get-Date) - (Get-Item $JSON_FILE).LastWriteTime).TotalSeconds
        if ($age -le $window) { $triggerMarket = $true }
    }
    if (-not $triggerMarket -and (Test-Path $HIST_FILE)) {
        $age = ((Get-Date) - (Get-Item $HIST_FILE).LastWriteTime).TotalSeconds
        if ($age -le $window) { $triggerMarket = $true }
    }

    # ── Detecta mudancas em reports ──────────────────────────────
    $triggerReports = $false
    if (Test-Path $REPORTS_DIR) {
        $recent = Get-ChildItem -Recurse -Path $REPORTS_DIR | Where-Object {
            -not $_.PSIsContainer -and
            ((Get-Date) - $_.LastWriteTime).TotalSeconds -le $window
        }
        if ($recent) { $triggerReports = $true }
    }
    # Tambem dispara se config.json ou presentations mudarem
    if (-not $triggerReports -and (Test-Path $CONFIG_FILE)) {
        $age = ((Get-Date) - (Get-Item $CONFIG_FILE).LastWriteTime).TotalSeconds
        if ($age -le $window) { $triggerReports = $true }
    }
    if (-not $triggerReports -and (Test-Path $PRES_DIR)) {
        $recent = Get-ChildItem -Path $PRES_DIR | Where-Object {
            ((Get-Date) - $_.LastWriteTime).TotalSeconds -le $window
        }
        if ($recent) { $triggerReports = $true }
    }

    # ── Publica market data ──────────────────────────────────────
    $cooldownMarket = ((Get-Date) - $lastPublishedMarket).TotalSeconds -ge $COOLDOWN
    if ($triggerMarket -and $cooldownMarket) {
        Start-Sleep -Seconds $GRACE_SEC
        PublishMarket
        $lastPublishedMarket = Get-Date
    }

    # ── Publica reports (com cooldown independente) ───────────────
    $cooldownReports = ((Get-Date) - $lastPublishedReports).TotalSeconds -ge $COOLDOWN
    if ($triggerReports -and $cooldownReports) {
        Start-Sleep -Seconds $GRACE_SEC
        PublishReports
        $lastPublishedReports = Get-Date
    }
}
