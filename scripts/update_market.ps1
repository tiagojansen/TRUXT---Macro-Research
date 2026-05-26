# ============================================================
# TRUXT Macro Research — Publicação de dados de mercado
# Agendado via Windows Task Scheduler: dias úteis 09:10 / 13:05 / 18:05
# Pré-requisito: VBA no Excel BBG já salvou data\market.json
# ============================================================

$SITE_DIR    = "S:\Macro\Site"
$JSON_FILE   = "$SITE_DIR\data\market.json"
$HIST_FILE   = "$SITE_DIR\data\market_history.json"
$LOG_FILE    = "$SITE_DIR\scripts\update_log.txt"

function Log($msg) {
    $line = "$(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')  $msg"
    Write-Host $line
    Add-Content -Path $LOG_FILE -Value $line -Encoding UTF8
}

Log "=== Iniciando publicação ==="

# Verifica se market.json existe
if (-not (Test-Path $JSON_FILE)) {
    Log "AVISO: $JSON_FILE não encontrado. VBA ainda não rodou?"
    Log "=== Concluído ===`n"
    exit 0
}

# market.json: publica se modificado nos últimos 10 min
$ageRecent = (Get-Date) - (Get-Item $JSON_FILE).LastWriteTime
$publishRecent = $ageRecent.TotalMinutes -le 10

# market_history.json: publica se existe e foi modificado nos últimos 60 min
$publishHistory = $false
if (Test-Path $HIST_FILE) {
    $ageHist = (Get-Date) - (Get-Item $HIST_FILE).LastWriteTime
    $publishHistory = $ageHist.TotalMinutes -le 60
}

if (-not $publishRecent -and -not $publishHistory) {
    Log "Sem atualizações recentes. Nada a publicar."
    Log "=== Concluído ===`n"
    exit 0
}

try {
    Set-Location $SITE_DIR

    if ($publishRecent) {
        Log "market.json atualizado há $([math]::Round($ageRecent.TotalMinutes,1)) min. Adicionando..."
        git add "data/market.json"
    }
    if ($publishHistory) {
        Log "market_history.json atualizado. Adicionando..."
        git add "data/market_history.json"
    }

    $commitMsg = "data: market $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
    $result = git commit -m $commitMsg 2>&1

    if ($result -match "nothing to commit") {
        Log "Sem alterações desde o último commit."
    } else {
        git push 2>&1 | ForEach-Object { Log $_ }
        Log "Publicado no GitHub com sucesso!"
    }
} catch {
    Log "ERRO: $_"
}

Log "=== Concluído ===`n"
