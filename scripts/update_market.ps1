# ============================================================
# TRUXT Macro Research — Publicação de dados de mercado
# Agendado via Windows Task Scheduler: dias úteis 09:10 / 13:05 / 18:05
# Pré-requisito: VBA no Excel BBG já salvou data\market.json
# ============================================================

$SITE_DIR  = "S:\Macro\Site"
$JSON_FILE = "$SITE_DIR\data\market.json"
$LOG_FILE  = "$SITE_DIR\scripts\update_log.txt"

function Log($msg) {
    $line = "$(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')  $msg"
    Write-Host $line
    Add-Content -Path $LOG_FILE -Value $line -Encoding UTF8
}

Log "=== Iniciando publicação market.json ==="

# Verifica se o arquivo existe
if (-not (Test-Path $JSON_FILE)) {
    Log "AVISO: $JSON_FILE não encontrado. VBA ainda não rodou?"
    Log "=== Concluído ===`n"
    exit 0
}

# Só publica se o arquivo foi modificado nos últimos 10 minutos
$age = (Get-Date) - (Get-Item $JSON_FILE).LastWriteTime
if ($age.TotalMinutes -gt 10) {
    Log "Sem atualização recente ($([int]$age.TotalMinutes) min atrás). Nada a publicar."
    Log "=== Concluído ===`n"
    exit 0
}

Log "market.json atualizado há $([math]::Round($age.TotalMinutes,1)) min. Publicando..."

try {
    Set-Location $SITE_DIR
    git add "data/market.json"

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
