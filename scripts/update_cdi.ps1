# ============================================================
# TRUXT Macro Research — Atualização diária do CDI
# Agendado via Windows Task Scheduler: dias úteis às 19:05
# ============================================================

$SITE_DIR = "S:\Macro\Site"
$CDI_FILE = "$SITE_DIR\data\cdi.json"
$LOG_FILE = "$SITE_DIR\scripts\update_log.txt"

function Log($msg) {
    $line = "$(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')  $msg"
    Write-Host $line
    Add-Content -Path $LOG_FILE -Value $line -Encoding UTF8
}

Log "=== Iniciando atualização CDI ==="

try {
    # Calcular data de início (10 anos atrás)
    $dataInicio = (Get-Date).AddYears(-10).ToString("dd/MM/yyyy")
    $url = "https://api.bcb.gov.br/dados/serie/bcdata.sgs.12/dados?formato=json&dataInicial=$dataInicio"

    Log "Buscando CDI do BCB: série 12 desde $dataInicio"
    $response = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 30
    $count = ($response.Content | ConvertFrom-Json).Count

    New-Item -ItemType Directory -Force -Path "$SITE_DIR\data" | Out-Null
    $response.Content | Out-File -FilePath $CDI_FILE -Encoding UTF8 -NoNewline
    Log "OK: $count registros salvos em $CDI_FILE"

    # Git commit e push
    Set-Location $SITE_DIR
    git add "data/cdi.json"

    $commitMsg = "data: CDI $(Get-Date -Format 'yyyy-MM-dd')"
    $result = git commit -m $commitMsg 2>&1

    if ($result -match "nothing to commit") {
        Log "Sem alterações no CDI desde ontem."
    } else {
        git push 2>&1 | ForEach-Object { Log $_ }
        Log "Publicado no GitHub com sucesso!"
    }

} catch {
    Log "ERRO: $_"
}

Log "=== Concluído ===`n"
