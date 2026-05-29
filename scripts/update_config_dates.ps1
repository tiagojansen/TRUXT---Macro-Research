# ============================================================
# TRUXT Macro — Atualiza datas do config.json
#
# Le a data real de modificacao de cada arquivo .html listado
# em config.json e atualiza o campo "updated" correspondente.
#
# Uso:
#   powershell -File "S:\Macro\Site\scripts\update_config_dates.ps1"
# ============================================================

param([string]$SiteDir = "S:\Macro\Site")

$configPath = Join-Path $SiteDir "config.json"

$raw    = Get-Content $configPath -Raw -Encoding UTF8
$config = $raw | ConvertFrom-Json

$changed = $false

foreach ($r in $config.reports) {
    $filePath = Join-Path $SiteDir $r.file
    if (Test-Path $filePath) {
        $fileDate = (Get-Item $filePath).LastWriteTime.ToString("yyyy-MM-dd HH:mm")
        if ($r.updated -ne $fileDate) {
            Write-Host "  $($r.id): $($r.updated) -> $fileDate"
            $r.updated = $fileDate
            $changed = $true
        }
    }
}

if ($changed) {
    $config | ConvertTo-Json -Depth 10 | Set-Content $configPath -Encoding UTF8
    Write-Host "config.json atualizado."
} else {
    Write-Host "Nenhuma mudanca detectada no config.json."
}
