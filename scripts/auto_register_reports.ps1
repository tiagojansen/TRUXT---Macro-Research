# auto_register_reports.ps1
# Detecta HTMLs novos em reports/ e os registra automaticamente no config.json.
# Executado pelo watcher antes de cada publicacao de reports.

$SITE_DIR = Split-Path $PSScriptRoot -Parent
$CONFIG   = "$SITE_DIR\config.json"
$REPORTS  = "$SITE_DIR\reports"

$folderToCountry = @{
    "brasil"    = "brasil"
    "us"        = "eua"
    "china"     = "china"
    "zona_euro" = "zona_euro"
}

$raw    = Get-Content $CONFIG -Raw -Encoding UTF8
$config = $raw | ConvertFrom-Json

$registeredFiles = [System.Collections.Generic.HashSet[string]]@(
    $config.reports | ForEach-Object { $_.file }
)

$newEntries = [System.Collections.ArrayList]::new()

Get-ChildItem -Path $REPORTS -Recurse -Filter "*.html" | ForEach-Object {
    $relPath = $_.FullName.Substring($SITE_DIR.Length + 1).Replace("\", "/")

    if ($registeredFiles.Contains($relPath)) { return }

    # Espera estrutura: reports/{countryFolder}/{section}/filename.html
    $parts = $relPath.Split("/")
    if ($parts.Count -lt 4) { return }

    $countryFolder = $parts[1]
    $section       = $parts[2]
    $filename      = [System.IO.Path]::GetFileNameWithoutExtension($_.Name)
    $country       = $folderToCountry[$countryFolder]

    if (-not $country) { return }  # pasta de pais desconhecida

    $id    = ($filename.ToLower() -replace "[-\s]", "_")
    $title = ($filename -replace "[_\-]", " ")
    $upd   = $_.LastWriteTime.ToString("yyyy-MM-dd HH:mm")

    $entry = [PSCustomObject]@{
        id      = $id
        title   = $title
        country = $country
        section = $section
        file    = $relPath
        updated = $upd
    }

    $newEntries.Add($entry) | Out-Null
    $registeredFiles.Add($relPath) | Out-Null
    Write-Host "  [auto] $relPath  →  id=$id, country=$country, section=$section"
}

if ($newEntries.Count -gt 0) {
    $allReports = [System.Collections.ArrayList]@($config.reports)
    foreach ($e in $newEntries) { $allReports.Add($e) | Out-Null }
    $config.reports = $allReports.ToArray()
    $config | ConvertTo-Json -Depth 10 | Set-Content $CONFIG -Encoding UTF8
    Write-Host "  $($newEntries.Count) report(s) adicionado(s) ao config.json"
}
