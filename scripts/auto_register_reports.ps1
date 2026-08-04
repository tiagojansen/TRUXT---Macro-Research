# auto_register_reports.ps1
# Detecta HTMLs novos em reports/ e PDFs novos em presentations/ e os registra
# automaticamente no config.json. Executado pelo watcher antes de cada publicacao.

$SITE_DIR = Split-Path $PSScriptRoot -Parent
$CONFIG   = "$SITE_DIR\config.json"
$REPORTS  = "$SITE_DIR\reports"

$folderToCountry = @{
    "brasil"    = "brasil"
    "us"        = "eua"
    "china"     = "china"
    "zona_euro" = "zona_euro"
    "mexico"    = "mexico"
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

$changed = $newEntries.Count -gt 0

if ($newEntries.Count -gt 0) {
    $allReports = [System.Collections.ArrayList]@($config.reports)
    foreach ($e in $newEntries) { $allReports.Add($e) | Out-Null }
    $config.reports = $allReports.ToArray()
    Write-Host "  $($newEntries.Count) report(s) adicionado(s) ao config.json"
}

# ── Apresentacoes (PDFs em presentations/) ───────────────────────────────────
$PRES = "$SITE_DIR\presentations"

if (Test-Path $PRES) {
    $registeredPres = [System.Collections.Generic.HashSet[string]]@(
        $config.presentations | ForEach-Object { $_.file }
    )

    $newPres = [System.Collections.ArrayList]::new()

    Get-ChildItem -Path $PRES -Filter "*.pdf" | ForEach-Object {
        $relPath = "presentations/" + $_.Name

        if ($registeredPres.Contains($relPath)) { return }

        $filename = [System.IO.Path]::GetFileNameWithoutExtension($_.Name)
        $id       = ($filename.ToLower() -replace "[\s\.\-]", "_" -replace "_+", "_")
        $title    = ($filename -replace "[_]", " ")
        $date     = $_.LastWriteTime.ToString("yyyy-MM-dd")

        $entry = [PSCustomObject]@{
            id          = $id
            title       = $title
            file        = $relPath
            date        = $date
            description = ""
        }

        $newPres.Add($entry) | Out-Null
        $registeredPres.Add($relPath) | Out-Null
        Write-Host "  [auto-ppt] $relPath  →  id=$id, date=$date"
    }

    if ($newPres.Count -gt 0) {
        $allPres = [System.Collections.ArrayList]@($config.presentations)
        foreach ($e in $newPres) { $allPres.Add($e) | Out-Null }
        $config.presentations = $allPres.ToArray()
        $changed = $true
        Write-Host "  $($newPres.Count) apresentacao(oes) adicionada(s) ao config.json"
    }
}

if ($changed) {
    $config | ConvertTo-Json -Depth 10 | Set-Content $CONFIG -Encoding UTF8
}
