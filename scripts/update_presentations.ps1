# ============================================================
# TRUXT Macro Research — Auto-publicação de apresentações
#
# Uso: powershell -ExecutionPolicy Bypass -File update_presentations.ps1
#
# O script escaneia presentations/ (incluindo subpastas),
# detecta PDFs novos, extrai data do nome do arquivo,
# atualiza config.json e faz git push automaticamente.
#
# Convenção de nome recomendada (opcional — fallback para data do arquivo):
#   YYYY-MM-DD_Titulo_Aqui.pdf        ex: 2026-05-09_Pre-IPCA-Maio.pdf
#   YYYY.MM.DD Titulo Aqui.pdf        ex: 2026.05.09 Pre-IPCA Maio.pdf
# ============================================================

$SITE_DIR  = "S:\Macro\Site"
$CONFIG    = "$SITE_DIR\config.json"
$PPTS_DIR  = "$SITE_DIR\presentations"
$LOG_FILE  = "$SITE_DIR\scripts\update_log.txt"

function Log($msg) {
    $line = "$(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')  $msg"
    Write-Host $line
    Add-Content -Path $LOG_FILE -Value $line -Encoding UTF8
}

# ── Extrai data do nome do arquivo ─────────────────────────────────────────
function Get-PptDate($file) {
    $name = $file.BaseName

    # YYYY.MM.DD ou YYYY-MM-DD
    if ($name -match '(\d{4})[.\-](\d{2})[.\-](\d{2})') {
        return "$($Matches[1])-$($Matches[2])-$($Matches[3])"
    }
    # YYYY.MM ou YYYY-MM
    if ($name -match '(\d{4})[.\-](\d{2})') {
        return "$($Matches[1])-$($Matches[2])-01"
    }
    # Mês português + 2 dígitos de ano  (ex: Abr26, Mai26)
    $ptMap = @{jan="01";fev="02";mar="03";abr="04";mai="05";jun="06";
               jul="07";ago="08";set="09";out="10";nov="11";dez="12"}
    foreach ($m in $ptMap.Keys) {
        if ($name -match "(?i)$m[_\-\.]?(\d{2})\b") {
            return "20$($Matches[1])-$($ptMap[$m])-01"
        }
    }
    # Fallback: data de modificação do arquivo
    return $file.LastWriteTime.ToString("yyyy-MM-dd")
}

# ── Gera título legível a partir do nome do arquivo ────────────────────────
function Get-PptTitle($basename) {
    $t = $basename
    $t = $t -replace '(\d{4})[.\-]\d{2}[.\-]\d{2}', ''   # remove YYYY.MM.DD
    $t = $t -replace '(\d{4})[.\-]\d{2}', ''              # remove YYYY.MM
    $t = $t -replace '[_\-\.]+', ' '
    $t = $t.Trim()
    if ($t.Length -gt 0) { $t = $t[0].ToString().ToUpper() + $t.Substring(1) }
    return $t
}

# ── Gera ID único a partir do nome do arquivo ──────────────────────────────
function Get-PptId($basename) {
    return ($basename -replace '[^a-zA-Z0-9]', '_' -replace '_+', '_').ToLower().Trim('_')
}

# ── Main ───────────────────────────────────────────────────────────────────
Log "=== Iniciando update_presentations ==="

# Carrega config.json
$raw    = Get-Content $CONFIG -Raw -Encoding UTF8
$config = $raw | ConvertFrom-Json

# Garante que presentations é array mutável
$existing = [System.Collections.Generic.List[object]]::new()
foreach ($p in $config.presentations) { $existing.Add($p) }

# Caminhos já cadastrados
$existingPaths = $existing | ForEach-Object { $_.file }

# Escaneia presentations/ (flat + 1 nível de subpasta)
$pdfs = Get-ChildItem "$PPTS_DIR\*.pdf" -ErrorAction SilentlyContinue
$pdfs += Get-ChildItem "$PPTS_DIR\*\*.pdf" -ErrorAction SilentlyContinue

$newCount = 0
foreach ($pdf in $pdfs) {
    # Caminho relativo ao SITE_DIR (ex: presentations/2026-05/arquivo.pdf)
    $relPath = $pdf.FullName.Replace("$SITE_DIR\", "").Replace("\", "/")

    if ($existingPaths -contains $relPath) {
        Log "Ja cadastrado: $relPath"
        continue
    }

    $date  = Get-PptDate $pdf
    $title = Get-PptTitle $pdf.BaseName
    $id    = Get-PptId $pdf.BaseName

    $entry = [PSCustomObject]@{
        id          = $id
        title       = $title
        file        = $relPath
        date        = $date
        description = "Buy Side · Brasil"
    }
    $existing.Add($entry)
    $newCount++
    Log "Novo: $relPath  |  data=$date  |  titulo=$title"
}

if ($newCount -eq 0) {
    Log "Nenhum PDF novo encontrado."
    Log "=== Concluido ===`n"
    exit 0
}

# Ordena por data decrescente
$sorted = $existing | Sort-Object date -Descending

# Recria config.json preservando todas as outras chaves
$config.presentations = $sorted
$json = $config | ConvertTo-Json -Depth 5 -Compress:$false
# Remove BOM e garante UTF-8
$utf8NoBom = [System.Text.UTF8Encoding]::new($false)
[System.IO.File]::WriteAllText($CONFIG, $json, $utf8NoBom)
Log "config.json atualizado ($newCount novo(s))."

# Git
Set-Location $SITE_DIR
git add "config.json"

# Adiciona os PDFs novos ao stage
foreach ($pdf in $pdfs) {
    $relPath = $pdf.FullName.Replace("$SITE_DIR\", "").Replace("\", "/")
    git add $relPath 2>$null
}

$commitMsg = "feat: add $newCount apresentacao(oes) $(Get-Date -Format 'yyyy-MM-dd')"
$result = git commit -m $commitMsg 2>&1
if ($result -match "nothing to commit") {
    Log "Nada para commitar."
} else {
    git push 2>&1 | ForEach-Object { Log $_ }
    Log "Publicado no GitHub!"
}

Log "=== Concluido ===`n"
