# ============================================================
# TRUXT Macro -- Exporta dados de mercado do Excel (horario)
# Agendado para rodar seg-sex das 8h as 19h, a cada 1h.
# Usa a instancia do Excel ja aberta, se houver.
# ============================================================

$xlFile  = "S:\Macro\Site\Market\market_data.xlsm"
$logFile = "S:\Macro\Site\scripts\update_log.txt"

function Log([string]$msg) {
    $ts   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$ts] [market_hourly] $msg"
    Write-Host $line
    Add-Content -Path $logFile -Value $line -Encoding UTF8
}

Log "Iniciando..."

$excel        = $null
$existingExcel = $false

try {
    # Reutiliza instancia do Excel ja aberta, se existir
    try {
        $excel        = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
        $existingExcel = $true
        Log "Excel ja esta aberto — usando instancia existente"
    } catch {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible        = $false
        $excel.DisplayAlerts  = $false
        Log "Nova instancia do Excel aberta"
    }

    # Verifica se o workbook ja esta aberto
    $wb        = $null
    $wbWasOpen = $false
    foreach ($w in $excel.Workbooks) {
        if ($w.FullName -ieq $xlFile) { $wb = $w; $wbWasOpen = $true; break }
    }

    if (-not $wb) {
        $wb = $excel.Workbooks.Open($xlFile, 0, $false)
        Log "Workbook aberto: $xlFile"
    } else {
        Log "Workbook ja estava aberto"
    }

    # Executa a macro de exportacao
    Log "Executando ExportarMercado..."
    $excel.Run("ExportarMercado")
    Log "ExportarMercado concluido"

    # Fecha o workbook apenas se foi nos quem abrimos
    if (-not $wbWasOpen) {
        $wb.Close($false)
        Log "Workbook fechado"
    }

} catch {
    Log "ERRO: $_"
} finally {
    # Fecha o Excel apenas se foi nos quem abrimos
    if (-not $existingExcel -and $null -ne $excel) {
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        Log "Instancia do Excel encerrada"
    }
}

Log "Concluido."
