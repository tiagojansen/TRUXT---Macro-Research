# ============================================================
# TRUXT Macro -- Atualizacao horaria de mercado
# Seg-sex, 8:59 ate 18:59: Bloomberg refresh + save + ExportarMercado
# Roda em background. Nao precisa de admin.
# ============================================================

$xlFile  = "S:\Macro\Site\Market\market_data.xlsm"
$logFile = "S:\Macro\Site\scripts\update_log.txt"

function Log($msg) {
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $logFile -Value "[$ts] [hourly] $msg" -Encoding UTF8
}

Log "Servico horario iniciado"
$lastRun = [datetime]::MinValue

while ($true) {
    $now  = Get-Date
    $dow  = [int]$now.DayOfWeek   # 0=Dom, 1=Seg...5=Sex, 6=Sab
    $h    = $now.Hour
    $m    = $now.Minute

    $isWeekday    = ($dow -ge 1 -and $dow -le 5)
    $isTargetTime = ($m -eq 59 -and $h -ge 8 -and $h -le 18)
    $cooldownOk   = (($now - $lastRun).TotalMinutes -gt 30)

    if ($isWeekday -and $isTargetTime -and $cooldownOk) {
        Log "=== Disparando $($h):$($m) ==="
        try {
            # Tenta pegar Excel ja aberto
            $excel = $null
            try {
                $excel = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
                Log "Excel encontrado"
            } catch {
                Log "AVISO: Excel nao estava aberto — abrindo..."
                $excel = New-Object -ComObject Excel.Application
                $excel.Visible       = $true
                $excel.DisplayAlerts = $false
            }

            # Localiza o workbook
            $wb = $null
            foreach ($w in $excel.Workbooks) {
                if ($w.FullName -ieq $xlFile) { $wb = $w; break }
            }
            if (-not $wb) {
                Log "Abrindo workbook..."
                $wb = $excel.Workbooks.Open($xlFile, 0, $false)
                Start-Sleep -Seconds 5
            }

            # Bloomberg refresh (equivalente a Ctrl+Alt+F9)
            Log "Bloomberg refresh..."
            $excel.CalculateFull()
            Start-Sleep -Seconds 30   # aguarda Bloomberg recalcular

            # Salva
            $wb.Save()
            Log "Workbook salvo"

            # Exporta (market.json + market_history.json)
            $excel.Run("ExportarMercado")
            Log "ExportarMercado concluido — watcher publicara em ~10s"

            $lastRun = Get-Date
        } catch {
            Log "ERRO: $_"
        }
        Log "=== Fim $($h):$($m) ===`n"
    }

    Start-Sleep -Seconds 30
}
