Attribute VB_Name = "ExportMercado"
' ============================================================
' TRUXT Macro Research — Exportação de dados de mercado
' Importar este módulo no workbook market_data.xlsx
'
' SETUP:
'   1. Importe este .bas via VBA Editor > File > Import File
'   2. Cole o evento Workbook_Open() no módulo "EstaPastaDeTrabalho"
'   3. Estruture a sheet EXPORT conforme documentado abaixo
'
' SHEET EXPORT — layout esperado (sem cabeçalho, a partir de A1):
'   Col A  Col B       Col C         Col D
'   di     Jan/26      13.250        156        (taxa, DU)
'   di     Jul/26      13.120        306
'   ...
'   ntnb   Mai/26      7.850
'   ntnb   Ago/28      7.960
'   ...
'   fx     usdbrl      5.7812
'   fx     eurbrl      6.4253
'   fx     eurusd      1.1119
'   treasury  2y       4.850
'   treasury  5y       4.620
'   treasury  10y      4.450
'   treasury  30y      4.630
'
' Deixe Col A vazia para encerrar a leitura.
' ============================================================

Option Explicit

Public Const JSON_PATH As String = "S:\Macro\Site\data\market.json"
Public Const HISTORY_PATH As String = "S:\Macro\Site\data\market_history.json"
Public Const MAX_SNAPSHOTS As Long = 90   ' ~30 dias x 3
Public Const EXPORT_SHEET As String = "EXPORT"

' ------------------------------------------------------------------
' Ponto de entrada principal — chamado por Application.OnTime
' ------------------------------------------------------------------
Public Sub ExportarMercado()

    Dim wsExp As Worksheet
    On Error GoTo ErrHandler

    ' Verifica sheet EXPORT
    On Error Resume Next
    Set wsExp = ThisWorkbook.Sheets(EXPORT_SHEET)
    On Error GoTo ErrHandler
    If wsExp Is Nothing Then
        MsgBox "Sheet '" & EXPORT_SHEET & "' não encontrada.", vbCritical, "ExportMercado"
        Exit Sub
    End If

    ' ── 1. Lê dados da sheet EXPORT ──────────────────────────────
    Dim diArr()     As String
    Dim ntnbArr()   As String
    Dim fxArr()     As String
    Dim treasArr()  As String
    Dim diCount     As Long
    Dim ntnbCount   As Long
    Dim fxCount     As Long
    Dim treasCount  As Long

    ReDim diArr(0 To 100)
    ReDim ntnbArr(0 To 50)
    ReDim fxArr(0 To 20)
    ReDim treasArr(0 To 20)

    Dim r As Long
    r = 1
    Do While wsExp.Cells(r, 1).Value <> ""
        Dim tipo   As String
        Dim label  As String
        Dim v1     As Double
        Dim v2     As Double

        tipo  = LCase(Trim(wsExp.Cells(r, 1).Value))
        label = Trim(wsExp.Cells(r, 2).Text)   ' .Text preserva o texto exato da célula
        v1    = 0 : v2 = 0
        If IsNumeric(wsExp.Cells(r, 3).Value) Then v1 = CDbl(wsExp.Cells(r, 3).Value)
        If IsNumeric(wsExp.Cells(r, 4).Value) Then v2 = CDbl(wsExp.Cells(r, 4).Value)
        ' Ignora linhas onde o valor principal é zero ou erro (dados ausentes no BBG)
        If tipo = "di" Or tipo = "ntnb" Or tipo = "treasury" Then
            If v1 = 0 Then GoTo NextRow
        End If

        Select Case tipo
            Case "ntnb"
                ntnbArr(ntnbCount) = "{""label"":""" & label & """," & _
                                      """yield"":" & Format(v1, "0.000") & "}"
                ntnbCount = ntnbCount + 1
            Case "fx"
                fxArr(fxCount) = """" & label & """:" & Format(v1, "0.0000")
                fxCount = fxCount + 1
            Case "treasury"
                treasArr(treasCount) = "{""label"":""" & label & """," & _
                                        """yield"":" & Format(v1, "0.000") & "}"
                treasCount = treasCount + 1
        End Select
NextRow:
        r = r + 1
    Loop

    ' ── 1b. DI: lê DI_Futuro diretamente (todos os vértices) ────
    Dim wsDI2 As Worksheet
    On Error Resume Next
    Set wsDI2 = ThisWorkbook.Sheets("DI_Futuro")
    On Error GoTo ErrHandler
    If Not wsDI2 Is Nothing Then
        Dim lastDIRow2 As Long
        lastDIRow2 = wsDI2.Cells(wsDI2.Rows.Count, 1).End(xlUp).Row
        Dim cDI As Long
        cDI = 2
        Do While wsDI2.Cells(4, cDI).Value <> ""
            Dim tickTxt As String
            tickTxt = Trim(wsDI2.Cells(4, cDI).Text)
            If Left(tickTxt, 2) = "OD" Then
                Dim diVal As Double
                diVal = 0
                If IsNumeric(wsDI2.Cells(lastDIRow2, cDI).Value) Then
                    diVal = CDbl(wsDI2.Cells(lastDIRow2, cDI).Value)
                End If
                If diVal <> 0 Then
                    diArr(diCount) = "{""label"":""" & ParseDILabel(tickTxt) & """," & _
                                      """du"":0," & _
                                      """taxa"":" & Format(diVal, "0.000") & "}"
                    diCount = diCount + 1
                End If
            End If
            cDI = cDI + 1
        Loop
    End If

    ' ── 1c. Bolsas: lê sheet Bolsas ─────────────────────────────
    Dim bolsasArr() As String
    Dim bolsasCount As Long
    ReDim bolsasArr(0 To 20)
    bolsasCount = 0

    Dim wsBolsas As Worksheet
    On Error Resume Next
    Set wsBolsas = ThisWorkbook.Sheets("Bolsas")
    On Error GoTo ErrHandler
    If Not wsBolsas Is Nothing Then
        Dim lastBolsasRow As Long
        lastBolsasRow = wsBolsas.Cells(wsBolsas.Rows.Count, 1).End(xlUp).Row
        Dim cBolsas As Long
        cBolsas = 2
        Do While wsBolsas.Cells(4, cBolsas).Value <> ""
            Dim bTickFull As String
            bTickFull = Trim(wsBolsas.Cells(4, cBolsas).Text)
            Dim bLabel As String, bGrupo As String
            bLabel = "" : bGrupo = ""
            Call BolsasMeta(bTickFull, bLabel, bGrupo)
            If bLabel <> "" Then
                Dim bVal As Double
                bVal = 0
                If IsNumeric(wsBolsas.Cells(lastBolsasRow, cBolsas).Value) Then
                    bVal = CDbl(wsBolsas.Cells(lastBolsasRow, cBolsas).Value)
                End If
                If bVal <> 0 Then
                    Dim bRoot As String
                    bRoot = Split(bTickFull, " ")(0)
                    bolsasArr(bolsasCount) = "{""label"":""" & bLabel & """," & _
                                             """ticker"":""" & bRoot & """," & _
                                             """grupo"":""" & bGrupo & """," & _
                                             """value"":" & Format(bVal, "0.00") & "}"
                    bolsasCount = bolsasCount + 1
                End If
            End If
            cBolsas = cBolsas + 1
        Loop
    End If

    ' ── 1d. Commodities: lê sheet Commodities ───────────────────
    Dim commArr() As String
    Dim commCount As Long
    ReDim commArr(0 To 20)
    commCount = 0

    Dim wsComm As Worksheet
    On Error Resume Next
    Set wsComm = ThisWorkbook.Sheets("Commodities")
    On Error GoTo ErrHandler
    If Not wsComm Is Nothing Then
        Dim lastCommRow As Long
        lastCommRow = wsComm.Cells(wsComm.Rows.Count, 1).End(xlUp).Row
        Dim cComm As Long
        cComm = 2
        Do While wsComm.Cells(4, cComm).Value <> ""
            Dim cTickFull As String
            cTickFull = Trim(wsComm.Cells(4, cComm).Text)
            Dim cLabel As String, cGrupo As String, cUnit As String, cTicker As String
            cLabel = "" : cGrupo = "" : cUnit = "" : cTicker = ""
            Call CommMeta(cTickFull, cLabel, cGrupo, cUnit, cTicker)
            If cLabel <> "" Then
                Dim cVal As Double
                cVal = 0
                If IsNumeric(wsComm.Cells(lastCommRow, cComm).Value) Then
                    cVal = CDbl(wsComm.Cells(lastCommRow, cComm).Value)
                End If
                If cVal <> 0 Then
                    commArr(commCount) = "{""label"":""" & cLabel & """," & _
                                         """ticker"":""" & cTicker & """," & _
                                         """grupo"":""" & cGrupo & """," & _
                                         """unit"":""" & cUnit & """," & _
                                         """value"":" & Format(cVal, "0.0000") & "}"
                    commCount = commCount + 1
                End If
            End If
            cComm = cComm + 1
        Loop
    End If

    ' ── 2. Monta JSON do snapshot atual ──────────────────────────
    Dim ts     As String
    Dim lbl    As String
    Dim ptMes(1 To 12) As String
    ptMes(1)="jan":ptMes(2)="fev":ptMes(3)="mar":ptMes(4)="abr"
    ptMes(5)="mai":ptMes(6)="jun":ptMes(7)="jul":ptMes(8)="ago"
    ptMes(9)="set":ptMes(10)="out":ptMes(11)="nov":ptMes(12)="dez"
    ts  = Format(Now, "yyyy-mm-ddThh:mm:ss")
    lbl = Format(Day(Now), "00") & "/" & ptMes(Month(Now)) & " " & _
          Format(Hour(Now), "00") & ":" & Format(Minute(Now), "00")

    Dim snapJson As String
    snapJson = "{" & vbCrLf & _
               "      ""ts"":""" & ts & """," & vbCrLf & _
               "      ""label"":""" & lbl & """," & vbCrLf

    ' DI array
    Dim i As Long
    snapJson = snapJson & "      ""di"":["
    For i = 0 To diCount - 1
        snapJson = snapJson & diArr(i)
        If i < diCount - 1 Then snapJson = snapJson & ","
    Next i
    snapJson = snapJson & "]," & vbCrLf

    ' NTN-B array
    snapJson = snapJson & "      ""ntnb"":["
    For i = 0 To ntnbCount - 1
        snapJson = snapJson & ntnbArr(i)
        If i < ntnbCount - 1 Then snapJson = snapJson & ","
    Next i
    snapJson = snapJson & "]," & vbCrLf

    ' FX object
    snapJson = snapJson & "      ""fx"":{"
    For i = 0 To fxCount - 1
        snapJson = snapJson & fxArr(i)
        If i < fxCount - 1 Then snapJson = snapJson & ","
    Next i
    snapJson = snapJson & "}," & vbCrLf

    ' Treasuries array
    snapJson = snapJson & "      ""treasuries"":["
    For i = 0 To treasCount - 1
        snapJson = snapJson & treasArr(i)
        If i < treasCount - 1 Then snapJson = snapJson & ","
    Next i
    snapJson = snapJson & "]," & vbCrLf

    ' Bolsas array
    snapJson = snapJson & "      ""bolsas"":["
    For i = 0 To bolsasCount - 1
        snapJson = snapJson & bolsasArr(i)
        If i < bolsasCount - 1 Then snapJson = snapJson & ","
    Next i
    snapJson = snapJson & "]," & vbCrLf

    ' Commodities array
    snapJson = snapJson & "      ""commodities"":["
    For i = 0 To commCount - 1
        snapJson = snapJson & commArr(i)
        If i < commCount - 1 Then snapJson = snapJson & ","
    Next i
    snapJson = snapJson & "]" & vbCrLf & "    }"

    ' ── 3. Carrega JSON existente, prepend, trunca ───────────────
    Dim existingSnaps As String
    existingSnaps = ""
    Dim snapCount As Long
    snapCount = 0

    If Dir(JSON_PATH) <> "" Then
        Dim fNum As Integer
        fNum = FreeFile
        Dim fileContent As String
        Open JSON_PATH For Input As #fNum
        fileContent = Input(LOF(fNum), fNum)
        Close #fNum

        ' Extrai o array interno de snapshots (entre os [ ])
        Dim startPos As Long
        Dim endPos   As Long
        startPos = InStr(fileContent, """snapshots"":[")
        If startPos > 0 Then
            startPos = InStr(startPos, fileContent, "[") + 1
            endPos   = Len(fileContent)
            ' Encontra o ] de fechamento do array (do fim para o início)
            Dim j As Long
            For j = endPos To 1 Step -1
                If Mid(fileContent, j, 1) = "]" Then
                    endPos = j - 1
                    Exit For
                End If
            Next j
            existingSnaps = Trim(Mid(fileContent, startPos, endPos - startPos + 1))

            ' Conta snapshots existentes (conta "{" no nível raiz de cada objeto)
            Dim depth As Long
            depth = 0
            For j = 1 To Len(existingSnaps)
                Dim ch As String
                ch = Mid(existingSnaps, j, 1)
                If ch = "{" Then
                    depth = depth + 1
                    If depth = 1 Then snapCount = snapCount + 1
                ElseIf ch = "}" Then
                    depth = depth - 1
                End If
            Next j
        End If
    End If

    ' ── 4. Trunca para MAX_SNAPSHOTS - 1 (para abrir espaço pro novo) ──
    If snapCount >= MAX_SNAPSHOTS And existingSnaps <> "" Then
        ' Remove o último objeto (o mais antigo está no final do array)
        Dim depth2  As Long
        Dim lastObjStart As Long
        depth2 = 0
        lastObjStart = 0
        For i = Len(existingSnaps) To 1 Step -1
            Dim c As String
            c = Mid(existingSnaps, i, 1)
            If c = "}" Then
                depth2 = depth2 + 1
                If depth2 = 1 Then ' fim do último objeto
                    ' continua buscando o {
                End If
            ElseIf c = "{" Then
                depth2 = depth2 - 1
                If depth2 = 0 Then
                    lastObjStart = i
                    Exit For
                End If
            End If
        Next i
        If lastObjStart > 1 Then
            existingSnaps = RTrim(Left(existingSnaps, lastObjStart - 1))
            ' Remove vírgula final se houver
            If Right(existingSnaps, 1) = "," Then
                existingSnaps = Left(existingSnaps, Len(existingSnaps) - 1)
            End If
        End If
    End If

    ' ── 5. Monta JSON final ──────────────────────────────────────
    Dim finalJson As String
    If existingSnaps <> "" Then
        finalJson = "{" & vbCrLf & "  ""snapshots"":[" & vbCrLf & _
                    "    " & snapJson & "," & vbCrLf & _
                    "    " & existingSnaps & vbCrLf & "  ]" & vbCrLf & "}"
    Else
        finalJson = "{" & vbCrLf & "  ""snapshots"":[" & vbCrLf & _
                    "    " & snapJson & vbCrLf & "  ]" & vbCrLf & "}"
    End If

    ' ── 6. Salva o arquivo ──────────────────────────────────────
    Dim fOut As Integer
    fOut = FreeFile
    Open JSON_PATH For Output As #fOut
    Print #fOut, finalJson
    Close #fOut

    ' Atualiza histórico incremental em silêncio (1 linha nova = rápido)
    Application.StatusBar = "Atualizando histórico..."
    ExportHistorico silencioso:=True

    Application.StatusBar = "Exportado em " & Format(Now, "hh:mm:ss") & " — site atualizará em segundos"
    Exit Sub

ErrHandler:
    MsgBox "Erro ao exportar: " & Err.Description, vbCritical, "ExportMercado"
End Sub

' ------------------------------------------------------------------
' ExportHistorico — exporta série BDH para market_history.json
' Primeira execução: exporta TUDO (oldest-first).
' Execuções seguintes: incremental — acrescenta só linhas novas.
' ------------------------------------------------------------------
Public Sub ExportHistorico(Optional silencioso As Boolean = False)

    Dim wsDI As Worksheet
    Dim wsFX As Worksheet
    Dim wsTR As Worksheet
    Dim wsBolsasH As Worksheet
    Dim wsCommodH As Worksheet

    On Error GoTo ErrHandler

    On Error Resume Next
    Set wsDI      = ThisWorkbook.Sheets("DI_Futuro")
    Set wsFX      = ThisWorkbook.Sheets("FX")
    Set wsTR      = ThisWorkbook.Sheets("Treasuries")
    Set wsBolsasH = ThisWorkbook.Sheets("Bolsas")
    Set wsCommodH = ThisWorkbook.Sheets("Commodities")
    On Error GoTo ErrHandler

    If wsDI Is Nothing Then MsgBox "Sheet 'DI_Futuro' não encontrada.", vbCritical, "ExportHistorico": Exit Sub
    If wsFX Is Nothing Then MsgBox "Sheet 'FX' não encontrada.", vbCritical, "ExportHistorico": Exit Sub
    If wsTR Is Nothing Then MsgBox "Sheet 'Treasuries' não encontrada.", vbCritical, "ExportHistorico": Exit Sub

    ' ── 1. Dicionários data→linha para FX e Treasuries ───────────────────────
    Dim dictFX As Object
    Dim dictTR As Object
    Set dictFX = CreateObject("Scripting.Dictionary")
    Set dictTR = CreateObject("Scripting.Dictionary")

    Dim r As Long
    r = 7
    Do While wsFX.Cells(r, 1).Value <> ""
        If IsDate(wsFX.Cells(r, 1).Value) Then
            dictFX(Format(CDate(wsFX.Cells(r, 1).Value), "yyyy-mm-dd")) = r
        End If
        r = r + 1
    Loop

    r = 7
    Do While wsTR.Cells(r, 1).Value <> ""
        If IsDate(wsTR.Cells(r, 1).Value) Then
            dictTR(Format(CDate(wsTR.Cells(r, 1).Value), "yyyy-mm-dd")) = r
        End If
        r = r + 1
    Loop

    Dim dictBolsas As Object
    Dim dictComm   As Object
    Set dictBolsas = CreateObject("Scripting.Dictionary")
    Set dictComm   = CreateObject("Scripting.Dictionary")

    If Not wsBolsasH Is Nothing Then
        r = 7
        Do While wsBolsasH.Cells(r, 1).Value <> ""
            If IsDate(wsBolsasH.Cells(r, 1).Value) Then
                dictBolsas(Format(CDate(wsBolsasH.Cells(r, 1).Value), "yyyy-mm-dd")) = r
            End If
            r = r + 1
        Loop
    End If

    If Not wsCommodH Is Nothing Then
        r = 7
        Do While wsCommodH.Cells(r, 1).Value <> ""
            If IsDate(wsCommodH.Cells(r, 1).Value) Then
                dictComm(Format(CDate(wsCommodH.Cells(r, 1).Value), "yyyy-mm-dd")) = r
            End If
            r = r + 1
        Loop
    End If

    ' ── 2. Mapeamentos FX e Treasuries (fixos) ───────────────────────────────
    Dim fxKeys(0 To 6) As String
    Dim fxCols(0 To 6) As Long
    fxKeys(0) = "eurbrl" : fxCols(0) = 2
    fxKeys(1) = "usdbrl" : fxCols(1) = 3
    fxKeys(2) = "eurusd" : fxCols(2) = 4
    fxKeys(3) = "gbpusd" : fxCols(3) = 5
    fxKeys(4) = "usdjpy" : fxCols(4) = 6
    fxKeys(5) = "usdcny" : fxCols(5) = 7
    fxKeys(6) = "dxy"    : fxCols(6) = 8

    Dim trLbls(0 To 6) As String
    Dim trCols(0 To 6) As Long
    trLbls(0) = "2y"  : trCols(0) = 2
    trLbls(1) = "3y"  : trCols(1) = 3
    trLbls(2) = "5y"  : trCols(2) = 4
    trLbls(3) = "7y"  : trCols(3) = 5
    trLbls(4) = "10y" : trCols(4) = 6
    trLbls(5) = "20y" : trCols(5) = 7
    trLbls(6) = "30y" : trCols(6) = 8

    ' Bolsas columns (dynamic, scanned from row 4)
    Dim bolsHCols(0 To 20)    As Long
    Dim bolsHLabels(0 To 20)  As String
    Dim bolsHGrupos(0 To 20)  As String
    Dim bolsHRoots(0 To 20)   As String
    Dim bolsHCount            As Long
    bolsHCount = 0
    If Not wsBolsasH Is Nothing Then
        Dim cBH As Long
        cBH = 2
        Do While wsBolsasH.Cells(4, cBH).Value <> ""
            Dim bhTick As String : bhTick = Trim(wsBolsasH.Cells(4, cBH).Text)
            Dim bhLbl  As String : bhLbl  = ""
            Dim bhGrp  As String : bhGrp  = ""
            Call BolsasMeta(bhTick, bhLbl, bhGrp)
            If bhLbl <> "" Then
                bolsHCols(bolsHCount)   = cBH
                bolsHLabels(bolsHCount) = bhLbl
                bolsHGrupos(bolsHCount) = bhGrp
                bolsHRoots(bolsHCount)  = Split(bhTick, " ")(0)
                bolsHCount = bolsHCount + 1
            End If
            cBH = cBH + 1
        Loop
    End If

    ' Commodities columns (dynamic, scanned from row 4)
    Dim commHCols(0 To 20)    As Long
    Dim commHLabels(0 To 20)  As String
    Dim commHGrupos(0 To 20)  As String
    Dim commHUnits(0 To 20)   As String
    Dim commHTickers(0 To 20) As String
    Dim commHCount            As Long
    commHCount = 0
    If Not wsCommodH Is Nothing Then
        Dim cCH As Long
        cCH = 2
        Do While wsCommodH.Cells(4, cCH).Value <> ""
            Dim chTick As String : chTick = Trim(wsCommodH.Cells(4, cCH).Text)
            Dim chLbl  As String : chLbl  = ""
            Dim chGrp  As String : chGrp  = ""
            Dim chUnit As String : chUnit = ""
            Dim chTkr  As String : chTkr  = ""
            Call CommMeta(chTick, chLbl, chGrp, chUnit, chTkr)
            If chLbl <> "" Then
                commHCols(commHCount)    = cCH
                commHLabels(commHCount)  = chLbl
                commHGrupos(commHCount)  = chGrp
                commHUnits(commHCount)   = chUnit
                commHTickers(commHCount) = chTkr
                commHCount = commHCount + 1
            End If
            cCH = cCH + 1
        Loop
    End If

    ' ── 3. Descobre colunas DI dinamicamente (Row 4 = tickers BDH) ──────────
    Dim diDynCols(0 To 100) As Long
    Dim diDynLbls(0 To 100) As String
    Dim diDynCount As Long
    diDynCount = 0
    Dim cScan As Long
    cScan = 2
    Dim tTick As String
    Do While wsDI.Cells(4, cScan).Value <> ""
        tTick = Trim(wsDI.Cells(4, cScan).Text)
        If Left(tTick, 2) = "OD" Then
            diDynCols(diDynCount) = cScan
            diDynLbls(diDynCount) = ParseDILabel(tTick)
            diDynCount = diDynCount + 1
        End If
        cScan = cScan + 1
    Loop

    ' ── 4. Detecta última data salva — lê só os últimos 2 KB (rápido) ─────────
    Dim lastSavedDate As String
    Dim histContent   As String
    lastSavedDate = ""
    histContent   = ""

    If Dir(HISTORY_PATH) <> "" Then
        Dim fIn As Integer
        fIn = FreeFile
        Dim fileLen As Long
        Open HISTORY_PATH For Binary As #fIn
        fileLen = LOF(fIn)
        Dim readLen As Long
        readLen = IIf(fileLen > 2048, 2048, fileLen)
        Dim tail As String
        tail = Space(readLen)
        Seek #fIn, fileLen - readLen + 1
        Get #fIn, , tail
        Close #fIn

        ' InStrRev encontra a última ocorrência de "ts":"YYYY-MM-DD na cauda
        Dim sKey As String
        sKey = """ts"":"""
        Dim lastPos As Long
        lastPos = InStrRev(tail, sKey)
        If lastPos > 0 Then
            lastSavedDate = Mid(tail, lastPos + Len(sKey), 10)  ' "YYYY-MM-DD"
        End If

        ' Para o modo incremental precisamos do conteúdo completo (para truncar o "]}")
        If lastSavedDate <> "" Then
            fIn = FreeFile
            Open HISTORY_PATH For Input As #fIn
            histContent = Input(LOF(fIn), fIn)
            Close #fIn
        End If
    End If

    Dim isIncremental As Boolean
    isIncremental = (lastSavedDate <> "")

    ' ── 5. Abre arquivo para escrita ─────────────────────────────────────────
    Dim fOut As Integer
    fOut = FreeFile

    Dim firstSnap As Boolean

    If Not isIncremental Then
        ' Primeira execução: escreve do zero
        Open HISTORY_PATH For Output As #fOut
        Print #fOut, "{""snapshots"":["
        firstSnap = True
    Else
        ' Incremental: remove o fechamento "]}" e reabre para appended entries
        ' Localiza o "]" de fechamento do array (de trás para frente)
        Dim endIdx As Long
        endIdx = Len(histContent)
        Do While endIdx > 0
            If Mid(histContent, endIdx, 1) = "}" Then Exit Do
            endIdx = endIdx - 1
        Loop
        Dim arrIdx As Long
        arrIdx = endIdx - 1
        Do While arrIdx > 0
            If Mid(histContent, arrIdx, 1) = "]" Then Exit Do
            arrIdx = arrIdx - 1
        Loop
        ' Regrava sem o fechamento "]}"
        Open HISTORY_PATH For Output As #fOut
        Print #fOut, Left(histContent, arrIdx - 1)
        firstSnap = False  ' próximo entry precisa de ","
    End If

    ' ── 6. Loop principal ────────────────────────────────────────────────────
    Dim ptMes2(1 To 12) As String
    ptMes2(1)="jan":ptMes2(2)="fev":ptMes2(3)="mar":ptMes2(4)="abr"
    ptMes2(5)="mai":ptMes2(6)="jun":ptMes2(7)="jul":ptMes2(8)="ago"
    ptMes2(9)="set":ptMes2(10)="out":ptMes2(11)="nov":ptMes2(12)="dez"

    Dim snapCount As Long
    snapCount = 0

    Dim lastDIRow As Long
    lastDIRow = wsDI.Cells(wsDI.Rows.Count, 1).End(xlUp).Row

    Dim dtDI       As Date
    Dim dateStr    As String
    Dim diJson     As String
    Dim diCount    As Long
    Dim fxJson     As String
    Dim trJson     As String
    Dim bolsasJson As String
    Dim commodJson As String
    Dim tsStr      As String
    Dim lblStr     As String
    Dim snap       As String
    Dim vDI        As Double
    Dim vFX        As Double
    Dim vTR        As Double
    Dim vBH        As Double
    Dim vCH        As Double
    Dim rFX        As Long
    Dim rTR        As Long
    Dim rBH        As Long
    Dim rCH        As Long
    Dim k          As Long
    Dim m          As Long

    For r = 7 To lastDIRow
        If wsDI.Cells(r, 1).Value = "" Then GoTo NextHistRow
        If Not IsDate(wsDI.Cells(r, 1).Value) Then GoTo NextHistRow

        dtDI    = CDate(wsDI.Cells(r, 1).Value)
        dateStr = Format(dtDI, "yyyy-mm-dd")

        ' Pula datas já exportadas (modo incremental)
        If isIncremental And dateStr <= lastSavedDate Then GoTo NextHistRow

        ' Monta DI array (dinâmico)
        diJson  = ""
        diCount = 0
        For k = 0 To diDynCount - 1
            vDI = 0
            If IsNumeric(wsDI.Cells(r, diDynCols(k)).Value) Then
                vDI = CDbl(wsDI.Cells(r, diDynCols(k)).Value)
            End If
            If vDI <> 0 Then
                If diJson <> "" Then diJson = diJson & ","
                diJson = diJson & "{""label"":""" & diDynLbls(k) & """,""du"":0,""taxa"":" & Format(vDI, "0.000") & "}"
                diCount = diCount + 1
            End If
        Next k

        If diCount = 0 Then GoTo NextHistRow

        ' Monta FX object
        fxJson = ""
        If dictFX.Exists(dateStr) Then
            rFX = dictFX(dateStr)
            For m = 0 To 6
                vFX = 0
                If IsNumeric(wsFX.Cells(rFX, fxCols(m)).Value) Then
                    vFX = CDbl(wsFX.Cells(rFX, fxCols(m)).Value)
                End If
                If vFX <> 0 Then
                    If fxJson <> "" Then fxJson = fxJson & ","
                    fxJson = fxJson & """" & fxKeys(m) & """:" & Format(vFX, "0.0000")
                End If
            Next m
        End If

        ' Monta Treasuries array
        trJson = ""
        If dictTR.Exists(dateStr) Then
            rTR = dictTR(dateStr)
            For m = 0 To 6
                vTR = 0
                If IsNumeric(wsTR.Cells(rTR, trCols(m)).Value) Then
                    vTR = CDbl(wsTR.Cells(rTR, trCols(m)).Value)
                End If
                If vTR <> 0 Then
                    If trJson <> "" Then trJson = trJson & ","
                    trJson = trJson & "{""label"":""" & trLbls(m) & """,""yield"":" & Format(vTR, "0.000") & "}"
                End If
            Next m
        End If

        ' Monta Bolsas array
        bolsasJson = ""
        If bolsHCount > 0 And dictBolsas.Exists(dateStr) Then
            rBH = dictBolsas(dateStr)
            For k = 0 To bolsHCount - 1
                vBH = 0
                If IsNumeric(wsBolsasH.Cells(rBH, bolsHCols(k)).Value) Then
                    vBH = CDbl(wsBolsasH.Cells(rBH, bolsHCols(k)).Value)
                End If
                If vBH <> 0 Then
                    If bolsasJson <> "" Then bolsasJson = bolsasJson & ","
                    bolsasJson = bolsasJson & "{""label"":""" & bolsHLabels(k) & """," & _
                                              """ticker"":""" & bolsHRoots(k) & """," & _
                                              """grupo"":""" & bolsHGrupos(k) & """," & _
                                              """value"":" & Format(vBH, "0.00") & "}"
                End If
            Next k
        End If

        ' Monta Commodities array
        commodJson = ""
        If commHCount > 0 And dictComm.Exists(dateStr) Then
            rCH = dictComm(dateStr)
            For k = 0 To commHCount - 1
                vCH = 0
                If IsNumeric(wsCommodH.Cells(rCH, commHCols(k)).Value) Then
                    vCH = CDbl(wsCommodH.Cells(rCH, commHCols(k)).Value)
                End If
                If vCH <> 0 Then
                    If commodJson <> "" Then commodJson = commodJson & ","
                    commodJson = commodJson & "{""label"":""" & commHLabels(k) & """," & _
                                              """ticker"":""" & commHTickers(k) & """," & _
                                              """grupo"":""" & commHGrupos(k) & """," & _
                                              """unit"":""" & commHUnits(k) & """," & _
                                              """value"":" & Format(vCH, "0.0000") & "}"
                End If
            Next k
        End If

        ' Timestamp e label (formato "22/mai/26")
        tsStr  = dateStr & "T18:00:00"
        lblStr = Format(Day(dtDI), "0") & "/" & ptMes2(Month(dtDI)) & "/" & _
                 Right(CStr(Year(dtDI)), 2)

        snap = "{""ts"":""" & tsStr & """," & _
               """label"":""" & lblStr & """," & _
               """di"":[" & diJson & "]," & _
               """ntnb"":[]," & _
               """fx"":{" & fxJson & "}," & _
               """treasuries"":[" & trJson & "]," & _
               """bolsas"":[" & bolsasJson & "]," & _
               """commodities"":[" & commodJson & "]}"

        If Not firstSnap Then Print #fOut, "  ,"
        Print #fOut, "  " & snap
        firstSnap = False
        snapCount = snapCount + 1

        If snapCount Mod 100 = 0 Then
            Application.StatusBar = "ExportHistorico: " & snapCount & " novos snapshots..."
        End If

NextHistRow:
    Next r

    Print #fOut, "]}"
    Close #fOut

    Dim modeStr As String
    modeStr = IIf(isIncremental, " (incremental)", " (completo)")
    Application.StatusBar = "market_history.json: " & snapCount & " snapshots" & modeStr

    If Not silencioso Then
        MsgBox snapCount & " snapshots exportados" & modeStr & ":" & vbCrLf & HISTORY_PATH, _
               vbInformation, "ExportHistorico"
    End If
    Exit Sub

ErrHandler:
    If fOut > 0 Then Close #fOut
    If Not silencioso Then
        MsgBox "Erro ao exportar histórico: " & Err.Description, vbCritical, "ExportHistorico"
    Else
        Application.StatusBar = "ERRO no histórico: " & Err.Description
    End If
End Sub

' ------------------------------------------------------------------
' BolsasMeta — retorna label e grupo para ticker de bolsa BBG
' ------------------------------------------------------------------
Private Sub BolsasMeta(tickerFull As String, ByRef outLabel As String, ByRef outGrupo As String)
    Dim t As String : t = UCase(tickerFull)
    If      InStr(t, "IBOV ")   = 1 Then outLabel = "Ibovespa"     : outGrupo = "América"
    ElseIf  InStr(t, "SPX ")    = 1 Then outLabel = "S&P 500"       : outGrupo = "América"
    ElseIf  InStr(t, "ES1 ")    = 1 Then outLabel = "S&P 500 Fut."  : outGrupo = "América"
    ElseIf  InStr(t, "NQ1 ")    = 1 Then outLabel = "Nasdaq Fut."   : outGrupo = "América"
    ElseIf  InStr(t, "INDU ")   = 1 Then outLabel = "Dow Jones"     : outGrupo = "América"
    ElseIf  InStr(t, "SX5E ")   = 1 Then outLabel = "Euro Stoxx 50" : outGrupo = "Europa"
    ElseIf  InStr(t, "NKY ")    = 1 Then outLabel = "Nikkei 225"    : outGrupo = "Ásia"
    ElseIf  InStr(t, "HSI ")    = 1 Then outLabel = "Hang Seng"     : outGrupo = "Ásia"
    ElseIf  InStr(t, "SHSZ300") = 1 Then outLabel = "CSI 300"       : outGrupo = "Ásia"
    ElseIf  InStr(t, "KOSPI")   = 1 Then outLabel = "KOSPI"         : outGrupo = "Ásia"
    ElseIf  InStr(t, "VIX ")    = 1 Then outLabel = "VIX"           : outGrupo = "Volatilidade"
    ElseIf  InStr(t, "MOVE ")   = 1 Then outLabel = "MOVE"          : outGrupo = "Volatilidade"
    Else                                  outLabel = ""              : outGrupo = ""
    End If
End Sub

' ------------------------------------------------------------------
' CommMeta — retorna label, grupo, unit e ticker para commodity BBG
' ------------------------------------------------------------------
Private Sub CommMeta(tickerFull As String, ByRef outLabel As String, ByRef outGrupo As String, ByRef outUnit As String, ByRef outTicker As String)
    Dim t As String : t = UCase(tickerFull)
    If      InStr(t, "CL1 ") = 1 Then outLabel = "WTI Crude" : outGrupo = "Energia"   : outUnit = "USD/bbl"   : outTicker = "CL1"
    ElseIf  InStr(t, "CO1 ") = 1 Then outLabel = "Brent"      : outGrupo = "Energia"   : outUnit = "USD/bbl"   : outTicker = "CO1"
    ElseIf  InStr(t, "NG1 ") = 1 Then outLabel = "Nat. Gas"   : outGrupo = "Energia"   : outUnit = "USD/MMBtu" : outTicker = "NG1"
    ElseIf  InStr(t, "GC1 ") = 1 Then outLabel = "Ouro"       : outGrupo = "Metais"    : outUnit = "USD/oz"    : outTicker = "GC1"
    ElseIf  InStr(t, "SI1 ") = 1 Then outLabel = "Prata"      : outGrupo = "Metais"    : outUnit = "USD/oz"    : outTicker = "SI1"
    ElseIf  InStr(t, "HG1 ") = 1 Then outLabel = "Cobre"      : outGrupo = "Metais"    : outUnit = "USD/lb"    : outTicker = "HG1"
    ElseIf  InStr(t, "S 1 ") = 1 Then outLabel = "Soja"       : outGrupo = "Agrícolas" : outUnit = "USD/bu"    : outTicker = "S 1"
    ElseIf  InStr(t, "C 1 ") = 1 Then outLabel = "Milho"      : outGrupo = "Agrícolas" : outUnit = "USD/bu"    : outTicker = "C 1"
    ElseIf  InStr(t, "LC1 ") = 1 Then outLabel = "Boi Gordo"  : outGrupo = "Agrícolas" : outUnit = "USD/lb"    : outTicker = "LC1"
    Else                               outLabel = ""           : outGrupo = ""          : outUnit = ""          : outTicker = ""
    End If
End Sub

' ------------------------------------------------------------------
' ParseDILabel — converte ticker Bloomberg → label legível
' Ex.: "ODM26 COMB Comdty" → "Jun/26"
' Códigos de mês BBG: F=Jan G=Feb H=Mar J=Apr K=Mai M=Jun
'                     N=Jul Q=Ago U=Set V=Out X=Nov Z=Dez
' ------------------------------------------------------------------
Public Function ParseDILabel(tickerText As String) As String
    Dim root      As String
    Dim monthCode As String
    Dim yearStr   As String
    Dim codes     As String
    Dim idx       As Long
    Dim names(1 To 12) As String

    root = Split(Trim(tickerText), " ")(0)      ' "ODM26"
    If Len(root) < 4 Or Left(root, 2) <> "OD" Then
        ParseDILabel = tickerText : Exit Function
    End If

    monthCode = Mid(root, 3, 1)                  ' "M"
    yearStr   = Right(root, 2)                    ' "26"
    codes     = "FGHJKMNQUVXZ"

    names(1) ="Jan": names(2) ="Fev": names(3) ="Mar": names(4) ="Abr"
    names(5) ="Mai": names(6) ="Jun": names(7) ="Jul": names(8) ="Ago"
    names(9) ="Set": names(10)="Out": names(11)="Nov": names(12)="Dez"

    idx = InStr(codes, UCase(monthCode))
    If idx = 0 Then ParseDILabel = tickerText : Exit Function

    ParseDILabel = names(idx) & "/" & yearStr
End Function

' ------------------------------------------------------------------
' Cole este código no módulo "EstaPastaDeTrabalho" (ThisWorkbook)
' ------------------------------------------------------------------
'
' Private Sub Workbook_Open()
'     ' Agenda exportações automáticas nos 3 horários
'     Dim horarios As Variant
'     Dim t As Variant
'     horarios = Array(TimeValue("09:05:00"), TimeValue("13:00:00"), TimeValue("18:00:00"))
'     For Each t In horarios
'         If Now < (Int(Now) + t) Then
'             Application.OnTime Int(Now) + t, "ExportarMercado"
'         End If
'     Next t
' End Sub
'
' ------------------------------------------------------------------
