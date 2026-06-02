Attribute VB_Name = "CalendarioExport"
' ============================================================
' TRUXT Macro Research -- Calendario Economico Bloomberg
' Importar via VBA Editor > File > Import File
'
' Macros publicas:
'   CriarSheetCalendario  -- cria a sheet CALENDARIO (rodar 1x)
'   ImportarDoWECO        -- le TODAS as planilhas Bloomberg abertas -> CALENDARIO -> JSON
'                            (abre EM e DM ao mesmo tempo, importa e ordena automaticamente)
'   ExportCalendario      -- exporta calendar.json
'
' Colunas Bloomberg WECO Export:
'   Date | Time | C (country) | Event | Period | Survey(M) | Prior | Actual
'   (Bloomberg mescla "Date Time C" em A1:C1, mas os dados ficam em cols separadas)
' ============================================================

Option Explicit

Public Const CAL_SHEET As String = "CALENDARIO"
Public Const CAL_PATH  As String = "S:\Macro\Site\data\calendar.json"

Private Function Navy() As Long:    Navy    = RGB(31, 56, 100):   End Function
Private Function NavyMid() As Long: NavyMid = RGB(31, 73, 125):   End Function
Private Function RowAlt() As Long:  RowAlt  = RGB(235, 243, 255): End Function

' -----------------------------------------------------------------------
' CriarSheetCalendario
' Layout: Data | Hora | Pais | Evento | Relevancia | Anterior | Estimativa | Realizado | Surpresa
' -----------------------------------------------------------------------
Public Sub CriarSheetCalendario()
    Dim wb As Workbook
    Set wb = ThisWorkbook

    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Sheets(CAL_SHEET).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    Dim wsCal As Worksheet
    Set wsCal = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    wsCal.Name = CAL_SHEET
    wsCal.Tab.Color = Navy()

    ' Linha 1: banner
    With wsCal.Range("A1:I1")
        .Merge
        .Value = "CALENDARIO ECONOMICO | Bloomberg WECO"
        .Font.Bold = True: .Font.Size = 14: .Font.Name = "Arial"
        .Font.Color = vbWhite
        .Interior.Color = Navy()
        .HorizontalAlignment = xlCenter
        .RowHeight = 30
    End With

    ' Linha 2: instrucao
    With wsCal.Range("A2:I2")
        .Merge
        .Value = "Para atualizar: Bloomberg Terminal -> WECO -> Export to Excel -> Rodar ImportarDoWECO"
        .Font.Name = "Arial": .Font.Size = 10: .Font.Italic = True
        .Interior.Color = NavyMid()
        .Font.Color = vbWhite
        .HorizontalAlignment = xlCenter
    End With

    ' Linha 3: cabecalhos (9 colunas — Data e Hora separados)
    Dim hdrs As Variant
    hdrs = Array("Data", "Hora", "Pais", "Evento / Indicador", "Relevancia", _
                 "Anterior", "Estimativa", "Realizado", "Surpresa")
    Dim wids As Variant
    wids = Array(12, 7, 7, 45, 12, 13, 13, 13, 11)

    Dim i As Integer
    For i = 0 To 8
        With wsCal.Cells(3, i + 1)
            .Value = hdrs(i)
            .Font.Bold = True: .Font.Name = "Arial": .Font.Size = 10
            .Font.Color = vbWhite
            .Interior.Color = Navy()
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        wsCal.Columns(i + 1).ColumnWidth = wids(i)
    Next i
    wsCal.Rows(3).RowHeight = 22

    ' Formatacao: linhas alternadas (A4:I203)
    Dim cfRng As Range
    Set cfRng = wsCal.Range("A4:I203")
    cfRng.FormatConditions.Delete
    Dim cf As Object
    Set cf = cfRng.FormatConditions.Add(Type:=xlExpression, Formula1:="=MOD(ROW()-3,2)=0")
    cf.Interior.Color = RowAlt()

    ' Bordas internas
    With wsCal.Range("A3:I203").Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous: .Weight = xlHairline: .Color = RGB(200, 215, 240)
    End With
    With wsCal.Range("A3:I203").Borders(xlInsideVertical)
        .LineStyle = xlContinuous: .Weight = xlHairline: .Color = RGB(200, 215, 240)
    End With

    ' Formatos
    wsCal.Range("A4:A203").NumberFormat = "mm/dd/yyyy"
    wsCal.Range("B4:B203").NumberFormat = "hh:mm"
    wsCal.Range("F4:I203").HorizontalAlignment = xlRight

    ' Freeze na linha 4
    wsCal.Activate
    wsCal.Range("A4").Select
    ActiveWindow.FreezePanes = True

    CriarBotaoCalendario wsCal

    MsgBox "Sheet CALENDARIO criada." & vbCrLf & vbCrLf & _
           "Para preencher:" & vbCrLf & _
           "  1. Bloomberg Terminal -> WECO -> Export to Excel" & vbCrLf & _
           "  2. Clicar o botao 'Atualizar Calendario' na sheet", vbInformation, "CalendarioExport"
End Sub

' -----------------------------------------------------------------------
' CriarBotaoCalendario
' Adiciona botao "Atualizar Calendario" na sheet.
' Pode ser chamado standalone (Ctrl+G -> CriarBotaoCalendario) se a sheet
' ja existir mas o botao nao tiver sido criado ainda.
' -----------------------------------------------------------------------
Public Sub CriarBotaoCalendario(Optional wsTarget As Worksheet = Nothing)
    Dim wsCal As Worksheet
    If wsTarget Is Nothing Then
        On Error Resume Next
        Set wsCal = ThisWorkbook.Sheets(CAL_SHEET)
        On Error GoTo 0
        If wsCal Is Nothing Then
            MsgBox "Sheet CALENDARIO nao encontrada. Execute CriarSheetCalendario primeiro.", _
                   vbExclamation, "CriarBotaoCalendario"
            Exit Sub
        End If
    Else
        Set wsCal = wsTarget
    End If

    ' Remove botao anterior se existir (Shape ou Button)
    Dim shp As Shape
    For Each shp In wsCal.Shapes
        If shp.Name = "btnAtualizarCalendario" Then shp.Delete
    Next shp

    ' Posicao: canto direito da linha 1 (fora das colunas de dados A-I)
    ' Usa coluna J como ancora para nao sobrepor o banner mesclado A1:I1
    Dim btnLeft As Double: btnLeft = wsCal.Columns(11).Left - 190
    Dim btnTop  As Double: btnTop  = wsCal.Rows(1).Top + 3
    Dim btnW    As Double: btnW    = 184
    Dim btnH    As Double: btnH    = wsCal.Rows(1).Height + wsCal.Rows(2).Height - 6

    ' Form Control Button (mais simples e confiavel que Shape)
    Dim btn As Button
    Set btn = wsCal.Buttons.Add(btnLeft, btnTop, btnW, btnH)
    With btn
        .Name      = "btnAtualizarCalendario"
        .Caption   = ChrW(8595) & " Atualizar Calendario"
        .OnAction  = "ImportarDoWECO"
        .Font.Name = "Arial"
        .Font.Size = 11
        .Font.Bold = True
    End With

    MsgBox "Botao criado na sheet " & CAL_SHEET & ".", vbInformation, "CriarBotaoCalendario"
End Sub

' -----------------------------------------------------------------------
' ImportarDoWECO
' Le TODAS as planilhas Bloomberg WECO abertas, preenche CALENDARIO,
' ordena por data+hora e exporta JSON.
'
' Fluxo: Bloomberg -> WECO EM -> Export to Excel
'         Bloomberg -> WECO DM -> Export to Excel
'         (ambas abertas) -> Rodar ImportarDoWECO -> importa as duas
'
' Bloomberg exporta "Date Time C" como celula mesclada A1:C1 mas os
' dados ficam em colunas separadas: col A=Date/hora combinado, B=Country.
' -----------------------------------------------------------------------
Public Sub ImportarDoWECO()

    ' ── 1. Coleta TODOS os workbooks Bloomberg abertos ───────────────────
    Dim bbgSheets(20) As Worksheet
    Dim bbgWBs(20)    As Workbook
    Dim bbgCount      As Long: bbgCount = 0

    Dim wbI As Workbook
    For Each wbI In Application.Workbooks
        If wbI.Name <> ThisWorkbook.Name Then
            Dim wsI As Worksheet
            For Each wsI In wbI.Worksheets
                If AcharCabecalho(wsI) > 0 Then
                    Set bbgSheets(bbgCount) = wsI
                    Set bbgWBs(bbgCount)    = wbI
                    bbgCount = bbgCount + 1
                    If bbgCount > 20 Then Exit For
                End If
            Next wsI
        End If
    Next wbI

    If bbgCount = 0 Then
        MsgBox "Nenhuma planilha Bloomberg WECO encontrada." & vbCrLf & vbCrLf & _
               "1. No terminal Bloomberg: WECO -> Export to Excel" & vbCrLf & _
               "2. Deixe a(s) planilha(s) aberta(s)" & vbCrLf & _
               "3. Rode ImportarDoWECO novamente", vbExclamation, "ImportarDoWECO"
        Exit Sub
    End If

    ' ── 2. Garante sheet CALENDARIO ──────────────────────────────────────
    Dim wsCal As Worksheet
    On Error Resume Next
    Set wsCal = ThisWorkbook.Sheets(CAL_SHEET)
    On Error GoTo 0
    If wsCal Is Nothing Then
        CriarSheetCalendario
        Set wsCal = ThisWorkbook.Sheets(CAL_SHEET)
    End If

    ' Limpa CALENDARIO uma unica vez antes de comecar
    Dim lastCalRow As Long
    lastCalRow = wsCal.Cells(wsCal.Rows.Count, 1).End(xlUp).Row
    If lastCalRow >= 4 Then wsCal.Rows("4:" & lastCalRow).ClearContents

    ' ── 3. Importa de cada fonte Bloomberg ───────────────────────────────
    Dim destRow  As Long: destRow  = 4
    Dim imported As Long: imported = 0

    Dim si   As Long
    Dim wsBBG As Worksheet
    Dim hRow  As Long

    ' Variaveis de mapeamento (reset a cada fonte)
    Dim colDate   As Long
    Dim colTime   As Long
    Dim colCtry   As Long
    Dim colEvent  As Long
    Dim colPeriod As Long
    Dim colSurv   As Long
    Dim colPrior  As Long
    Dim colActual As Long
    Dim colRel    As Long
    Dim lastCol   As Long
    Dim c         As Long
    Dim h         As String
    Dim cc        As Long
    Dim hm        As String
    Dim hDateTxt  As String

    ' Variaveis de linha
    Dim rw       As Long
    Dim dtRaw    As String
    Dim dtTxt    As String
    Dim tmTxt    As String
    Dim spPos    As Long
    Dim tmSep    As String
    Dim evTxt    As String
    Dim perTxt   As String
    Dim ctryTxt  As String
    Dim relTxt   As String
    Dim priorTxt As String
    Dim survTxt  As String
    Dim actualTxt As String
    Dim surpresa  As String
    Dim dtDate    As Date

    For si = 0 To bbgCount - 1
        Set wsBBG = bbgSheets(si)
        hRow = AcharCabecalho(wsBBG)

        ' Reset mapeamento de colunas para esta fonte
        colDate = 0: colTime = 0: colCtry = 0: colEvent = 0: colPeriod = 0
        colSurv = 0: colPrior = 0: colActual = 0: colRel = 0

        lastCol = wsBBG.Cells(hRow, wsBBG.Columns.Count).End(xlToLeft).Column

        For c = 1 To lastCol
            h = LCase(Trim(wsBBG.Cells(hRow, c).Text))
            Select Case True
                Case h Like "*date*" Or h = "data":                                   colDate   = c
                Case h = "time" Or h = "hora":                                         colTime   = c
                Case h = "c" Or h = "ctry" Or h Like "*countr*" Or h = "pais":        colCtry   = c
                Case h Like "*event*" Or h Like "*release*" Or h Like "*indicator*":  colEvent  = c
                Case h Like "*period*":                                                colPeriod = c
                Case h Like "*surv*" Or h Like "*median*" Or h Like "*forecast*":     colSurv   = c
                Case h = "prior" Or h Like "*anterior*" Or h Like "*previous*":       colPrior  = c
                Case h = "actual" Or h Like "*realiz*":                                colActual = c
            End Select
        Next c

        ' Bloomberg WECO: "Date Time C" pode ser celula mesclada
        If colDate > 0 And colCtry = 0 Then
            hDateTxt = LCase(Trim(wsBBG.Cells(hRow, colDate).Text))
            If InStr(hDateTxt, "time") > 0 Or InStr(hDateTxt, " c") > 0 Then
                colCtry = colDate + 1
            End If
        End If

        ' Fallback: varre por celula com "date" no texto
        If colDate = 0 Then
            For cc = 1 To lastCol
                hm = LCase(Trim(wsBBG.Cells(hRow, cc).Text))
                If InStr(hm, "date") > 0 Then
                    colDate = cc
                    If colCtry = 0 Then colCtry = cc + 1
                    Exit For
                End If
            Next cc
        End If

        If colDate = 0 Or colEvent = 0 Then
            MsgBox "Planilha '" & wsBBG.Name & "': nao foi possivel identificar colunas." & vbCrLf & _
                   "Cabecalho na linha " & hRow & ". Esperado: Date | Time | C | Event | ...", _
                   vbCritical, "ImportarDoWECO"
            GoTo ProxFonte
        End If

        ' Le linhas de dados
        Dim lastRow As Long: lastRow = wsBBG.Cells(wsBBG.Rows.Count, colDate).End(xlUp).Row

        For rw = hRow + 1 To lastRow

            ' Data+Hora: Bloomberg exporta combinado ("6/30/2026 8:30")
            dtRaw = Trim(wsBBG.Cells(rw, colDate).Text)
            If dtRaw = "" Then GoTo ProxLinha
            dtTxt = dtRaw: tmTxt = ""
            spPos = InStr(dtRaw, " ")
            If spPos > 0 Then
                dtTxt = Left(dtRaw, spPos - 1)   ' "6/30/2026"
                tmTxt = Mid(dtRaw, spPos + 1)     ' "8:30"
            End If
            If colTime > 0 Then
                tmSep = Trim(wsBBG.Cells(rw, colTime).Text)
                If tmSep <> "" And tmSep <> "N/A" And tmSep <> "--" Then tmTxt = tmSep
            End If

            evTxt = Trim(wsBBG.Cells(rw, colEvent).Text)
            If evTxt = "" Then GoTo ProxLinha

            If colPeriod > 0 Then
                perTxt = Trim(wsBBG.Cells(rw, colPeriod).Text)
                If perTxt <> "" And perTxt <> "N/A" And perTxt <> "--" Then
                    evTxt = evTxt & "  [" & perTxt & "]"
                End If
            End If

            ctryTxt = ""
            If colCtry > 0 Then ctryTxt = Trim(wsBBG.Cells(rw, colCtry).Text)

            relTxt = ""
            If colRel > 0 Then relTxt = MapRel(Trim(wsBBG.Cells(rw, colRel).Text))

            priorTxt = LimparNum(wsBBG.Cells(rw, colPrior).Text)
            survTxt  = "": actualTxt = ""
            If colSurv   > 0 Then survTxt   = LimparNum(wsBBG.Cells(rw, colSurv).Text)
            If colActual > 0 Then actualTxt = LimparNum(wsBBG.Cells(rw, colActual).Text)

            surpresa = ""
            If actualTxt <> "" And survTxt <> "" Then
                If IsNumeric(actualTxt) And IsNumeric(survTxt) Then
                    surpresa = CStr(Round(CDbl(actualTxt) - CDbl(survTxt), 4))
                End If
            End If

            ' Data como valor Date (locale-safe) para ordenacao e formato DD/MM/YYYY no JSON
            dtDate = ParseBBGDate(dtTxt)
            wsCal.Cells(destRow, 1).Value = IIf(dtDate > 0, dtDate, dtTxt)
            wsCal.Cells(destRow, 2).Value = tmTxt
            wsCal.Cells(destRow, 3).Value = ctryTxt
            wsCal.Cells(destRow, 4).Value = evTxt
            wsCal.Cells(destRow, 5).Value = relTxt
            wsCal.Cells(destRow, 6).Value = priorTxt
            wsCal.Cells(destRow, 7).Value = survTxt
            wsCal.Cells(destRow, 8).Value = actualTxt
            wsCal.Cells(destRow, 9).Value = surpresa

            destRow  = destRow  + 1
            imported = imported + 1

ProxLinha:
        Next rw

ProxFonte:
    Next si

    ' ── 4. Ordena CALENDARIO por Data + Hora ─────────────────────────────
    ' Funciona porque datas sao valores Date (serials), nao texto
    Dim sortEnd As Long: sortEnd = destRow - 1
    If sortEnd >= 5 Then
        wsCal.Sort.SortFields.Clear
        wsCal.Sort.SortFields.Add Key:=wsCal.Range("A4:A" & sortEnd), Order:=xlAscending
        wsCal.Sort.SortFields.Add Key:=wsCal.Range("B4:B" & sortEnd), Order:=xlAscending
        With wsCal.Sort
            .SetRange wsCal.Range("A4:I" & sortEnd)
            .Header = xlNo
            .Apply
        End With
    End If

    ' ── 5. Pergunta sobre fechar planilhas Bloomberg ──────────────────────
    Dim nomes As String: nomes = ""
    For si = 0 To bbgCount - 1
        nomes = nomes & "  - " & bbgWBs(si).Name & vbCrLf
    Next si

    Dim resp As Integer
    resp = MsgBox(imported & " eventos importados de " & bbgCount & " planilha(s):" & _
                  vbCrLf & nomes & vbCrLf & "Fechar estas planilhas?", _
                  vbQuestion + vbYesNo, "ImportarDoWECO")
    If resp = vbYes Then
        For si = 0 To bbgCount - 1
            bbgWBs(si).Close SaveChanges:=False
        Next si
    End If

    ' ── 6. Exporta JSON ──────────────────────────────────────────────────
    ExportCalendario silencioso:=False

End Sub

' -----------------------------------------------------------------------
' ExportCalendario
' Le a sheet CALENDARIO e grava S:\Macro\Site\data\calendar.json
' JSON: { date, time, country, event, relevance, prior, estimate, actual, surprise }
' -----------------------------------------------------------------------
Public Sub ExportCalendario(Optional silencioso As Boolean = False)

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(CAL_SHEET)
    On Error GoTo ErrHandler

    If ws Is Nothing Then
        If Not silencioso Then
            MsgBox "Sheet '" & CAL_SHEET & "' nao encontrada." & vbCrLf & _
                   "Execute CriarSheetCalendario() primeiro.", vbCritical
        End If
        Exit Sub
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 4 Then
        Application.StatusBar = "Calendario: sem dados para exportar"
        Exit Sub
    End If

    ' Carrega whitelist de Important Data.xlsx (nome exato BBG + pais + bold)
    Dim dictWL   As Object: Set dictWL   = CreateObject("Scripting.Dictionary")
    Dim dictBold As Object: Set dictBold = CreateObject("Scripting.Dictionary")
    CarregarWhitelist dictWL, dictBold

    Dim eventsJson As String: eventsJson = ""
    Dim evCount    As Long:   evCount    = 0
    Dim fOut       As Integer

    Dim rw As Long
    For rw = 4 To lastRow
        Dim dtTxt  As String: dtTxt  = Trim(ws.Cells(rw, 1).Text)
        Dim evName As String: evName = Trim(ws.Cells(rw, 4).Text)
        If dtTxt = "" Or evName = "" Then GoTo NextRow

        Dim tmTxt  As String: tmTxt  = Trim(ws.Cells(rw, 2).Text)
        Dim ctry   As String: ctry   = Trim(ws.Cells(rw, 3).Text)

        ' Strip sufixo [Period] para matching com whitelist
        ' Ex: "IBGE Inflation IPCA MoM  [May]" -> "IBGE Inflation IPCA MoM"
        Dim baseName As String
        Dim bPos     As Long: bPos = InStr(evName, "  [")
        If bPos = 0 Then bPos = InStr(evName, " [")
        baseName = IIf(bPos > 0, Trim(Left(evName, bPos - 1)), Trim(evName))

        ' Normaliza pais: Bloomberg WECO usa EC/CH, Important Data usa EA/CN
        Dim ctryN As String: ctryN = UCase(ctry)
        If ctryN = "EC" Then ctryN = "EA"
        If ctryN = "CH" Then ctryN = "CN"

        ' Filtra por whitelist — pula se nao estiver em Important Data.xlsx
        Dim wlKey As String: wlKey = ctryN & "|" & UCase(baseName)
        If dictWL.Count > 0 And Not dictWL.Exists(wlKey) Then GoTo NextRow

        ' Relevancia vem da whitelist (Bold = HIGH)
        Dim relOut As String
        relOut = IIf(dictBold.Exists(wlKey), "HIGH", "")

        Dim prior  As String: prior  = LimparNum(ws.Cells(rw, 6).Text)
        Dim estim  As String: estim  = LimparNum(ws.Cells(rw, 7).Text)
        Dim actual As String: actual = LimparNum(ws.Cells(rw, 8).Text)
        Dim surpr  As String: surpr  = LimparNum(ws.Cells(rw, 9).Text)

        evName = Replace(evName, """", "'")
        ctry   = Replace(ctry,   """", "'")

        If eventsJson <> "" Then eventsJson = eventsJson & "," & vbCrLf
        eventsJson = eventsJson & "    {" & _
            """date"":""" & dtTxt & """," & _
            """time"":""" & tmTxt & """," & _
            """country"":""" & ctry & """," & _
            """event"":""" & evName & """," & _
            """relevance"":""" & relOut & """," & _
            """prior"":" & IIf(prior  <> "", """" & prior  & """", "null") & "," & _
            """estimate"":" & IIf(estim  <> "", """" & estim  & """", "null") & "," & _
            """actual"":" & IIf(actual <> "", """" & actual & """", "null") & "," & _
            """surprise"":" & IIf(surpr  <> "", surpr, "null") & "}"
        evCount = evCount + 1
NextRow:
    Next rw

    Dim ts As String: ts = Format(Now, "yyyy-mm-ddThh:mm:ss")
    Dim finalJson As String
    finalJson = "{" & vbCrLf & _
                "  ""updated"":""" & ts & """," & vbCrLf & _
                "  ""count"":" & evCount & "," & vbCrLf & _
                "  ""events"":[" & vbCrLf & _
                eventsJson & vbCrLf & _
                "  ]" & vbCrLf & "}"

    fOut = FreeFile
    Open CAL_PATH For Output As #fOut
    Print #fOut, finalJson
    Close #fOut

    Application.StatusBar = "calendar.json: " & evCount & " eventos | " & Format(Now, "hh:mm")
    If Not silencioso Then
        MsgBox evCount & " eventos exportados para:" & vbCrLf & CAL_PATH, _
               vbInformation, "ExportCalendario"
    End If
    Exit Sub

ErrHandler:
    If fOut > 0 Then Close #fOut
    If Not silencioso Then
        MsgBox "Erro: " & Err.Description, vbCritical, "ExportCalendario"
    Else
        Application.StatusBar = "ERRO calendario: " & Err.Description
    End If
End Sub

' -----------------------------------------------------------------------
' Auxiliares privados
' -----------------------------------------------------------------------

' Encontra linha de cabecalho na planilha Bloomberg
Private Function AcharCabecalho(ws As Worksheet) As Long
    Dim r As Long, c As Long, txt As String
    For r = 1 To 10
        For c = 1 To 15
            txt = LCase(Trim(ws.Cells(r, c).Text))
            If txt = "date" Or txt = "data" Or txt Like "*event*" Or _
               txt Like "*release*" Or (InStr(txt, "date") > 0 And InStr(txt, "time") > 0) Then
                AcharCabecalho = r
                Exit Function
            End If
        Next c
    Next r
    AcharCabecalho = 0
End Function

' Mapeia marcador de relevancia Bloomberg -> HIGH/MEDIUM/LOW
Private Function MapRel(raw As String) As String
    Select Case UCase(Trim(raw))
        Case "***", "3", "H", "HIGH":           MapRel = "HIGH"
        Case "**",  "2", "M", "MED", "MEDIUM":  MapRel = "MEDIUM"
        Case "*",   "1", "L", "LOW":            MapRel = "LOW"
        Case Else:                               MapRel = raw
    End Select
End Function

' -----------------------------------------------------------------------
' CarregarWhitelist
' Le Important Data.xlsx (mesma pasta do market_data.xlsm) e monta
' dois dicionarios: dictWL (eventos permitidos) e dictBold (bold=HIGH).
' Chave: "PAIS|NOMEEVENTOMAIUS" (pais ja normalizado: EA, CN...)
' -----------------------------------------------------------------------
Private Sub CarregarWhitelist(dictWL As Object, dictBold As Object)
    Dim wlPath As String
    wlPath = ThisWorkbook.Path & "\Important Data.xlsx"

    If Dir(wlPath) = "" Then
        ' Arquivo nao encontrado — exporta sem filtro
        Application.StatusBar = "AVISO: Important Data.xlsx nao encontrado — exportando sem filtro"
        Exit Sub
    End If

    Dim wbWL As Workbook
    Dim wsWL As Worksheet
    On Error Resume Next
    Application.ScreenUpdating = False
    Application.DisplayAlerts  = False
    Set wbWL = Workbooks.Open(wlPath, ReadOnly:=True, UpdateLinks:=False)
    Application.ScreenUpdating = True
    Application.DisplayAlerts  = True
    On Error GoTo 0

    If wbWL Is Nothing Then
        Application.StatusBar = "AVISO: nao foi possivel abrir Important Data.xlsx"
        Exit Sub
    End If

    Set wsWL = wbWL.Sheets(1)
    Dim r As Long
    Dim lastWLRow As Long: lastWLRow = wsWL.Cells(wsWL.Rows.Count, 1).End(xlUp).Row

    For r = 3 To lastWLRow   ' linha 1 = titulo, linha 2 = cabecalho
        Dim evN  As String: evN  = Trim(wsWL.Cells(r, 1).Value)
        Dim ctN  As String: ctN  = UCase(Trim(wsWL.Cells(r, 2).Value))
        Dim bold As String: bold = UCase(Trim(wsWL.Cells(r, 3).Value))
        If evN <> "" And ctN <> "" Then
            Dim key As String: key = ctN & "|" & UCase(evN)
            dictWL(key) = True
            If bold = "YES" Then dictBold(key) = True
        End If
    Next r

    wbWL.Close False
End Sub

' Converte data Bloomberg "M/D/YYYY" em Date (locale-safe, nao usa CDate)
' Evita problemas de locale pt-BR onde CDate("6/30/2026") falharia
Private Function ParseBBGDate(dtTxt As String) As Date
    On Error Resume Next
    Dim p() As String: p = Split(Trim(dtTxt), "/")
    If UBound(p) = 2 Then
        ParseBBGDate = DateSerial(CInt(p(2)), CInt(p(0)), CInt(p(1)))
    End If
    On Error GoTo 0
End Function

' Retorna "" para valores vazios / N/A / erros
Private Function LimparNum(txt As String) As String
    Dim t As String: t = Trim(txt)
    If t = "" Or t = "N/A" Or t = "--" Or t = "N.A." Or Left(t, 1) = "#" Then
        LimparNum = ""
    Else
        LimparNum = t
    End If
End Function
