Attribute VB_Name = "CalendarioExport"
' ============================================================
' TRUXT Macro Research -- Calendario Economico Bloomberg
' Importar via VBA Editor > File > Import File
'
' Macros publicas:
'   CriarSheetCalendario  -- cria a sheet CALENDARIO (rodar 1x)
'   ImportarDoWECO        -- le planilha Bloomberg aberta -> CALENDARIO -> JSON
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
    wsCal.Range("A4:A203").NumberFormat = "dd/mm/yyyy"
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

    ' Remove botao anterior se existir
    Dim shp As Shape
    For Each shp In wsCal.Shapes
        If shp.Name = "btnAtualizarCalendario" Then shp.Delete
    Next shp

    ' Cria botao no canto superior direito (coluna I, linha 1-2)
    Dim btn As Shape
    Set btn = wsCal.Shapes.AddShape(msoShapeRoundedRectangle, _
                Left:=wsCal.Columns(6).Left + 4, _
                Top:=wsCal.Rows(1).Top + 4, _
                Width:=180, Height:=44)

    btn.Name = "btnAtualizarCalendario"

    With btn.Fill
        .ForeColor.RGB = RGB(0, 130, 180)   ' azul TRUXT
        .BackColor.RGB = RGB(0, 130, 180)
        .Solid
    End With

    With btn.Line
        .Visible = msoFalse
    End With

    With btn.TextFrame2
        .TextRange.Text = Chr(8595) & "  Atualizar Calendario"
        With .TextRange.Font
            .Name = "Arial"
            .Size = 11
            .Bold = msoTrue
            .Fill.ForeColor.RGB = vbWhite
        End With
        .VerticalAnchor = msoAnchorMiddle
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With

    btn.OnAction = "ImportarDoWECO"

    Application.StatusBar = "Botao criado na sheet " & CAL_SHEET
End Sub

' -----------------------------------------------------------------------
' ImportarDoWECO
' Le a planilha Bloomberg WECO aberta, preenche CALENDARIO, exporta JSON.
'
' Bloomberg exporta "Date Time C" como celula mesclada A1:C1 mas os
' dados ficam em colunas separadas: col A=Date, col B=Time, col C=Country.
' -----------------------------------------------------------------------
Public Sub ImportarDoWECO()

    ' ── 1. Encontra planilha Bloomberg aberta ────────────────────────────
    Dim wbBBG As Workbook
    Dim wsBBG As Worksheet
    Dim hRow  As Long

    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If wb.Name <> ThisWorkbook.Name Then
            Dim ws As Worksheet
            For Each ws In wb.Worksheets
                hRow = AcharCabecalho(ws)
                If hRow > 0 Then
                    Set wbBBG = wb
                    Set wsBBG = ws
                    Exit For
                End If
            Next ws
            If Not wsBBG Is Nothing Then Exit For
        End If
    Next wb

    If wsBBG Is Nothing Then
        MsgBox "Nenhuma planilha Bloomberg WECO encontrada." & vbCrLf & vbCrLf & _
               "1. No terminal Bloomberg: WECO -> Export to Excel" & vbCrLf & _
               "2. Deixe a planilha aberta" & vbCrLf & _
               "3. Rode ImportarDoWECO novamente", vbExclamation, "ImportarDoWECO"
        Exit Sub
    End If

    ' ── 2. Mapeia colunas ────────────────────────────────────────────────
    Dim colDate   As Long: colDate   = 0
    Dim colTime   As Long: colTime   = 0
    Dim colCtry   As Long: colCtry   = 0
    Dim colEvent  As Long: colEvent  = 0
    Dim colPeriod As Long: colPeriod = 0
    Dim colSurv   As Long: colSurv   = 0
    Dim colPrior  As Long: colPrior  = 0
    Dim colActual As Long: colActual = 0
    Dim colRel    As Long: colRel    = 0

    Dim lastCol As Long
    lastCol = wsBBG.Cells(hRow, wsBBG.Columns.Count).End(xlToLeft).Column

    Dim c As Long
    For c = 1 To lastCol
        Dim h As String
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

    ' Bloomberg WECO: "Date Time C" pode ser celula mesclada cobrindo A1:B1 ou A1:C1.
    ' Nesse caso colDate e encontrado mas colCtry fica vazio (header mesclado).
    ' Inferir: pais fica na coluna logo apos o datetime.
    If colDate > 0 And colCtry = 0 Then
        Dim hDateTxt As String
        hDateTxt = LCase(Trim(wsBBG.Cells(hRow, colDate).Text))
        If InStr(hDateTxt, "time") > 0 Or InStr(hDateTxt, " c") > 0 Then
            colCtry = colDate + 1
        End If
    End If

    ' Fallback: se ainda nao achou colDate, varre por celula com "date" no texto
    If colDate = 0 Then
        Dim cc As Long
        For cc = 1 To lastCol
            Dim hm As String: hm = LCase(Trim(wsBBG.Cells(hRow, cc).Text))
            If InStr(hm, "date") > 0 Then
                colDate = cc
                If colCtry = 0 Then colCtry = cc + 1
                Exit For
            End If
        Next cc
    End If

    If colDate = 0 Or colEvent = 0 Then
        MsgBox "Nao foi possivel identificar as colunas da planilha Bloomberg." & vbCrLf & _
               "Cabecalho encontrado na linha " & hRow & " da sheet '" & wsBBG.Name & "'." & vbCrLf & vbCrLf & _
               "Colunas esperadas: Date | Time | C | Event | Period | Survey(M) | Prior | Actual", _
               vbCritical, "ImportarDoWECO"
        Exit Sub
    End If

    ' ── 3. Garante sheet CALENDARIO ──────────────────────────────────────
    Dim wsCal As Worksheet
    On Error Resume Next
    Set wsCal = ThisWorkbook.Sheets(CAL_SHEET)
    On Error GoTo 0
    If wsCal Is Nothing Then
        CriarSheetCalendario
        Set wsCal = ThisWorkbook.Sheets(CAL_SHEET)
    End If

    Dim lastCalRow As Long
    lastCalRow = wsCal.Cells(wsCal.Rows.Count, 1).End(xlUp).Row
    If lastCalRow >= 4 Then wsCal.Rows("4:" & lastCalRow).ClearContents

    ' ── 4. Le e escreve os dados ─────────────────────────────────────────
    Dim lastRow  As Long: lastRow  = wsBBG.Cells(wsBBG.Rows.Count, colDate).End(xlUp).Row
    Dim destRow  As Long: destRow  = 4
    Dim imported As Long: imported = 0

    Dim rw As Long
    For rw = hRow + 1 To lastRow

        ' Data+Hora: Bloomberg exporta combinado ("6/30/2026 8:30") — separar aqui
        Dim dtRaw As String: dtRaw = Trim(wsBBG.Cells(rw, colDate).Text)
        If dtRaw = "" Then GoTo ProxLinha
        Dim dtTxt As String: dtTxt = dtRaw
        Dim tmTxt As String: tmTxt = ""
        Dim spPos As Long:   spPos = InStr(dtRaw, " ")
        If spPos > 0 Then
            dtTxt = Left(dtRaw, spPos - 1)   ' "6/30/2026"
            tmTxt = Mid(dtRaw, spPos + 1)     ' "8:30"
        End If
        ' Coluna Time separada tem prioridade (caso exista em algum export futuro)
        If colTime > 0 Then
            Dim tmSep As String: tmSep = Trim(wsBBG.Cells(rw, colTime).Text)
            If tmSep <> "" And tmSep <> "N/A" And tmSep <> "--" Then tmTxt = tmSep
        End If

        ' Evento
        Dim evTxt As String: evTxt = Trim(wsBBG.Cells(rw, colEvent).Text)
        If evTxt = "" Then GoTo ProxLinha

        ' Periodo: adiciona ao nome do evento (ex: "IPCA MoM  [May]")
        If colPeriod > 0 Then
            Dim perTxt As String: perTxt = Trim(wsBBG.Cells(rw, colPeriod).Text)
            If perTxt <> "" And perTxt <> "N/A" And perTxt <> "--" Then
                evTxt = evTxt & "  [" & perTxt & "]"
            End If
        End If

        ' Pais
        Dim ctryTxt As String: ctryTxt = ""
        If colCtry > 0 Then ctryTxt = Trim(wsBBG.Cells(rw, colCtry).Text)

        ' Relevancia (opcional)
        Dim relTxt As String: relTxt = ""
        If colRel > 0 Then relTxt = MapRel(Trim(wsBBG.Cells(rw, colRel).Text))

        ' Valores numericos
        Dim priorTxt  As String: priorTxt  = LimparNum(wsBBG.Cells(rw, colPrior).Text)
        Dim survTxt   As String: survTxt   = ""
        Dim actualTxt As String: actualTxt = ""
        If colSurv   > 0 Then survTxt   = LimparNum(wsBBG.Cells(rw, colSurv).Text)
        If colActual > 0 Then actualTxt = LimparNum(wsBBG.Cells(rw, colActual).Text)

        ' Surpresa = Actual - Survey(M)
        Dim surpresa As String: surpresa = ""
        If actualTxt <> "" And survTxt <> "" Then
            If IsNumeric(actualTxt) And IsNumeric(survTxt) Then
                surpresa = CStr(Round(CDbl(actualTxt) - CDbl(survTxt), 4))
            End If
        End If

        ' Escreve na sheet CALENDARIO (9 colunas)
        wsCal.Cells(destRow, 1).Value = dtTxt    ' Data
        wsCal.Cells(destRow, 2).Value = tmTxt    ' Hora
        wsCal.Cells(destRow, 3).Value = ctryTxt  ' Pais
        wsCal.Cells(destRow, 4).Value = evTxt    ' Evento
        wsCal.Cells(destRow, 5).Value = relTxt   ' Relevancia
        wsCal.Cells(destRow, 6).Value = priorTxt ' Anterior
        wsCal.Cells(destRow, 7).Value = survTxt  ' Estimativa
        wsCal.Cells(destRow, 8).Value = actualTxt' Realizado
        wsCal.Cells(destRow, 9).Value = surpresa ' Surpresa

        destRow  = destRow  + 1
        imported = imported + 1

ProxLinha:
    Next rw

    ' ── 5. Fecha planilha Bloomberg ──────────────────────────────────────
    Dim resp As Integer
    resp = MsgBox(imported & " eventos importados." & vbCrLf & vbCrLf & _
                  "Fechar a planilha Bloomberg (" & wbBBG.Name & ")?", _
                  vbQuestion + vbYesNo, "ImportarDoWECO")
    If resp = vbYes Then wbBBG.Close SaveChanges:=False

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

' Retorna "" para valores vazios / N/A / erros
Private Function LimparNum(txt As String) As String
    Dim t As String: t = Trim(txt)
    If t = "" Or t = "N/A" Or t = "--" Or t = "N.A." Or Left(t, 1) = "#" Then
        LimparNum = ""
    Else
        LimparNum = t
    End If
End Function
