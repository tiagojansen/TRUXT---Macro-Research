Attribute VB_Name = "CalendarioExport"
' ============================================================
' TRUXT Macro Research -- Calendario Economico Bloomberg
' Importar via VBA Editor > File > Import File
'
' Macros publicas:
'   CriarSheetCalendario  -- cria a sheet CALENDARIO (rodar 1x)
'   ImportarDoWECO        -- le planilha Bloomberg aberta -> CALENDARIO -> JSON
'   ExportCalendario      -- exporta calendar.json (chamado pelo ImportarDoWECO)
' ============================================================

Option Explicit

Public Const CAL_SHEET As String = "CALENDARIO"
Public Const CAL_PATH  As String = "S:\Macro\Site\data\calendar.json"

Private Function Navy() As Long:    Navy    = RGB(31, 56, 100):   End Function
Private Function NavyMid() As Long: NavyMid = RGB(31, 73, 125):   End Function
Private Function RowAlt() As Long:  RowAlt  = RGB(235, 243, 255): End Function

' -----------------------------------------------------------------------
' CriarSheetCalendario
' Cria a sheet CALENDARIO vazia (tabela destino para ImportarDoWECO).
' Rodar UMA vez apos importar o modulo.
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
    With wsCal.Range("A1:H1")
        .Merge
        .Value = "CALENDARIO ECONOMICO | Bloomberg WECO"
        .Font.Bold = True: .Font.Size = 14: .Font.Name = "Arial"
        .Font.Color = vbWhite
        .Interior.Color = Navy()
        .HorizontalAlignment = xlCenter
        .RowHeight = 30
    End With

    ' Linha 2: instrucao
    With wsCal.Range("A2:H2")
        .Merge
        .Value = "Para atualizar: Bloomberg Terminal -> WECO -> Export to Excel -> Rodar ImportarDoWECO"
        .Font.Name = "Arial": .Font.Size = 10: .Font.Italic = True
        .Interior.Color = NavyMid()
        .Font.Color = vbWhite
        .HorizontalAlignment = xlCenter
    End With

    ' Linha 3: cabecalhos
    Dim hdrs As Variant
    hdrs = Array("Data / Hora", "Pais", "Evento / Indicador", "Relevancia", _
                 "Anterior", "Estimativa", "Realizado", "Surpresa")
    Dim wids As Variant
    wids = Array(20, 7, 45, 12, 13, 13, 13, 11)

    Dim i As Integer
    For i = 0 To 7
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

    ' Formatacao: linhas alternadas
    Dim cfRng As Range
    Set cfRng = wsCal.Range("A4:H203")
    cfRng.FormatConditions.Delete
    Dim cf As Object
    Set cf = cfRng.FormatConditions.Add(Type:=xlExpression, Formula1:="=MOD(ROW()-3,2)=0")
    cf.Interior.Color = RowAlt()

    ' Bordas internas
    With wsCal.Range("A3:H203").Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous: .Weight = xlHairline: .Color = RGB(200, 215, 240)
    End With
    With wsCal.Range("A3:H203").Borders(xlInsideVertical)
        .LineStyle = xlContinuous: .Weight = xlHairline: .Color = RGB(200, 215, 240)
    End With

    ' Formato e alinhamentos
    wsCal.Range("A4:A203").NumberFormat = "dd/mm/yyyy hh:mm"
    wsCal.Range("E4:H203").HorizontalAlignment = xlRight

    ' Freeze na linha 4
    wsCal.Activate
    wsCal.Range("A4").Select
    ActiveWindow.FreezePanes = True

    MsgBox "Sheet CALENDARIO criada." & vbCrLf & vbCrLf & _
           "Para preencher:" & vbCrLf & _
           "  1. Bloomberg Terminal -> WECO -> Export to Excel" & vbCrLf & _
           "  2. Rodar ImportarDoWECO", vbInformation, "CalendarioExport"
End Sub

' -----------------------------------------------------------------------
' ImportarDoWECO
' Le a planilha Bloomberg WECO aberta, preenche CALENDARIO, exporta JSON.
' Workflow:
'   1. Bloomberg Terminal -> WECO -> Export to Excel  (planilha fica aberta)
'   2. Rodar esta macro no market_data.xlsm
' -----------------------------------------------------------------------
Public Sub ImportarDoWECO()

    ' ── 1. Encontra a planilha Bloomberg aberta ──────────────────────────
    Dim wbBBG  As Workbook
    Dim wsBBG  As Worksheet
    Dim hRow   As Long

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

    ' ── 2. Mapeia colunas pelo cabecalho ─────────────────────────────────
    ' Colunas esperadas do Bloomberg WECO: Date | Time | C (country) | Event | Period | Surv(M) | Prior | Actual
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
            Case h = "date" Or h = "data":                                       colDate   = c
            Case h = "time" Or h = "hora":                                        colTime   = c
            Case h = "c" Or h = "ctry" Or h Like "*countr*" Or h = "pais":       colCtry   = c
            Case h Like "*event*" Or h Like "*release*" Or h Like "*indicator*": colEvent  = c
            Case h Like "*period*":                                               colPeriod = c
            Case h Like "*surv*" Or h Like "*median*" Or h Like "*forecast*":    colSurv   = c
            Case h = "prior" Or h Like "*anterior*" Or h Like "*previous*":      colPrior  = c
            Case h = "actual" Or h Like "*realiz*":                               colActual = c
            Case h Like "*imp*" Or h = "rel" Or h = "relevance" Or h = "x":     colRel    = c
        End Select
    Next c

    If colDate = 0 Or colEvent = 0 Then
        MsgBox "Colunas 'Date' e 'Event' nao encontradas na planilha Bloomberg." & vbCrLf & _
               "Cabecalho encontrado na linha " & hRow & " da sheet '" & wsBBG.Name & "'.", _
               vbCritical, "ImportarDoWECO"
        Exit Sub
    End If

    ' ── 3. Garante que CALENDARIO existe ─────────────────────────────────
    Dim wsCal As Worksheet
    On Error Resume Next
    Set wsCal = ThisWorkbook.Sheets(CAL_SHEET)
    On Error GoTo 0
    If wsCal Is Nothing Then
        CriarSheetCalendario
        Set wsCal = ThisWorkbook.Sheets(CAL_SHEET)
    End If

    ' Limpa dados anteriores (linha 4 em diante)
    Dim lastCalRow As Long
    lastCalRow = wsCal.Cells(wsCal.Rows.Count, 1).End(xlUp).Row
    If lastCalRow >= 4 Then wsCal.Rows("4:" & lastCalRow).ClearContents

    ' ── 4. Le e escreve os dados ─────────────────────────────────────────
    Dim lastRow  As Long
    lastRow = wsBBG.Cells(wsBBG.Rows.Count, colDate).End(xlUp).Row

    Dim destRow  As Long: destRow = 4
    Dim imported As Long: imported = 0

    Dim rw As Long
    For rw = hRow + 1 To lastRow

        ' Data
        Dim dtTxt As String: dtTxt = Trim(wsBBG.Cells(rw, colDate).Text)
        If dtTxt = "" Then GoTo ProxLinha

        ' Combina data + hora
        Dim tmTxt As String: tmTxt = ""
        If colTime > 0 Then tmTxt = Trim(wsBBG.Cells(rw, colTime).Text)
        Dim dtFull As String
        If tmTxt <> "" And tmTxt <> "N/A" And tmTxt <> "--" And tmTxt <> "00:00" Then
            dtFull = dtTxt & " " & tmTxt
        Else
            dtFull = dtTxt
        End If

        ' Evento
        Dim evTxt As String: evTxt = Trim(wsBBG.Cells(rw, colEvent).Text)
        If evTxt = "" Then GoTo ProxLinha

        ' Adiciona periodo ao nome (ex: "IPCA MoM  [May]")
        If colPeriod > 0 Then
            Dim perTxt As String: perTxt = Trim(wsBBG.Cells(rw, colPeriod).Text)
            If perTxt <> "" And perTxt <> "N/A" Then
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

        ' Surpresa = Actual - Surv(M)
        Dim surpresa As String: surpresa = ""
        If actualTxt <> "" And survTxt <> "" Then
            If IsNumeric(actualTxt) And IsNumeric(survTxt) Then
                surpresa = CStr(CDbl(actualTxt) - CDbl(survTxt))
            End If
        End If

        ' Escreve na sheet CALENDARIO
        wsCal.Cells(destRow, 1).Value = dtFull
        wsCal.Cells(destRow, 2).Value = ctryTxt
        wsCal.Cells(destRow, 3).Value = evTxt
        wsCal.Cells(destRow, 4).Value = relTxt
        wsCal.Cells(destRow, 5).Value = priorTxt
        wsCal.Cells(destRow, 6).Value = survTxt
        wsCal.Cells(destRow, 7).Value = actualTxt
        wsCal.Cells(destRow, 8).Value = surpresa

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

    Dim eventsJson As String: eventsJson = ""
    Dim evCount    As Long:   evCount    = 0

    Dim rw As Long
    For rw = 4 To lastRow
        Dim dtTxt  As String: dtTxt  = Trim(ws.Cells(rw, 1).Text)
        Dim evName As String: evName = Trim(ws.Cells(rw, 3).Text)
        If dtTxt = "" Or evName = "" Then GoTo NextRow

        Dim ctry   As String: ctry   = Trim(ws.Cells(rw, 2).Text)
        Dim relOut As String: relOut = Trim(ws.Cells(rw, 4).Text)
        Dim prior  As String: prior  = LimparNum(ws.Cells(rw, 5).Text)
        Dim estim  As String: estim  = LimparNum(ws.Cells(rw, 6).Text)
        Dim actual As String: actual = LimparNum(ws.Cells(rw, 7).Text)
        Dim surpr  As String: surpr  = LimparNum(ws.Cells(rw, 8).Text)

        evName = Replace(evName, """", "'")
        ctry   = Replace(ctry,   """", "'")

        If eventsJson <> "" Then eventsJson = eventsJson & "," & vbCrLf
        eventsJson = eventsJson & "    {" & _
            """dt"":""" & dtTxt & """," & _
            """country"":""" & ctry & """," & _
            """event"":""" & evName & """," & _
            """relevance"":""" & relOut & """," & _
            """prior"":" & IIf(prior  <> "", """" & prior  & """", "null") & "," & _
            """estimate"":" & IIf(estim  <> "", """" & estim  & """", "null") & "," & _
            """actual"":" & IIf(actual <> "", """" & actual & """", "null") & "," & _
            """surprise"":" & IIf(surpr  <> "", """" & surpr  & """", "null") & "}"
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

    Dim fOut As Integer: fOut = FreeFile
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

' Encontra linha de cabecalho na planilha Bloomberg (procura "Date" ou "Event")
Private Function AcharCabecalho(ws As Worksheet) As Long
    Dim r As Long, c As Long, txt As String
    For r = 1 To 10
        For c = 1 To 15
            txt = LCase(Trim(ws.Cells(r, c).Text))
            If txt = "date" Or txt = "data" Or txt Like "*event*" Or txt Like "*release*" Then
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
        Case "***", "3", "H", "HIGH":          MapRel = "HIGH"
        Case "**",  "2", "M", "MED", "MEDIUM": MapRel = "MEDIUM"
        Case "*",   "1", "L", "LOW":           MapRel = "LOW"
        Case Else:                              MapRel = raw
    End Select
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
