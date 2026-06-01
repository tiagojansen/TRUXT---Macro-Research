Attribute VB_Name = "CalendarioExport"
' ============================================================
' TRUXT Macro Research -- Calendario Economico Bloomberg
' Importar via VBA Editor > File > Import File
'
' Macros publicas:
'   CriarSheetCalendario  -- cria/recria as sheets (rodar 1x)
'   ExportCalendario      -- exporta calendar.json
' ============================================================

Option Explicit

Public Const CAL_SHEET    As String = "CALENDARIO"
Public Const CFG_SHEET    As String = "CONFIG"
Public Const CAL_PATH     As String = "S:\Macro\Site\data\calendar.json"

Private Function Navy() As Long:    Navy    = RGB(31, 56, 100):   End Function
Private Function NavyMid() As Long: NavyMid = RGB(31, 73, 125):   End Function
Private Function RowAlt() As Long:  RowAlt  = RGB(235, 243, 255): End Function

' -----------------------------------------------------------------------
' CriarSheetCalendario
' Cria as sheets CALENDARIO, CONFIG e INSTRUCOES no workbook atual.
' Rodar UMA vez apos importar o modulo.
' -----------------------------------------------------------------------
Public Sub CriarSheetCalendario()
    Dim wb As Workbook
    Set wb = ThisWorkbook

    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Sheets(CAL_SHEET).Delete
    wb.Sheets(CFG_SHEET).Delete
    wb.Sheets("INSTRUCOES").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    ' ── Sheet CONFIG ────────────────────────────────────────────────────
    Dim wsCfg As Worksheet
    Set wsCfg = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
    wsCfg.Name = CFG_SHEET
    wsCfg.Tab.Color = RGB(0, 112, 192)

    With wsCfg.Range("B1:D1")
        .Merge
        .Value = "CONFIGURACAO -- CALENDARIO ECONOMICO"
        .Font.Bold = True: .Font.Size = 12: .Font.Name = "Arial"
        .Font.Color = vbWhite
        .Interior.Color = Navy()
        .HorizontalAlignment = xlCenter
        .RowHeight = 26
    End With

    Dim cfgRows As Variant
    cfgRows = Array( _
        Array("Ticker",      "WECO Index",   False), _
        Array("Data Inicio", "=TEXT(DATE(YEAR(TODAY()),MONTH(TODAY()),1),""MM/DD/YYYY"")", True), _
        Array("Data Fim",    "=TEXT(EOMONTH(TODAY(),0),""MM/DD/YYYY"")",                  True), _
        Array("Paises",      "BZ,US,EC,CH",  True), _
        Array("Max Linhas",  200,             False) _
    )

    Dim i As Integer
    For i = 0 To 4
        Dim r As Integer: r = i + 2
        With wsCfg.Cells(r, 2)
            .Value = cfgRows(i)(0)
            .Font.Bold = True: .Font.Name = "Arial"
            .Interior.Color = RGB(217, 217, 217)
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
        End With
        With wsCfg.Cells(r, 3)
            If cfgRows(i)(2) Then
                If Left(CStr(cfgRows(i)(1)), 1) = "=" Then
                    .Formula = cfgRows(i)(1)
                Else
                    .Value = cfgRows(i)(1)
                End If
                .Interior.Color = RGB(255, 255, 153)
                .Font.Bold = True
            Else
                .Value = cfgRows(i)(1)
                .Interior.Color = RGB(242, 242, 242)
            End If
            .Font.Name = "Arial"
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
        End With
        With wsCfg.Cells(r, 4)
            If cfgRows(i)(2) Then
                .Value = "<-- editavel"
                .Font.Color = RGB(150, 150, 150)
                .Font.Italic = True
                .Font.Size = 9
            End If
        End With
    Next i

    wsCfg.Columns("A").ColumnWidth = 2
    wsCfg.Columns("B").ColumnWidth = 14
    wsCfg.Columns("C").ColumnWidth = 28
    wsCfg.Columns("D").ColumnWidth = 16

    ' ── Sheet CALENDARIO ────────────────────────────────────────────────
    Dim wsCal As Worksheet
    Set wsCal = wb.Sheets.Add(Before:=wsCfg)
    wsCal.Name = CAL_SHEET
    wsCal.Tab.Color = Navy()

    ' Linha 1: banner
    With wsCal.Range("A1:H1")
        .Merge
        .Value = "CALENDARIO ECONOMICO | Bloomberg Data"
        .Font.Bold = True: .Font.Size = 14: .Font.Name = "Arial"
        .Font.Color = vbWhite
        .Interior.Color = Navy()
        .HorizontalAlignment = xlCenter
        .RowHeight = 30
    End With

    ' Linha 2: info
    With wsCal.Range("A2:H2")
        .Merge
        .Value = "Paises: Brasil (BZ)  |  EUA (US)  |  Zona do Euro (EC)  |  China (CH)  |  Periodo: Mes Atual"
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
    wids = Array(20, 7, 42, 12, 13, 13, 13, 11)

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

    ' Linha 4: formulas BDS (concatenacao dinamica com sheet CONFIG)
    Dim ov As String
    ov = """cols=1;rows=200;startDate=""&CONFIG!C3&"";endDate=""&CONFIG!C4&"";country=""&CONFIG!C5"
    Dim flds As Variant
    flds = Array("ECO_RELEASE_DT", "ECO_COUNTRY", "ECO_FUTURE_RELEASE_NM", "ECO_RELEVANCE", _
                 "ECO_PRIOR_REVISION", "ECO_MEDIAN_FORECAST", "ECO_ACTUAL_RELEASE", "ECO_SURPRISE")

    For i = 0 To 7
        wsCal.Cells(4, i + 1).Formula = _
            "=BDS(""WECO Index"",""" & flds(i) & """," & ov & ")"
    Next i

    ' Freeze panes na linha 4
    wsCal.Activate
    wsCal.Range("A4").Select
    ActiveWindow.FreezePanes = True

    ' Formatacao condicional: linhas alternadas (A4:H203)
    Dim cfRng As Range
    Set cfRng = wsCal.Range("A4:H203")
    cfRng.FormatConditions.Delete
    Dim cf As Object
    Set cf = cfRng.FormatConditions.Add(Type:=xlExpression, Formula1:="=MOD(ROW()-3,2)=0")
    cf.Interior.Color = RowAlt()

    ' Bordas internas leves
    With wsCal.Range("A3:H203").Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlHairline
        .Color = RGB(200, 215, 240)
    End With
    With wsCal.Range("A3:H203").Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlHairline
        .Color = RGB(200, 215, 240)
    End With

    ' Coluna A: formato data/hora; colunas numericas: alinhar direita
    wsCal.Range("A4:A203").NumberFormat = "dd/mm/yyyy hh:mm"
    wsCal.Range("E4:H203").HorizontalAlignment = xlRight

    ' ── Sheet INSTRUCOES ────────────────────────────────────────────────
    Dim wsInst As Worksheet
    Set wsInst = wb.Sheets.Add(After:=wsCfg)
    wsInst.Name = "INSTRUCOES"
    wsInst.Tab.Color = RGB(0, 176, 80)

    With wsInst.Range("A1:D1")
        .Merge
        .Value = "INSTRUCOES DE USO"
        .Font.Bold = True: .Font.Size = 13: .Font.Name = "Arial"
        .Font.Color = vbWhite
        .Interior.Color = RGB(0, 112, 0)
        .HorizontalAlignment = xlCenter
        .RowHeight = 26
    End With

    ' Escreve linhas sem array grande (evita limite de 24 continuacoes do VBA)
    Dim iRow As Integer: iRow = 2
    Call AddInstTitle(wsInst, iRow, "CONFIGURACAO INICIAL")
    Call AddInstLine(wsInst, iRow, "1. Certifique-se de que o Bloomberg Add-In esta instalado e ativo (aba Bloomberg no menu).")
    Call AddInstLine(wsInst, iRow, "2. Abra o market_data.xlsm com o Bloomberg Terminal aberto e conectado.")
    Call AddInstLine(wsInst, iRow, "3. Execute a macro CriarSheetCalendario() uma unica vez para criar as sheets.")
    Call AddInstLine(wsInst, iRow, "4. Va para a sheet CALENDARIO e aguarde os dados carregarem (5-30 segundos).")
    Call AddInstLine(wsInst, iRow, "")
    Call AddInstTitle(wsInst, iRow, "COMO ATUALIZAR OS DADOS")
    Call AddInstLine(wsInst, iRow, "5. Para forcar atualizacao: Bloomberg (menu) > Refresh Worksheets (ou F9).")
    Call AddInstLine(wsInst, iRow, "6. A macro ExportCalendario() exporta os dados para o site automaticamente.")
    Call AddInstLine(wsInst, iRow, "7. O ExportarMercado ja chama ExportCalendario ao rodar -- nao precisa fazer manualmente.")
    Call AddInstLine(wsInst, iRow, "")
    Call AddInstTitle(wsInst, iRow, "CONFIGURACAO DE PAISES E DATAS (sheet CONFIG)")
    Call AddInstLine(wsInst, iRow, "8. Paises aceitos: BZ=Brasil, US=EUA, EC=Zona Euro, CH=China (separados por virgula).")
    Call AddInstLine(wsInst, iRow, "9. Data Inicio e Data Fim sao calculadas automaticamente para o mes atual.")
    Call AddInstLine(wsInst, iRow, "10. Para ver outro mes: altere as formulas em C3 e C4 na sheet CONFIG.")
    Call AddInstLine(wsInst, iRow, "")
    Call AddInstTitle(wsInst, iRow, "RELEVANCIA DOS EVENTOS")
    Call AddInstLine(wsInst, iRow, "HIGH   = evento de alta importancia (ex: IPCA, PIB, Payroll, decisao de juros).")
    Call AddInstLine(wsInst, iRow, "MEDIUM = relevancia moderada.")
    Call AddInstLine(wsInst, iRow, "LOW    = baixa relevancia.")
    Call AddInstLine(wsInst, iRow, "")
    Call AddInstTitle(wsInst, iRow, "CAMPOS DO CALENDARIO")
    Call AddInstLine(wsInst, iRow, "Data/Hora  : data e horario previsto de divulgacao.")
    Call AddInstLine(wsInst, iRow, "Pais       : codigo do pais (BZ, US, EC, CH).")
    Call AddInstLine(wsInst, iRow, "Evento     : nome do indicador economico.")
    Call AddInstLine(wsInst, iRow, "Anterior   : valor da divulgacao anterior (revisado).")
    Call AddInstLine(wsInst, iRow, "Estimativa : mediana das expectativas de mercado.")
    Call AddInstLine(wsInst, iRow, "Realizado  : valor efetivamente divulgado.")
    Call AddInstLine(wsInst, iRow, "Surpresa   : Realizado menos Estimativa.")

    wsInst.Columns("A").ColumnWidth = 90
    wsInst.Range("A1:A50").WrapText = False

    ' ── Finaliza ────────────────────────────────────────────────────────
    wsCal.Activate

    MsgBox "Sheets criadas com sucesso!" & vbCrLf & vbCrLf & _
           "Aguarde o Bloomberg carregar os dados na sheet CALENDARIO." & vbCrLf & _
           "Se nao carregar, va em Bloomberg > Refresh Worksheets.", _
           vbInformation, "CalendarioExport"
End Sub

' -----------------------------------------------------------------------
' ExportCalendario
' Le a sheet CALENDARIO e exporta S:\Macro\Site\data\calendar.json
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

    ' Mapeamento de relevancia (Bloomberg retorna 1/2/3 ou texto)
    Dim relMap As Object
    Set relMap = CreateObject("Scripting.Dictionary")
    relMap("1") = "HIGH" : relMap("2") = "MEDIUM" : relMap("3") = "LOW"
    relMap("HIGH") = "HIGH" : relMap("MEDIUM") = "MEDIUM"
    relMap("MED") = "MEDIUM" : relMap("LOW") = "LOW"

    Dim eventsJson As String
    eventsJson = ""
    Dim evCount   As Long
    evCount = 0

    Dim rw As Long
    For rw = 4 To lastRow
        Dim dtTxt  As String: dtTxt  = Trim(ws.Cells(rw, 1).Text)
        Dim ctry   As String: ctry   = Trim(ws.Cells(rw, 2).Text)
        Dim evName As String: evName = Trim(ws.Cells(rw, 3).Text)

        ' Pula linhas vazias ou com erro Bloomberg
        If dtTxt = "" Or dtTxt = "0" Or evName = "" Then GoTo NextRow
        If Left(dtTxt, 2) = "#N" Or Left(evName, 2) = "#N" Then GoTo NextRow

        ' Relevancia
        Dim relRaw As String: relRaw = UCase(Trim(ws.Cells(rw, 4).Text))
        Dim relOut As String
        If relMap.Exists(relRaw) Then relOut = relMap(relRaw) Else relOut = relRaw

        ' Valores numericos (null se vazio ou N.A.)
        Dim prior  As String: prior  = CleanVal(ws.Cells(rw, 5).Text)
        Dim estim  As String: estim  = CleanVal(ws.Cells(rw, 6).Text)
        Dim actual As String: actual = CleanVal(ws.Cells(rw, 7).Text)
        Dim surpr  As String: surpr  = CleanVal(ws.Cells(rw, 8).Text)

        ' Escapa aspas no nome do evento
        evName = Replace(evName, """", "'")
        ctry   = Replace(ctry,   """", "'")

        If eventsJson <> "" Then eventsJson = eventsJson & "," & vbCrLf
        eventsJson = eventsJson & "    {" & _
            """dt"":""" & dtTxt & """," & _
            """country"":""" & ctry & """," & _
            """event"":""" & evName & """," & _
            """relevance"":""" & relOut & """," & _
            """prior"":" & IIf(prior <> "", """" & prior & """", "null") & "," & _
            """estimate"":" & IIf(estim <> "", """" & estim & """", "null") & "," & _
            """actual"":" & IIf(actual <> "", """" & actual & """", "null") & "," & _
            """surprise"":" & IIf(surpr <> "", """" & surpr & """", "null") & "}"
        evCount = evCount + 1

NextRow:
    Next rw

    ' Monta JSON final
    Dim ts As String: ts = Format(Now, "yyyy-mm-ddThh:mm:ss")

    Dim finalJson As String
    finalJson = "{" & vbCrLf & _
                "  ""updated"":""" & ts & """," & vbCrLf & _
                "  ""count"":" & evCount & "," & vbCrLf & _
                "  ""events"":[" & vbCrLf & _
                eventsJson & vbCrLf & _
                "  ]" & vbCrLf & "}"

    Dim fOut As Integer
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
        MsgBox "Erro ao exportar calendario: " & Err.Description, vbCritical, "ExportCalendario"
    Else
        Application.StatusBar = "ERRO calendario: " & Err.Description
    End If
End Sub

' -----------------------------------------------------------------------
' Auxiliares para sheet INSTRUCOES
' -----------------------------------------------------------------------
Private Sub AddInstTitle(ws As Worksheet, ByRef rw As Integer, txt As String)
    With ws.Cells(rw, 1)
        .Value = txt
        .Font.Bold = True: .Font.Name = "Arial": .Font.Size = 10
        .Interior.Color = RGB(198, 224, 180)
        .Font.Color = RGB(0, 97, 0)
    End With
    rw = rw + 1
End Sub

Private Sub AddInstLine(ws As Worksheet, ByRef rw As Integer, txt As String)
    If txt = "" Then
        rw = rw + 1
        Exit Sub
    End If
    ws.Cells(rw, 1).Value = txt
    ws.Cells(rw, 1).Font.Name = "Arial"
    ws.Cells(rw, 1).Font.Size = 10
    rw = rw + 1
End Sub

' Auxiliar: retorna string limpa ou "" se vazio/N.A./erro
Private Function CleanVal(txt As String) As String
    Dim t As String: t = Trim(txt)
    If t = "" Or t = "N.A." Or t = "N/A" Or Left(t, 2) = "#N" Or t = "0" Then
        CleanVal = ""
    Else
        CleanVal = t
    End If
End Function
