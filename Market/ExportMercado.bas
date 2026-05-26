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

    ReDim diArr(0 To 50)
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
            Case "di"
                diArr(diCount) = "{""label"":""" & label & """," & _
                                  """du"":" & CLng(v2) & "," & _
                                  """taxa"":" & Format(v1, "0.000") & "}"
                diCount = diCount + 1
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

    ' Debug: confirma no rodapé do Excel
    Application.StatusBar = "market.json exportado em " & Format(Now, "hh:mm:ss")
    Exit Sub

ErrHandler:
    MsgBox "Erro ao exportar: " & Err.Description, vbCritical, "ExportMercado"
End Sub

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
