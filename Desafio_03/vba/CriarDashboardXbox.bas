Attribute VB_Name = "CriarDashboardXbox"
Option Explicit

' =====================================================================
' Cria as abas "Cálculos" e "Dashboard" (Xbox Game Pass) em um clique.
' Macro principal: CriarDashboard
' Requisitos: base na aba "Bases" (A1:M..), colunas na ordem original.
' Os nomes das abas usam o caractere combinado U+0333, por isso são
' montados com ChrW (evita problemas de codificação do editor VBA).
' =====================================================================

Private Const VERDE_ESCURO As Long = &H417C10   ' #107C41 (BGR)
Private Const VERDE_CLARO  As Long = &H5EC522   ' #22C55E
Private Const CINZA_CLARO  As Long = &HE9E6E8   ' #E8E6E9
Private Const CINZA_TEXTO  As Long = &H37291F   ' #1F2937
Private Const CINZA_NO     As Long = &HA6A09A   ' #9AA0A6 (barra "No")
Private Const FUNDO_CARD   As Long = &HF5F6F4   ' #F4F6F5
Private Const ULT_LINHA    As String = "5000"

Private Function NomeBases() As String
    NomeBases = "B" & ChrW(819) & "ases"
End Function
Private Function NomeCalc() As String
    NomeCalc = "C" & ChrW(819) & ChrW(225) & "lculos"
End Function
Private Function NomeDash() As String
    NomeDash = "D" & ChrW(819) & "ashboard"
End Function

' Intervalo de uma coluna da base, ex.: 'Bases'!$I$2:$I$5000
Private Function Rg(col As String) As String
    Rg = "'" & NomeBases() & "'!$" & col & "$2:$" & col & "$" & ULT_LINHA
End Function

Public Sub CriarDashboard()
    Dim wsC As Worksheet, wsD As Worksheet
    On Error GoTo Falha
    Application.ScreenUpdating = False

    Set wsC = ObterAba(NomeCalc())
    Set wsD = ObterAba(NomeDash())
    LimparAba wsC
    LimparAba wsD

    MontarCalculos wsC
    MontarDashboard wsD, wsC

    wsD.Activate
    wsD.Range("D5").Select
    Application.ScreenUpdating = True
    MsgBox "Dashboard criado com sucesso!", vbInformation, "Xbox Game Pass"
    Exit Sub
Falha:
    Application.ScreenUpdating = True
    MsgBox "Erro ao criar o dashboard: " & Err.Description, vbCritical
End Sub

Private Function ObterAba(nome As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = nome Then Set ObterAba = ws: Exit Function
    Next ws
    Set ObterAba = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    ObterAba.Name = nome
End Function

Private Sub LimparAba(ws As Worksheet)
    Dim co As ChartObject
    For Each co In ws.ChartObjects: co.Delete: Next co
    ws.Cells.UnMerge
    ws.Cells.Clear
    ws.Cells.Validation.Delete
    ws.Cells.FormatConditions.Delete
    ws.Cells.Font.Name = "Arial"
    ws.Cells.Font.Size = 10
    ws.Cells.Font.Color = CINZA_TEXTO
    ws.Activate
    ActiveWindow.DisplayGridlines = False
End Sub

Private Sub Cabecalho(ws As Worksheet, endereco As String, texto As String)
    With ws.Range(endereco)
        .Interior.Color = VERDE_ESCURO
        .Font.Color = vbWhite
        .Font.Bold = True
    End With
    ws.Range(endereco).Cells(1, 1).Value = texto
End Sub

Private Sub SubCabecalho(ws As Worksheet, endereco As String)
    With ws.Range(endereco)
        .Interior.Color = CINZA_CLARO
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    ws.Range(endereco).Cells(1, 1).HorizontalAlignment = xlLeft
End Sub

' ---------------------------------------------------------------------
' ABA DE CÁLCULOS
' ---------------------------------------------------------------------
Private Sub MontarCalculos(ws As Worksheet)
    Dim i As Long, r As Long, tipos As Variant
    Dim brl As String: brl = """R$"" #,##0.00"

    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 44
    ws.Columns("C:G").ColumnWidth = 20

    ws.Range("B2").Value = "Cálculos " & ChrW(8212) & " tabelas de apoio ao Dashboard"
    ws.Range("B2").Font.Size = 16: ws.Range("B2").Font.Bold = True: ws.Range("B2").Font.Color = VERDE_ESCURO

    ' 1. Parâmetro do filtro
    Cabecalho ws, "B4:G4", "1. Parâmetro do filtro (vem da aba Dashboard)"
    ws.Range("B5").Value = "Filtro selecionado"
    ws.Range("C5").Formula = "='" & NomeDash() & "'!$D$5"
    ws.Range("B6").Value = "Critério usado nas fórmulas (""*"" = todos)"
    ws.Range("C6").Formula = "=IF(C5=""Todos"",""*"",C5)"
    ws.Range("C5:C6").HorizontalAlignment = xlCenter: ws.Range("C5:C6").Font.Bold = True

    ' 2. KPIs filtrados
    Cabecalho ws, "B8:G8", "2. KPIs " & ChrW(8212) & " respondem ao filtro"
    ws.Range("B9:D9").Value = Array("Métrica", "Valor", "Lógica")
    SubCabecalho ws, "B9:D9"
    ws.Range("B10").Value = "Total R$ " & ChrW(8212) & " EA Play Season Pass"
    ws.Range("C10").Formula = "=SUMIFS(" & Rg("I") & "," & Rg("H") & ",""Yes""," & Rg("G") & ",$C$6)"
    ws.Range("B11").Value = "Total R$ " & ChrW(8212) & " Minecraft Season Pass"
    ws.Range("C11").Formula = "=SUMIFS(" & Rg("K") & "," & Rg("J") & ",""Yes""," & Rg("G") & ",$C$6)"
    ws.Range("B12").Value = "Game Pass " & ChrW(8212) & " Auto Renewal YES (qtd)"
    ws.Range("C12").Formula = "=COUNTIFS(" & Rg("E") & ",""Yes""," & Rg("G") & ",$C$6)"
    ws.Range("B13").Value = "Game Pass " & ChrW(8212) & " Auto Renewal NO (qtd)"
    ws.Range("C13").Formula = "=COUNTIFS(" & Rg("E") & ",""No""," & Rg("G") & ",$C$6)"
    ws.Range("B14").Value = "Total de assinantes no filtro"
    ws.Range("C14").Formula = "=COUNTIFS(" & Rg("G") & ",$C$6)"
    ws.Range("B15").Value = "% Auto Renewal YES"
    ws.Range("C15").Formula = "=IFERROR(C12/C14,0)"
    ws.Range("B16").Value = "% Auto Renewal NO"
    ws.Range("C16").Formula = "=IFERROR(C13/C14,0)"
    ws.Range("C10:C11").NumberFormat = brl
    ws.Range("C12:C14").NumberFormat = "#,##0"
    ws.Range("C15:C16").NumberFormat = "0.0%"
    ws.Range("C10:C16").HorizontalAlignment = xlCenter: ws.Range("C10:C16").Font.Bold = True
    ws.Range("D10").Value = "SOMASES(Preço EA; EA=""Yes""; Tipo=critério)"
    ws.Range("D11").Value = "SOMASES(Preço Minecraft; Minecraft=""Yes""; Tipo=critério)"
    ws.Range("D12").Value = "CONT.SES(Auto Renewal=""Yes""; Tipo=critério)"
    ws.Range("D13").Value = "CONT.SES(Auto Renewal=""No""; Tipo=critério)"
    ws.Range("D14").Value = "CONT.SES(Tipo=critério)"
    ws.Range("D10:D16").Font.Italic = True: ws.Range("D10:D16").Font.Color = RGB(107, 114, 128)

    ' 3. Resumo por tipo (sem filtro)
    Cabecalho ws, "B18:G18", "3. Resumo por tipo de faturamento " & ChrW(8212) & " sem filtro"
    ws.Range("B19:G19").Value = Array("Subscription Type", "EA Play (R$)", "Minecraft (R$)", _
        "Auto Renewal YES", "Auto Renewal NO", "Total assinantes")
    SubCabecalho ws, "B19:G19"
    tipos = Array("Annual", "Monthly", "Quarterly")
    For i = 0 To 2
        r = 20 + i
        ws.Cells(r, 2).Value = tipos(i)
        ws.Cells(r, 3).Formula = "=SUMIFS(" & Rg("I") & "," & Rg("H") & ",""Yes""," & Rg("G") & ",$B" & r & ")"
        ws.Cells(r, 4).Formula = "=SUMIFS(" & Rg("K") & "," & Rg("J") & ",""Yes""," & Rg("G") & ",$B" & r & ")"
        ws.Cells(r, 5).Formula = "=COUNTIFS(" & Rg("E") & ",""Yes""," & Rg("G") & ",$B" & r & ")"
        ws.Cells(r, 6).Formula = "=COUNTIFS(" & Rg("E") & ",""No""," & Rg("G") & ",$B" & r & ")"
        ws.Cells(r, 7).Formula = "=COUNTIFS(" & Rg("G") & ",$B" & r & ")"
    Next i
    ws.Range("B23").Value = "Total"
    ws.Range("C23:G23").Formula = "=SUM(C20:C22)"
    ws.Range("B23:G23").Font.Bold = True
    ws.Range("C20:D23").NumberFormat = brl
    ws.Range("E20:G23").NumberFormat = "#,##0"
    ws.Range("C20:G23").HorizontalAlignment = xlCenter

    ' 4. Dados do gráfico
    Cabecalho ws, "B25:G25", "4. Dados do gráfico (respondem ao filtro)"
    ws.Range("B26:C26").Value = Array("Categoria", "Qtd")
    SubCabecalho ws, "B26:C26"
    ws.Range("B27").Value = "Yes (renovação ativa)": ws.Range("C27").Formula = "=C12"
    ws.Range("B28").Value = "No (sem renovação)":    ws.Range("C28").Formula = "=C13"
    ws.Range("C27:C28").HorizontalAlignment = xlCenter

    ' 5. Checagens
    Cabecalho ws, "B30:G30", "5. Checagens de consistência"
    ws.Range("B31").Value = "Soma EA por tipo = soma da base"
    ws.Range("C31").Formula = "=IF(ROUND(C23-SUM(" & Rg("I") & "),2)=0,""OK"",""VERIFICAR"")"
    ws.Range("B32").Value = "Soma Minecraft por tipo = soma da base"
    ws.Range("C32").Formula = "=IF(ROUND(D23-SUM(" & Rg("K") & "),2)=0,""OK"",""VERIFICAR"")"
    ws.Range("B33").Value = "Total de assinantes = linhas da base"
    ws.Range("C33").Formula = "=IF(G23=COUNTA(" & Rg("A") & "),""OK"",""VERIFICAR"")"
    ws.Range("C31:C33").HorizontalAlignment = xlCenter: ws.Range("C31:C33").Font.Bold = True

    ws.Range("B35").Value = "Obs.: as fórmulas leem as linhas 2 a " & ULT_LINHA & " da base; novas linhas entram automaticamente."
    ws.Range("B35").Font.Italic = True: ws.Range("B35").Font.Size = 9
End Sub

' ---------------------------------------------------------------------
' ABA DASHBOARD
' ---------------------------------------------------------------------
Private Sub MontarDashboard(ws As Worksheet, wsC As Worksheet)
    Dim c As String: c = "='" & NomeCalc() & "'!"
    Dim brl As String: brl = """R$"" #,##0.00"
    Dim i As Long, r As Long

    ws.Tab.Color = VERDE_ESCURO
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("N").ColumnWidth = 2
    ws.Columns("B:M").ColumnWidth = 12.5
    ws.Rows("1:40").RowHeight = 20
    ws.Rows(1).RowHeight = 8

    ' Cabeçalho
    ws.Range("B2:M3").Interior.Color = VERDE_ESCURO
    ws.Rows(2).RowHeight = 26
    ws.Range("B2:I3").Merge
    With ws.Range("B2")
        .Value = "Xbox Game Pass  |  Dashboard de Assinaturas"
        .Font.Size = 20: .Font.Bold = True: .Font.Color = vbWhite
        .VerticalAlignment = xlCenter: .IndentLevel = 1
    End With
    ws.Range("J2:M3").Merge
    With ws.Range("J2")
        .Value = "EA Play & Minecraft Season Pass"
        .Font.Color = RGB(209, 250, 229): .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter: .IndentLevel = 1
    End With
    ws.Range("B4:M4").Interior.Color = VERDE_CLARO
    ws.Rows(4).RowHeight = 4

    ' Filtro (lista suspensa)
    ws.Rows(5).RowHeight = 26
    ws.Range("B5:C5").Merge
    With ws.Range("B5")
        .Value = "Tipo de faturamento " & ChrW(9656)
        .Font.Bold = True: .Font.Color = VERDE_ESCURO: .HorizontalAlignment = xlRight: .VerticalAlignment = xlCenter
    End With
    ws.Range("D5:E5").Merge
    With ws.Range("D5")
        .Value = "Todos"
        .Font.Bold = True: .Font.Size = 12: .Font.Color = VERDE_ESCURO
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .Interior.Color = vbWhite
        With .Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Todos,Annual,Monthly,Quarterly"
            .InCellDropdown = True
            .ErrorTitle = "Filtro inválido"
            .ErrorMessage = "Escolha: Todos, Annual, Monthly ou Quarterly."
        End With
    End With
    With ws.Range("D5:E5").Borders
        .LineStyle = xlContinuous: .Weight = xlMedium: .Color = VERDE_ESCURO
    End With
    ws.Range("F5:H5").Merge
    ws.Range("F5").Value = ChrW(9666) & " clique na seta para filtrar"
    ws.Range("F5").Font.Italic = True: ws.Range("F5").Font.Size = 9: ws.Range("F5").VerticalAlignment = xlCenter
    ws.Range("J5:L5").Merge
    ws.Range("J5").Value = "Assinantes no filtro:"
    ws.Range("J5").Font.Bold = True: ws.Range("J5").HorizontalAlignment = xlRight: ws.Range("J5").VerticalAlignment = xlCenter
    ws.Range("M5").Formula = c & "$C$14"
    ws.Range("M5").Font.Bold = True: ws.Range("M5").Font.Size = 12: ws.Range("M5").Font.Color = VERDE_ESCURO
    ws.Range("M5").NumberFormat = "#,##0": ws.Range("M5").HorizontalAlignment = xlLeft: ws.Range("M5").IndentLevel = 1

    ' Cartões de KPI
    ws.Rows(7).RowHeight = 22: ws.Rows(8).RowHeight = 26: ws.Rows(9).RowHeight = 26
    Dim ini As Variant, fim As Variant, rot As Variant, fonte As Variant, fmt As Variant, cor As Variant
    ini = Array("B", "E", "H", "K")
    fim = Array("D", "G", "J", "M")
    rot = Array("TOTAL R$ " & ChrW(8212) & " EA PLAY SEASON PASS", "TOTAL R$ " & ChrW(8212) & " MINECRAFT SEASON PASS", _
                "GAME PASS " & ChrW(8212) & " AUTO RENEWAL: YES", "GAME PASS " & ChrW(8212) & " AUTO RENEWAL: NO")
    fonte = Array("$C$10", "$C$11", "$C$12", "$C$13")
    fmt = Array(brl, brl, "#,##0", "#,##0")
    cor = Array(VERDE_ESCURO, VERDE_ESCURO, VERDE_ESCURO, RGB(95, 99, 104))
    For i = 0 To 3
        ws.Range(ini(i) & "7:" & fim(i) & "10").Interior.Color = FUNDO_CARD
        ws.Range(ini(i) & "7:" & fim(i) & "7").Interior.Color = cor(i)
        ws.Range(ini(i) & "7:" & fim(i) & "7").Merge
        With ws.Range(ini(i) & "7")
            .Value = rot(i): .Font.Bold = True: .Font.Size = 9: .Font.Color = vbWhite
            .VerticalAlignment = xlCenter: .IndentLevel = 1
        End With
        Dim ultCol As String
        ultCol = IIf(i < 2, Chr(Asc(ini(i)) + 1), fim(i))   ' cartões 1 e 2 deixam a 3ª coluna para o ícone
        ws.Range(ini(i) & "8:" & ultCol & "9").Merge
        With ws.Range(ini(i) & "8")
            .Formula = c & fonte(i)
            .NumberFormat = fmt(i)
            .Font.Bold = True: .Font.Size = 22: .Font.Color = cor(i)
            .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter: .IndentLevel = 1
        End With
        ws.Range(ini(i) & "10:" & fim(i) & "10").Merge
        With ws.Range(ini(i) & "10")
            Select Case i
                Case 0, 1: .Value = "Valor acumulado no filtro"
                Case 2:    .Formula = "=TEXT(" & Mid(c, 2) & "$C$15,""0.0%"")&"" dos assinantes no filtro"""
                Case 3:    .Formula = "=TEXT(" & Mid(c, 2) & "$C$16,""0.0%"")&"" dos assinantes no filtro"""
            End Select
            .Font.Size = 9: .Font.Color = RGB(107, 114, 128): .VerticalAlignment = xlCenter: .IndentLevel = 1
        End With
        ' separação visual entre cartões
        With ws.Range(ini(i) & "7:" & ini(i) & "10").Borders(xlEdgeLeft)
            .LineStyle = xlContinuous: .Weight = xlThick: .Color = vbWhite
        End With
        With ws.Range(fim(i) & "7:" & fim(i) & "10").Borders(xlEdgeRight)
            .LineStyle = xlContinuous: .Weight = xlThick: .Color = vbWhite
        End With
    Next i

    ' Seções
    Secao ws, "B12:G12", "Distribuição do Xbox Game Pass " & ChrW(8212) & " Auto Renewal"
    Secao ws, "H12:M12", "Detalhe por status"

    ' Tabela de status (filtrada)
    ws.Range("H13:I13").Merge: ws.Range("J13:K13").Merge: ws.Range("L13:M13").Merge
    ws.Range("H13").Value = "Status": ws.Range("J13").Value = "Qtd": ws.Range("L13").Value = "% do total"
    ws.Range("H13:M13").Interior.Color = CINZA_CLARO: ws.Range("H13:M13").Font.Bold = True
    ws.Range("J13:M13").HorizontalAlignment = xlCenter: ws.Range("H13").IndentLevel = 1
    Dim lbl As Variant, q As Variant, p As Variant
    lbl = Array("Yes " & ChrW(8212) & " renovação ativa", "No " & ChrW(8212) & " sem renovação", "Total")
    q = Array("$C$12", "$C$13", "$C$14")
    p = Array("$C$15", "$C$16", "")
    For i = 0 To 2
        r = 14 + i
        ws.Rows(r).RowHeight = 24
        ws.Range("H" & r & ":I" & r).Merge: ws.Range("J" & r & ":K" & r).Merge: ws.Range("L" & r & ":M" & r).Merge
        ws.Range("H" & r).Value = lbl(i): ws.Range("H" & r).IndentLevel = 1
        ws.Range("J" & r).Formula = c & q(i): ws.Range("J" & r).NumberFormat = "#,##0"
        If i < 2 Then ws.Range("L" & r).Formula = c & p(i) Else ws.Range("L" & r).Formula = "=IF(" & Mid(c, 2) & "$C$14=0,0,1)"
        ws.Range("L" & r).NumberFormat = "0.0%"
        ws.Range("J" & r & ":M" & r).HorizontalAlignment = xlCenter
        ws.Range("H" & r & ":M" & r).VerticalAlignment = xlCenter
        ws.Range("H" & r & ":M" & r).Borders(xlEdgeBottom).LineStyle = xlContinuous
        ws.Range("H" & r & ":M" & r).Borders(xlEdgeBottom).Color = RGB(208, 208, 208)
        If i = 2 Then ws.Range("H" & r & ":M" & r).Font.Bold = True
    Next i
    ws.Range("H14:M14").Font.Color = VERDE_ESCURO
    ws.Range("H14:M14").Font.Bold = True

    ' Notas de uso
    ws.Range("H19").Value = "Como usar": ws.Range("H19").Font.Bold = True: ws.Range("H19").Font.Color = VERDE_ESCURO
    ws.Range("H20").Value = ChrW(8226) & " O filtro atualiza os cartões, o gráfico e a tabela de status."
    ws.Range("H21").Value = ChrW(8226) & " Escolha ""Todos"" para ver o consolidado da base."
    ws.Range("H22").Value = ChrW(8226) & " A tabela inferior é fixa e destaca a linha do tipo escolhido."
    ws.Range("H20:H22").Font.Size = 9: ws.Range("H20:H22").Font.Color = RGB(107, 114, 128)

    ' Gráfico de colunas Yes x No
    Dim co As ChartObject
    Set co = ws.ChartObjects.Add(Left:=ws.Range("B13").Left, Top:=ws.Range("B13").Top + 2, _
                                 Width:=ws.Range("B13:G13").Width, Height:=ws.Range("B13:B25").Height)
    co.Name = "grfAutoRenewal"
    With co.Chart
        .ChartType = xlColumnClustered
        Do While .SeriesCollection.Count > 0: .SeriesCollection(1).Delete: Loop
        With .SeriesCollection.NewSeries
            .Name = "Auto Renewal"
            .Values = wsC.Range("C27:C28")
            .XValues = wsC.Range("B27:B28")
            .Format.Fill.ForeColor.RGB = VERDE_CLARO
            .Points(2).Format.Fill.ForeColor.RGB = CINZA_NO
            .HasDataLabels = True
            .DataLabels.Font.Size = 14: .DataLabels.Font.Bold = True
            .DataLabels.Position = xlLabelPositionOutsideEnd
        End With
        .HasLegend = False
        .HasTitle = False
        .ChartGroups(1).GapWidth = 60
        .Axes(xlValue).MinimumScale = 0
        .Axes(xlValue).HasMajorGridlines = False
        .Axes(xlValue).Delete
        .ChartArea.Format.Line.Visible = msoFalse
    End With

    ' Tabela por tipo (sem filtro)
    Secao ws, "B26:M26", "Totais por tipo de faturamento (sem filtro)"
    Dim h As Variant, a As Variant, b As Variant, cols As Variant
    h = Array("Subscription Type", "EA Play (R$)", "Minecraft (R$)", "Auto Renewal YES", "Auto Renewal NO", "Total assinantes")
    a = Array("B", "D", "F", "H", "J", "L")
    b = Array("C", "E", "G", "I", "K", "M")
    cols = Array("B", "C", "D", "E", "F", "G")
    For i = 0 To 5
        ws.Range(a(i) & "27:" & b(i) & "27").Merge
        ws.Range(a(i) & "27").Value = h(i)
        ws.Range(a(i) & "27:" & b(i) & "27").Interior.Color = VERDE_ESCURO
        ws.Range(a(i) & "27").Font.Bold = True: ws.Range(a(i) & "27").Font.Color = vbWhite
        ws.Range(a(i) & "27").HorizontalAlignment = IIf(i = 0, xlLeft, xlCenter)
        For r = 28 To 31
            ws.Range(a(i) & r & ":" & b(i) & r).Merge
            ws.Range(a(i) & r).Formula = c & cols(i) & (r - 8)
            ws.Range(a(i) & r).NumberFormat = IIf(i = 1 Or i = 2, brl, "#,##0")
            ws.Range(a(i) & r).HorizontalAlignment = IIf(i = 0, xlLeft, xlCenter)
            ws.Range(a(i) & r & ":" & b(i) & r).Borders(xlEdgeBottom).LineStyle = xlContinuous
            ws.Range(a(i) & r & ":" & b(i) & r).Borders(xlEdgeBottom).Color = RGB(208, 208, 208)
        Next r
    Next i
    ws.Range("B31:M31").Font.Bold = True
    ws.Range("B27").IndentLevel = 1: ws.Range("B28:B31").IndentLevel = 1

    ' Destaque da linha do tipo selecionado
    With ws.Range("B28:M30").FormatConditions.Add(Type:=xlExpression, Formula1:="=$B28=$D$5")
        .Interior.Color = RGB(209, 250, 229)
        .Font.Bold = True
        .Font.Color = VERDE_ESCURO
    End With

    ws.Range("B33:M33").Merge
    ws.Range("B33").Formula = "=""Fonte: aba base de dados (""&TEXT(COUNTA(" & Rg("A") & "),""#,##0"")&"" assinantes). Valores em R$ " & ChrW(183) & " atualiza automaticamente."""
    ws.Range("B33").Font.Italic = True: ws.Range("B33").Font.Size = 9

    ' Impressão
    With ws.PageSetup
        .Orientation = xlLandscape
        .PrintArea = "$A$1:$N$34"
        .Zoom = False: .FitToPagesWide = 1: .FitToPagesTall = 1
    End With
End Sub

Private Sub Secao(ws As Worksheet, endereco As String, texto As String)
    With ws.Range(endereco)
        .Merge
        .Font.Bold = True: .Font.Size = 12: .Font.Color = VERDE_ESCURO
        .VerticalAlignment = xlCenter
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Borders(xlEdgeBottom).Color = VERDE_CLARO
    End With
    ws.Range(endereco).Cells(1, 1).Value = texto
End Sub
