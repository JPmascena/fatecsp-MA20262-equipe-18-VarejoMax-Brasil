Attribute VB_Name = "modResumoVendas"
Option Explicit

' ==============================================================================
' Módulo: modResumoVendas
' Projeto: VarejoMax Brasil (FATEC SP - Microinformática Aplicada)
' Objetivo: Gerar o Resumo Executivo de Vendas por Categoria e por Canal na aba
'           "Resumo_Categorias" a partir de "Base_Vendas" e "Cad_Produtos",
'           aplicando filtros opcionais de "Menu_Operacional" (C6, C7, C8).
' ==============================================================================

Public Sub GerarResumoVendasCategoria()
    Dim calcPrev As XlCalculation
    Dim scrPrev As Boolean
    Dim evtPrev As Boolean

    Dim wsMenu As Worksheet
    Dim wsBase As Worksheet
    Dim wsProd As Worksheet
    Dim wsOut As Worksheet

    Dim dtIni As Variant
    Dim dtFim As Variant
    Dim fCanal As String
    Dim hasIni As Boolean
    Dim hasFim As Boolean
    Dim hasCanal As Boolean

    Dim arrProd As Variant
    Dim arrBase As Variant
    Dim lastRowProd As Long
    Dim lastRowBase As Long
    Dim lastColBase As Long

    Dim dictProdCat As Object
    Dim dictProdCusto As Object
    Dim colIdx As Object

    ' Dicionários de agregação para Categorias e Canais
    Dim dictCatBruto As Object
    Dim dictCatLiq As Object
    Dim dictCatUnid As Object
    Dim dictCatCusto As Object
    Dim dictCatPed As Object

    Dim dictCanBruto As Object
    Dim dictCanLiq As Object
    Dim dictCanUnid As Object
    Dim dictCanCusto As Object
    Dim dictCanPed As Object

    Dim dictAllValidPed As Object

    Dim cIdPed As Long
    Dim cDataPed As Long
    Dim cStatusPed As Long
    Dim cCanal As Long
    Dim cIdProd As Long
    Dim cQtd As Long
    Dim cBruto As Long
    Dim cLiq As Long

    Dim r As Long
    Dim idProd As String
    Dim cat As String
    Dim canal As String
    Dim stPed As String
    Dim idPed As String
    Dim dtPed As Date
    Dim qtd As Long
    Dim vBruto As Double
    Dim vLiq As Double
    Dim cRef As Double

    Dim totBruto As Double
    Dim totLiq As Double
    Dim totUnid As Long
    Dim totCusto As Double
    Dim totPedVal As Long

    On Error GoTo ErrHandler

    scrPrev = Application.ScreenUpdating
    evtPrev = Application.EnableEvents
    calcPrev = Application.Calculation

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Set wsMenu = ThisWorkbook.Worksheets("Menu_Operacional")
    Set wsBase = ThisWorkbook.Worksheets("Base_Vendas")
    Set wsProd = ThisWorkbook.Worksheets("Cad_Produtos")
    Set wsOut = ThisWorkbook.Worksheets("Resumo_Categorias")

    ' 1. Ler filtros opcionais de Menu_Operacional (C6=Data Ini, C7=Data Fim, C8=Canal)
    hasIni = TryGetDateValue(wsMenu.Range("C6").Value, dtIni)
    hasFim = TryGetDateValue(wsMenu.Range("C7").Value, dtFim)
    fCanal = Trim$(CStr(wsMenu.Range("C8").Value))
    hasCanal = (Len(fCanal) > 0 And UCase$(fCanal) <> "TODOS")

    ' 2. Carregar Cad_Produtos em memória (ID_Produto -> Categoria e Custo_Ref)
    lastRowProd = wsProd.Cells(wsProd.Rows.Count, 1).End(xlUp).Row
    arrProd = wsProd.Range("A1:G" & lastRowProd).Value

    Set dictProdCat = CreateObject("Scripting.Dictionary")
    Set dictProdCusto = CreateObject("Scripting.Dictionary")
    dictProdCat.CompareMode = vbTextCompare
    dictProdCusto.CompareMode = vbTextCompare

    Dim pColId As Long, pColCat As Long, pColCusto As Long
    pColId = FindColInArray(arrProd, "ID_Produto")
    pColCat = FindColInArray(arrProd, "Categoria")
    pColCusto = FindColInArray(arrProd, "Custo_Ref")

    For r = 2 To UBound(arrProd, 1)
        idProd = Trim$(CStr(arrProd(r, pColId)))
        If Len(idProd) > 0 Then
            dictProdCat(idProd) = Trim$(CStr(arrProd(r, pColCat)))
            dictProdCusto(idProd) = ParseNum(arrProd(r, pColCusto))
        End If
    Next r

    ' 3. Carregar Base_Vendas inteira em array VBA
    lastRowBase = wsBase.Cells(wsBase.Rows.Count, 1).End(xlUp).Row
    lastColBase = wsBase.Cells(1, wsBase.Columns.Count).End(xlToLeft).Column
    arrBase = wsBase.Range(wsBase.Cells(1, 1), wsBase.Cells(lastRowBase, lastColBase)).Value

    Set colIdx = BuildHeaderMap(arrBase)
    cIdPed = colIdx("ID_Pedido")
    cDataPed = colIdx("Data_Pedido")
    cStatusPed = colIdx("Status_Pedido")
    cCanal = colIdx("Canal")
    cIdProd = colIdx("ID_Produto")
    cQtd = colIdx("Quantidade")
    cBruto = colIdx("Valor_Bruto")
    cLiq = colIdx("Valor_Liquido")

    ' 4. Inicializar dicionários de agregação
    Set dictCatBruto = CreateObject("Scripting.Dictionary")
    Set dictCatLiq = CreateObject("Scripting.Dictionary")
    Set dictCatUnid = CreateObject("Scripting.Dictionary")
    Set dictCatCusto = CreateObject("Scripting.Dictionary")
    Set dictCatPed = CreateObject("Scripting.Dictionary")

    Set dictCanBruto = CreateObject("Scripting.Dictionary")
    Set dictCanLiq = CreateObject("Scripting.Dictionary")
    Set dictCanUnid = CreateObject("Scripting.Dictionary")
    Set dictCanCusto = CreateObject("Scripting.Dictionary")
    Set dictCanPed = CreateObject("Scripting.Dictionary")

    Set dictAllValidPed = CreateObject("Scripting.Dictionary")

    dictCatBruto.CompareMode = vbTextCompare
    dictCatLiq.CompareMode = vbTextCompare
    dictCatUnid.CompareMode = vbTextCompare
    dictCatCusto.CompareMode = vbTextCompare
    dictCatPed.CompareMode = vbTextCompare

    dictCanBruto.CompareMode = vbTextCompare
    dictCanLiq.CompareMode = vbTextCompare
    dictCanUnid.CompareMode = vbTextCompare
    dictCanCusto.CompareMode = vbTextCompare
    dictCanPed.CompareMode = vbTextCompare

    dictAllValidPed.CompareMode = vbTextCompare

    ' 5. Processar as 13.613 linhas em memória (< 1s)
    For r = 2 To UBound(arrBase, 1)
        stPed = Trim$(CStr(arrBase(r, cStatusPed)))
        If stPed = "Cancelado" Then GoTo NextRow

        ' Filtro de Data_Pedido
        If hasIni Or hasFim Then
            Dim tmpDt As Variant
            If Not TryGetDateValue(arrBase(r, cDataPed), tmpDt) Then GoTo NextRow
            dtPed = CDate(tmpDt)
            If hasIni Then If dtPed < CDate(dtIni) Then GoTo NextRow
            If hasFim Then If dtPed > CDate(dtFim) Then GoTo NextRow
        End If

        canal = Trim$(CStr(arrBase(r, cCanal)))
        If hasCanal Then
            If StrComp(canal, fCanal, vbTextCompare) <> 0 Then GoTo NextRow
        End If

        idProd = Trim$(CStr(arrBase(r, cIdProd)))
        If dictProdCat.Exists(idProd) Then
            cat = dictProdCat(idProd)
            cRef = dictProdCusto(idProd)
        Else
            cat = "Sem Categoria"
            cRef = 0#
        End If

        idPed = Trim$(CStr(arrBase(r, cIdPed)))
        qtd = CLng(ParseNum(arrBase(r, cQtd)))
        vBruto = ParseNum(arrBase(r, cBruto))
        vLiq = ParseNum(arrBase(r, cLiq))

        ' Garantir chaves nos dicionários
        EnsureAggKey dictCatBruto, dictCatLiq, dictCatUnid, dictCatCusto, dictCatPed, cat
        EnsureAggKey dictCanBruto, dictCanLiq, dictCanUnid, dictCanCusto, dictCanPed, canal

        ' Regra oficial: Faturamento Bruto exclui apenas Cancelado
        dictCatBruto(cat) = dictCatBruto(cat) + vBruto
        dictCanBruto(canal) = dictCanBruto(canal) + vBruto
        totBruto = totBruto + vBruto

        ' Regra oficial: Líquido, Pedidos Válidos, Unidades e Custo Ref. excluem Cancelado e Devolvido
        If stPed <> "Devolvido" Then
            dictCatLiq(cat) = dictCatLiq(cat) + vLiq
            dictCatUnid(cat) = dictCatUnid(cat) + qtd
            dictCatCusto(cat) = dictCatCusto(cat) + (cRef * qtd)
            If Not dictCatPed(cat).Exists(idPed) Then dictCatPed(cat).Add idPed, True

            dictCanLiq(canal) = dictCanLiq(canal) + vLiq
            dictCanUnid(canal) = dictCanUnid(canal) + qtd
            dictCanCusto(canal) = dictCanCusto(canal) + (cRef * qtd)
            If Not dictCanPed(canal).Exists(idPed) Then dictCanPed(canal).Add idPed, True

            If Not dictAllValidPed.Exists(idPed) Then dictAllValidPed.Add idPed, True

            totLiq = totLiq + vLiq
            totUnid = totUnid + qtd
            totCusto = totCusto + (cRef * qtd)
        End If

NextRow:
    Next r

    totPedVal = dictAllValidPed.Count

    ' 6. Escrever e formatar saída na aba Resumo_Categorias
    wsOut.Cells.FormatConditions.Delete
    Dim maxClearRow As Long
    maxClearRow = wsOut.Cells(wsOut.Rows.Count, 1).End(xlUp).Row
    If maxClearRow < 40 Then maxClearRow = 40
    wsOut.Range("A5:I" & maxClearRow).Clear

    ' Cabeçalho executivo
    With wsOut.Range("A1:I1")
        .Merge
        .Value = "VAREJOMAX BRASIL — RESUMO EXECUTIVO DE VENDAS POR CATEGORIA E CANAL"
        .Font.Name = "Calibri"
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(27, 54, 93)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .RowHeight = 30
    End With

    With wsOut.Range("A2:I2")
        .Merge
        .Value = "Gerado em: " & Format$(Now, "dd/mm/yyyy hh:nn:ss") & " | Regra Oficial: Fat. Bruto exclui Cancelado; Fat. Líquido, Pedidos, Unidades e Custo Ref. excluem Cancelado e Devolvido"
        .Font.Name = "Calibri"
        .Font.Size = 9
        .Font.Italic = True
        .Font.Color = RGB(60, 60, 60)
        .Interior.Color = RGB(238, 242, 248)
        .RowHeight = 20
    End With

    Dim txtFiltro As String
    txtFiltro = "Filtros Aplicados: Período = " & _
        IIf(hasIni, Format$(CDate(dtIni), "dd/mm/yyyy"), "Início") & " até " & _
        IIf(hasFim, Format$(CDate(dtFim), "dd/mm/yyyy"), "Fim") & _
        " | Canal = " & IIf(hasCanal, fCanal, "Todos")

    With wsOut.Range("A3:I3")
        .Merge
        .Value = txtFiltro
        .Font.Name = "Calibri"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = RGB(27, 54, 93)
        .Interior.Color = RGB(247, 249, 252)
        .RowHeight = 20
    End With

    ' Quadro 1: Categorias
    Dim nextRow As Long
    nextRow = WriteSummaryTable(wsOut, 5, _
        "QUADRO 1 — DESEMPENHO POR CATEGORIA (ORDENADO POR FATURAMENTO LÍQUIDO DESC)", _
        "Categoria", dictCatPed, dictCatUnid, dictCatBruto, dictCatLiq, dictCatCusto, _
        totPedVal, totUnid, totBruto, totLiq, totCusto, True)

    ' Nota explicativa de contagem distinta
    With wsOut.Range("A" & nextRow & ":I" & nextRow)
        .Merge
        .Value = "* Nota: Pedidos Válidos por categoria representam pedidos distintos que contêm ao menos 1 item da categoria; TOTAL GERAL exibe os " & Format$(totPedVal, "#,##0") & " pedidos únicos válidos da base."
        .Font.Name = "Calibri"
        .Font.Size = 9
        .Font.Italic = True
        .Font.Color = RGB(90, 90, 90)
    End With

    ' Quadro 2: Canais de Venda
    nextRow = WriteSummaryTable(wsOut, nextRow + 2, _
        "QUADRO 2 — DESEMPENHO POR CANAL DE VENDA (ORDENADO POR FATURAMENTO LÍQUIDO DESC)", _
        "Canal de Venda", dictCanPed, dictCanUnid, dictCanBruto, dictCanLiq, dictCanCusto, _
        totPedVal, totUnid, totBruto, totLiq, totCusto, False)

    wsOut.Columns("A:I").AutoFit
    If wsOut.Columns("A").ColumnWidth < 24 Then wsOut.Columns("A").ColumnWidth = 24

CleanExit:
    Application.Calculation = calcPrev
    Application.EnableEvents = evtPrev
    Application.ScreenUpdating = scrPrev
    Exit Sub

ErrHandler:
    MsgBox "Erro em GerarResumoVendasCategoria (" & Err.Number & "): " & Err.Description, vbCritical, "VarejoMax VBA"
    Resume CleanExit
End Sub

Private Function WriteSummaryTable(ByVal ws As Worksheet, ByVal startRow As Long, _
    ByVal titleText As String, ByVal firstColTitle As String, _
    ByVal dictPed As Object, ByVal dictUnid As Object, ByVal dictBruto As Object, _
    ByVal dictLiq As Object, ByVal dictCusto As Object, _
    ByVal totPedVal As Long, ByVal totUnid As Long, ByVal totBruto As Double, _
    ByVal totLiq As Double, ByVal totCusto As Double, ByVal highlightLowMargin As Boolean) As Long

    Dim keys As Variant
    Dim n As Long
    Dim i As Long, j As Long
    Dim tmpKey As Variant
    Dim curRow As Long
    Dim dataStart As Long, dataEnd As Long

    ' Banner do quadro
    With ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, 9))
        .Merge
        .Value = titleText
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(43, 84, 126)
        .RowHeight = 22
    End With

    ' Cabeçalho das colunas
    curRow = startRow + 1
    Dim headers As Variant
    headers = Array(firstColTitle, "Pedidos Válidos", "Unidades Vendidas", _
                    "Faturamento Bruto (R$)", "Faturamento Líquido (R$)", _
                    "Ticket Médio (R$)", "Custo Ref. Total (R$)", _
                    "Margem Bruta (%)", "Participação (%)")
    For i = 0 To 8
        With ws.Cells(curRow, i + 1)
            .Value = headers(i)
            .Font.Name = "Calibri"
            .Font.Size = 10
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(31, 78, 120)
            .HorizontalAlignment = IIf(i = 0, xlLeft, xlRight)
            .VerticalAlignment = xlCenter
        End With
    Next i
    ws.Rows(curRow).RowHeight = 22

    n = dictLiq.Count
    dataStart = curRow + 1

    If n > 0 Then
        keys = dictLiq.keys
        ' Ordenar chaves por Faturamento Líquido DESC
        For i = LBound(keys) To UBound(keys) - 1
            For j = i + 1 To UBound(keys)
                If dictLiq(keys(j)) > dictLiq(keys(i)) Then
                    tmpKey = keys(i)
                    keys(i) = keys(j)
                    keys(j) = tmpKey
                End If
            Next j
        Next i

        For i = LBound(keys) To UBound(keys)
            curRow = curRow + 1
            Dim k As String
            Dim peds As Long, unids As Long
            Dim vB As Double, vL As Double, vC As Double
            Dim tm As Double, mb As Double, part As Double

            k = CStr(keys(i))
            peds = dictPed(k).Count
            unids = dictUnid(k)
            vB = dictBruto(k)
            vL = dictLiq(k)
            vC = dictCusto(k)

            tm = IIf(peds > 0, vL / peds, 0#)
            mb = IIf(vL > 0, (vL - vC) / vL, 0#)
            part = IIf(totLiq > 0, vL / totLiq, 0#)

            ws.Cells(curRow, 1).Value = k
            ws.Cells(curRow, 2).Value = peds
            ws.Cells(curRow, 3).Value = unids
            ws.Cells(curRow, 4).Value = Round(vB, 2)
            ws.Cells(curRow, 5).Value = Round(vL, 2)
            ws.Cells(curRow, 6).Value = Round(tm, 2)
            ws.Cells(curRow, 7).Value = Round(vC, 2)
            ws.Cells(curRow, 8).Value = mb
            ws.Cells(curRow, 9).Value = part

            ' Zebra striping
            If (i Mod 2) = 1 Then
                ws.Range(ws.Cells(curRow, 1), ws.Cells(curRow, 9)).Interior.Color = RGB(244, 247, 251)
            End If

            ' Destaque direto se Margem < 20%
            If highlightLowMargin And mb < 0.2 Then
                With ws.Cells(curRow, 8)
                    .Interior.Color = RGB(255, 217, 217)
                    .Font.Color = RGB(156, 0, 6)
                    .Font.Bold = True
                End With
            End If
        Next i
        dataEnd = curRow
    Else
        curRow = curRow + 1
        ws.Cells(curRow, 1).Value = "(Sem registros para os filtros informados)"
        dataEnd = curRow
    End If

    ' Linha TOTAL GERAL
    curRow = curRow + 1
    Dim totTM As Double, totMB As Double
    totTM = IIf(totPedVal > 0, totLiq / totPedVal, 0#)
    totMB = IIf(totLiq > 0, (totLiq - totCusto) / totLiq, 0#)

    ws.Cells(curRow, 1).Value = "TOTAL GERAL"
    ws.Cells(curRow, 2).Value = totPedVal
    ws.Cells(curRow, 3).Value = totUnid
    ws.Cells(curRow, 4).Value = Round(totBruto, 2)
    ws.Cells(curRow, 5).Value = Round(totLiq, 2)
    ws.Cells(curRow, 6).Value = Round(totTM, 2)
    ws.Cells(curRow, 7).Value = Round(totCusto, 2)
    ws.Cells(curRow, 8).Value = totMB
    ws.Cells(curRow, 9).Value = IIf(totLiq > 0, 1#, 0#)

    With ws.Range(ws.Cells(curRow, 1), ws.Cells(curRow, 9))
        .Font.Bold = True
        .Interior.Color = RGB(217, 225, 242)
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeBottom).LineStyle = xlDouble
    End With

    ' Formatação numérica e bordas
    With ws.Range(ws.Cells(dataStart, 1), ws.Cells(curRow, 9))
        .Font.Name = "Calibri"
        .Font.Size = 10
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).Color = RGB(217, 217, 217)
    End With

    ws.Range(ws.Cells(dataStart, 2), ws.Cells(curRow, 3)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(dataStart, 4), ws.Cells(curRow, 7)).NumberFormat = "R$ #,##0.00"
    ws.Range(ws.Cells(dataStart, 8), ws.Cells(curRow, 9)).NumberFormat = "0.00%"

    ' Formatação condicional adicional na coluna Margem Bruta (< 20%)
    If highlightLowMargin And n > 0 Then
        Dim rngMargin As Range
        Dim fc As FormatCondition
        Set rngMargin = ws.Range(ws.Cells(dataStart, 8), ws.Cells(dataEnd, 8))
        Set fc = rngMargin.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:=CStr(0.2))
        fc.Interior.Color = RGB(255, 217, 217)
        fc.Font.Color = RGB(156, 0, 6)
        fc.Font.Bold = True
    End If

    WriteSummaryTable = curRow + 1
End Function

Private Sub EnsureAggKey(ByVal dBruto As Object, ByVal dLiq As Object, ByVal dUnid As Object, _
    ByVal dCusto As Object, ByVal dPed As Object, ByVal keyName As String)
    If Not dBruto.Exists(keyName) Then
        dBruto.Add keyName, 0#
        dLiq.Add keyName, 0#
        dUnid.Add keyName, 0&
        dCusto.Add keyName, 0#
        Dim subDict As Object
        Set subDict = CreateObject("Scripting.Dictionary")
        subDict.CompareMode = vbTextCompare
        dPed.Add keyName, subDict
    End If
End Sub

Private Function BuildHeaderMap(ByVal arr As Variant) As Object
    Dim d As Object
    Dim c As Long
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    For c = LBound(arr, 2) To UBound(arr, 2)
        d(Trim$(CStr(arr(1, c)))) = c
    Next c
    Set BuildHeaderMap = d
End Function

Private Function FindColInArray(ByVal arr As Variant, ByVal headerName As String) As Long
    Dim c As Long
    For c = LBound(arr, 2) To UBound(arr, 2)
        If StrComp(Trim$(CStr(arr(1, c))), headerName, vbTextCompare) = 0 Then
            FindColInArray = c
            Exit Function
        End If
    Next c
    Err.Raise vbObjectError + 101, "modResumoVendas", "Coluna não encontrada: " & headerName
End Function

Private Function ParseNum(ByVal v As Variant) As Double
    If IsEmpty(v) Or IsNull(v) Then
        ParseNum = 0#
    ElseIf IsNumeric(v) Then
        ParseNum = CDbl(v)
    Else
        Dim s As String
        s = Trim$(CStr(v))
        If Len(s) = 0 Then
            ParseNum = 0#
        Else
            If InStr(s, ",") > 0 Then
                s = Replace(s, ".", "")
                s = Replace(s, ",", ".")
            End If
            ParseNum = Val(s)
        End If
    End If
End Function

Private Function TryGetDateValue(ByVal v As Variant, ByRef outDate As Variant) As Boolean
    On Error GoTo Fail
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    If IsDate(v) Then
        outDate = CDate(Int(CDbl(CDate(v))))
        TryGetDateValue = True
        Exit Function
    End If
    If IsNumeric(v) Then
        outDate = CDate(Int(CDbl(v)))
        TryGetDateValue = True
        Exit Function
    End If
    Dim s As String
    s = Trim$(CStr(v))
    If Len(s) = 0 Then Exit Function
    Dim parts() As String
    parts = Split(s, "/")
    If UBound(parts) = 2 Then
        outDate = DateSerial(CInt(parts(2)), CInt(parts(1)), CInt(parts(0)))
        TryGetDateValue = True
        Exit Function
    End If
Fail:
    TryGetDateValue = False
End Function
