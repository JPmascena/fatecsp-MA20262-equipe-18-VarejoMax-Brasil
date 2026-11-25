Attribute VB_Name = "modAtrasosLogistica"
Option Explicit

' ==============================================================================
' Módulo: modAtrasosLogistica
' Projeto: VarejoMax Brasil (FATEC SP - Microinformática Aplicada)
' Objetivo: Gerar o Relatório de Desempenho por Transportadora e Pedidos
'           Críticos Atrasados na aba "Relatorio_Atrasos" a partir de
'           "Base_Vendas" e "Cad_Cidades", com filtros em "Menu_Operacional"
'           (C9 = Transportadora, C10 = UF ou Região, além de C6/C7 se preenchidos).
' ==============================================================================

Public Sub GerarRelatorioAtrasosTransportadora()
    Dim calcPrev As XlCalculation
    Dim scrPrev As Boolean
    Dim evtPrev As Boolean

    Dim wsMenu As Worksheet
    Dim wsBase As Worksheet
    Dim wsCid As Worksheet
    Dim wsOut As Worksheet

    Dim dtIni As Variant
    Dim dtFim As Variant
    Dim hasIni As Boolean
    Dim hasFim As Boolean
    Dim fTransp As String
    Dim fUfReg As String
    Dim hasTransp As Boolean
    Dim hasUfReg As Boolean

    Dim arrCid As Variant
    Dim arrBase As Variant
    Dim lastRowCid As Long
    Dim lastRowBase As Long
    Dim lastColBase As Long

    Dim dictCidUF As Object
    Dim dictCidReg As Object
    Dim colIdx As Object
    Dim dictOrders As Object

    Dim cIdPed As Long, cDataPed As Long, cStatusPed As Long
    Dim cStatusAtr As Long, cCanal As Long, cCidade As Long
    Dim cTransp As Long, cPrev As Long, cReal As Long
    Dim cAtrDias As Long, cAval As Long

    Dim strNaoInf As String
    Dim r As Long
    Dim stPed As String, stAtr As String, idPed As String
    Dim trItem As String, ordInfo As Variant
    Dim avalVal As Double, avalInt As Long

    On Error GoTo ErrHandler

    scrPrev = Application.ScreenUpdating
    evtPrev = Application.EnableEvents
    calcPrev = Application.Calculation

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    strNaoInf = "N" & ChrW(227) & "o informado"

    Set wsMenu = ThisWorkbook.Worksheets("Menu_Operacional")
    Set wsBase = ThisWorkbook.Worksheets("Base_Vendas")
    Set wsCid = ThisWorkbook.Worksheets("Cad_Cidades")
    Set wsOut = ThisWorkbook.Worksheets("Relatorio_Atrasos")

    ' 1. Ler filtros de Menu_Operacional
    hasIni = TryGetDateValue(wsMenu.Range("C6").Value, dtIni)
    hasFim = TryGetDateValue(wsMenu.Range("C7").Value, dtFim)
    fTransp = Trim$(CStr(wsMenu.Range("C9").Value))
    fUfReg = Trim$(CStr(wsMenu.Range("C10").Value))
    hasTransp = (Len(fTransp) > 0 And UCase$(fTransp) <> "TODAS")
    hasUfReg = (Len(fUfReg) > 0 And UCase$(fUfReg) <> "TODAS" And UCase$(fUfReg) <> "TODOS")

    ' 2. Carregar Cad_Cidades (Cidade -> UF, Regiao)
    lastRowCid = wsCid.Cells(wsCid.Rows.Count, 1).End(xlUp).Row
    arrCid = wsCid.Range("A1:F" & lastRowCid).Value

    Set dictCidUF = CreateObject("Scripting.Dictionary")
    Set dictCidReg = CreateObject("Scripting.Dictionary")
    dictCidUF.CompareMode = vbTextCompare
    dictCidReg.CompareMode = vbTextCompare

    Dim cidColNome As Long, cidColUF As Long, cidColReg As Long
    cidColNome = FindColInArray(arrCid, "Cidade")
    cidColUF = FindColInArray(arrCid, "UF")
    cidColReg = FindColInArray(arrCid, "Regiao")

    For r = 2 To UBound(arrCid, 1)
        Dim cidKey As String
        cidKey = Trim$(CStr(arrCid(r, cidColNome)))
        If Len(cidKey) > 0 Then
            dictCidUF(cidKey) = Trim$(CStr(arrCid(r, cidColUF)))
            dictCidReg(cidKey) = Trim$(CStr(arrCid(r, cidColReg)))
        End If
    Next r

    ' 3. Carregar Base_Vendas inteira em array
    lastRowBase = wsBase.Cells(wsBase.Rows.Count, 1).End(xlUp).Row
    lastColBase = wsBase.Cells(1, wsBase.Columns.Count).End(xlToLeft).Column
    arrBase = wsBase.Range(wsBase.Cells(1, 1), wsBase.Cells(lastRowBase, lastColBase)).Value

    Set colIdx = BuildHeaderMap(arrBase)
    cIdPed = colIdx("ID_Pedido")
    cDataPed = colIdx("Data_Pedido")
    cStatusPed = colIdx("Status_Pedido")
    cStatusAtr = colIdx("Status_Atraso")
    cCanal = colIdx("Canal")
    cCidade = colIdx("Cidade")
    cTransp = colIdx("Transportadora")
    cPrev = colIdx("Prazo_Previsto_Dias")
    cReal = colIdx("Prazo_Real_Dias")
    cAtrDias = colIdx("Atraso_Dias")
    cAval = colIdx("Avaliacao_Cliente")

    ' 4. Consolidar por ID_Pedido (evita duplicidade nos 358 pedidos com variação entre itens)
    ' Escopo: Status_Pedido = "Entregue" (exclui Cancelado, Devolvido, Em trânsito e Em separação)
    Set dictOrders = CreateObject("Scripting.Dictionary")
    dictOrders.CompareMode = vbTextCompare

    For r = 2 To UBound(arrBase, 1)
        stPed = Trim$(CStr(arrBase(r, cStatusPed)))
        If stPed <> "Entregue" Then GoTo NextBaseRow

        If hasIni Or hasFim Then
            Dim tmpDt As Variant
            If Not TryGetDateValue(arrBase(r, cDataPed), tmpDt) Then GoTo NextBaseRow
            If hasIni Then If CDate(tmpDt) < CDate(dtIni) Then GoTo NextBaseRow
            If hasFim Then If CDate(tmpDt) > CDate(dtFim) Then GoTo NextBaseRow
        End If

        idPed = Trim$(CStr(arrBase(r, cIdPed)))
        stAtr = Trim$(CStr(arrBase(r, cStatusAtr)))
        trItem = Trim$(CStr(arrBase(r, cTransp)))
        If Len(trItem) = 0 Then trItem = strNaoInf

        If Not dictOrders.Exists(idPed) Then
            ReDim ordInfo(0 To 9) As Variant
            ordInfo(0) = arrBase(r, cDataPed)
            ordInfo(1) = trItem
            ordInfo(2) = Trim$(CStr(arrBase(r, cCidade)))
            ordInfo(3) = Trim$(CStr(arrBase(r, cCanal)))
            ordInfo(4) = IIf(IsEmpty(arrBase(r, cPrev)) Or Trim$(CStr(arrBase(r, cPrev))) = "", Empty, CLng(ParseNum(arrBase(r, cPrev))))
            ordInfo(5) = Empty
            ordInfo(6) = False
            ordInfo(7) = 0&
            ordInfo(8) = Empty
            ordInfo(9) = False
            dictOrders.Add idPed, ordInfo
        Else
            ordInfo = dictOrders(idPed)
            ' Preferir transportadora informada quando outro item do pedido tem "Não informado"
            If StrComp(CStr(ordInfo(1)), strNaoInf, vbTextCompare) = 0 And StrComp(trItem, strNaoInf, vbTextCompare) <> 0 Then
                ordInfo(1) = trItem
            End If
        End If

        ' Prazo_Real_Dias válido (desconsidera linhas Inconsistente)
        If stAtr = "Inconsistente" Then
            ordInfo(9) = True
        ElseIf IsEmpty(ordInfo(5)) And Not IsEmpty(arrBase(r, cReal)) And Trim$(CStr(arrBase(r, cReal))) <> "" Then
            ordInfo(5) = CLng(ParseNum(arrBase(r, cReal)))
        End If

        ' Atraso
        If stAtr = "Atrasado" Then
            ordInfo(6) = True
            If Not IsEmpty(arrBase(r, cAtrDias)) And Trim$(CStr(arrBase(r, cAtrDias))) <> "" Then
                Dim dAtr As Long
                dAtr = CLng(ParseNum(arrBase(r, cAtrDias)))
                If dAtr > CLng(ordInfo(7)) Then ordInfo(7) = dAtr
            End If
        End If

        ' Avaliação válida: estritamente 1 <= nota <= 5 (filtra notas 0 e 6..10)
        If IsEmpty(ordInfo(8)) And Not IsEmpty(arrBase(r, cAval)) And Trim$(CStr(arrBase(r, cAval))) <> "" Then
            avalVal = ParseNum(arrBase(r, cAval))
            If avalVal >= 1 And avalVal <= 5 And avalVal = Int(avalVal) Then
                avalInt = CLng(avalVal)
                ordInfo(8) = avalInt
            End If
        End If

        dictOrders(idPed) = ordInfo
NextBaseRow:
    Next r

    ' 5. Aplicar filtros de Transportadora (C9) e UF/Região (C10) e agregar
    Dim dictTr As Object
    Set dictTr = CreateObject("Scripting.Dictionary")
    dictTr.CompareMode = vbTextCompare

    ' Lista dinâmica de pedidos atrasados críticos
    Dim critCount As Long
    Dim critArr() As Variant
    ReDim critArr(1 To Application.Max(1, dictOrders.Count), 1 To 10)

    Dim totEnt As Long, totAtr As Long, totInc As Long
    Dim sumAllPrev As Double, cntAllPrev As Long
    Dim sumAllReal As Double, cntAllReal As Long
    Dim sumAllAtrDias As Double, cntAllAtrDias As Long
    Dim sumAllNota As Double, cntAllNota As Long

    Dim kOrder As Variant
    Dim cid As String, uf As String, reg As String, tr As String
    Dim trStats As Variant

    For Each kOrder In dictOrders.keys
        ordInfo = dictOrders(kOrder)
        tr = CStr(ordInfo(1))
        cid = CStr(ordInfo(2))
        uf = IIf(dictCidUF.Exists(cid), CStr(dictCidUF(cid)), "")
        reg = IIf(dictCidReg.Exists(cid), CStr(dictCidReg(cid)), "")

        If hasTransp Then
            If StrComp(tr, fTransp, vbTextCompare) <> 0 Then GoTo NextOrder
        End If
        If hasUfReg Then
            If StrComp(uf, fUfReg, vbTextCompare) <> 0 And StrComp(reg, fUfReg, vbTextCompare) <> 0 Then GoTo NextOrder
        End If

        If Not dictTr.Exists(tr) Then
            ' 0:Entregues, 1:Atrasados, 2:SumPrev, 3:CntPrev, 4:SumReal, 5:CntReal, 6:SumAtrDias, 7:CntAtrDias, 8:SumNota, 9:CntNota
            dictTr.Add tr, Array(0&, 0&, 0#, 0&, 0#, 0&, 0#, 0&, 0#, 0&)
        End If
        trStats = dictTr(tr)

        trStats(0) = trStats(0) + 1
        totEnt = totEnt + 1
        If CBool(ordInfo(9)) Then totInc = totInc + 1

        If Not IsEmpty(ordInfo(4)) Then
            trStats(2) = trStats(2) + CDbl(ordInfo(4))
            trStats(3) = trStats(3) + 1
            sumAllPrev = sumAllPrev + CDbl(ordInfo(4))
            cntAllPrev = cntAllPrev + 1
        End If

        If Not IsEmpty(ordInfo(5)) Then
            trStats(4) = trStats(4) + CDbl(ordInfo(5))
            trStats(5) = trStats(5) + 1
            sumAllReal = sumAllReal + CDbl(ordInfo(5))
            cntAllReal = cntAllReal + 1
        End If

        If CBool(ordInfo(6)) Then
            trStats(1) = trStats(1) + 1
            totAtr = totAtr + 1
            trStats(6) = trStats(6) + CDbl(ordInfo(7))
            trStats(7) = trStats(7) + 1
            sumAllAtrDias = sumAllAtrDias + CDbl(ordInfo(7))
            cntAllAtrDias = cntAllAtrDias + 1

            critCount = critCount + 1
            critArr(critCount, 1) = CStr(kOrder)
            critArr(critCount, 2) = ordInfo(0)
            critArr(critCount, 3) = tr
            critArr(critCount, 4) = cid
            critArr(critCount, 5) = uf
            critArr(critCount, 6) = CStr(ordInfo(3))
            critArr(critCount, 7) = ordInfo(4)
            critArr(critCount, 8) = ordInfo(5)
            critArr(critCount, 9) = CLng(ordInfo(7))
            critArr(critCount, 10) = ordInfo(8)
        End If

        If Not IsEmpty(ordInfo(8)) Then
            trStats(8) = trStats(8) + CDbl(ordInfo(8))
            trStats(9) = trStats(9) + 1
            sumAllNota = sumAllNota + CDbl(ordInfo(8))
            cntAllNota = cntAllNota + 1
        End If

        dictTr(tr) = trStats
NextOrder:
    Next kOrder

    ' 6. Preparar aba Relatorio_Atrasos
    wsOut.Cells.FormatConditions.Delete
    Dim maxClearRow As Long
    maxClearRow = wsOut.Cells(wsOut.Rows.Count, 1).End(xlUp).Row
    If maxClearRow < 700 Then maxClearRow = 700
    wsOut.Range("A6:J" & maxClearRow).Clear

    With wsOut.Range("A1:J1")
        .Merge
        .Value = "VAREJOMAX BRASIL — RELATÓRIO DE ATRASOS E DESEMPENHO POR TRANSPORTADORA"
        .Font.Name = "Calibri"
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(27, 54, 93)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .RowHeight = 30
    End With

    With wsOut.Range("A2:J2")
        .Merge
        .Value = "Gerado em: " & Format$(Now, "dd/mm/yyyy hh:nn:ss") & " | Escopo: Pedidos com Status_Pedido = Entregue (exclui Cancelado, Devolvido, Em trânsito e Em separação)"
        .Font.Name = "Calibri"
        .Font.Size = 9
        .Font.Italic = True
        .Font.Color = RGB(60, 60, 60)
        .Interior.Color = RGB(238, 242, 248)
        .RowHeight = 20
    End With

    With wsOut.Range("A3:J3")
        .Merge
        .Value = "Filtros Aplicados: Transportadora = " & IIf(hasTransp, fTransp, "Todas") & _
                 " | UF/Região = " & IIf(hasUfReg, fUfReg, "Todas") & _
                 " | Período = " & IIf(hasIni, Format$(CDate(dtIni), "dd/mm/yyyy"), "Início") & " até " & IIf(hasFim, Format$(CDate(dtFim), "dd/mm/yyyy"), "Fim")
        .Font.Name = "Calibri"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = RGB(27, 54, 93)
        .Interior.Color = RGB(247, 249, 252)
        .RowHeight = 20
    End With

    Dim pctTotAtr As Double
    pctTotAtr = IIf(totEnt > 0, totAtr / totEnt, 0#)
    With wsOut.Range("A4:J4")
        .Merge
        .Value = "Resumo Logístico: Entregues = " & Format$(totEnt, "#,##0") & _
                 " | Atrasados = " & Format$(totAtr, "#,##0") & " (" & Format$(pctTotAtr, "0.00%") & ")" & _
                 " | Entregas com item Inconsistente excluído dos prazos = " & Format$(totInc, "#,##0") & _
                 " | Nota Média filtrada estritamente para 1 a 5"
        .Font.Name = "Calibri"
        .Font.Size = 9
        .Font.Color = RGB(50, 50, 50)
        .Interior.Color = RGB(247, 249, 252)
        .RowHeight = 19
    End With

    ' 7. Escrever Quadro 1 — Desempenho por Transportadora
    Dim curRow As Long
    curRow = 6
    With wsOut.Range("A" & curRow & ":H" & curRow)
        .Merge
        .Value = "QUADRO 1 — DESEMPENHO POR TRANSPORTADORA (ORDENADO POR PEDIDOS ATRASADOS DESC)"
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(43, 84, 126)
        .RowHeight = 22
    End With

    curRow = 7
    Dim h1 As Variant
    Dim i As Long, j As Long
    h1 = Array("Transportadora", "Pedidos Entregues", "Pedidos Atrasados", "% Atraso", _
               "Prazo Previsto Médio (dias)", "Prazo Real Médio (dias)", _
               "Atraso Médio nos Atrasados (dias)", "Nota Média (1 a 5)")
    For i = 0 To 7
        With wsOut.Cells(curRow, i + 1)
            .Value = h1(i)
            .Font.Name = "Calibri"
            .Font.Size = 10
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(31, 78, 120)
            .HorizontalAlignment = IIf(i = 0, xlLeft, xlRight)
            .VerticalAlignment = xlCenter
        End With
    Next i
    wsOut.Rows(curRow).RowHeight = 22

    Dim q1Start As Long
    q1Start = 8
    If dictTr.Count > 0 Then
        Dim trKeys As Variant
        Dim tmpK As Variant
        trKeys = dictTr.keys
        For i = LBound(trKeys) To UBound(trKeys) - 1
            For j = i + 1 To UBound(trKeys)
                Dim aJ As Long, aI As Long, eJ As Long, eI As Long
                aJ = dictTr(trKeys(j))(1): aI = dictTr(trKeys(i))(1)
                eJ = dictTr(trKeys(j))(0): eI = dictTr(trKeys(i))(0)
                If aJ > aI Or (aJ = aI And eJ > eI) Then
                    tmpK = trKeys(i)
                    trKeys(i) = trKeys(j)
                    trKeys(j) = tmpK
                End If
            Next j
        Next i

        For i = LBound(trKeys) To UBound(trKeys)
            curRow = curRow + 1
            tr = CStr(trKeys(i))
            trStats = dictTr(tr)

            wsOut.Cells(curRow, 1).Value = tr
            wsOut.Cells(curRow, 2).Value = trStats(0)
            wsOut.Cells(curRow, 3).Value = trStats(1)
            wsOut.Cells(curRow, 4).Value = IIf(trStats(0) > 0, trStats(1) / trStats(0), 0#)
            wsOut.Cells(curRow, 5).Value = IIf(trStats(3) > 0, trStats(2) / trStats(3), Empty)
            wsOut.Cells(curRow, 6).Value = IIf(trStats(5) > 0, trStats(4) / trStats(5), Empty)
            wsOut.Cells(curRow, 7).Value = IIf(trStats(7) > 0, trStats(6) / trStats(7), Empty)
            wsOut.Cells(curRow, 8).Value = IIf(trStats(9) > 0, trStats(8) / trStats(9), Empty)

            If (i Mod 2) = 1 Then
                wsOut.Range(wsOut.Cells(curRow, 1), wsOut.Cells(curRow, 8)).Interior.Color = RGB(244, 247, 251)
            End If
        Next i
    Else
        curRow = curRow + 1
        wsOut.Cells(curRow, 1).Value = "(Sem registros para os filtros informados)"
    End If

    ' Total Geral Quadro 1
    curRow = curRow + 1
    wsOut.Cells(curRow, 1).Value = "TOTAL GERAL"
    wsOut.Cells(curRow, 2).Value = totEnt
    wsOut.Cells(curRow, 3).Value = totAtr
    wsOut.Cells(curRow, 4).Value = pctTotAtr
    wsOut.Cells(curRow, 5).Value = IIf(cntAllPrev > 0, sumAllPrev / cntAllPrev, Empty)
    wsOut.Cells(curRow, 6).Value = IIf(cntAllReal > 0, sumAllReal / cntAllReal, Empty)
    wsOut.Cells(curRow, 7).Value = IIf(cntAllAtrDias > 0, sumAllAtrDias / cntAllAtrDias, Empty)
    wsOut.Cells(curRow, 8).Value = IIf(cntAllNota > 0, sumAllNota / cntAllNota, Empty)

    With wsOut.Range(wsOut.Cells(curRow, 1), wsOut.Cells(curRow, 8))
        .Font.Bold = True
        .Interior.Color = RGB(217, 225, 242)
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeBottom).LineStyle = xlDouble
    End With

    With wsOut.Range(wsOut.Cells(q1Start, 1), wsOut.Cells(curRow, 8))
        .Font.Name = "Calibri"
        .Font.Size = 10
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).Color = RGB(217, 217, 217)
    End With
    wsOut.Range(wsOut.Cells(q1Start, 2), wsOut.Cells(curRow, 3)).NumberFormat = "#,##0"
    wsOut.Range(wsOut.Cells(q1Start, 4), wsOut.Cells(curRow, 4)).NumberFormat = "0.00%"
    wsOut.Range(wsOut.Cells(q1Start, 5), wsOut.Cells(curRow, 8)).NumberFormat = "0.00"

    ' 8. Escrever Quadro 2 — Pedidos Críticos Atrasados
    curRow = curRow + 2
    With wsOut.Range("A" & curRow & ":J" & curRow)
        .Merge
        .Value = "QUADRO 2 — PEDIDOS CRÍTICOS ATRASADOS (ORDENADO POR ATRASO_DIAS DESC | DESTAQUE VERMELHO: ATRASO >= 5 DIAS)"
        .Font.Name = "Calibri"
        .Font.Size = 11
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(43, 84, 126)
        .RowHeight = 22
    End With

    curRow = curRow + 1
    Dim h2 As Variant
    h2 = Array("ID_Pedido", "Data_Pedido", "Transportadora", "Cidade", "UF", "Canal", _
               "Prazo_Previsto_Dias", "Prazo_Real_Dias", "Atraso_Dias", "Avaliacao_Cliente")
    For i = 0 To 9
        With wsOut.Cells(curRow, i + 1)
            .Value = h2(i)
            .Font.Name = "Calibri"
            .Font.Size = 10
            .Font.Bold = True
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(31, 78, 120)
            .HorizontalAlignment = IIf(i >= 6, xlRight, xlLeft)
            .VerticalAlignment = xlCenter
        End With
    Next i
    wsOut.Rows(curRow).RowHeight = 22

    Dim q2Start As Long
    q2Start = curRow + 1

    If critCount > 0 Then
        QuickSortCrit critArr, 1, critCount

        Dim outCrit() As Variant
        ReDim outCrit(1 To critCount, 1 To 10)
        For r = 1 To critCount
            For j = 1 To 10
                outCrit(r, j) = critArr(r, j)
            Next j
        Next r

        wsOut.Range(wsOut.Cells(q2Start, 1), wsOut.Cells(q2Start + critCount - 1, 10)).Value = outCrit

        With wsOut.Range(wsOut.Cells(q2Start, 1), wsOut.Cells(q2Start + critCount - 1, 10))
            .Font.Name = "Calibri"
            .Font.Size = 10
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).Color = RGB(230, 230, 230)
        End With

        wsOut.Range(wsOut.Cells(q2Start, 2), wsOut.Cells(q2Start + critCount - 1, 2)).NumberFormat = "dd/mm/yyyy"
        wsOut.Range(wsOut.Cells(q2Start, 7), wsOut.Cells(q2Start + critCount - 1, 10)).NumberFormat = "#,##0"

        ' Destacar em vermelho claro linhas com Atraso_Dias >= 5
        For r = 1 To critCount
            If CLng(outCrit(r, 9)) >= 5 Then
                With wsOut.Range(wsOut.Cells(q2Start + r - 1, 1), wsOut.Cells(q2Start + r - 1, 10))
                    .Interior.Color = RGB(255, 217, 217)
                    .Font.Color = RGB(156, 0, 6)
                    .Font.Bold = True
                End With
            End If
        Next r

        ' Regra de formatação condicional na coluna Atraso_Dias (>= 5)
        Dim rngCrit As Range
        Dim fc As FormatCondition
        Set rngCrit = wsOut.Range(wsOut.Cells(q2Start, 9), wsOut.Cells(q2Start + critCount - 1, 9))
        Set fc = rngCrit.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="5")
        fc.Interior.Color = RGB(255, 217, 217)
        fc.Font.Color = RGB(156, 0, 6)
        fc.Font.Bold = True
    Else
        wsOut.Cells(q2Start, 1).Value = "(Nenhum pedido atrasado encontrado para os filtros informados)"
    End If

    wsOut.Columns("A:J").AutoFit
    If wsOut.Columns("A").ColumnWidth < 24 Then wsOut.Columns("A").ColumnWidth = 24
    If wsOut.Columns("C").ColumnWidth < 22 Then wsOut.Columns("C").ColumnWidth = 22

CleanExit:
    Application.Calculation = calcPrev
    Application.EnableEvents = evtPrev
    Application.ScreenUpdating = scrPrev
    Exit Sub

ErrHandler:
    MsgBox "Erro em GerarRelatorioAtrasosTransportadora (" & Err.Number & "): " & Err.Description, vbCritical, "VarejoMax VBA"
    Resume CleanExit
End Sub

Private Sub QuickSortCrit(ByRef arr() As Variant, ByVal low As Long, ByVal high As Long)
    Dim i As Long, j As Long, c As Long
    Dim pivotAtr As Long, pivotId As String
    Dim tmp As Variant

    If low >= high Then Exit Sub
    i = low
    j = high
    pivotAtr = CLng(arr((low + high) \ 2, 9))
    pivotId = CStr(arr((low + high) \ 2, 1))

    Do While i <= j
        Do While (CLng(arr(i, 9)) > pivotAtr) Or (CLng(arr(i, 9)) = pivotAtr And StrComp(CStr(arr(i, 1)), pivotId, vbTextCompare) < 0)
            i = i + 1
        Loop
        Do While (CLng(arr(j, 9)) < pivotAtr) Or (CLng(arr(j, 9)) = pivotAtr And StrComp(CStr(arr(j, 1)), pivotId, vbTextCompare) > 0)
            j = j - 1
        Loop
        If i <= j Then
            For c = 1 To 10
                tmp = arr(i, c)
                arr(i, c) = arr(j, c)
                arr(j, c) = tmp
            Next c
            i = i + 1
            j = j - 1
        End If
    Loop

    If low < j Then QuickSortCrit arr, low, j
    If i < high Then QuickSortCrit arr, i, high
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
    Err.Raise vbObjectError + 102, "modAtrasosLogistica", "Coluna não encontrada: " & headerName
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
