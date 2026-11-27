Attribute VB_Name = "modUtilitariosExportacao"
Option Explicit

' ==============================================================================
' Módulo: modUtilitariosExportacao
' Projeto: VarejoMax Brasil (FATEC SP - Microinformática Aplicada)
' Objetivo: Orquestrar atualização geral com conferência de batimento, exportação
'           de relatórios em PDF (paisagem, ajustado à largura) e limpeza de
'           filtros/relatórios mantendo cabeçalho e layout executivo.
' ==============================================================================

Private Const GABARITO_FAT_LIQUIDO As Double = 3448872.47
Private Const GABARITO_PEDIDOS_VALIDOS As Long = 8291
Private Const GABARITO_ATRASADOS_VALIDOS As Long = 628

Public Sub AtualizarBasesERelatorios()
    Dim calcPrev As XlCalculation
    Dim scrPrev As Boolean
    Dim evtPrev As Boolean
    Dim tInicio As Double

    Dim wsMenu As Worksheet
    Dim wsRes As Worksheet
    Dim wsAtr As Worksheet

    Dim semFiltros As Boolean
    Dim rowTotRes As Long
    Dim rowTotAtr As Long
    Dim fatLiqApurado As Double
    Dim pedValApurado As Long
    Dim pedAtrApurado As Long
    Dim diffFat As Double
    Dim msgStatus As String

    On Error GoTo ErrHandler

    tInicio = Timer
    scrPrev = Application.ScreenUpdating
    evtPrev = Application.EnableEvents
    calcPrev = Application.Calculation

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ThisWorkbook.RefreshAll

    ' Re-aplicar modo manual caso RefreshAll altere estado de eventos/tela
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    modResumoVendas.GerarResumoVendasCategoria
    modAtrasosLogistica.GerarRelatorioAtrasosTransportadora

    Set wsMenu = ThisWorkbook.Worksheets("Menu_Operacional")
    Set wsRes = ThisWorkbook.Worksheets("Resumo_Categorias")
    Set wsAtr = ThisWorkbook.Worksheets("Relatorio_Atrasos")

    semFiltros = IsEmptyOrAll(wsMenu.Range("C6").Value) And _
                 IsEmptyOrAll(wsMenu.Range("C7").Value) And _
                 IsEmptyOrAll(wsMenu.Range("C8").Value) And _
                 IsEmptyOrAll(wsMenu.Range("C9").Value) And _
                 IsEmptyOrAll(wsMenu.Range("C10").Value)

    rowTotRes = FindTotalRow(wsRes, 1, 6, 40)
    rowTotAtr = FindTotalRow(wsAtr, 1, 7, 40)

    If rowTotRes > 0 Then
        pedValApurado = CLng(wsRes.Cells(rowTotRes, 2).Value)
        fatLiqApurado = CDbl(wsRes.Cells(rowTotRes, 5).Value)
    End If
    If rowTotAtr > 0 Then
        pedAtrApurado = CLng(wsAtr.Cells(rowTotAtr, 3).Value)
    End If

    If semFiltros Then
        diffFat = Abs(fatLiqApurado - GABARITO_FAT_LIQUIDO)
        If diffFat < 0.01 And pedValApurado = GABARITO_PEDIDOS_VALIDOS And pedAtrApurado = GABARITO_ATRASADOS_VALIDOS Then
            msgStatus = "RECONCILIAÇÃO OFICIAL: APROVADA (Diferença = R$ 0,00)"
        Else
            msgStatus = "RECONCILIAÇÃO: DIVERGÊNCIA DETECTADA (Diferença = R$ " & Format$(diffFat, "#,##0.00") & ")"
        End If
    Else
        msgStatus = "Modo Filtrado ativo — reconciliação integral não aplicável."
    End If

    MsgBox "Relatórios atualizados com sucesso em " & Format$(Timer - tInicio, "0.00") & " s." & vbCrLf & vbCrLf & _
           "• Pedidos Válidos: " & Format$(pedValApurado, "#,##0") & vbCrLf & _
           "• Faturamento Líquido: R$ " & Format$(fatLiqApurado, "#,##0.00") & vbCrLf & _
           "• Pedidos Atrasados (Entregues): " & Format$(pedAtrApurado, "#,##0") & vbCrLf & vbCrLf & _
           msgStatus, vbInformation, "VarejoMax Brasil — Atualização Concluída"

CleanExit:
    Application.Calculation = calcPrev
    Application.EnableEvents = evtPrev
    Application.ScreenUpdating = scrPrev
    Exit Sub

ErrHandler:
    MsgBox "Erro em AtualizarBasesERelatorios (" & Err.Number & "): " & Err.Description, vbCritical, "VarejoMax VBA"
    Resume CleanExit
End Sub

Public Sub ExportarRelatoriosPDF()
    Dim calcPrev As XlCalculation
    Dim scrPrev As Boolean
    Dim evtPrev As Boolean

    Dim baseDir As String
    Dim outDir As String
    Dim stamp As String
    Dim pdfRes As String
    Dim pdfAtr As String

    Dim wsRes As Worksheet
    Dim wsAtr As Worksheet

    On Error GoTo ErrHandler

    scrPrev = Application.ScreenUpdating
    evtPrev = Application.EnableEvents
    calcPrev = Application.Calculation

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    baseDir = ThisWorkbook.Path
    If Len(baseDir) = 0 Then
        Err.Raise vbObjectError + 201, "ExportarRelatoriosPDF", _
            "Salve a pasta de trabalho (.xlsm) antes de exportar os PDFs."
    End If

    outDir = baseDir & "\relatorios_gerados"
    If Dir(outDir, vbDirectory) = "" Then
        MkDir outDir
    End If

    stamp = Format$(Now, "yyyymmdd_hhnn")
    pdfRes = outDir & "\Resumo_Categorias_" & stamp & ".pdf"
    pdfAtr = outDir & "\Relatorio_Atrasos_" & stamp & ".pdf"

    Set wsRes = ThisWorkbook.Worksheets("Resumo_Categorias")
    Set wsAtr = ThisWorkbook.Worksheets("Relatorio_Atrasos")

    ConfigureLandscapePage wsRes, True
    ConfigureLandscapePage wsAtr, False

    wsRes.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfRes, _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, OpenAfterPublish:=False

    wsAtr.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfAtr, _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, OpenAfterPublish:=False

    MsgBox "PDFs exportados com sucesso em:" & vbCrLf & vbCrLf & _
           "• " & pdfRes & vbCrLf & _
           "• " & pdfAtr, vbInformation, "VarejoMax Brasil — Exportação PDF"

CleanExit:
    Application.Calculation = calcPrev
    Application.EnableEvents = evtPrev
    Application.ScreenUpdating = scrPrev
    Exit Sub

ErrHandler:
    MsgBox "Erro em ExportarRelatoriosPDF (" & Err.Number & "): " & Err.Description, vbCritical, "VarejoMax VBA"
    Resume CleanExit
End Sub

Public Sub LimparFiltrosERelatorios()
    Dim calcPrev As XlCalculation
    Dim scrPrev As Boolean
    Dim evtPrev As Boolean

    Dim wsMenu As Worksheet
    Dim wsRes As Worksheet
    Dim wsAtr As Worksheet
    Dim maxRowRes As Long
    Dim maxRowAtr As Long

    On Error GoTo ErrHandler

    scrPrev = Application.ScreenUpdating
    evtPrev = Application.EnableEvents
    calcPrev = Application.Calculation

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Set wsMenu = ThisWorkbook.Worksheets("Menu_Operacional")
    Set wsRes = ThisWorkbook.Worksheets("Resumo_Categorias")
    Set wsAtr = ThisWorkbook.Worksheets("Relatorio_Atrasos")

    ' Limpar filtros opcionais de Menu_Operacional (C6:C10)
    wsMenu.Range("C6:C10").ClearContents

    ' Limpar quadros gerados preservando cabeçalho executivo (linhas 1 a 3/4)
    wsRes.Cells.FormatConditions.Delete
    maxRowRes = wsRes.Cells(wsRes.Rows.Count, 1).End(xlUp).Row
    If maxRowRes < 40 Then maxRowRes = 40
    wsRes.Range("A5:I" & maxRowRes).Clear
    wsRes.Range("A3").Value = "Filtros Aplicados: Período = Início até Fim | Canal = Todos (relatório limpo; execute AtualizarBasesERelatorios)"

    wsAtr.Cells.FormatConditions.Delete
    maxRowAtr = wsAtr.Cells(wsAtr.Rows.Count, 1).End(xlUp).Row
    If maxRowAtr < 700 Then maxRowAtr = 700
    wsAtr.Range("A6:J" & maxRowAtr).Clear
    wsAtr.Range("A3").Value = "Filtros Aplicados: Transportadora = Todas | UF/Região = Todas | Período = Completo (relatório limpo)"
    wsAtr.Range("A4").Value = "Resumo Logístico: (aguardando execução da macro GerarRelatorioAtrasosTransportadora)"

    MsgBox "Filtros de Menu_Operacional e quadros dos relatórios limpos com sucesso.", _
           vbInformation, "VarejoMax Brasil — Limpeza Concluída"

CleanExit:
    Application.Calculation = calcPrev
    Application.EnableEvents = evtPrev
    Application.ScreenUpdating = scrPrev
    Exit Sub

ErrHandler:
    MsgBox "Erro em LimparFiltrosERelatorios (" & Err.Number & "): " & Err.Description, vbCritical, "VarejoMax VBA"
    Resume CleanExit
End Sub

Private Sub ConfigureLandscapePage(ByVal ws As Worksheet, ByVal fitSingleTall As Boolean)
    With ws.PageSetup
        .Orientation = xlLandscape
        .PaperSize = xlPaperA4
        .Zoom = False
        .FitToPagesWide = 1
        If fitSingleTall Then
            .FitToPagesTall = 1
        Else
            .FitToPagesTall = False
        End If
        .CenterHorizontally = True
    End With
End Sub

Private Function IsEmptyOrAll(ByVal v As Variant) As Boolean
    Dim s As String
    s = UCase$(Trim$(CStr(v)))
    IsEmptyOrAll = (Len(s) = 0 Or s = "TODOS" Or s = "TODAS")
End Function

Private Function FindTotalRow(ByVal ws As Worksheet, ByVal colNum As Long, ByVal startRow As Long, ByVal endRow As Long) As Long
    Dim r As Long
    For r = startRow To endRow
        If StrComp(Trim$(CStr(ws.Cells(r, colNum).Value)), "TOTAL GERAL", vbTextCompare) = 0 Then
            FindTotalRow = r
            Exit Function
        End If
    Next r
    FindTotalRow = 0
End Function
