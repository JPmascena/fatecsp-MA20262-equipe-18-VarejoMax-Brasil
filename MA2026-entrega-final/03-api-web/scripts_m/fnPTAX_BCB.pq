// ============================================================================
// fnPTAX_BCB.pq — Consumo da API Olinda PTAX (Banco Central do Brasil)
// VarejoMax Brasil | MA2026
// ============================================================================
// Retorna serie diaria continua de 2025-01-01 a 2026-08-31 (608 dias),
// mantendo o ultimo boletim do dia util e preenchendo fins de semana/feriados
// via carry-forward (Table.FillDown) a partir de 2024-12-20.
// ============================================================================

let
    // 1. Consulta PTAX a partir de 12-20-2024 para garantir cotacao util antes de 2025-01-01
    FonteJson = Json.Document(
        Web.Contents(
            "https://olinda.bcb.gov.br",
            [
                RelativePath = "olinda/servico/PTAX/versao/v1/odata/CotacaoDolarPeriodo(dataInicial=@dataInicial,dataFinalCotacao=@dataFinalCotacao)",
                Query = [
                    #"@dataInicial" = "'12-20-2024'",
                    #"@dataFinalCotacao" = "'08-31-2026'",
                    #"$format" = "json",
                    #"$select" = "cotacaoCompra,cotacaoVenda,dataHoraCotacao"
                ]
            ]
        )
    ),
    ListaValores = FonteJson[value],
    TabelaBruta = Table.FromList(ListaValores, Splitter.SplitByNothing(), {"Registro"}, null, ExtraValues.Error),
    Expandido = Table.ExpandRecordColumn(
        TabelaBruta,
        "Registro",
        {"cotacaoCompra", "cotacaoVenda", "dataHoraCotacao"},
        {"cotacao_compra", "cotacao_venda", "data_hora_cotacao"}
    ),
    TiposIniciais = Table.TransformColumnTypes(
        Expandido,
        {
            {"cotacao_compra", type number},
            {"cotacao_venda", type number},
            {"data_hora_cotacao", type datetime}
        },
        "en-US"
    ),
    ComData = Table.AddColumn(TiposIniciais, "data", each Date.From([data_hora_cotacao]), type date),

    // 2. Mantem apenas o ultimo boletim (maior data_hora_cotacao) de cada dia util
    OrdenadoDesc = Table.Buffer(Table.Sort(ComData, {{"data_hora_cotacao", Order.Descending}})),
    UltimoBoletimDia = Table.Distinct(OrdenadoDesc, {"data"}),
    SelecionaBoletim = Table.SelectColumns(UltimoBoletimDia, {"data", "cotacao_compra", "cotacao_venda"}),

    // 3. Gera calendario continuo de 2024-12-20 a 2026-08-31 para suportar carry-forward inicial
    DataSeed = #date(2024, 12, 20),
    DataFim = #date(2026, 8, 31),
    QtdDiasSeed = Duration.Days(DataFim - DataSeed) + 1,
    ListaDatas = List.Dates(DataSeed, QtdDiasSeed, #duration(1, 0, 0, 0)),
    TabelaCalendario = Table.FromList(ListaDatas, Splitter.SplitByNothing(), {"data"}, null, ExtraValues.Error),
    CalendarioTipado = Table.TransformColumnTypes(TabelaCalendario, {{"data", type date}}),

    // 4. Left Join + classificacao de origem antes do FillDown
    Juncao = Table.NestedJoin(CalendarioTipado, {"data"}, SelecionaBoletim, {"data"}, "PTAX", JoinKind.LeftOuter),
    ExpandidoCal = Table.ExpandTableColumn(Juncao, "PTAX", {"cotacao_compra", "cotacao_venda"}, {"cotacao_compra", "cotacao_venda"}),
    OrdenadoAsc = Table.Sort(ExpandidoCal, {{"data", Order.Ascending}}),
    ComOrigem = Table.AddColumn(
        OrdenadoAsc,
        "boletim_origem",
        each if [cotacao_venda] <> null then "Oficial BCB" else "Carry-forward",
        type text
    ),

    // 5. Carry-forward (preenche fins de semana e feriados com o ultimo dia util anterior)
    Preenchido = Table.FillDown(ComOrigem, {"cotacao_compra", "cotacao_venda"}),

    // 6. Recorta exatamente o periodo analitico (2025-01-01 a 2026-08-31 -> 608 dias)
    PeriodoFinal = Table.SelectRows(Preenchido, each [data] >= #date(2025, 1, 1) and [data] <= #date(2026, 8, 31)),

    // 7. Arredondamento para 4 casas decimais e faixa analitica do dolar
    Arredondado = Table.TransformColumns(
        PeriodoFinal,
        {
            {"cotacao_compra", each Number.Round(_, 4), type number},
            {"cotacao_venda", each Number.Round(_, 4), type number}
        }
    ),
    ComFaixa = Table.AddColumn(
        Arredondado,
        "faixa_dolar",
        each
            if [cotacao_venda] < 5.10 then "1-Abaixo de R$ 5,10"
            else if [cotacao_venda] < 5.40 then "2-R$ 5,10 a R$ 5,39"
            else if [cotacao_venda] < 5.70 then "3-R$ 5,40 a R$ 5,69"
            else if [cotacao_venda] < 6.00 then "4-R$ 5,70 a R$ 5,99"
            else "5-R$ 6,00 ou mais",
        type text
    )
in
    ComFaixa
