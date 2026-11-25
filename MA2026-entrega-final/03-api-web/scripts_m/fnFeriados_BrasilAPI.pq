// ============================================================================
// fnFeriados_BrasilAPI.pq — Consumo da API de Feriados Nacionais (BrasilAPI)
// VarejoMax Brasil | MA2026
// ============================================================================
// Coleta os anos 2025 e 2026 e recorta o periodo analitico 2025-01-01 a 2026-08-31.
// ============================================================================

let
    fnBuscaAno = (ano as number) as table =>
        let
            Fonte = Json.Document(
                Web.Contents(
                    "https://brasilapi.com.br",
                    [
                        RelativePath = "api/feriados/v1/" & Text.From(ano)
                    ]
                )
            ),
            Tabela = Table.FromList(Fonte, Splitter.SplitByNothing(), {"Registro"}, null, ExtraValues.Error),
            Expandido = Table.ExpandRecordColumn(
                Tabela,
                "Registro",
                {"date", "name", "type"},
                {"data", "nome_feriado", "tipo_raw"}
            )
        in
            Expandido,

    Consolidado = Table.Combine({fnBuscaAno(2025), fnBuscaAno(2026)}),
    Tipado = Table.TransformColumnTypes(
        Consolidado,
        {
            {"data", type date},
            {"nome_feriado", type text},
            {"tipo_raw", type text}
        },
        "en-US"
    ),
    PeriodoRecorte = Table.SelectRows(
        Tipado,
        each [data] >= #date(2025, 1, 1) and [data] <= #date(2026, 8, 31)
    ),
    ComTipoPadronizado = Table.AddColumn(
        PeriodoRecorte,
        "tipo",
        each if [tipo_raw] = "national" then "Nacional" else [tipo_raw],
        type text
    ),
    Final = Table.SelectColumns(
        Table.Sort(ComTipoPadronizado, {{"data", Order.Ascending}}),
        {"data", "nome_feriado", "tipo"}
    )
in
    Final
