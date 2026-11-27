// ============================================================================
// fnClima_OpenMeteo.pq — Funcao Power Query M para Open-Meteo Archive API
// VarejoMax Brasil | MA2026
// ============================================================================
// Invocada a partir da dimensao dCidade (24 cidades) com Latitude e Longitude.
// Retorna 608 dias (2025-01-01 a 2026-08-31) com fuso America/Sao_Paulo.
// ============================================================================

(cidade as text, lat as number, lon as number) as table =>
let
    // 1. Requisicao parametrizada com RelativePath e Query (compativel com Refresh no Power BI Service)
    FonteJson = Json.Document(
        Web.Contents(
            "https://archive-api.open-meteo.com",
            [
                RelativePath = "v1/archive",
                Query = [
                    latitude = Number.ToText(lat, "en-US"),
                    longitude = Number.ToText(lon, "en-US"),
                    start_date = "2025-01-01",
                    end_date = "2026-08-31",
                    daily = "temperature_2m_max,temperature_2m_min,precipitation_sum",
                    timezone = "America/Sao_Paulo"
                ]
            ]
        )
    ),
    Daily = FonteJson[daily],

    // 2. Combina as listas paralelas retornadas pela API em tabela colunar
    TabelaColunas = Table.FromColumns(
        {
            List.Repeat({cidade}, List.Count(Daily[time])),
            Daily[time],
            Daily[temperature_2m_max],
            Daily[temperature_2m_min],
            Daily[precipitation_sum]
        },
        {"cidade", "data_ref", "temp_max", "temp_min", "precipitacao"}
    ),
    Tipos = Table.TransformColumnTypes(
        TabelaColunas,
        {
            {"cidade", type text},
            {"data_ref", type date},
            {"temp_max", type number},
            {"temp_min", type number},
            {"precipitacao", type number}
        },
        "en-US"
    ),

    // 3. Fallback controlado para eventual null em dias recentes + arredondamento (1 casa)
    PreenchidoTemp = Table.FillDown(Tipos, {"temp_max", "temp_min"}),
    PrecipSemNull = Table.ReplaceValue(PreenchidoTemp, null, 0.0, Replacer.ReplaceValue, {"precipitacao"}),
    Arredondado = Table.TransformColumns(
        PrecipSemNull,
        {
            {"temp_max", each Number.Round(_, 1), type number},
            {"temp_min", each Number.Round(_, 1), type number},
            {"precipitacao", each Number.Round(_, 1), type number}
        }
    ),

    // 4. Classificacao da precipitacao diaria em 4 faixas operacionais
    ComFaixa = Table.AddColumn(
        Arredondado,
        "faixa_precipitacao",
        each
            if [precipitacao] = 0 then "0-Sem chuva (0 mm)"
            else if [precipitacao] <= 10 then "1-Chuva fraca (0,1-10 mm)"
            else if [precipitacao] <= 30 then "2-Chuva moderada (10-30 mm)"
            else "3-Chuva forte (>30 mm)",
        type text
    )
in
    ComFaixa
