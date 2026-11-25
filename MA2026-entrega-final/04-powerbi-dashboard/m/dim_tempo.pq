let
    Fonte = Odbc.Query(
        pConexaoODBC,
        "SELECT data, ano, mes, nome_mes, mes_abrev, ano_mes, mes_inicio, trimestre, dia_semana, nome_dia, fim_de_semana, feriado, nome_feriado FROM dim_tempo"
    ),
    VaziosParaNull = Table.ReplaceValue(
        Fonte,
        "",
        null,
        Replacer.ReplaceValue,
        {"nome_feriado"}
    ),
    Tipos = Table.TransformColumnTypes(
        VaziosParaNull,
        {
            {"data", type date},
            {"ano", Int64.Type},
            {"mes", Int64.Type},
            {"nome_mes", type text},
            {"mes_abrev", type text},
            {"ano_mes", type text},
            {"mes_inicio", type date},
            {"trimestre", type text},
            {"dia_semana", Int64.Type},
            {"nome_dia", type text},
            {"fim_de_semana", type text},
            {"feriado", type text},
            {"nome_feriado", type text}
        },
        "en-US"
    )
in
    Tipos
