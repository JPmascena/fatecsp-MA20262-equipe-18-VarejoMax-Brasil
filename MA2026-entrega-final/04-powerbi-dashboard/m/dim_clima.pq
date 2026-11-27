let
    Fonte = Odbc.Query(
        pConexaoODBC,
        "SELECT chave_clima, cidade, data, temp_max, temp_min, precipitacao, faixa_precipitacao FROM dim_clima"
    ),
    Tipos = Table.TransformColumnTypes(
        Fonte,
        {
            {"chave_clima", type text},
            {"cidade", type text},
            {"data", type date},
            {"temp_max", type number},
            {"temp_min", type number},
            {"precipitacao", type number},
            {"faixa_precipitacao", type text}
        },
        "en-US"
    )
in
    Tipos
