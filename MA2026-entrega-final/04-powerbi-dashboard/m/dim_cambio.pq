let
    Fonte = Odbc.Query(
        pConexaoODBC,
        "SELECT data, cotacao_compra, cotacao_venda, boletim_origem, faixa_dolar FROM dim_cambio"
    ),
    Tipos = Table.TransformColumnTypes(
        Fonte,
        {
            {"data", type date},
            {"cotacao_compra", type number},
            {"cotacao_venda", type number},
            {"boletim_origem", type text},
            {"faixa_dolar", type text}
        },
        "en-US"
    )
in
    Tipos
