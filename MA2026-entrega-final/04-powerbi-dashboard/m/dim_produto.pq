let
    Fonte = Odbc.Query(
        pConexaoODBC,
        "SELECT id_produto, produto, categoria, marca, preco_ref_2025, preco_ref_2026, custo_ref FROM dim_produto"
    ),
    Tipos = Table.TransformColumnTypes(
        Fonte,
        {
            {"id_produto", type text},
            {"produto", type text},
            {"categoria", type text},
            {"marca", type text},
            {"preco_ref_2025", Currency.Type},
            {"preco_ref_2026", Currency.Type},
            {"custo_ref", Currency.Type}
        },
        "en-US"
    )
in
    Tipos
