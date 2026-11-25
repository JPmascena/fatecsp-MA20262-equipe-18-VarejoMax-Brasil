let
    Fonte = Odbc.Query(
        pConexaoODBC,
        "SELECT transportadora, tipo FROM dim_transportadora"
    ),
    Tipos = Table.TransformColumnTypes(
        Fonte,
        {
            {"transportadora", type text},
            {"tipo", type text}
        },
        "en-US"
    )
in
    Tipos
