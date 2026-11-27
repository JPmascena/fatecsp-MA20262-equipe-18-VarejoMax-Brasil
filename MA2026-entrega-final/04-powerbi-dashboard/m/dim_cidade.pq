let
    Fonte = Odbc.Query(
        pConexaoODBC,
        "SELECT cidade, uf, regiao, latitude, longitude, tem_loja_fisica FROM dim_cidade"
    ),
    Tipos = Table.TransformColumnTypes(
        Fonte,
        {
            {"cidade", type text},
            {"uf", type text},
            {"regiao", type text},
            {"latitude", type number},
            {"longitude", type number},
            {"tem_loja_fisica", type text}
        },
        "en-US"
    )
in
    Tipos
