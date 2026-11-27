let
    Fonte = Odbc.Query(
        pConexaoODBC,
        "SELECT chave_canal_loja, canal, loja, tipo_canal, cod_loja, cidade_loja, uf_loja FROM dim_canal_loja"
    ),
    VaziosParaNull = Table.ReplaceValue(
        Fonte,
        "",
        null,
        Replacer.ReplaceValue,
        {"cod_loja", "cidade_loja", "uf_loja"}
    ),
    Tipos = Table.TransformColumnTypes(
        VaziosParaNull,
        {
            {"chave_canal_loja", type text},
            {"canal", type text},
            {"loja", type text},
            {"tipo_canal", type text},
            {"cod_loja", type text},
            {"cidade_loja", type text},
            {"uf_loja", type text}
        },
        "en-US"
    )
in
    Tipos
