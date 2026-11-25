let
    Fonte = Odbc.Query(
        pConexaoODBC,
        "SELECT id_cliente, nome_cliente, email, data_cadastro, cidade, uf, regiao FROM dim_cliente"
    ),
    VaziosParaNull = Table.ReplaceValue(
        Fonte,
        "",
        null,
        Replacer.ReplaceValue,
        {"email", "data_cadastro"}
    ),
    Tipos = Table.TransformColumnTypes(
        VaziosParaNull,
        {
            {"id_cliente", type text},
            {"nome_cliente", type text},
            {"email", type text},
            {"data_cadastro", type date},
            {"cidade", type text},
            {"uf", type text},
            {"regiao", type text}
        },
        "en-US"
    )
in
    Tipos
