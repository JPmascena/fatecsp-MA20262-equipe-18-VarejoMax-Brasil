// Consulta de diagnóstico da conexão ODBC (habilitar carga = Desabilitado)
// Verifica a versão do SQLite e a contagem de registros das 8 views dimensionais + fato_vendas.
let
    SqlDiagnostico = "
        SELECT 'sqlite_version' AS objeto, sqlite_version() AS valor
        UNION ALL SELECT 'dim_produto', CAST(COUNT(*) AS TEXT) FROM dim_produto
        UNION ALL SELECT 'dim_cliente', CAST(COUNT(*) AS TEXT) FROM dim_cliente
        UNION ALL SELECT 'dim_cidade', CAST(COUNT(*) AS TEXT) FROM dim_cidade
        UNION ALL SELECT 'dim_canal_loja', CAST(COUNT(*) AS TEXT) FROM dim_canal_loja
        UNION ALL SELECT 'dim_transportadora', CAST(COUNT(*) AS TEXT) FROM dim_transportadora
        UNION ALL SELECT 'dim_tempo', CAST(COUNT(*) AS TEXT) FROM dim_tempo
        UNION ALL SELECT 'dim_cambio', CAST(COUNT(*) AS TEXT) FROM dim_cambio
        UNION ALL SELECT 'dim_clima', CAST(COUNT(*) AS TEXT) FROM dim_clima
        UNION ALL SELECT 'fato_vendas', CAST(COUNT(*) AS TEXT) FROM fato_vendas
    ",
    Fonte = Odbc.Query(pConexaoODBC, SqlDiagnostico),
    Tipos = Table.TransformColumnTypes(
        Fonte,
        {
            {"objeto", type text},
            {"valor", type text}
        },
        "en-US"
    )
in
    Tipos
