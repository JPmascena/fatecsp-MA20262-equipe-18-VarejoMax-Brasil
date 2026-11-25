let
    Fonte = Odbc.Query(
        pConexaoODBC,
        "SELECT id_pedido, item, data_pedido, data_envio, data_entrega_prevista, data_entrega_real, prazo_previsto_dias, prazo_real_dias, atraso_dias, status_atraso, status_pedido, canal, loja, chave_canal_loja, id_cliente, cidade, cep, id_produto, quantidade, preco_unitario, custo_unitario, desconto_pct, valor_bruto, valor_liquido, custo_ref_total, custo_real_total, frete_valor, forma_pagamento, parcelas, transportadora, avaliacao_cliente, chave_clima FROM fato_vendas"
    ),
    VaziosParaNull = Table.ReplaceValue(
        Fonte,
        "",
        null,
        Replacer.ReplaceValue,
        {
            "data_envio",
            "data_entrega_prevista",
            "data_entrega_real",
            "prazo_previsto_dias",
            "prazo_real_dias",
            "atraso_dias",
            "avaliacao_cliente"
        }
    ),
    Tipos = Table.TransformColumnTypes(
        VaziosParaNull,
        {
            {"id_pedido", type text},
            {"item", Int64.Type},
            {"data_pedido", type date},
            {"data_envio", type date},
            {"data_entrega_prevista", type date},
            {"data_entrega_real", type date},
            {"prazo_previsto_dias", Int64.Type},
            {"prazo_real_dias", Int64.Type},
            {"atraso_dias", Int64.Type},
            {"status_atraso", type text},
            {"status_pedido", type text},
            {"canal", type text},
            {"loja", type text},
            {"chave_canal_loja", type text},
            {"id_cliente", type text},
            {"cidade", type text},
            {"cep", type text},
            {"id_produto", type text},
            {"quantidade", Int64.Type},
            {"preco_unitario", Currency.Type},
            {"custo_unitario", Currency.Type},
            {"desconto_pct", type number},
            {"valor_bruto", Currency.Type},
            {"valor_liquido", Currency.Type},
            {"custo_ref_total", Currency.Type},
            {"custo_real_total", Currency.Type},
            {"frete_valor", Currency.Type},
            {"forma_pagamento", type text},
            {"parcelas", Int64.Type},
            {"transportadora", type text},
            {"avaliacao_cliente", Int64.Type},
            {"chave_clima", type text}
        },
        "en-US"
    )
in
    Tipos
