-- =============================================================
-- varejomax_conferencia.sql -- Validacao e Batimento K1 a K5
-- Projeto: VarejoMax Brasil (ILP590 -- Fatec Sao Paulo)
-- Regra de ouro: coluna "diferenca" / "ocorrencias" = 0 em K1, K2, K3 e K4.
-- =============================================================

-- === K1 ===
-- K1) Batimento de volumetria e totais financeiros: staging x tabelas 3FN / fato_vendas
SELECT 'Itens de pedido (linhas)' AS indicador,
       (SELECT COUNT(*) FROM stg_vendas) AS esperado_staging,
       (SELECT COUNT(*) FROM item_pedido) AS banco,
       (SELECT COUNT(*) FROM item_pedido) - (SELECT COUNT(*) FROM stg_vendas) AS diferenca
UNION ALL
SELECT 'Pedidos totais distintos',
       (SELECT COUNT(DISTINCT TRIM(ID_Pedido)) FROM stg_vendas),
       (SELECT COUNT(*) FROM pedido),
       (SELECT COUNT(*) FROM pedido) - (SELECT COUNT(DISTINCT TRIM(ID_Pedido)) FROM stg_vendas)
UNION ALL
SELECT 'Pedidos validos (sem Cancelado/Devolvido)',
       (SELECT COUNT(DISTINCT TRIM(ID_Pedido)) FROM stg_vendas WHERE TRIM(Status_Pedido) NOT IN ('Cancelado', 'Devolvido')),
       (SELECT COUNT(DISTINCT id_pedido) FROM fato_vendas WHERE status_pedido NOT IN ('Cancelado', 'Devolvido')),
       (SELECT COUNT(DISTINCT id_pedido) FROM fato_vendas WHERE status_pedido NOT IN ('Cancelado', 'Devolvido'))
       - (SELECT COUNT(DISTINCT TRIM(ID_Pedido)) FROM stg_vendas WHERE TRIM(Status_Pedido) NOT IN ('Cancelado', 'Devolvido'))
UNION ALL
SELECT 'Unidades vendidas (validos)',
       (SELECT SUM(CAST(TRIM(Quantidade) AS INTEGER)) FROM stg_vendas WHERE TRIM(Status_Pedido) NOT IN ('Cancelado', 'Devolvido')),
       (SELECT SUM(quantidade) FROM fato_vendas WHERE status_pedido NOT IN ('Cancelado', 'Devolvido')),
       (SELECT SUM(quantidade) FROM fato_vendas WHERE status_pedido NOT IN ('Cancelado', 'Devolvido'))
       - (SELECT SUM(CAST(TRIM(Quantidade) AS INTEGER)) FROM stg_vendas WHERE TRIM(Status_Pedido) NOT IN ('Cancelado', 'Devolvido'))
UNION ALL
SELECT 'Faturamento Liquido (validos)',
       (SELECT ROUND(SUM(CAST(REPLACE(TRIM(Valor_Liquido), ',', '.') AS REAL)), 2)
        FROM stg_vendas WHERE TRIM(Status_Pedido) NOT IN ('Cancelado', 'Devolvido')),
       (SELECT ROUND(SUM(valor_liquido), 2)
        FROM fato_vendas WHERE status_pedido NOT IN ('Cancelado', 'Devolvido')),
       ROUND(
           (SELECT SUM(valor_liquido) FROM fato_vendas WHERE status_pedido NOT IN ('Cancelado', 'Devolvido'))
           - (SELECT SUM(CAST(REPLACE(TRIM(Valor_Liquido), ',', '.') AS REAL))
              FROM stg_vendas WHERE TRIM(Status_Pedido) NOT IN ('Cancelado', 'Devolvido')),
           2
       )
UNION ALL
SELECT 'Faturamento Bruto (nao Cancelado)',
       (SELECT ROUND(SUM(CAST(REPLACE(TRIM(Valor_Bruto), ',', '.') AS REAL)), 2)
        FROM stg_vendas WHERE TRIM(Status_Pedido) <> 'Cancelado'),
       (SELECT ROUND(SUM(valor_bruto), 2)
        FROM fato_vendas WHERE status_pedido <> 'Cancelado'),
       ROUND(
           (SELECT SUM(valor_bruto) FROM fato_vendas WHERE status_pedido <> 'Cancelado')
           - (SELECT SUM(CAST(REPLACE(TRIM(Valor_Bruto), ',', '.') AS REAL))
              FROM stg_vendas WHERE TRIM(Status_Pedido) <> 'Cancelado'),
           2
       )
UNION ALL
SELECT 'Frete Total (validos)',
       (SELECT ROUND(SUM(CAST(REPLACE(TRIM(Frete_Valor), ',', '.') AS REAL)), 2)
        FROM stg_vendas WHERE TRIM(Status_Pedido) NOT IN ('Cancelado', 'Devolvido')),
       (SELECT ROUND(SUM(frete_valor), 2)
        FROM fato_vendas WHERE status_pedido NOT IN ('Cancelado', 'Devolvido')),
       ROUND(
           (SELECT SUM(frete_valor) FROM fato_vendas WHERE status_pedido NOT IN ('Cancelado', 'Devolvido'))
           - (SELECT SUM(CAST(REPLACE(TRIM(Frete_Valor), ',', '.') AS REAL))
              FROM stg_vendas WHERE TRIM(Status_Pedido) NOT IN ('Cancelado', 'Devolvido')),
           2
       )
UNION ALL
SELECT 'Produtos (dim_produto)',
       (SELECT COUNT(*) FROM stg_produto),
       (SELECT COUNT(*) FROM dim_produto),
       (SELECT COUNT(*) FROM dim_produto) - (SELECT COUNT(*) FROM stg_produto)
UNION ALL
SELECT 'Clientes (dim_cliente)',
       (SELECT COUNT(*) FROM stg_cliente),
       (SELECT COUNT(*) FROM dim_cliente),
       (SELECT COUNT(*) FROM dim_cliente) - (SELECT COUNT(*) FROM stg_cliente)
UNION ALL
SELECT 'Cidades (dim_cidade)',
       (SELECT COUNT(*) FROM stg_cidade),
       (SELECT COUNT(*) FROM dim_cidade),
       (SELECT COUNT(*) FROM dim_cidade) - (SELECT COUNT(*) FROM stg_cidade)
UNION ALL
SELECT 'Canais/Lojas (dim_canal_loja)',
       (SELECT COUNT(*) FROM stg_canal_loja),
       (SELECT COUNT(*) FROM dim_canal_loja),
       (SELECT COUNT(*) FROM dim_canal_loja) - (SELECT COUNT(*) FROM stg_canal_loja)
UNION ALL
SELECT 'Transportadoras (dim_transportadora)',
       (SELECT COUNT(*) FROM stg_transportadora),
       (SELECT COUNT(*) FROM dim_transportadora),
       (SELECT COUNT(*) FROM dim_transportadora) - (SELECT COUNT(*) FROM stg_transportadora)
UNION ALL
SELECT 'Calendario (dim_tempo)',
       608,
       (SELECT COUNT(*) FROM dim_tempo),
       (SELECT COUNT(*) FROM dim_tempo) - 608
UNION ALL
SELECT 'Cambio PTAX (dim_cambio)',
       (SELECT COUNT(*) FROM stg_cambio_ptax),
       (SELECT COUNT(*) FROM dim_cambio),
       (SELECT COUNT(*) FROM dim_cambio) - (SELECT COUNT(*) FROM stg_cambio_ptax)
UNION ALL
SELECT 'Clima 24 Cidades (dim_clima)',
       (SELECT COUNT(*) FROM stg_clima_cidade),
       (SELECT COUNT(*) FROM dim_clima),
       (SELECT COUNT(*) FROM dim_clima) - (SELECT COUNT(*) FROM stg_clima_cidade);

-- === K2 ===
-- K2) Orfaos de integridade referencial em todas as FKs (esperado = 0 em todas)
SELECT 'Itens sem pedido' AS teste_integridade, COUNT(*) AS orfaos
FROM item_pedido i LEFT JOIN pedido p ON p.id_pedido = i.id_pedido
WHERE p.id_pedido IS NULL
UNION ALL
SELECT 'Itens sem produto', COUNT(*)
FROM item_pedido i LEFT JOIN produto pr ON pr.id_produto = i.id_produto
WHERE pr.id_produto IS NULL
UNION ALL
SELECT 'Itens sem transportadora', COUNT(*)
FROM item_pedido i LEFT JOIN transportadora t ON t.transportadora = i.transportadora
WHERE t.transportadora IS NULL
UNION ALL
SELECT 'Pedidos sem item', COUNT(*)
FROM pedido p LEFT JOIN item_pedido i ON i.id_pedido = p.id_pedido
WHERE i.id_pedido IS NULL
UNION ALL
SELECT 'Pedidos sem cliente', COUNT(*)
FROM pedido p LEFT JOIN cliente c ON c.id_cliente = p.id_cliente
WHERE c.id_cliente IS NULL
UNION ALL
SELECT 'Pedidos sem canal_loja', COUNT(*)
FROM pedido p LEFT JOIN canal_loja cl ON cl.chave_canal_loja = p.chave_canal_loja
WHERE cl.chave_canal_loja IS NULL
UNION ALL
SELECT 'Pedidos sem cidade', COUNT(*)
FROM pedido p LEFT JOIN cidade cid ON cid.cidade = p.cidade
WHERE cid.cidade IS NULL
UNION ALL
SELECT 'Pedidos fora do calendario', COUNT(*)
FROM pedido p LEFT JOIN calendario cal ON cal.data = p.data_pedido
WHERE cal.data IS NULL
UNION ALL
SELECT 'Clientes (cidade nao nula) sem cidade', COUNT(*)
FROM cliente c LEFT JOIN cidade cid ON cid.cidade = c.cidade
WHERE c.cidade IS NOT NULL AND cid.cidade IS NULL
UNION ALL
SELECT 'Cambio fora do calendario', COUNT(*)
FROM cambio_ptax cb LEFT JOIN calendario cal ON cal.data = cb.data_ref
WHERE cal.data IS NULL
UNION ALL
SELECT 'Clima sem cidade', COUNT(*)
FROM clima_cidade cl LEFT JOIN cidade cid ON cid.cidade = cl.cidade
WHERE cid.cidade IS NULL
UNION ALL
SELECT 'Clima fora do calendario', COUNT(*)
FROM clima_cidade cl LEFT JOIN calendario cal ON cal.data = cl.data_ref
WHERE cal.data IS NULL;

-- === K3 ===
-- K3) Checagens de regras de dominio e negocio (esperado = 0 violacoes)
SELECT 'Quantidade <= 0' AS regra_negocio, COUNT(*) AS violacoes
FROM item_pedido WHERE quantidade <= 0
UNION ALL
SELECT 'Preco unitario <= 0 ou custo < 0', COUNT(*)
FROM item_pedido WHERE preco_unitario <= 0 OR custo_unitario < 0
UNION ALL
SELECT 'Desconto fora de [0, 1]', COUNT(*)
FROM item_pedido WHERE desconto_pct < 0 OR desconto_pct > 1
UNION ALL
SELECT 'Frete negativo ou presente em item <> 1 na fato_vendas', COUNT(*)
FROM fato_vendas WHERE frete_valor < 0 OR (item <> 1 AND frete_valor > 0)
UNION ALL
SELECT 'Calendario fim_de_semana/feriado fora de Sim/Nao', COUNT(*)
FROM dim_tempo WHERE fim_de_semana NOT IN ('Sim', 'Não') OR feriado NOT IN ('Sim', 'Não')
UNION ALL
SELECT 'Divergencia formula valor_bruto ou valor_liquido vs staging', COUNT(*)
FROM fato_vendas fv
JOIN stg_vendas sv ON TRIM(sv.ID_Pedido) = fv.id_pedido AND CAST(TRIM(sv.Item) AS INTEGER) = fv.item
WHERE ABS(fv.valor_bruto - CAST(REPLACE(TRIM(sv.Valor_Bruto), ',', '.') AS REAL)) > 0.001
   OR ABS(fv.valor_liquido - CAST(REPLACE(TRIM(sv.Valor_Liquido), ',', '.') AS REAL)) > 0.001
UNION ALL
SELECT 'Chave primaria duplicada em fato_vendas (id_pedido, item)', COUNT(*)
FROM (SELECT id_pedido, item FROM fato_vendas GROUP BY id_pedido, item HAVING COUNT(*) > 1);

-- === K4 ===
-- K4) Paridade temporal e cobertura 100% das APIs externas nas Views
SELECT 'fato_vendas sem match em dim_tempo (data_pedido)' AS cobertura_view, COUNT(*) AS sem_match
FROM fato_vendas fv LEFT JOIN dim_tempo dt ON dt.data = fv.data_pedido
WHERE dt.data IS NULL
UNION ALL
SELECT 'fato_vendas sem match em dim_cambio (data_pedido)', COUNT(*)
FROM fato_vendas fv LEFT JOIN dim_cambio dc ON dc.data = fv.data_pedido
WHERE dc.data IS NULL
UNION ALL
SELECT 'fato_vendas sem match em dim_clima (chave_clima)', COUNT(*)
FROM fato_vendas fv LEFT JOIN dim_clima dcl ON dcl.chave_clima = fv.chave_clima
WHERE dcl.chave_clima IS NULL
UNION ALL
SELECT 'dim_tempo sem match em dim_cambio', COUNT(*)
FROM dim_tempo dt LEFT JOIN dim_cambio dc ON dc.data = dt.data
WHERE dc.data IS NULL
UNION ALL
SELECT 'dim_tempo x dim_cidade sem match em dim_clima (14.592 combinacoes)', COUNT(*)
FROM dim_tempo dt
CROSS JOIN dim_cidade dc
LEFT JOIN dim_clima dcl ON dcl.data = dt.data AND dcl.cidade = dc.cidade
WHERE dcl.chave_clima IS NULL;

-- === K5 ===
-- K5) KPIs Executivos Oficiais (Reproducao exata da Etapa 3)
WITH base_valida AS (
    SELECT fv.*, dcl.tipo_canal
    FROM fato_vendas fv
    JOIN dim_canal_loja dcl ON dcl.chave_canal_loja = fv.chave_canal_loja
    WHERE fv.status_pedido NOT IN ('Cancelado', 'Devolvido')
),
pedidos_entregues AS (
    SELECT id_pedido,
           MAX(CASE WHEN status_atraso = 'Atrasado' THEN 1 ELSE 0 END) AS teve_atraso,
           AVG(CASE WHEN avaliacao_cliente BETWEEN 1 AND 5 THEN avaliacao_cliente END) AS nota_pedido
    FROM fato_vendas
    WHERE status_pedido = 'Entregue'
    GROUP BY id_pedido
),
pedidos_status AS (
    SELECT id_pedido, MAX(status_pedido) AS status_pedido
    FROM fato_vendas
    GROUP BY id_pedido
)
SELECT
    (SELECT COUNT(*) FROM pedidos_status) AS qtd_pedidos_total,
    COUNT(DISTINCT id_pedido) AS qtd_pedidos_validos,
    SUM(quantidade) AS qtd_unidades_validas,
    ROUND((SELECT SUM(valor_bruto) FROM fato_vendas WHERE status_pedido <> 'Cancelado'), 2) AS fat_bruto_nao_canc,
    ROUND(SUM(valor_liquido), 2) AS fat_liquido_valido,
    ROUND(SUM(valor_liquido) / COUNT(DISTINCT id_pedido), 2) AS ticket_medio,
    ROUND(1.0 * SUM(quantidade) / COUNT(DISTINCT id_pedido), 2) AS itens_por_pedido,
    ROUND(SUM(custo_ref_total), 2) AS custo_ref_total,
    ROUND(100.0 * (SUM(valor_liquido) - SUM(custo_ref_total)) / SUM(valor_liquido), 2) AS margem_bruta_ref_pct,
    ROUND(SUM(custo_real_total), 2) AS custo_real_total,
    ROUND(100.0 * (SUM(valor_liquido) - SUM(custo_real_total)) / SUM(valor_liquido), 2) AS margem_real_pct,
    ROUND(SUM(frete_valor), 2) AS frete_total_valido,
    ROUND(100.0 * SUM(CASE WHEN tipo_canal = 'Online (e-commerce)' THEN valor_liquido ELSE 0 END)
          / SUM(valor_liquido), 2) AS pct_vendas_online,
    (SELECT COUNT(*) FROM pedidos_entregues) AS qtd_pedidos_entregues,
    (SELECT SUM(teve_atraso) FROM pedidos_entregues) AS qtd_pedidos_atrasados,
    ROUND(100.0 * (SELECT SUM(teve_atraso) FROM pedidos_entregues)
          / (SELECT COUNT(*) FROM pedidos_entregues), 2) AS pct_atraso_entregues,
    ROUND(100.0 * (SELECT COUNT(*) FROM pedidos_status WHERE status_pedido = 'Cancelado')
          / (SELECT COUNT(*) FROM pedidos_status), 2) AS pct_cancelados,
    ROUND(100.0 * (SELECT COUNT(*) FROM pedidos_status WHERE status_pedido = 'Devolvido')
          / (SELECT COUNT(*) FROM pedidos_status), 2) AS pct_devolvidos,
    ROUND(AVG(CASE WHEN avaliacao_cliente BETWEEN 1 AND 5 THEN avaliacao_cliente END), 2) AS nota_media_1_a_5,
    ROUND((SELECT AVG(CASE WHEN avaliacao_cliente BETWEEN 1 AND 5 THEN avaliacao_cliente END) FROM fato_vendas), 2) AS nota_media_item_todas,
    ROUND(AVG(CASE WHEN avaliacao_cliente BETWEEN 1 AND 5 THEN avaliacao_cliente END), 2) AS nota_media_item_valida,
    ROUND((SELECT AVG(nota_pedido) FROM pedidos_entregues), 2) AS nota_media_pedido_entregue
FROM base_valida;
