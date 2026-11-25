-- =============================================================
-- varejomax_diagnostico.sql -- Diagnostico de Anomalias D1 a D4
-- Projeto: VarejoMax Brasil (ILP590 -- Fatec Sao Paulo)
-- Documenta os casos de borda saneados entre staging e modelo 3FN.
-- =============================================================

-- === D1 ===
-- D1) 358 pedidos com variacao entre itens nos 5 atributos mantidos em item_pedido
WITH por_pedido AS (
    SELECT
        TRIM(ID_Pedido) AS id_pedido,
        COUNT(*) AS qtd_itens_linha,
        CASE WHEN COUNT(DISTINCT COALESCE(TRIM(CEP), '')) > 1 THEN 1 ELSE 0 END AS var_cep,
        CASE WHEN COUNT(DISTINCT TRIM(Transportadora)) > 1 THEN 1 ELSE 0 END AS var_transp,
        CASE WHEN COUNT(DISTINCT TRIM(Forma_Pagamento)) > 1 THEN 1 ELSE 0 END AS var_pagto,
        CASE WHEN COUNT(DISTINCT COALESCE(TRIM(Data_Entrega_Real), '')) > 1 THEN 1 ELSE 0 END AS var_entrega,
        CASE WHEN COUNT(DISTINCT COALESCE(TRIM(Avaliacao_Cliente), '')) > 1 THEN 1 ELSE 0 END AS var_avaliacao
    FROM stg_vendas
    GROUP BY TRIM(ID_Pedido)
)
SELECT
    SUM(var_cep)       AS pedidos_var_cep,
    SUM(var_transp)    AS pedidos_var_transportadora,
    SUM(var_pagto)     AS pedidos_var_forma_pagamento,
    SUM(var_entrega)   AS pedidos_var_data_entrega_real,
    SUM(var_avaliacao) AS pedidos_var_avaliacao_cliente,
    SUM(CASE WHEN (var_cep + var_transp + var_pagto + var_entrega + var_avaliacao) > 0 THEN 1 ELSE 0 END)
                       AS total_pedidos_distintos_com_variacao
FROM por_pedido;

-- === D2 ===
-- D2) Anomalias de qualidade preservadas da Etapa 2 (avaliacao fora 1..5 e entregas Inconsistente)
SELECT
    'Avaliacao fora de 1..5 (notas 0 e 6..10)' AS anomalia,
    COUNT(*) AS ocorrencias,
    GROUP_CONCAT(id_pedido || '#item' || item || '=' || avaliacao_cliente, ', ') AS exemplos
FROM item_pedido
WHERE avaliacao_cliente IS NOT NULL AND (avaliacao_cliente < 1 OR avaliacao_cliente > 5)
UNION ALL
SELECT
    'Entrega Inconsistente (data_entrega_real < data_envio)',
    COUNT(*),
    SUBSTR(GROUP_CONCAT(id_pedido || '#item' || item, ', '), 1, 140)
FROM item_pedido
WHERE status_atraso = 'Inconsistente';

-- === D3 ===
-- D3) Campos nulos / ausentes e regra de frete no item = 1
SELECT 'Clientes sem cidade (NULL em cliente.cidade)' AS diagnostico, COUNT(*) AS ocorrencias
FROM cliente WHERE cidade IS NULL
UNION ALL
SELECT 'Clientes sem e-mail (NULL em cliente.email)', COUNT(*)
FROM cliente WHERE email IS NULL
UNION ALL
SELECT 'Itens sem CEP (lojas fisicas / cancelados)', COUNT(*)
FROM item_pedido WHERE cep IS NULL
UNION ALL
SELECT 'Pedidos sem data_envio (Cancelado / Em separacao)', COUNT(*)
FROM pedido WHERE data_envio IS NULL
UNION ALL
SELECT 'Itens sem data_entrega_real', COUNT(*)
FROM item_pedido WHERE data_entrega_real IS NULL
UNION ALL
SELECT 'Itens sem avaliacao_cliente', COUNT(*)
FROM item_pedido WHERE avaliacao_cliente IS NULL
UNION ALL
SELECT 'Itens com Frete_Valor > 0 em Item = 1 (staging)', COUNT(*)
FROM stg_vendas
WHERE CAST(TRIM(Item) AS INTEGER) = 1
  AND CAST(REPLACE(TRIM(Frete_Valor), ',', '.') AS REAL) > 0
UNION ALL
SELECT 'Itens com Frete_Valor > 0 em Item <> 1 (staging)', COUNT(*)
FROM stg_vendas
WHERE CAST(TRIM(Item) AS INTEGER) <> 1
  AND CAST(REPLACE(TRIM(Frete_Valor), ',', '.') AS REAL) > 0;

-- === D4 ===
-- D4) Cobertura temporal (justificativa da ausencia de FK em data_entrega_prevista)
SELECT
    MIN(data_pedido) AS min_data_pedido,
    MAX(data_pedido) AS max_data_pedido,
    MIN(data_entrega_prevista) AS min_entrega_prevista,
    MAX(data_entrega_prevista) AS max_entrega_prevista,
    SUM(CASE WHEN data_entrega_prevista > (SELECT MAX(data) FROM calendario) THEN 1 ELSE 0 END)
        AS pedidos_prevista_apos_calendario,
    MIN(data_entrega_real) AS min_entrega_real,
    MAX(data_entrega_real) AS max_entrega_real
FROM pedido p
LEFT JOIN item_pedido i ON i.id_pedido = p.id_pedido;
