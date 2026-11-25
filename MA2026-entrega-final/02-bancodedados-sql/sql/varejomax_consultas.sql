-- =============================================================
-- varejomax_consultas.sql -- 10 Consultas Analiticas (C1 a C10)
-- Projeto: VarejoMax Brasil (ILP590 -- Fatec Sao Paulo)
-- Responde as 4 hipoteses e perguntas de negocio usando Views Star Schema.
-- =============================================================

-- === C1 ===
-- C1) Evolucao mensal: Faturamento Bruto, Faturamento Liquido, Pedidos Validos, Ticket Medio e MoM % (LAG)
WITH mensal AS (
    SELECT
        dt.ano_mes,
        dt.ano,
        dt.mes_abrev,
        COUNT(DISTINCT CASE WHEN fv.status_pedido <> 'Cancelado' THEN fv.id_pedido END) AS pedidos_nao_canc,
        COUNT(DISTINCT CASE WHEN fv.status_pedido NOT IN ('Cancelado', 'Devolvido') THEN fv.id_pedido END) AS pedidos_validos,
        SUM(CASE WHEN fv.status_pedido NOT IN ('Cancelado', 'Devolvido') THEN fv.quantidade ELSE 0 END) AS unidades_validas,
        ROUND(SUM(CASE WHEN fv.status_pedido <> 'Cancelado' THEN fv.valor_bruto ELSE 0 END), 2) AS fat_bruto,
        ROUND(SUM(CASE WHEN fv.status_pedido NOT IN ('Cancelado', 'Devolvido') THEN fv.valor_liquido ELSE 0 END), 2) AS fat_liquido,
        ROUND(SUM(CASE WHEN fv.status_pedido NOT IN ('Cancelado', 'Devolvido') THEN fv.custo_real_total ELSE 0 END), 2) AS custo_real
    FROM fato_vendas fv
    JOIN dim_tempo dt ON dt.data = fv.data_pedido
    GROUP BY dt.ano_mes, dt.ano, dt.mes_abrev
)
SELECT
    ano_mes,
    mes_abrev || '/' || SUBSTR(CAST(ano AS TEXT), 3, 2) AS periodo,
    pedidos_validos,
    unidades_validas,
    fat_bruto,
    fat_liquido,
    ROUND(1.0 * fat_liquido / pedidos_validos, 2) AS ticket_medio,
    ROUND(100.0 * (fat_liquido - custo_real) / fat_liquido, 2) AS margem_real_pct,
    ROUND(LAG(fat_liquido) OVER (ORDER BY ano_mes), 2) AS fat_mes_anterior,
    ROUND(100.0 * (fat_liquido - LAG(fat_liquido) OVER (ORDER BY ano_mes))
          / NULLIF(LAG(fat_liquido) OVER (ORDER BY ano_mes), 0), 2) AS variacao_mom_pct,
    CASE WHEN ano_mes = '2025-11' THEN 'Pico Black Friday'
         WHEN ano_mes = '2025-12' THEN 'Pico Natal'
         ELSE 'Regular' END AS destaque_sazonal
FROM mensal
ORDER BY ano_mes;

-- === C2 ===
-- C2) Desempenho por Categoria: Faturamento, Margem Ref. vs Realizada e % Participacao (SUM OVER)
WITH por_cat AS (
    SELECT
        dp.categoria,
        COUNT(DISTINCT fv.id_pedido) AS pedidos_validos,
        SUM(fv.quantidade) AS unidades_vendidas,
        ROUND(SUM(fv.valor_bruto), 2) AS fat_bruto,
        ROUND(SUM(fv.valor_liquido), 2) AS fat_liquido,
        ROUND(SUM(fv.custo_ref_total), 2) AS custo_ref,
        ROUND(SUM(fv.custo_real_total), 2) AS custo_real
    FROM fato_vendas fv
    JOIN dim_produto dp ON dp.id_produto = fv.id_produto
    WHERE fv.status_pedido NOT IN ('Cancelado', 'Devolvido')
    GROUP BY dp.categoria
)
SELECT
    RANK() OVER (ORDER BY fat_liquido DESC) AS rank_fat,
    categoria,
    pedidos_validos,
    unidades_vendidas,
    fat_bruto,
    fat_liquido,
    ROUND(100.0 * fat_liquido / SUM(fat_liquido) OVER (), 2) AS participacao_pct,
    ROUND(100.0 * (fat_liquido - custo_ref) / fat_liquido, 2) AS margem_ref_pct,
    ROUND(100.0 * (fat_liquido - custo_real) / fat_liquido, 2) AS margem_real_pct,
    ROUND(100.0 * (fat_liquido - custo_real) / fat_liquido
          - 100.0 * (fat_liquido - custo_ref) / fat_liquido, 2) AS delta_margem_pp
FROM por_cat
ORDER BY fat_liquido DESC;

-- === C3 ===
-- C3) Canais de Venda e Lojas: Online vs Fisica, Participacao %, Ticket Medio e Nota Media (1..5)
WITH por_canal AS (
    SELECT
        dcl.tipo_canal,
        dcl.canal,
        dcl.loja,
        COUNT(DISTINCT fv.id_pedido) AS pedidos_validos,
        SUM(fv.quantidade) AS unidades_vendidas,
        ROUND(SUM(fv.valor_liquido), 2) AS fat_liquido,
        ROUND(SUM(fv.custo_real_total), 2) AS custo_real,
        ROUND(AVG(CASE WHEN fv.avaliacao_cliente BETWEEN 1 AND 5 THEN fv.avaliacao_cliente END), 2) AS nota_media
    FROM fato_vendas fv
    JOIN dim_canal_loja dcl ON dcl.chave_canal_loja = fv.chave_canal_loja
    WHERE fv.status_pedido NOT IN ('Cancelado', 'Devolvido')
    GROUP BY dcl.tipo_canal, dcl.canal, dcl.loja
)
SELECT
    tipo_canal,
    canal,
    loja,
    pedidos_validos,
    unidades_vendidas,
    fat_liquido,
    ROUND(100.0 * fat_liquido / SUM(fat_liquido) OVER (), 2) AS participacao_rede_pct,
    ROUND(1.0 * fat_liquido / pedidos_validos, 2) AS ticket_medio,
    ROUND(100.0 * (fat_liquido - custo_real) / fat_liquido, 2) AS margem_real_pct,
    nota_media
FROM por_canal
ORDER BY fat_liquido DESC;

-- === C4 ===
-- C4) Performance por Regiao e Top 10 Cidades (RANK), comparando cidades com vs sem loja fisica
WITH cidade_vendas AS (
    SELECT
        dc.cidade,
        dc.uf,
        dc.regiao,
        dc.tem_loja_fisica,
        COUNT(DISTINCT CASE WHEN fv.status_pedido NOT IN ('Cancelado', 'Devolvido') THEN fv.id_pedido END) AS pedidos_validos,
        ROUND(SUM(CASE WHEN fv.status_pedido NOT IN ('Cancelado', 'Devolvido') THEN fv.valor_liquido ELSE 0 END), 2) AS fat_liquido,
        COUNT(DISTINCT CASE WHEN fv.status_pedido = 'Entregue' THEN fv.id_pedido END) AS pedidos_entregues,
        COUNT(DISTINCT CASE WHEN fv.status_pedido = 'Entregue' AND fv.status_atraso = 'Atrasado' THEN fv.id_pedido END) AS pedidos_atrasados
    FROM fato_vendas fv
    JOIN dim_cidade dc ON dc.cidade = fv.cidade
    GROUP BY dc.cidade, dc.uf, dc.regiao, dc.tem_loja_fisica
),
ranking AS (
    SELECT
        RANK() OVER (ORDER BY fat_liquido DESC) AS rank_nacional,
        cidade,
        uf,
        regiao,
        tem_loja_fisica,
        pedidos_validos,
        fat_liquido,
        ROUND(100.0 * fat_liquido / SUM(fat_liquido) OVER (), 2) AS participacao_pct,
        ROUND(1.0 * fat_liquido / pedidos_validos, 2) AS ticket_medio,
        pedidos_entregues,
        pedidos_atrasados,
        ROUND(100.0 * pedidos_atrasados / NULLIF(pedidos_entregues, 0), 2) AS pct_atraso
    FROM cidade_vendas
)
SELECT *
FROM ranking
WHERE rank_nacional <= 10
ORDER BY rank_nacional;

-- === C5 ===
-- C5) Auditoria Logistica por Transportadora (1 linha por ID_Pedido Entregue, 100% aderente ao VBA) + Nota Media (Hipoteses 1 e 3)
WITH entrega_pedido AS (
    SELECT
        fv.id_pedido,
        COALESCE(
            MAX(CASE WHEN fv.transportadora <> 'Não informado' THEN fv.transportadora END),
            'Não informado'
        ) AS transportadora,
        MAX(CASE WHEN fv.transportadora <> 'Retirada em loja' THEN fv.prazo_previsto_dias END) AS prazo_previsto_dias,
        AVG(CASE WHEN fv.transportadora <> 'Retirada em loja'
                  AND fv.status_atraso IN ('Adiantado', 'No prazo', 'Atrasado')
                 THEN fv.prazo_real_dias END) AS prazo_real_dias,
        MAX(CASE WHEN fv.status_atraso IN ('Adiantado', 'No prazo', 'Atrasado') THEN fv.atraso_dias END) AS atraso_dias_geral,
        MAX(CASE WHEN fv.status_atraso = 'Atrasado' THEN fv.atraso_dias END) AS atraso_dias,
        MAX(CASE WHEN fv.status_atraso = 'Atrasado' THEN 1 ELSE 0 END) AS flag_atraso,
        AVG(CASE WHEN fv.avaliacao_cliente BETWEEN 1 AND 5 THEN fv.avaliacao_cliente END) AS nota_pedido
    FROM fato_vendas fv
    WHERE fv.status_pedido = 'Entregue'
    GROUP BY fv.id_pedido
)
SELECT
    RANK() OVER (ORDER BY 1.0 * SUM(ep.flag_atraso) / COUNT(*) DESC) AS rank_maior_atraso,
    ep.transportadora,
    dt.tipo AS tipo_transp,
    COUNT(*) AS entregas_consolidadas,
    SUM(ep.flag_atraso) AS entregas_atrasadas,
    ROUND(100.0 * SUM(ep.flag_atraso) / COUNT(*), 2) AS pct_atraso,
    ROUND(AVG(ep.prazo_previsto_dias), 2) AS prazo_previsto_medio,
    ROUND(AVG(ep.prazo_real_dias), 2) AS prazo_real_medio,
    ROUND(AVG(ep.atraso_dias_geral), 2) AS desvio_medio_dias,
    ROUND(AVG(CASE WHEN ep.flag_atraso = 1 THEN ep.atraso_dias END), 2) AS atraso_medio_quando_atrasa,
    ROUND(AVG(ep.nota_pedido), 2) AS nota_media_geral,
    ROUND(AVG(CASE WHEN ep.flag_atraso = 1 THEN ep.nota_pedido END), 2) AS nota_media_com_atraso,
    ROUND(AVG(CASE WHEN ep.flag_atraso = 0 THEN ep.nota_pedido END), 2) AS nota_media_sem_atraso
FROM entrega_pedido ep
JOIN dim_transportadora dt ON dt.transportadora = ep.transportadora
GROUP BY ep.transportadora, dt.tipo
ORDER BY pct_atraso DESC, entregas_consolidadas DESC;

-- === C6 ===
-- C6) Excecoes e Pedidos Criticos em Atraso (Atraso >= 5 dias) com detalhamento geografico e transportadora (Hipotese 2)
WITH pedidos_criticos AS (
    SELECT
        fv.id_pedido,
        fv.data_pedido,
        fv.data_entrega_prevista,
        MAX(fv.data_entrega_real) AS data_entrega_real,
        dc.regiao,
        dc.uf,
        dc.cidade,
        GROUP_CONCAT(DISTINCT fv.transportadora) AS transportadoras,
        MAX(fv.atraso_dias) AS max_atraso_dias,
        ROUND(SUM(fv.valor_liquido), 2) AS valor_liquido_pedido,
        ROUND(AVG(CASE WHEN fv.avaliacao_cliente BETWEEN 1 AND 5 THEN fv.avaliacao_cliente END), 2) AS nota_cliente
    FROM fato_vendas fv
    JOIN dim_cidade dc ON dc.cidade = fv.cidade
    WHERE fv.status_pedido = 'Entregue'
      AND fv.status_atraso = 'Atrasado'
      AND fv.atraso_dias >= 5
    GROUP BY fv.id_pedido, fv.data_pedido, fv.data_entrega_prevista, dc.regiao, dc.uf, dc.cidade
)
SELECT
    RANK() OVER (ORDER BY max_atraso_dias DESC, valor_liquido_pedido DESC) AS rank_criticidade,
    id_pedido,
    data_pedido,
    data_entrega_prevista,
    data_entrega_real,
    regiao,
    uf,
    cidade,
    transportadoras,
    max_atraso_dias,
    valor_liquido_pedido,
    nota_cliente
FROM pedidos_criticos
ORDER BY rank_criticidade, id_pedido
LIMIT 25;

-- === C7 ===
-- C7) Segmentacao RFV / Concentracao de Clientes (Frequencia e Valor com Pareto / SUM OVER)
WITH base_cliente AS (
    SELECT
        fv.id_cliente,
        COUNT(DISTINCT fv.id_pedido) AS freq_pedidos,
        CAST(JULIANDAY('2026-08-31') - JULIANDAY(MAX(fv.data_pedido)) AS INTEGER) AS recencia_dias,
        ROUND(SUM(fv.valor_liquido), 2) AS valor_total
    FROM fato_vendas fv
    WHERE fv.status_pedido NOT IN ('Cancelado', 'Devolvido')
    GROUP BY fv.id_cliente
),
segmentos AS (
    SELECT
        CASE WHEN freq_pedidos = 1 THEN '1-Novo (1 pedido)'
             WHEN freq_pedidos BETWEEN 2 AND 4 THEN '2-Recorrente (2 a 4 pedidos)'
             WHEN freq_pedidos BETWEEN 5 AND 7 THEN '3-Fiel (5 a 7 pedidos)'
             ELSE '4-VIP (8+ pedidos)' END AS segmento_rfv,
        id_cliente,
        freq_pedidos,
        recencia_dias,
        valor_total
    FROM base_cliente
)
SELECT
    segmento_rfv,
    COUNT(*) AS qtd_clientes,
    ROUND(100.0 * COUNT(*) / SUM(COUNT(*)) OVER (), 2) AS pct_clientes,
    SUM(freq_pedidos) AS total_pedidos,
    ROUND(AVG(freq_pedidos), 2) AS freq_media,
    ROUND(AVG(recencia_dias), 1) AS recencia_media_dias,
    ROUND(SUM(valor_total), 2) AS faturamento_liquido,
    ROUND(100.0 * SUM(valor_total) / SUM(SUM(valor_total)) OVER (), 2) AS pct_faturamento,
    ROUND(SUM(valor_total) / COUNT(*), 2) AS ltv_medio_cliente,
    ROUND(SUM(valor_total) / SUM(freq_pedidos), 2) AS ticket_medio_pedido
FROM segmentos
GROUP BY segmento_rfv
ORDER BY segmento_rfv;

-- === C8 ===
-- C8) Impacto de Descontos, Devolucoes e Cancelamentos por Categoria e Tipo de Canal
SELECT
    dp.categoria,
    dcl.tipo_canal,
    COUNT(DISTINCT fv.id_pedido) AS pedidos_totais,
    COUNT(DISTINCT CASE WHEN fv.status_pedido = 'Cancelado' THEN fv.id_pedido END) AS pedidos_cancelados,
    COUNT(DISTINCT CASE WHEN fv.status_pedido = 'Devolvido' THEN fv.id_pedido END) AS pedidos_devolvidos,
    ROUND(100.0 * COUNT(DISTINCT CASE WHEN fv.status_pedido = 'Cancelado' THEN fv.id_pedido END)
          / COUNT(DISTINCT fv.id_pedido), 2) AS pct_cancelamento,
    ROUND(100.0 * COUNT(DISTINCT CASE WHEN fv.status_pedido = 'Devolvido' THEN fv.id_pedido END)
          / COUNT(DISTINCT fv.id_pedido), 2) AS pct_devolucao,
    ROUND(SUM(CASE WHEN fv.status_pedido IN ('Cancelado', 'Devolvido') THEN fv.valor_liquido ELSE 0 END), 2)
        AS receita_perdida_canc_dev,
    ROUND(100.0 * AVG(CASE WHEN fv.status_pedido NOT IN ('Cancelado', 'Devolvido') THEN fv.desconto_pct END), 2)
        AS desconto_medio_valido_pct,
    ROUND(SUM(CASE WHEN fv.status_pedido NOT IN ('Cancelado', 'Devolvido') THEN fv.valor_bruto ELSE 0 END), 2)
        AS fat_bruto_valido,
    ROUND(SUM(CASE WHEN fv.status_pedido NOT IN ('Cancelado', 'Devolvido') THEN fv.valor_liquido ELSE 0 END), 2)
        AS fat_liquido_valido,
    ROUND(100.0 * (
            SUM(CASE WHEN fv.status_pedido NOT IN ('Cancelado', 'Devolvido') THEN fv.valor_liquido - fv.custo_real_total ELSE 0 END)
          ) / NULLIF(SUM(CASE WHEN fv.status_pedido NOT IN ('Cancelado', 'Devolvido') THEN fv.valor_liquido ELSE 0 END), 0), 2)
        AS margem_real_pct
FROM fato_vendas fv
JOIN dim_produto dp ON dp.id_produto = fv.id_produto
JOIN dim_canal_loja dcl ON dcl.chave_canal_loja = fv.chave_canal_loja
GROUP BY dp.categoria, dcl.tipo_canal
ORDER BY dp.categoria, dcl.tipo_canal;

-- === C9 ===
-- C9) Analise Externa 1: Dolar PTAX x Faturamento e Margem de Eletronicos vs Demais Categorias
WITH vendas_cambio AS (
    SELECT
        dc.faixa_dolar,
        CASE WHEN dp.categoria = 'Eletrônicos' THEN 'Eletrônicos' ELSE 'Demais Categorias' END AS grupo_categoria,
        fv.data_pedido,
        fv.id_pedido,
        fv.quantidade,
        dc.cotacao_venda,
        fv.valor_liquido,
        fv.custo_ref_total,
        fv.custo_real_total
    FROM fato_vendas fv
    JOIN dim_produto dp ON dp.id_produto = fv.id_produto
    JOIN dim_cambio dc ON dc.data = fv.data_pedido
    WHERE fv.status_pedido NOT IN ('Cancelado', 'Devolvido')
)
SELECT
    faixa_dolar,
    grupo_categoria,
    COUNT(DISTINCT data_pedido) AS dias_com_venda,
    COUNT(DISTINCT id_pedido) AS pedidos_validos,
    SUM(quantidade) AS unidades_vendidas,
    ROUND(AVG(cotacao_venda), 4) AS ptax_venda_medio,
    ROUND(SUM(valor_liquido), 2) AS fat_liquido,
    ROUND(1.0 * SUM(valor_liquido) / COUNT(DISTINCT id_pedido), 2) AS ticket_medio,
    ROUND(100.0 * (SUM(valor_liquido) - SUM(custo_ref_total)) / SUM(valor_liquido), 2) AS margem_ref_pct,
    ROUND(100.0 * (SUM(valor_liquido) - SUM(custo_real_total)) / SUM(valor_liquido), 2) AS margem_real_pct,
    ROUND(100.0 * (SUM(valor_liquido) - SUM(custo_real_total)) / SUM(valor_liquido)
          - 100.0 * (SUM(valor_liquido) - SUM(custo_ref_total)) / SUM(valor_liquido), 2) AS compressao_margem_pp
FROM vendas_cambio
GROUP BY faixa_dolar, grupo_categoria
ORDER BY faixa_dolar, grupo_categoria;

-- === C10 ===
-- C10) Analise Externa 2: Clima (faixa_precipitacao) e Calendario (Feriado/Fim de Semana) x SLA e Vendas
WITH clima_entrega AS (
    SELECT
        fv.id_pedido,
        dcl.faixa_precipitacao,
        CASE WHEN dt.feriado = 'Sim' THEN 'Feriado Nacional'
             WHEN dt.fim_de_semana = 'Sim' THEN 'Fim de Semana'
             ELSE 'Dia Útil' END AS tipo_dia_pedido,
        MAX(CASE WHEN fv.status_pedido = 'Entregue' AND fv.status_atraso IN ('Adiantado', 'No prazo', 'Atrasado') THEN 1 ELSE 0 END) AS eh_entregue_valido,
        MAX(CASE WHEN fv.status_pedido = 'Entregue' AND fv.status_atraso = 'Atrasado' THEN 1 ELSE 0 END) AS eh_atrasado,
        AVG(CASE WHEN fv.status_pedido = 'Entregue' AND fv.status_atraso IN ('Adiantado', 'No prazo', 'Atrasado') THEN fv.atraso_dias END) AS atraso_dias_pedido,
        MAX(CASE WHEN fv.status_pedido NOT IN ('Cancelado', 'Devolvido') THEN 1 ELSE 0 END) AS eh_pedido_valido,
        SUM(CASE WHEN fv.status_pedido NOT IN ('Cancelado', 'Devolvido') THEN fv.valor_liquido ELSE 0 END) AS valor_liquido_pedido,
        AVG(dcl.precipitacao) AS precipitacao_mm,
        AVG(dcl.temp_max) AS temp_max_media
    FROM fato_vendas fv
    JOIN dim_clima dcl
      ON dcl.cidade = fv.cidade
     AND dcl.data   = COALESCE(fv.data_entrega_real, fv.data_pedido)
    JOIN dim_tempo dt ON dt.data = fv.data_pedido
    GROUP BY fv.id_pedido, dcl.faixa_precipitacao,
             CASE WHEN dt.feriado = 'Sim' THEN 'Feriado Nacional'
                  WHEN dt.fim_de_semana = 'Sim' THEN 'Fim de Semana'
                  ELSE 'Dia Útil' END
)
SELECT
    faixa_precipitacao,
    tipo_dia_pedido,
    SUM(eh_pedido_valido) AS pedidos_validos,
    ROUND(SUM(valor_liquido_pedido), 2) AS fat_liquido,
    SUM(eh_entregue_valido) AS pedidos_entregues,
    SUM(eh_atrasado) AS pedidos_atrasados,
    ROUND(100.0 * SUM(eh_atrasado) / NULLIF(SUM(eh_entregue_valido), 0), 2) AS pct_atraso,
    ROUND(AVG(atraso_dias_pedido), 2) AS desvio_medio_dias,
    ROUND(AVG(precipitacao_mm), 2) AS precipitacao_media_mm,
    ROUND(AVG(temp_max_media), 1) AS temp_max_media_c
FROM clima_entrega
GROUP BY faixa_precipitacao, tipo_dia_pedido
ORDER BY faixa_precipitacao, tipo_dia_pedido;
