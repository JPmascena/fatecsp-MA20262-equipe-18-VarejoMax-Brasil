-- =============================================================
-- varejomax_views.sql -- Camada Analitica Star Schema (9 Views)
-- Projeto: VarejoMax Brasil (ILP590 -- Fatec Sao Paulo)
-- Contrato 100% aderente aos scripts .pq de 04-powerbi-dashboard/m/
-- =============================================================

DROP VIEW IF EXISTS fato_vendas;
DROP VIEW IF EXISTS dim_clima;
DROP VIEW IF EXISTS dim_cambio;
DROP VIEW IF EXISTS dim_tempo;
DROP VIEW IF EXISTS dim_transportadora;
DROP VIEW IF EXISTS dim_canal_loja;
DROP VIEW IF EXISTS dim_cidade;
DROP VIEW IF EXISTS dim_cliente;
DROP VIEW IF EXISTS dim_produto;

-- 1) dim_produto (49 linhas)
CREATE VIEW dim_produto AS
SELECT
    id_produto,
    produto,
    categoria,
    marca,
    preco_ref_2025,
    preco_ref_2026,
    custo_ref
FROM produto;

-- 2) dim_cliente (1.607 linhas; LEFT JOIN preserva cadastro sem cidade)
CREATE VIEW dim_cliente AS
SELECT
    c.id_cliente,
    c.nome_cliente,
    c.email,
    c.data_cadastro,
    c.cidade,
    cid.uf,
    cid.regiao
FROM cliente c
LEFT JOIN cidade cid ON cid.cidade = c.cidade;

-- 3) dim_cidade (24 linhas)
CREATE VIEW dim_cidade AS
SELECT
    cidade,
    uf,
    regiao,
    latitude,
    longitude,
    tem_loja_fisica
FROM cidade;

-- 4) dim_canal_loja (10 linhas)
CREATE VIEW dim_canal_loja AS
SELECT
    chave_canal_loja,
    canal,
    loja,
    tipo_canal,
    cod_loja,
    cidade_loja,
    uf_loja
FROM canal_loja;

-- 5) dim_transportadora (6 linhas)
CREATE VIEW dim_transportadora AS
SELECT
    transportadora,
    tipo
FROM transportadora;

-- 6) dim_tempo (608 linhas: 2025-01-01 a 2026-08-31)
CREATE VIEW dim_tempo AS
SELECT
    data,
    ano,
    mes,
    nome_mes,
    mes_abrev,
    ano_mes,
    mes_inicio,
    trimestre,
    dia_semana,
    nome_dia,
    fim_de_semana,
    feriado,
    nome_feriado
FROM calendario;

-- 7) dim_cambio (608 linhas: PTAX BCB com carry-forward)
CREATE VIEW dim_cambio AS
SELECT
    data_ref AS data,
    cotacao_compra,
    cotacao_venda,
    boletim_origem,
    faixa_dolar
FROM cambio_ptax;

-- 8) dim_clima (14.592 linhas: 24 cidades x 608 dias)
CREATE VIEW dim_clima AS
SELECT
    cidade || '|' || data_ref AS chave_clima,
    cidade,
    data_ref AS data,
    temp_max,
    temp_min,
    precipitacao,
    faixa_precipitacao
FROM clima_cidade;

-- 9) fato_vendas (13.613 linhas no grao id_pedido + item)
-- Nota: +1e-9 garante arredondamento comercial half-up em IEEE-754 (paridade 100% com Etapa 2).
-- Nota: frete_valor e exposto exclusivamente no item = 1 para evitar dupla contagem.
CREATE VIEW fato_vendas AS
SELECT
    i.id_pedido,
    i.item,
    p.data_pedido,
    p.data_envio,
    p.data_entrega_prevista,
    i.data_entrega_real,
    p.prazo_previsto_dias,
    i.prazo_real_dias,
    i.atraso_dias,
    i.status_atraso,
    p.status_pedido,
    cl.canal,
    cl.loja,
    p.chave_canal_loja,
    p.id_cliente,
    p.cidade,
    i.cep,
    i.id_produto,
    i.quantidade,
    i.preco_unitario,
    i.custo_unitario,
    i.desconto_pct,
    ROUND(i.quantidade * i.preco_unitario + 1e-9, 2) AS valor_bruto,
    ROUND(i.quantidade * i.preco_unitario * (1.0 - i.desconto_pct) + 1e-9, 2) AS valor_liquido,
    ROUND(i.quantidade * pr.custo_ref + 1e-9, 2) AS custo_ref_total,
    ROUND(i.quantidade * i.custo_unitario + 1e-9, 2) AS custo_real_total,
    CASE WHEN i.item = 1 THEN p.frete_valor ELSE 0.0 END AS frete_valor,
    i.forma_pagamento,
    p.parcelas,
    i.transportadora,
    i.avaliacao_cliente,
    p.cidade || '|' || COALESCE(i.data_entrega_real, p.data_pedido) AS chave_clima
FROM item_pedido i
JOIN pedido     p  ON p.id_pedido        = i.id_pedido
JOIN canal_loja cl ON cl.chave_canal_loja = p.chave_canal_loja
JOIN produto    pr ON pr.id_produto      = i.id_produto;
