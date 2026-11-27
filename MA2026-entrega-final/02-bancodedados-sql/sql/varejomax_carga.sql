-- =============================================================
-- varejomax_carga.sql -- Carga idempotente das tabelas 3FN e APIs
-- Projeto: VarejoMax Brasil (ILP590 -- Fatec Sao Paulo)
-- Requisito: pode ser executado N vezes seguidas com resultado identico.
-- =============================================================

PRAGMA foreign_keys = ON;

-- -------------------------------------------------------------
-- 0) Limpeza previa em ordem inversa das FKs (idempotencia)
-- -------------------------------------------------------------
DELETE FROM item_pedido;
DELETE FROM pedido;
DELETE FROM clima_cidade;
DELETE FROM cambio_ptax;
DELETE FROM cliente;
DELETE FROM transportadora;
DELETE FROM canal_loja;
DELETE FROM produto;
DELETE FROM cidade;
DELETE FROM calendario;

-- -------------------------------------------------------------
-- 1) Calendario (608 dias: 2025-01-01 a 2026-08-31) + BrasilAPI
-- -------------------------------------------------------------
WITH RECURSIVE datas(d) AS (
    SELECT '2025-01-01'
    UNION ALL
    SELECT DATE(d, '+1 day')
    FROM datas
    WHERE d < '2026-08-31'
),
feriados_unicos AS (
    SELECT TRIM(data) AS data,
           MIN(TRIM(nome_feriado)) AS nome_feriado
    FROM stg_feriados
    WHERE TRIM(data) <> ''
    GROUP BY TRIM(data)
)
INSERT INTO calendario (
    data, ano, mes, nome_mes, mes_abrev, ano_mes, mes_inicio,
    trimestre, dia_semana, nome_dia, fim_de_semana, feriado, nome_feriado
)
SELECT
    d AS data,
    CAST(STRFTIME('%Y', d) AS INTEGER) AS ano,
    CAST(STRFTIME('%m', d) AS INTEGER) AS mes,
    CASE STRFTIME('%m', d)
        WHEN '01' THEN 'Janeiro'   WHEN '02' THEN 'Fevereiro'
        WHEN '03' THEN 'Março'     WHEN '04' THEN 'Abril'
        WHEN '05' THEN 'Maio'      WHEN '06' THEN 'Junho'
        WHEN '07' THEN 'Julho'     WHEN '08' THEN 'Agosto'
        WHEN '09' THEN 'Setembro'  WHEN '10' THEN 'Outubro'
        WHEN '11' THEN 'Novembro'  ELSE 'Dezembro'
    END AS nome_mes,
    CASE STRFTIME('%m', d)
        WHEN '01' THEN 'Jan' WHEN '02' THEN 'Fev' WHEN '03' THEN 'Mar'
        WHEN '04' THEN 'Abr' WHEN '05' THEN 'Mai' WHEN '06' THEN 'Jun'
        WHEN '07' THEN 'Jul' WHEN '08' THEN 'Ago' WHEN '09' THEN 'Set'
        WHEN '10' THEN 'Out' WHEN '11' THEN 'Nov' ELSE 'Dez'
    END AS mes_abrev,
    STRFTIME('%Y-%m', d) AS ano_mes,
    STRFTIME('%Y-%m-01', d) AS mes_inicio,
    'T' || CAST(((CAST(STRFTIME('%m', d) AS INTEGER) - 1) / 3) + 1 AS TEXT) AS trimestre,
    CASE CAST(STRFTIME('%w', d) AS INTEGER)
        WHEN 0 THEN 7
        ELSE CAST(STRFTIME('%w', d) AS INTEGER)
    END AS dia_semana,
    CASE CAST(STRFTIME('%w', d) AS INTEGER)
        WHEN 0 THEN 'Domingo'
        WHEN 1 THEN 'Segunda-feira'
        WHEN 2 THEN 'Terça-feira'
        WHEN 3 THEN 'Quarta-feira'
        WHEN 4 THEN 'Quinta-feira'
        WHEN 5 THEN 'Sexta-feira'
        ELSE 'Sábado'
    END AS nome_dia,
    CASE WHEN STRFTIME('%w', d) IN ('0', '6') THEN 'Sim' ELSE 'Não' END AS fim_de_semana,
    CASE WHEN f.data IS NOT NULL THEN 'Sim' ELSE 'Não' END AS feriado,
    f.nome_feriado AS nome_feriado
FROM datas
LEFT JOIN feriados_unicos f ON f.data = datas.d;

-- -------------------------------------------------------------
-- 2) Dimensoes / Cadastros (3FN)
-- -------------------------------------------------------------
INSERT INTO cidade (cidade, uf, regiao, latitude, longitude, tem_loja_fisica)
SELECT
    TRIM(Cidade),
    TRIM(UF),
    TRIM(Regiao),
    CAST(CASE WHEN INSTR(Latitude, ',') > 0
              THEN REPLACE(REPLACE(TRIM(Latitude), '.', ''), ',', '.')
              ELSE TRIM(Latitude) END AS REAL),
    CAST(CASE WHEN INSTR(Longitude, ',') > 0
              THEN REPLACE(REPLACE(TRIM(Longitude), '.', ''), ',', '.')
              ELSE TRIM(Longitude) END AS REAL),
    TRIM(Tem_Loja_Fisica)
FROM stg_cidade;

INSERT INTO canal_loja (chave_canal_loja, canal, loja, tipo_canal, cod_loja, cidade_loja, uf_loja)
SELECT
    TRIM(Chave_Canal_Loja),
    TRIM(Canal),
    TRIM(Loja),
    TRIM(Tipo_Canal),
    NULLIF(TRIM(Cod_Loja), ''),
    NULLIF(TRIM(Cidade_Loja), ''),
    NULLIF(TRIM(UF_Loja), '')
FROM stg_canal_loja;

INSERT INTO produto (id_produto, produto, categoria, marca, preco_ref_2025, preco_ref_2026, custo_ref)
SELECT
    TRIM(ID_Produto),
    TRIM(Produto),
    TRIM(Categoria),
    TRIM(Marca),
    CAST(CASE WHEN INSTR(Preco_Ref_2025, ',') > 0
              THEN REPLACE(REPLACE(TRIM(Preco_Ref_2025), '.', ''), ',', '.')
              ELSE TRIM(Preco_Ref_2025) END AS REAL),
    CAST(CASE WHEN INSTR(Preco_Ref_2026, ',') > 0
              THEN REPLACE(REPLACE(TRIM(Preco_Ref_2026), '.', ''), ',', '.')
              ELSE TRIM(Preco_Ref_2026) END AS REAL),
    CAST(CASE WHEN INSTR(Custo_Ref, ',') > 0
              THEN REPLACE(REPLACE(TRIM(Custo_Ref), '.', ''), ',', '.')
              ELSE TRIM(Custo_Ref) END AS REAL)
FROM stg_produto;

INSERT INTO cliente (id_cliente, nome_cliente, email, data_cadastro, cidade)
SELECT
    TRIM(ID_Cliente),
    TRIM(Nome_Cliente),
    NULLIF(TRIM(Email), ''),
    CASE WHEN INSTR(TRIM(Data_Cadastro), '/') > 0
         THEN SUBSTR(TRIM(Data_Cadastro), 7, 4) || '-' || SUBSTR(TRIM(Data_Cadastro), 4, 2) || '-' || SUBSTR(TRIM(Data_Cadastro), 1, 2)
         ELSE TRIM(Data_Cadastro) END,
    NULLIF(TRIM(Cidade), '')
FROM stg_cliente;

INSERT INTO transportadora (transportadora, tipo)
SELECT
    TRIM(Transportadora),
    TRIM(Tipo)
FROM stg_transportadora;

-- -------------------------------------------------------------
-- 3) Tabelas Operacionais: pedido (8.996) e item_pedido (13.613)
-- -------------------------------------------------------------
INSERT INTO pedido (
    id_pedido, data_pedido, data_envio, data_entrega_prevista,
    prazo_previsto_dias, status_pedido, id_cliente, chave_canal_loja,
    cidade, parcelas, frete_valor
)
SELECT
    TRIM(ID_Pedido) AS id_pedido,
    MIN(CASE WHEN INSTR(TRIM(Data_Pedido), '/') > 0
             THEN SUBSTR(TRIM(Data_Pedido), 7, 4) || '-' || SUBSTR(TRIM(Data_Pedido), 4, 2) || '-' || SUBSTR(TRIM(Data_Pedido), 1, 2)
             ELSE TRIM(Data_Pedido) END) AS data_pedido,
    MIN(CASE WHEN NULLIF(TRIM(Data_Envio), '') IS NULL THEN NULL
             WHEN INSTR(TRIM(Data_Envio), '/') > 0
             THEN SUBSTR(TRIM(Data_Envio), 7, 4) || '-' || SUBSTR(TRIM(Data_Envio), 4, 2) || '-' || SUBSTR(TRIM(Data_Envio), 1, 2)
             ELSE TRIM(Data_Envio) END) AS data_envio,
    MIN(CASE WHEN NULLIF(TRIM(Data_Entrega_Prevista), '') IS NULL THEN NULL
             WHEN INSTR(TRIM(Data_Entrega_Prevista), '/') > 0
             THEN SUBSTR(TRIM(Data_Entrega_Prevista), 7, 4) || '-' || SUBSTR(TRIM(Data_Entrega_Prevista), 4, 2) || '-' || SUBSTR(TRIM(Data_Entrega_Prevista), 1, 2)
             ELSE TRIM(Data_Entrega_Prevista) END) AS data_entrega_prevista,
    MIN(CASE WHEN NULLIF(TRIM(Prazo_Previsto_Dias), '') IS NULL THEN NULL
             ELSE CAST(TRIM(Prazo_Previsto_Dias) AS INTEGER) END) AS prazo_previsto_dias,
    MIN(TRIM(Status_Pedido)) AS status_pedido,
    MIN(TRIM(ID_Cliente)) AS id_cliente,
    MIN(TRIM(Chave_Canal_Loja)) AS chave_canal_loja,
    MIN(TRIM(Cidade)) AS cidade,
    MIN(CAST(TRIM(Parcelas) AS INTEGER)) AS parcelas,
    MAX(CAST(CASE WHEN INSTR(Frete_Valor, ',') > 0
                  THEN REPLACE(REPLACE(TRIM(Frete_Valor), '.', ''), ',', '.')
                  ELSE TRIM(Frete_Valor) END AS REAL)) AS frete_valor
FROM stg_vendas
GROUP BY TRIM(ID_Pedido);

INSERT INTO item_pedido (
    id_pedido, item, id_produto, quantidade, preco_unitario,
    custo_unitario, desconto_pct, cep, forma_pagamento,
    transportadora, data_entrega_real, prazo_real_dias,
    atraso_dias, status_atraso, avaliacao_cliente
)
SELECT
    TRIM(ID_Pedido),
    CAST(TRIM(Item) AS INTEGER),
    TRIM(ID_Produto),
    CAST(TRIM(Quantidade) AS INTEGER),
    CAST(CASE WHEN INSTR(Preco_Unitario, ',') > 0
              THEN REPLACE(REPLACE(TRIM(Preco_Unitario), '.', ''), ',', '.')
              ELSE TRIM(Preco_Unitario) END AS REAL),
    CAST(CASE WHEN INSTR(Custo_Unitario, ',') > 0
              THEN REPLACE(REPLACE(TRIM(Custo_Unitario), '.', ''), ',', '.')
              ELSE TRIM(Custo_Unitario) END AS REAL),
    CAST(CASE WHEN INSTR(Desconto_Pct, ',') > 0
              THEN REPLACE(REPLACE(TRIM(Desconto_Pct), '.', ''), ',', '.')
              ELSE TRIM(Desconto_Pct) END AS REAL),
    NULLIF(TRIM(CEP), ''),
    TRIM(Forma_Pagamento),
    TRIM(Transportadora),
    CASE WHEN NULLIF(TRIM(Data_Entrega_Real), '') IS NULL THEN NULL
         WHEN INSTR(TRIM(Data_Entrega_Real), '/') > 0
         THEN SUBSTR(TRIM(Data_Entrega_Real), 7, 4) || '-' || SUBSTR(TRIM(Data_Entrega_Real), 4, 2) || '-' || SUBSTR(TRIM(Data_Entrega_Real), 1, 2)
         ELSE TRIM(Data_Entrega_Real) END,
    CASE WHEN NULLIF(TRIM(Prazo_Real_Dias), '') IS NULL THEN NULL
         ELSE CAST(TRIM(Prazo_Real_Dias) AS INTEGER) END,
    CASE WHEN NULLIF(TRIM(Atraso_Dias), '') IS NULL THEN NULL
         ELSE CAST(TRIM(Atraso_Dias) AS INTEGER) END,
    TRIM(Status_Atraso),
    CASE WHEN NULLIF(TRIM(Avaliacao_Cliente), '') IS NULL THEN NULL
         ELSE CAST(TRIM(Avaliacao_Cliente) AS INTEGER) END
FROM stg_vendas;

-- -------------------------------------------------------------
-- 4) Tabelas Externas de APIs (PTAX + Clima Open-Meteo)
-- -------------------------------------------------------------
INSERT INTO cambio_ptax (data_ref, cotacao_compra, cotacao_venda, boletim_origem, faixa_dolar)
SELECT
    TRIM(data),
    CAST(CASE WHEN INSTR(cotacao_compra, ',') > 0
              THEN REPLACE(REPLACE(TRIM(cotacao_compra), '.', ''), ',', '.')
              ELSE TRIM(cotacao_compra) END AS REAL),
    CAST(CASE WHEN INSTR(cotacao_venda, ',') > 0
              THEN REPLACE(REPLACE(TRIM(cotacao_venda), '.', ''), ',', '.')
              ELSE TRIM(cotacao_venda) END AS REAL),
    TRIM(boletim_origem),
    TRIM(faixa_dolar)
FROM stg_cambio_ptax;

INSERT INTO clima_cidade (cidade, data_ref, temp_max, temp_min, precipitacao, faixa_precipitacao)
SELECT
    TRIM(cidade),
    TRIM(data_ref),
    CAST(CASE WHEN NULLIF(TRIM(temp_max), '') IS NULL THEN NULL
              WHEN INSTR(temp_max, ',') > 0
              THEN REPLACE(REPLACE(TRIM(temp_max), '.', ''), ',', '.')
              ELSE TRIM(temp_max) END AS REAL),
    CAST(CASE WHEN NULLIF(TRIM(temp_min), '') IS NULL THEN NULL
              WHEN INSTR(temp_min, ',') > 0
              THEN REPLACE(REPLACE(TRIM(temp_min), '.', ''), ',', '.')
              ELSE TRIM(temp_min) END AS REAL),
    CAST(CASE WHEN NULLIF(TRIM(precipitacao), '') IS NULL THEN NULL
              WHEN INSTR(precipitacao, ',') > 0
              THEN REPLACE(REPLACE(TRIM(precipitacao), '.', ''), ',', '.')
              ELSE TRIM(precipitacao) END AS REAL),
    TRIM(faixa_precipitacao)
FROM stg_clima_cidade;
