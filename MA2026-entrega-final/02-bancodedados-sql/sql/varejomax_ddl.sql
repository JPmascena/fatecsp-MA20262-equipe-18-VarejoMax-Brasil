-- =============================================================
-- varejomax_ddl.sql -- Estrutura relacional 3FN do banco varejomax.db
-- Projeto: VarejoMax Brasil (ILP590 -- Fatec Sao Paulo)
-- SGBD: SQLite 3
-- =============================================================

PRAGMA foreign_keys = ON;

-- Limpeza em ordem inversa das dependencias de FK
DROP TABLE IF EXISTS item_pedido;
DROP TABLE IF EXISTS pedido;
DROP TABLE IF EXISTS clima_cidade;
DROP TABLE IF EXISTS cambio_ptax;
DROP TABLE IF EXISTS cliente;
DROP TABLE IF EXISTS transportadora;
DROP TABLE IF EXISTS canal_loja;
DROP TABLE IF EXISTS produto;
DROP TABLE IF EXISTS cidade;
DROP TABLE IF EXISTS calendario;

DROP TABLE IF EXISTS stg_vendas;
DROP TABLE IF EXISTS stg_produto;
DROP TABLE IF EXISTS stg_cliente;
DROP TABLE IF EXISTS stg_cidade;
DROP TABLE IF EXISTS stg_canal_loja;
DROP TABLE IF EXISTS stg_transportadora;
DROP TABLE IF EXISTS stg_cambio_ptax;
DROP TABLE IF EXISTS stg_clima_cidade;
DROP TABLE IF EXISTS stg_feriados;

-- -------------------------------------------------------------
-- 1) Tabelas de Dimensao / Cadastro (3FN)
-- -------------------------------------------------------------
CREATE TABLE cidade (
    cidade           TEXT NOT NULL,
    uf               TEXT NOT NULL,
    regiao           TEXT NOT NULL,
    latitude         REAL NOT NULL,
    longitude        REAL NOT NULL,
    tem_loja_fisica  TEXT NOT NULL,
    CONSTRAINT pk_cidade PRIMARY KEY (cidade),
    CONSTRAINT ck_cidade_uf CHECK (LENGTH(uf) = 2),
    CONSTRAINT ck_cidade_regiao CHECK (
        regiao IN ('Centro-Oeste', 'Nordeste', 'Norte', 'Sudeste', 'Sul')
    ),
    CONSTRAINT ck_cidade_loja_fisica CHECK (tem_loja_fisica IN ('Sim', 'Não'))
);

CREATE TABLE canal_loja (
    chave_canal_loja TEXT NOT NULL,
    canal            TEXT NOT NULL,
    loja             TEXT NOT NULL,
    tipo_canal       TEXT NOT NULL,
    cod_loja         TEXT,
    cidade_loja      TEXT,
    uf_loja          TEXT,
    CONSTRAINT pk_canal_loja PRIMARY KEY (chave_canal_loja),
    CONSTRAINT ck_canal_nome CHECK (
        canal IN ('App', 'Loja Física', 'Marketplace', 'Site')
    ),
    CONSTRAINT ck_canal_tipo CHECK (
        tipo_canal IN ('Loja física', 'Online (e-commerce)')
    )
);

CREATE TABLE produto (
    id_produto     TEXT NOT NULL,
    produto        TEXT NOT NULL,
    categoria      TEXT NOT NULL,
    marca          TEXT NOT NULL,
    preco_ref_2025 REAL NOT NULL,
    preco_ref_2026 REAL NOT NULL,
    custo_ref      REAL NOT NULL,
    CONSTRAINT pk_produto PRIMARY KEY (id_produto),
    CONSTRAINT ck_produto_categoria CHECK (
        categoria IN (
            'Beleza e Saúde',
            'Casa e Decoração',
            'Eletrônicos',
            'Esporte e Lazer',
            'Livros e Papelaria',
            'Moda'
        )
    ),
    CONSTRAINT ck_produto_valores CHECK (
        preco_ref_2025 > 0 AND preco_ref_2026 > 0 AND custo_ref >= 0
    )
);

-- Nota: cliente.cidade e anulavel (ha registro de cadastro sem cidade na Etapa 2)
CREATE TABLE cliente (
    id_cliente    TEXT NOT NULL,
    nome_cliente  TEXT NOT NULL,
    email         TEXT,
    data_cadastro TEXT NOT NULL,
    cidade        TEXT,
    CONSTRAINT pk_cliente PRIMARY KEY (id_cliente),
    CONSTRAINT fk_cliente_cidade FOREIGN KEY (cidade)
        REFERENCES cidade (cidade)
);

CREATE TABLE transportadora (
    transportadora TEXT NOT NULL,
    tipo           TEXT NOT NULL,
    CONSTRAINT pk_transportadora PRIMARY KEY (transportadora),
    CONSTRAINT ck_transportadora_tipo CHECK (
        tipo IN ('Não informado', 'Retirada', 'Transportadora')
    )
);

CREATE TABLE calendario (
    data          TEXT    NOT NULL,
    ano           INTEGER NOT NULL,
    mes           INTEGER NOT NULL,
    nome_mes      TEXT    NOT NULL,
    mes_abrev     TEXT    NOT NULL,
    ano_mes       TEXT    NOT NULL,
    mes_inicio    TEXT    NOT NULL,
    trimestre     TEXT    NOT NULL,
    dia_semana    INTEGER NOT NULL,
    nome_dia      TEXT    NOT NULL,
    fim_de_semana TEXT    NOT NULL,
    feriado       TEXT    NOT NULL,
    nome_feriado  TEXT,
    CONSTRAINT pk_calendario PRIMARY KEY (data),
    CONSTRAINT ck_cal_mes CHECK (mes BETWEEN 1 AND 12),
    CONSTRAINT ck_cal_trimestre CHECK (trimestre IN ('T1', 'T2', 'T3', 'T4')),
    CONSTRAINT ck_cal_dia_semana CHECK (dia_semana BETWEEN 1 AND 7),
    CONSTRAINT ck_cal_fim_de_semana CHECK (fim_de_semana IN ('Sim', 'Não')),
    CONSTRAINT ck_cal_feriado CHECK (feriado IN ('Sim', 'Não'))
);

-- -------------------------------------------------------------
-- 2) Tabelas Operacionais (3FN: pedido + item_pedido)
-- -------------------------------------------------------------
-- Nota: data_entrega_prevista nao recebe FK para calendario porque
-- existem entregas previstas ate 2026-09-11 (apos 2026-08-31).
CREATE TABLE pedido (
    id_pedido             TEXT    NOT NULL,
    data_pedido           TEXT    NOT NULL,
    data_envio            TEXT,
    data_entrega_prevista TEXT,
    prazo_previsto_dias   INTEGER,
    status_pedido         TEXT    NOT NULL,
    id_cliente            TEXT    NOT NULL,
    chave_canal_loja      TEXT    NOT NULL,
    cidade                TEXT    NOT NULL,
    parcelas              INTEGER NOT NULL,
    frete_valor           REAL    NOT NULL DEFAULT 0,
    CONSTRAINT pk_pedido PRIMARY KEY (id_pedido),
    CONSTRAINT fk_pedido_calendario FOREIGN KEY (data_pedido)
        REFERENCES calendario (data),
    CONSTRAINT fk_pedido_cliente FOREIGN KEY (id_cliente)
        REFERENCES cliente (id_cliente),
    CONSTRAINT fk_pedido_canal_loja FOREIGN KEY (chave_canal_loja)
        REFERENCES canal_loja (chave_canal_loja),
    CONSTRAINT fk_pedido_cidade FOREIGN KEY (cidade)
        REFERENCES cidade (cidade),
    CONSTRAINT ck_pedido_status CHECK (
        status_pedido IN (
            'Cancelado',
            'Devolvido',
            'Em separação',
            'Em trânsito',
            'Entregue'
        )
    ),
    CONSTRAINT ck_pedido_parcelas CHECK (parcelas >= 1),
    CONSTRAINT ck_pedido_frete CHECK (frete_valor >= 0)
);

-- Nota: Os 5 atributos que variam entre itens do mesmo pedido (cep,
-- transportadora, forma_pagamento, data_entrega_real, avaliacao_cliente)
-- permanecem em item_pedido, junto com prazo_real_dias, atraso_dias e
-- status_atraso, garantindo paridade 100% com fVendas.
-- Nao aplicamos CHECK 1..5 em avaliacao_cliente nem data_entrega_real >= data_envio
-- para preservar os registros aceitos na Etapa 2 (tratados nas queries/DAX).
CREATE TABLE item_pedido (
    id_pedido         TEXT    NOT NULL,
    item              INTEGER NOT NULL,
    id_produto        TEXT    NOT NULL,
    quantidade        INTEGER NOT NULL,
    preco_unitario    REAL    NOT NULL,
    custo_unitario    REAL    NOT NULL,
    desconto_pct      REAL    NOT NULL DEFAULT 0,
    cep               TEXT,
    forma_pagamento   TEXT    NOT NULL,
    transportadora    TEXT    NOT NULL,
    data_entrega_real TEXT,
    prazo_real_dias   INTEGER,
    atraso_dias       INTEGER,
    status_atraso     TEXT    NOT NULL,
    avaliacao_cliente INTEGER,
    CONSTRAINT pk_item_pedido PRIMARY KEY (id_pedido, item),
    CONSTRAINT uq_item_pedido_produto UNIQUE (id_pedido, id_produto),
    CONSTRAINT fk_item_pedido FOREIGN KEY (id_pedido)
        REFERENCES pedido (id_pedido),
    CONSTRAINT fk_item_produto FOREIGN KEY (id_produto)
        REFERENCES produto (id_produto),
    CONSTRAINT fk_item_transportadora FOREIGN KEY (transportadora)
        REFERENCES transportadora (transportadora),
    CONSTRAINT ck_item_num CHECK (item >= 1),
    CONSTRAINT ck_item_quantidade CHECK (quantidade > 0),
    CONSTRAINT ck_item_valores CHECK (preco_unitario > 0 AND custo_unitario >= 0),
    CONSTRAINT ck_item_desconto CHECK (desconto_pct >= 0 AND desconto_pct <= 1),
    CONSTRAINT ck_item_status_atraso CHECK (
        status_atraso IN ('-', 'Adiantado', 'Atrasado', 'Inconsistente', 'No prazo')
    )
);

-- -------------------------------------------------------------
-- 3) Tabelas Externas de APIs (PTAX BCB + Open-Meteo)
-- -------------------------------------------------------------
CREATE TABLE cambio_ptax (
    data_ref       TEXT NOT NULL,
    cotacao_compra REAL NOT NULL,
    cotacao_venda  REAL NOT NULL,
    boletim_origem TEXT NOT NULL,
    faixa_dolar    TEXT NOT NULL,
    CONSTRAINT pk_cambio_ptax PRIMARY KEY (data_ref),
    CONSTRAINT fk_cambio_calendario FOREIGN KEY (data_ref)
        REFERENCES calendario (data),
    CONSTRAINT ck_cambio_cotacao CHECK (cotacao_compra > 0 AND cotacao_venda > 0),
    CONSTRAINT ck_cambio_origem CHECK (boletim_origem IN ('Carry-forward', 'Oficial BCB'))
);

CREATE TABLE clima_cidade (
    cidade             TEXT NOT NULL,
    data_ref           TEXT NOT NULL,
    temp_max           REAL,
    temp_min           REAL,
    precipitacao       REAL,
    faixa_precipitacao TEXT NOT NULL,
    CONSTRAINT pk_clima_cidade PRIMARY KEY (cidade, data_ref),
    CONSTRAINT fk_clima_cidade FOREIGN KEY (cidade)
        REFERENCES cidade (cidade),
    CONSTRAINT fk_clima_calendario FOREIGN KEY (data_ref)
        REFERENCES calendario (data),
    CONSTRAINT ck_clima_precip CHECK (precipitacao IS NULL OR precipitacao >= 0)
);

-- -------------------------------------------------------------
-- 4) Indices para Consultas Analiticas e Joins
-- -------------------------------------------------------------
CREATE INDEX ix_pedido_data         ON pedido (data_pedido);
CREATE INDEX ix_pedido_cliente      ON pedido (id_cliente);
CREATE INDEX ix_pedido_canal_loja   ON pedido (chave_canal_loja);
CREATE INDEX ix_pedido_cidade       ON pedido (cidade);
CREATE INDEX ix_pedido_status       ON pedido (status_pedido);
CREATE INDEX ix_item_pedido_produto ON item_pedido (id_produto);
CREATE INDEX ix_item_pedido_transp  ON item_pedido (transportadora);
CREATE INDEX ix_clima_cidade_data   ON clima_cidade (data_ref);
CREATE INDEX ix_calendario_ano_mes  ON calendario (ano_mes);

-- -------------------------------------------------------------
-- 5) Tabelas de Staging (carga bruta dos 9 CSVs como TEXT)
-- -------------------------------------------------------------
CREATE TABLE stg_cidade (
    Cidade          TEXT,
    UF              TEXT,
    Regiao          TEXT,
    Latitude        TEXT,
    Longitude       TEXT,
    Tem_Loja_Fisica TEXT
);

CREATE TABLE stg_canal_loja (
    Chave_Canal_Loja TEXT,
    Canal            TEXT,
    Loja             TEXT,
    Tipo_Canal       TEXT,
    Cod_Loja         TEXT,
    Cidade_Loja      TEXT,
    UF_Loja          TEXT
);

CREATE TABLE stg_produto (
    ID_Produto     TEXT,
    Produto        TEXT,
    Categoria      TEXT,
    Marca          TEXT,
    Preco_Ref_2025 TEXT,
    Preco_Ref_2026 TEXT,
    Custo_Ref      TEXT
);

CREATE TABLE stg_cliente (
    ID_Cliente    TEXT,
    Nome_Cliente  TEXT,
    Email         TEXT,
    Data_Cadastro TEXT,
    Cidade        TEXT,
    UF            TEXT,
    Regiao        TEXT
);

CREATE TABLE stg_transportadora (
    Transportadora TEXT,
    Tipo           TEXT
);

CREATE TABLE stg_vendas (
    ID_Pedido             TEXT,
    Item                  TEXT,
    Data_Pedido           TEXT,
    Data_Envio            TEXT,
    Data_Entrega_Prevista TEXT,
    Data_Entrega_Real     TEXT,
    Prazo_Previsto_Dias   TEXT,
    Prazo_Real_Dias       TEXT,
    Atraso_Dias           TEXT,
    Status_Atraso         TEXT,
    Status_Pedido         TEXT,
    Canal                 TEXT,
    Loja                  TEXT,
    Chave_Canal_Loja      TEXT,
    ID_Cliente            TEXT,
    Cidade                TEXT,
    CEP                   TEXT,
    ID_Produto            TEXT,
    Quantidade            TEXT,
    Preco_Unitario        TEXT,
    Custo_Unitario        TEXT,
    Desconto_Pct          TEXT,
    Valor_Bruto           TEXT,
    Valor_Liquido         TEXT,
    Valor_Item_Original   TEXT,
    Frete_Valor           TEXT,
    Forma_Pagamento       TEXT,
    Parcelas              TEXT,
    Transportadora        TEXT,
    Avaliacao_Cliente     TEXT,
    Arquivo_Origem        TEXT,
    Linha_Origem          TEXT
);

CREATE TABLE stg_cambio_ptax (
    data           TEXT,
    cotacao_compra TEXT,
    cotacao_venda  TEXT,
    boletim_origem TEXT,
    faixa_dolar    TEXT
);

CREATE TABLE stg_clima_cidade (
    cidade             TEXT,
    data_ref           TEXT,
    temp_max           TEXT,
    temp_min           TEXT,
    precipitacao       TEXT,
    faixa_precipitacao TEXT
);

CREATE TABLE stg_feriados (
    data         TEXT,
    nome_feriado TEXT,
    tipo         TEXT
);
