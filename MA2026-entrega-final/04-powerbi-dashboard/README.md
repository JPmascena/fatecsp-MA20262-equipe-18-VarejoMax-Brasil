# 04 — Power BI Dashboard (`VarejoMax Brasil`)

Camada de visualização e inteligência de negócios do projeto **VarejoMax Brasil** (FATEC SP · MA2026 · Prof. Veríssimo).  
Esta pasta contém toda a infraestrutura reproduzível — tema corporativo, consultas Power Query M, medidas DAX, bateria de validação 1-para-1 com o SQL e documentação de montagem — para gerar o relatório **`VarejoMax_Dashboard.pbix`** e sua exportação **`VarejoMax_Dashboard.pdf`** conectados ao banco SQLite **`../02-bancodedados-sql/varejomax.db`** (SQLite **`3.50.4`**) via ODBC (`dsn=varejomax`).

---

## Estrutura da Pasta

```text
04-powerbi-dashboard/
├── README.md
├── tema/
│   └── tema_varejomax.json          # Tema corporativo (paleta, tipografia, fundo #FAF7F2)
├── m/
│   ├── pConexaoODBC.pq              # Parâmetro "dsn=varejomax"
│   ├── tst_Conexao.pq               # Diagnóstico SQLite 3.50.4 + contagem das 9 views
│   ├── dim_produto.pq               # Dimensão Produtos (49 linhas, en-US)
│   ├── dim_cliente.pq               # Dimensão Clientes (1.607 linhas, en-US)
│   ├── dim_cidade.pq                # Dimensão Cidades e Coordenadas (24 linhas, en-US)
│   ├── dim_canal_loja.pq            # Dimensão Canais e Lojas (10 linhas, en-US)
│   ├── dim_transportadora.pq        # Dimensão Transportadoras (6 linhas, en-US)
│   ├── dim_tempo.pq                 # Dimensão Calendário oficial (608 linhas, en-US)
│   ├── dim_cambio.pq                # Dimensão Câmbio BCB PTAX (608 linhas, en-US)
│   ├── dim_clima.pq                 # Dimensão Clima Open-Meteo (14.592 linhas, en-US)
│   └── fato_vendas.pq               # Tabela Fato Vendas (13.613 linhas, en-US)
├── dax/
│   ├── medidas_finais.dax           # 76 medidas em 8 pastas + bloco EVALUATE rápido
│   └── validacao.dax                # Bateria V-K4, V-K5, V-C1..V-C10 (1-para-1 SQL), V-EXTRA1/2 e V-CONF
├── docs/
│   ├── diagrama_modelo.png          # Diagrama Star Schema em alta resolução (1600x1000)
│   ├── modelo_powerbi.md            # Contrato do modelo, relacionamentos R1..R8 e regras DAX
│   └── guia_montagem_pbix.md        # Roteiro passo a passo das 5 páginas + 3 de apoio
└── evidencias/
    └── validacao_resultados.md      # Resultados verificados no SQLite + checklist do .pbix
```

---

## Views SQLite (`../02-bancodedados-sql/varejomax.db` — SQLite `3.50.4`)

| View SQL | Registros Verificados | Script Power Query M |
|---|---:|---|
| `dim_produto` | `49` | [m/dim_produto.pq](m/dim_produto.pq) |
| `dim_cliente` | `1.607` | [m/dim_cliente.pq](m/dim_cliente.pq) |
| `dim_cidade` | `24` | [m/dim_cidade.pq](m/dim_cidade.pq) |
| `dim_canal_loja` | `10` | [m/dim_canal_loja.pq](m/dim_canal_loja.pq) |
| `dim_transportadora` | `6` | [m/dim_transportadora.pq](m/dim_transportadora.pq) |
| `dim_tempo` | `608` | [m/dim_tempo.pq](m/dim_tempo.pq) |
| `dim_cambio` | `608` | [m/dim_cambio.pq](m/dim_cambio.pq) |
| `dim_clima` | `14.592` | [m/dim_clima.pq](m/dim_clima.pq) |
| `fato_vendas` | `13.613` | [m/fato_vendas.pq](m/fato_vendas.pq) |

---

## Como Montar o `VarejoMax_Dashboard.pbix`

1. No **Administrador de Fonte de Dados ODBC (64-bit)** do Windows (`odbcad32.exe`), crie um **DSN de Sistema** usando o *SQLite3 ODBC Driver (64-bit)* com o nome **`varejomax`** apontando para **`../02-bancodedados-sql/varejomax.db`** (caminho absoluto completo até `MA2026-entrega-final/02-bancodedados-sql/varejomax.db`). Detalhes completos em [docs/guia_montagem_pbix.md](docs/guia_montagem_pbix.md).
2. Importe os **11 scripts `.pq`** da pasta [m/](m/) no Editor do Power Query (mantendo `tst_Conexao` com carga desabilitada). Todos usam cultura `"en-US"` para evitar inversão decimal em Windows `pt-BR`.
3. Configure os **8 relacionamentos (`R1`..`R8`)** (`*:1`, ativos, direção única) conforme [docs/modelo_powerbi.md](docs/modelo_powerbi.md) e [docs/diagrama_modelo.png](docs/diagrama_modelo.png), marcando `dim_tempo` como **Tabela de Datas**.
4. Crie a tabela `_Medidas = { BLANK() }`, aplique [dax/medidas_finais.dax](dax/medidas_finais.dax) via **DAX Query View** e importe [tema/tema_varejomax.json](tema/tema_varejomax.json).
5. Monte as **5 páginas principais** (`Visão Executiva`, `Produtos e Categorias`, `Clientes`, `Logística e Clima`, `Canais e Lojas`) e as **3 páginas de apoio** (`Detalhe Transportadora / Loja` drill-through, `Dica Resumo` tooltip `320x240` oculta e `Conferência` oculta).
6. Execute [dax/validacao.dax](dax/validacao.dax), compare com a coluna SQLite em [evidencias/validacao_resultados.md](evidencias/validacao_resultados.md) e exporte **`VarejoMax_Dashboard.pbix`** e **`VarejoMax_Dashboard.pdf`**.

---

## Mapeamento 1-para-1 da Bateria de Validação (`dax/validacao.dax` ↔ SQL)

| Bloco DAX | Espelho SQL | Escopo |
|---|---|---|
| `V-K4` / `V-K5` | `K1`–`K5` | KPIs executivos, contagens de views e integridade |
| `V-C1` | `C1` | Evolução mensal (`ano_mes`) e crescimento MoM |
| `V-C2` | `C2` | Categorias, Margem Ref. vs Realizada e `% Participação` |
| `V-C3` | `C3` | Canais de venda e lojas (`tipo_canal`, `canal`, `loja`) |
| `V-C4` | `C4` | Geográfico — Regiões e Top 10 Cidades |
| `V-C5` | `C5` | Desempenho logístico por transportadora e nota média |
| `V-C6` | `C6` | Pedidos críticos em atraso (`atraso_dias >= 5`) |
| `V-C7` | `C7` | Clientes ativos, recorrência/RFV e concentração regional |
| `V-C8` | `C8` | Impacto de descontos, cancelamentos e devoluções |
| `V-C9` | `C9` | API 1 — Câmbio PTAX × `Eletrônicos` vs `Demais Categorias` |
| `V-C10` | `C10` | API 2 — Clima (`faixa_precipitacao`) & Calendário × atrasos e vendas |
| `V-EXTRA1` / `V-EXTRA2` | Complemento BI | Top 10 Produtos e Mix de Formas de Pagamento |
| `V-CONF` | Página `Conferência` | Checksum cruzado final do modelo |

---

## Números de Referência Verificados (`varejomax.db`)

| KPI | Valor Verificado (SQLite `3.50.4`) |
|---|---:|
| Linhas `fato_vendas` | `13.613` |
| Pedidos Totais / Válidos | `8.996` / `8.291` |
| Faturamento Bruto (não Cancelado) | `R$ 3.797.167,40` |
| Faturamento Líquido (Válidos) | `R$ 3.448.872,47` |
| Ticket Médio | `R$ 415,98` |
| Frete Total (Todos Status / Válidos) | `R$ 97.922,36` / `R$ 93.707,67` |
| Margem Bruta % (Ref.) / Realizada % | `31,20%` / `30,33%` |
| Pedidos Atrasados (`628` s/ `8.291` válidos vs. `8.180` entregues) | `7,57%` / `7,68%` |
| Nota Média Cliente (Itens 1–5 / Itens Válidos / Pedidos Entregues) | `4,11` / `4,18` / `4,21` |
| Clientes Ativos / Novos / Recorrentes (`% Recompra`) | `1.571` / `449` / `1.122` (`71,42%`) |
