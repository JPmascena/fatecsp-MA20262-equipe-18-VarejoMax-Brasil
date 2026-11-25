# VarejoMax Brasil — Inteligência e Análise de Dados para Varejo e E-commerce

> Projeto Prático **MA2026** · FATEC SP · Orientação: Prof. Veríssimo

Plataforma de análise de dados construída de forma evolutiva, que parte de uma planilha bruta e inconsistente de vendas e chega a um **dashboard interativo no Power BI**, passando por Excel avançado, Power Query, Power Pivot/DAX, VBA, SQL e consumo de APIs REST.

> **Nota:** a VarejoMax Brasil é uma empresa **fictícia** e todos os dados são **simulados**.

---

## Sumário

- [Contexto](#contexto)
- [Problema e pergunta central](#problema-e-pergunta-central)
- [Objetivos](#objetivos)
- [Arquitetura e pipeline](#arquitetura-e-pipeline)
- [Fontes de dados](#fontes-de-dados)
- [Modelo dimensional](#modelo-dimensional)
- [Indicadores (KPIs)](#indicadores-kpis)
- [Premissas de tratamento](#premissas-de-tratamento)
- [Estrutura do repositório](#estrutura-do-repositório)
- [Cronograma e entregas](#cronograma-e-entregas)
- [Resultados esperados](#resultados-esperados)
- [Validação das 4 hipóteses — principais achados N2](#validação-das-4-hipóteses--principais-achados-n2)
- [Status do projeto](#status-do-projeto)
- [Equipe](#equipe)

---

## Contexto

A **VarejoMax Brasil** é uma rede omnichannel que vende pelo site, pelo aplicativo, por marketplace e em lojas físicas (São Paulo, Rio de Janeiro, Belo Horizonte, Curitiba, Porto Alegre e Brasília).

| Item | Detalhe |
|---|---|
| Catálogo | 49 produtos em 6 categorias |
| Volume | ~9 mil pedidos (~13,7 mil linhas de itens, 30 colunas) |
| Período | janeiro/2025 a agosto/2026 |
| Abrangência | clientes de 24 cidades brasileiras |

Os dados chegam em **uma única planilha bruta** exportada do sistema, misturando pedido, cliente, produto, pagamento e logística em cada linha. Não há padronização, e os problemas incluem:

- datas gravadas como texto;
- categorias e cidades escritas de formas diferentes;
- linhas duplicadas;
- valores que não batem com preço × quantidade;
- CEPs sem o zero à esquerda;
- registros de teste.

Como consequência, cada área chega a números diferentes e não existe uma visão consolidada por canal, região, produto ou transportadora.

## Problema e pergunta central

A gestão não consegue cruzar, de forma confiável, **vendas, logística e comportamento dos clientes**.

> **Como os dados de vendas, logística e clientes podem apoiar decisões no varejo e no e-commerce?**

### Perguntas de apoio

- Quais produtos, categorias e canais concentram faturamento e margem, e como isso evolui ao longo do tempo?
- Onde e com quais transportadoras o prazo real excede o previsto, e qual o efeito sobre avaliação e devoluções?
- Como se comportam os clientes novos e os recorrentes (ticket médio, frequência e recompra)?
- Câmbio e clima se relacionam com as vendas de determinadas categorias e com os atrasos de entrega?

### Hipóteses a validar

1. O atraso na entrega reduz a nota do cliente.
2. Regiões mais distantes do centro de distribuição concentram os maiores atrasos.
3. O desempenho varia entre transportadoras e entre canais.
4. Picos de demanda, como a Black Friday, pioram o prazo.

## Objetivos

**Geral:** construir uma plataforma de inteligência e análise de dados para a VarejoMax Brasil, conectando fontes heterogêneas (planilhas, APIs externas e banco relacional) a um dashboard interativo no Power BI.

**Específicos:**

1. **Excel avançado:** estruturar, padronizar e auditar a planilha `Vendas_Bruto` com PROCX, validações e colunas de auditoria, sem erros de fórmula, registrando cada inconsistência e a decisão de tratamento.
2. **Power Query (M):** limpar, transformar e consolidar as fontes (tipos, textos, datas, duplicatas, registros de teste e valores ausentes).
3. **Power Pivot e DAX:** modelar um Star Schema e criar KPIs (faturamento, ticket médio, margem, prazo, atraso e recompra).
4. **VBA:** automatizar relatórios operacionais, como o resumo periódico de vendas por categoria e o relatório de pedidos atrasados por transportadora.
5. **SQL:** modelar e implementar um banco relacional (DDL e DML) para consultas analíticas.
6. **APIs REST:** enriquecer a base com câmbio e clima, relacionando fatores externos às vendas e aos atrasos.
7. **Power BI:** desenvolver um dashboard com navegação interativa e filtros dinâmicos para apoio à decisão.

## Arquitetura e pipeline

```text
Dado Bruto → Excel Avançado → Power Query → Power Pivot/DAX → VBA → SQL → API/Web → Power BI
```

| Camada | Papel no projeto |
|---|---|
| Excel Avançado | Auditoria, validações e registro de inconsistências da planilha bruta |
| Power Query | Limpeza, padronização e consolidação (linguagem M) |
| Power Pivot / DAX | Modelo estrela e medidas analíticas |
| VBA | Macros de relatórios operacionais e reconciliação automática (`01-vba`) |
| SQL | Persistência relacional 3FN (`varejomax.db` — SQLite 3.50.4), views Star Schema e consultas `K1`–`K5`, `D1`–`D4` e `C1`–`C10` (`02-bancodedados-sql`) |
| API / Web | Enriquecimento com câmbio, clima e feriados nacionais via 3 APIs REST (`03-api-web`) |
| Power BI | Dashboard final conectado via ODBC 64-bit (`dsn=varejomax`) para tomada de decisão (`04-powerbi-dashboard`) |

## Fontes de dados

| Tipo | Origem | Finalidade |
|---|---|---|
| **Dataset principal** | Planilha bruta simulada (`.xlsx`), aba `Vendas_Bruto` (`13.613` itens tratados, `8.996` pedidos totais, `8.291` válidos). Aba de apoio `Ref_Cidades` (região e coordenadas de 24 cidades). | Histórico de pedidos com cliente, produto, valores, pagamento, entrega, status e avaliação |
| **API 1: câmbio** | [PTAX do Banco Central do Brasil](https://dadosabertos.bcb.gov.br/) (serviço Olinda — `CotacaoDolarPeriodo`), retorno em JSON (`608` cotações diárias com carry-forward para fins de semana/feriados: `418` oficiais + `190` carry-forward) | Relacionar o dólar (`R$ 4,8973` a `R$ 6,2086`) às vendas, ao ticket médio e à compressão de margem em eletrônicos |
| **API 2: clima** | [Open-Meteo Archive API](https://open-meteo.com/en/docs/historical-weather-api) (`/v1/archive`, `timezone=America/Sao_Paulo`), retorno em JSON (`14.592` registros = `24` cidades × `608` dias) | Cruzar precipitação diária e temperatura por cidade/dia com o atraso das entregas |
| **API 3: feriados** | [BrasilAPI — Feriados Nacionais](https://brasilapi.com.br/) (`/api/feriados/v1/{ano}`), retorno em JSON (`22` feriados nacionais em `2025–2026`: `14` em 2025 + `8` em jan–ago/2026) | Enriquecer `calendario`/`dim_tempo` (dia útil vs feriado) e auditar dias de `Carry-forward` no câmbio |
| **Banco relacional** | **SQLite 3 (`varejomax.db`, SQLite 3.50.4 + ODBC 64-bit DSN `varejomax`)** | Persistência normalizada em 3FN (`10` tabelas), views Star Schema (`9` views) e consultas analíticas (`K1`–`K5`, `D1`–`D4`, `C1`–`C10`) |

**Por que câmbio?** Eletrônicos concentram grande parte do faturamento (`42,69%`) e têm a menor margem do catálogo (`16,78%` ref. / `15,70%` realizada), o que os torna sensíveis ao dólar. A cotação (`dim_cambio`) é ligada à fato por `data_pedido`.

**Por que clima e feriados?** Para testar se a chuva forte (`>30 mm`) e o calendário de feriados explicam atrasos (`data_entrega_real` > `data_entrega_prevista`) por região. A precipitação é obtida pelas coordenadas de `Ref_Cidades` (`dCidade.csv`) e cruzada via chave composta `(cidade, data)` com `dim_clima`.

**Complementares:** a API de Localidades do IBGE e o ViaCEP permanecem como referências conceituais para dados cadastrais reais (os CEPs da base são simulados), enquanto a **BrasilAPI** foi efetivamente integrada em produção para o calendário oficial de feriados.

## Modelo dimensional

Modelo estrela evolutivo: iniciou na N1 com uma tabela fato e seis dimensões e foi expandido na entrega final (N2) para **1 tabela fato e 8 dimensões** (materializadas como views SQLite em [02-bancodedados-sql](MA2026-entrega-final/02-bancodedados-sql/README.md) e consumidas via Power Query M em [04-powerbi-dashboard](MA2026-entrega-final/04-powerbi-dashboard/README.md)).

**Tabela fato — `fato_vendas`** (uma linha por item de pedido: `id_pedido` + `item`, `13.613` linhas)
quantidade, desconto, valor bruto/líquido, custo referência/realizado, frete, datas de pedido, envio, entrega prevista e entrega real, status, atraso, avaliação e `chave_clima`.

**Dimensões (8 tabelas)**

| Dimensão | Conteúdo |
|---|---|
| `dim_produto` | produto, categoria, marca, preços de referência (`2025`/`2026`) e `custo_ref` (`49` linhas) |
| `dim_cliente` | cadastro, e-mail, data de cadastro, cidade, UF e região (`1.607` linhas) |
| `dim_cidade` | UF, região, latitude, longitude e indicador de loja física (`24` linhas) |
| `dim_tempo` | calendário completo de `2025-01-01` a `2026-08-31` com feriados BrasilAPI (`608` linhas) |
| `dim_canal_loja` | tipo de canal, canal e loja (`10` linhas: `3` online + `7` lojas físicas) |
| `dim_transportadora` | transportadora e tipo (`6` linhas) |
| `dim_cambio` | cotação diária PTAX compra/venda, origem do boletim e `faixa_dolar` (`608` linhas) |
| `dim_clima` | temperatura máx/mín, precipitação e `faixa_precipitacao` por cidade/dia (`14.592` linhas) |

## Indicadores (KPIs)

Implementados em **DAX** e conferidos contra consultas **SQL**.

| Indicador | Definição |
|---|---|
| Faturamento | Soma do valor dos itens dos pedidos não cancelados |
| Quantidade de vendas | Unidades vendidas e número de pedidos distintos |
| Ticket médio | Faturamento ÷ contagem distinta de pedidos |
| Margem bruta | Valor dos itens − (custo unitário × quantidade) |
| Produtos mais vendidos | Ranking por quantidade e por faturamento |
| Participação percentual | Peso de cada categoria e de cada canal sobre o total |
| Prazo de entrega | Média de dias entre envio e entrega real; % de entregas atrasadas (real > prevista) |
| Cancelamento e devolução | % de pedidos cancelados e devolvidos |
| Clientes | Novos e recorrentes, taxa de recompra e avaliação média (nota de 1 a 5) |

## Premissas de tratamento

- Pedidos **cancelados** não entram no faturamento.
- **Devoluções** são analisadas à parte.
- O **frete** é lançado apenas no item 1 de cada pedido e tratado separadamente do faturamento de produtos.
- **Nota de avaliação em branco** significa que o cliente não avaliou.

## Estrutura do repositório

```text
MA2026-projeto-pratico/
├── README.md                                  # Visão geral executiva do projeto
├── MA2026-entrega-parcial/                    # N1 concluída e congelada (referência histórica)
│   ├── 01-excel/                              # Planilha auditada, PROCX e log de inconsistências
│   ├── 02-powerquery/                         # Consultas em M e fontes tratadas
│   ├── 03-powerquery/                         # Artefatos complementares N1
│   └── 04-documentação/                       # Relatório parcial
└── MA2026-entrega-final/                      # N2 — Entrega Final integrada
    ├── README.md                              # Índice executivo, reconciliação e checklist final
    ├── ROTEIRO_DEFESA_FINAL.md                # Roteiro 15 min, Q&A da banca e plano de commits
    ├── 01-vba/                                # Template .xlsx, 3 macros .bas, CSVs ref. e conferência
    ├── 02-bancodedados-sql/                   # varejomax.db (SQLite 3), DDL/carga/views/K1-K5/D1-D4/C1-C10, diagrama ER
    ├── 03-api-web/                            # 3 APIs REST (PTAX, Open-Meteo, BrasilAPI), scripts Python/M e .xlsx
    └── 04-powerbi-dashboard/                  # Consultas M (.pq), DAX, tema JSON, diagrama e guia de montagem .pbix
```

- Índice da Entrega Final (N2): [MA2026-entrega-final/README.md](MA2026-entrega-final/README.md)
- Roteiro de Defesa Final (15 min + Q&A): [MA2026-entrega-final/ROTEIRO_DEFESA_FINAL.md](MA2026-entrega-final/ROTEIRO_DEFESA_FINAL.md)
- Módulos N2: [01-vba](MA2026-entrega-final/01-vba/README.md) · [02-bancodedados-sql](MA2026-entrega-final/02-bancodedados-sql/README.md) · [03-api-web](MA2026-entrega-final/03-api-web/README.md) · [04-powerbi-dashboard](MA2026-entrega-final/04-powerbi-dashboard/README.md)

## Cronograma e entregas

| Etapa | Foco técnico | Entregável principal | Aula final |
|---|---|---|---|
| Etapa 1 | Excel avançado, PROCX, auditoria e Power Query | Planilha limpa e tratada (`01-excel`, `02-powerquery`) com registro dos problemas e decisões | Aula 5 |
| Etapa 2 | Power Pivot (Star Schema) e DAX inicial | Modelo dimensional e medidas iniciais (`03-powerpivot`) | Aula 7 |
| **Parcial 1** | Defesa parcial e relatório | Apresentação oral (10–15 min), repositório congelado e relatório em `04-documentacao` | Aula 8 |
| Etapa 3 | Automação VBA | Macros de resumo de vendas e pedidos atrasados (`01-vba`) | Aula 11 |
| Etapa 4 | Modelagem de BD (DDL) e consultas (DML) | Script `.sql` com tabelas e consultas analíticas (`02-bancodedados-sql`) | Aula 13 |
| Etapa 5 | Consumo de API REST | Conexão documentada com câmbio e clima (`03-api-web`) | Aula 15 |
| Checkpoint | Validação técnica | Feedback do docente, sem nota isolada | Aula 16 |
| Etapa 6 | Dashboard Power BI e relatório | Painel interativo (`04-powerbi-dashboard`) | Aula 18 |
| **Final** | Entrega final e demonstração ao vivo | Defesa completa das 4 ferramentas ao vivo | Aula 20 |

## Resultados esperados

- Planilha auditada e tratada, com log dos problemas encontrados e das decisões de tratamento.
- Modelo estrela com uma tabela fato e ao menos cinco dimensões, com medidas DAX conferidas contra consultas SQL.
- Ao menos três macros VBA funcionais e documentadas.
- Banco relacional com script DDL e no mínimo dez consultas DML analíticas.
- Consumo de ao menos uma API pública, integrada ao modelo e documentada (a segunda API é meta opcional).
- Dashboard no Power BI com ao menos cinco páginas: Visão Executiva, Produtos e Categorias, Clientes, Logística e Canais.

## Validação das 4 hipóteses — principais achados N2

Os dados consolidados em [varejomax_consultas.sql](MA2026-entrega-final/02-bancodedados-sql/sql/varejomax_consultas.sql) e conferidos em [conferencia_resultados_vba.md](MA2026-entrega-final/01-vba/evidencias/conferencia_resultados_vba.md) e [resumo_validacao.md](MA2026-entrega-final/02-bancodedados-sql/evidencias/resumo_validacao.md) (`R$ 3.448.872,47` de faturamento líquido em `8.291` pedidos válidos) validam as 4 frentes investigadas:

- **H1 (Logística e Experiência):** Dos `8.180` pedidos com status `Entregue`, `628` sofreram atraso (**7,68%** sobre entregues; **7,57%** sobre os `8.291` pedidos válidos). A transportadora `EntregaJá` lidera os atrasos (**15,75%**, `306/1.943`), enquanto a `Veloz Express` apresenta apenas **2,79%** (`61/2.187`). O atraso derruba a avaliação do cliente em todas as transportadoras — na `EntregaJá`, a nota média cai de **4,27** (no prazo) para **3,41** (com atraso).
- **H2 (Regional e Expansão Física):** `São Paulo` concentra **26,14%** da receita (**R$ 901.513,58**, apenas **3,27%** de atraso). Em contrapartida, entre as cidades do Top 10 nacional localizadas no Nordeste sem loja física, os atrasos disparam: `Recife` (**24,43%**), `Salvador` (**20,59%**) e `Fortaleza` (**20,10%**). Todos os **15** pedidos críticos com atraso `>= 5` dias (máximo de **11 dias** no pedido `PED-005672`, Manaus/AM) estão concentrados no Norte, Nordeste e Centro-Oeste.
- **H3 (Mix, Margem e Câmbio):** `Eletrônicos` responde por **42,69%** do faturamento líquido (**R$ 1.472.456,60**), mas registra a pior margem do portfólio (**16,78%** de referência / **15,70%** realizada, abaixo do limiar de `20%`), em contraste com `Moda` (**48,32%** ref. / **47,64%** realizada). Cruzando com a **API PTAX (BCB)**, a margem realizada de `Eletrônicos` sofre compressão cambial de **17,03%** (dólar `< R$ 5,10`) para **14,37%** (dólar `>= R$ 6,00`).
- **H4 (Omnichannel, Fidelização e Fatores Externos):** Canais online (`Site` **46,07%**, `App` **24,87%**, `Marketplace` **18,97%**) geram **89,91%** da receita frente a **10,09%** das `7` lojas físicas (cancelamentos giram em torno de ~`4%` no online contra **0,00%** nas lojas físicas). Na base de **1.571** clientes ativos (**71,4%** de taxa de recompra), os **278** clientes `VIP (8+ pedidos)` (**17,70%**) geram **57,06%** do faturamento (**R$ 1.968.032,01**), com pico sazonal em Nov/2025 (`Black Friday`, `+99,50%` MoM). Cruzando com **Open-Meteo + BrasilAPI**, chuva forte (`>30 mm`) eleva a taxa de atraso de **4,67%** (dia útil sem chuva) para **14,04%** (dia útil) e **16,67%** (feriado).

## Status do projeto

### Entregas e artefatos construídos

- [x] **Etapa 1 & 2 / Parcial 1 (N1)** — Planilha tratada, Power Query, Power Pivot/DAX e relatório parcial concluídos e congelados em `MA2026-entrega-parcial/`.
- [x] **Etapa 3 (VBA — [01-vba](MA2026-entrega-final/01-vba/README.md))** — Template `VarejoMax_04_VBA_Template.xlsx` (6 abas), 3 módulos `.bas` em `Windows-1252`/CRLF (`modResumoVendas.bas`, `modAtrasosLogistica.bas`, `modUtilitariosExportacao.bas`), CSVs de referência e [conferencia_resultados_vba.md](MA2026-entrega-final/01-vba/evidencias/conferencia_resultados_vba.md) (`Diferença = R$ 0,00`).
- [x] **Etapa 4 (SQL — [02-bancodedados-sql](MA2026-entrega-final/02-bancodedados-sql/README.md))** — Banco SQLite 3 (`varejomax.db`, SQLite 3.50.4), scripts `varejomax_ddl.sql` (3FN), `varejomax_carga.sql` (idempotente 2×), `varejomax_views.sql` (9 views Star Schema), `K1`–`K5`, `D1`–`D4`, `C1`–`C10`, `diagrama_er.png` e [resumo_validacao.md](MA2026-entrega-final/02-bancodedados-sql/evidencias/resumo_validacao.md).
- [x] **Etapa 5 (APIs REST — [03-api-web](MA2026-entrega-final/03-api-web/README.md))** — Coleta real e scripts Python/Power Query M para BCB/PTAX (`608` dias), Open-Meteo Archive (`14.592` registros) e BrasilAPI (`22` feriados), planilha `VarejoMax_05_APIs_Web.xlsx` e [especificacao_apis.md](MA2026-entrega-final/03-api-web/docs/especificacao_apis.md).
- [x] **Etapa 6 (Infraestrutura Power BI — [04-powerbi-dashboard](MA2026-entrega-final/04-powerbi-dashboard/README.md))** — `11` scripts Power Query M (`.pq`, cultura `"en-US"`), `medidas_finais.dax` (8 pastas + `TREATAS`), `validacao.dax`, `tema_varejomax.json`, `diagrama_modelo.png`, [modelo_powerbi.md](MA2026-entrega-final/04-powerbi-dashboard/docs/modelo_powerbi.md) e [guia_montagem_pbix.md](MA2026-entrega-final/04-powerbi-dashboard/docs/guia_montagem_pbix.md).
- [x] **Documentação Final e Roteiro de Defesa** — [MA2026-entrega-final/README.md](MA2026-entrega-final/README.md) e [MA2026-entrega-final/ROTEIRO_DEFESA_FINAL.md](MA2026-entrega-final/ROTEIRO_DEFESA_FINAL.md).

### Passos finais de compilação em GUI Windows (pré-apresentação)

- [ ] **Compilar o `.xlsm` no Excel Desktop:** abrir `MA2026-entrega-final/01-vba/VarejoMax_04_VBA_Template.xlsx`, importar os 3 módulos `.bas` de `01-vba/macros/` via `Alt + F11`, salvar como `VarejoMax_04_VBA.xlsm`, rodar `AtualizarBasesERelatorios` e exportar os PDFs via `ExportarRelatoriosPDF`.
- [ ] **Configurar ODBC 64-bit e montar o `.pbix`:** instalar o SQLite3 ODBC Driver 64-bit, criar o DSN `varejomax` apontando para `varejomax.db`, importar os `.pq` no Power BI Desktop, configurar `R1`..`R8`, aplicar `medidas_finais.dax` e `tema_varejomax.json`, montar as `5 + 3` páginas, executar `validacao.dax` preenchendo [validacao_resultados.md](MA2026-entrega-final/04-powerbi-dashboard/evidencias/validacao_resultados.md) e salvar `VarejoMax_Dashboard.pbix` + `VarejoMax_Dashboard.pdf`.

## Equipe

| Nome | Responsabilidade |
|---|---|
| _Nome do integrante_ | _Papel / etapas_ |
| _Nome do integrante_ | _Papel / etapas_ |
| _Nome do integrante_ | _Papel / etapas_ |

**Orientador / Docente:** Prof. Veríssimo

---

<sub>Projeto acadêmico da FATEC SP. Empresa e dados fictícios, sem relação com organizações reais.</sub>
