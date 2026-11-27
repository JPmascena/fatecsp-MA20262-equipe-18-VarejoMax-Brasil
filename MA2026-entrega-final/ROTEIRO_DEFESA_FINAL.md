# Roteiro de Defesa Final — VarejoMax Brasil (15 min)

> **Projeto Prático MA2026** · FATEC SP — Microinformática Aplicada · **Banca / Docente:** Prof. Veríssimo  
> **Documento de apoio:** [README.md (Entrega Final)](README.md) · [README.md (Raiz)](../README.md)  
> **Duração total:** **15 minutos** de demonstração ao vivo + espaço para arguição técnica.

---

## 1. Divisão de Papéis e Cronograma (15 min)

> *Nota:* Os nomes abaixo (`Diego`, `Daniela`, `Matheus`, `Vitor`) representam a divisão recomendada em 4 frentes técnicas para a apresentação — ajuste conforme os integrantes listados na seção `Equipe` de [README.md](../README.md).

| Bloco | Tempo | Duração | Responsável | Foco da Demonstração ao Vivo |
|---|---:|---:|---|---|
| **1. Contexto, Problema & Arquitetura** | 00:00–02:00 | 2 min | **Diego** | Dor de negócio da VarejoMax, 4 hipóteses e pipeline N1 → N2 |
| **2. Excel VBA & Reconciliação** | 02:00–05:00 | 3 min | **Daniela** | `VarejoMax_04_VBA.xlsm`, execução `< 1s`, filtros, PDF e batimento `R$ 0,00` |
| **3. APIs REST & SQLite 3FN** | 05:00–09:00 | 4 min | **Matheus** | 3 APIs (`608` / `14.592` / `22`), `construir_banco.py`, `K1`–`K5` e `C1`–`C10` |
| **4. Dashboard Power BI (`5 + 3` págs.)** | 09:00–13:00 | 4 min | **Vitor** | Navegação interativa, `TREATAS`, Drill-through, Tooltip e `Conferência` |
| **5. Recomendações & Conclusão** | 13:00–15:00 | 2 min | **Diego** *(+ equipe)* | Plano de ação baseado em H1–H4 e fechamento executivo |

---

## 2. Roteiro Detalhado por Bloco (O que Mostrar e Falar)

### Bloco 1 — Contexto, Problema e Arquitetura (00:00–02:00 · 2 min · Diego)
- **O que mostrar na tela:**
  1. [README.md](../README.md) (seção *Contexto* e *Problema*) e o diagrama de arquitetura em [MA2026-entrega-final/README.md](README.md).
  2. Breve menção à pasta congelada `MA2026-entrega-parcial/` (N1 concluída: saneamento de `13.613` linhas de itens / `8.996` pedidos / `24` cidades).
- **Fala objetiva:**
  - *"A VarejoMax Brasil é uma rede omnichannel simulada (`jan/2025` a `ago/2026`, `49` produtos em `6` categorias, `24` cidades). Na N1, recebemos uma planilha bruta com datas como texto, categorias inconsistentes, valores divergentes e registros de teste."*
  - *"Na N2, evoluímos para uma plataforma completa com quatro frentes integradas sob uma fonte única da verdade: automação operacional em Excel VBA, enriquecimento via 3 APIs REST públicas, banco relacional SQLite 3FN com views Star Schema e dashboard no Power BI via ODBC 64-bit, para responder às 4 hipóteses de vendas, logística, margem/câmbio e clima."*

---

### Bloco 2 — Demonstração ao Vivo: Excel VBA (`01-vba/`) (02:00–05:00 · 3 min · Daniela)
- **Pré-condição:** `VarejoMax_04_VBA.xlsm` aberto na aba `Menu_Operacional`, com os 3 módulos de [`01-vba/macros/`](01-vba/README.md) importados.
- **O que mostrar na tela:**
  1. Aba `Menu_Operacional` (limpa ou pronta para rodar) → clicar no botão / executar `AtualizarBasesERelatorios`.
  2. Mostrar o `MsgBox` instantâneo: **execução `< 1s`** com **`RECONCILIAÇÃO OFICIAL: APROVADA (Diferença = R$ 0,00)`**.
  3. Navegar para `Resumo_Categorias`: destacar **Quadro 1** (`Eletrônicos` com `R$ 1.472.456,60` / `42,69%` de share e alerta vermelho em `16,78% < 20%`) e **Quadro 2** (`Site 46,07%`, `App 24,87%`, `Marketplace 18,97%`, `Loja Física 10,09%`).
  4. Navegar para `Relatorio_Atrasos`: mostrar **Quadro 1** (`8.180` entregues, `628` atrasados = `7,68%`; `EntregaJá` com `15,75%` vs `Veloz Express` com `2,79%`) e **Quadro 2** (`15` pedidos críticos com `>= 5` dias destacados, topo `PED-005672` com `11` dias).
  5. *(Demonstração de filtro + PDF):* Filtrar `Menu_Operacional!C10 = "Nordeste"` ou `C9 = "EntregaJá"`, rodar `GerarRelatorioAtrasosTransportadora` e mostrar a exportação em PDF paisagem via `ExportarRelatoriosPDF`.
- **Fala objetiva:**
  - *"Processamos os 13.613 itens inteiramente em arrays na memória (`Scripting.Dictionary` e `QuickSort`), aplicando `Option Explicit`, `CleanExit` e as regras oficiais de negócio — evitando dupla contagem nos 358 pedidos com variação entre itens e filtrando notas fora de 1–5."*

---

### Bloco 3 — Engenharia de Dados: 3 APIs REST + SQLite 3FN (`03-api-web/` & `02-bancodedados-sql/`) (05:00–09:00 · 4 min · Matheus)
- **O que mostrar na tela:**
  1. [`03-api-web/docs/especificacao_apis.md`](03-api-web/docs/especificacao_apis.md) e [`03-api-web/evidencias/resumo_coleta_apis.md`](03-api-web/evidencias/resumo_coleta_apis.md) (ou rodar `python "MA2026-entrega-final/02-bancodedados-sql/scripts_build/construir_banco.py"` no terminal para provar reprodutibilidade).
  2. [`02-bancodedados-sql/docs/diagrama_er.png`](02-bancodedados-sql/docs/diagrama_er.png) mostrando as **10 tabelas 3FN** + **9 views Star Schema**.
  3. [`02-bancodedados-sql/evidencias/resumo_validacao.md`](02-bancodedados-sql/evidencias/resumo_validacao.md) e os CSVs [`k1.csv`](02-bancodedados-sql/evidencias/k1.csv), [`c5.csv`](02-bancodedados-sql/evidencias/c5.csv), [`c9.csv`](02-bancodedados-sql/evidencias/c9.csv) e [`c10.csv`](02-bancodedados-sql/evidencias/c10.csv).
- **Fala objetiva:**
  - *"Integramos 3 APIs REST públicas sem chave: **BCB PTAX** (`608` dias, sendo `418` boletins oficiais + `190` carry-forward em fins de semana/feriados), **Open-Meteo Archive** (`14.592` registros = `24` cidades × `608` dias, zero nulos) e **BrasilAPI** (`22` feriados nacionais de 2025–2026)."*
  - *"No SQLite 3 (`varejomax.db`), garantimos 3FN com `PRAGMA foreign_keys = ON`, carga **idempotente** (rodada 2× com contagens idênticas), zero órfãos em `K1`–`K4` e índice em todas as consultas `C1`–`C10` (`EXPLAIN QUERY PLAN`)."*
  - *"Nos achados SQL: `C9` mostra a margem realizada de `Eletrônicos` caindo de `17,03%` (dólar `< R$ 5,10`) para `14,37%` (dólar `>= R$ 6,00`); e `C10` comprova que chuva forte (`>30 mm`) eleva atrasos de `4,67%` (dia útil seco) para `14,04%` (dia útil) e `16,67%` (feriado)."*

---

### Bloco 4 — Painel Executivo Power BI (`04-powerbi-dashboard/`) (09:00–13:00 · 4 min · Vitor)
- **O que mostrar na tela:**
  1. Abrir `VarejoMax_Dashboard.pbix` (ou exibição **Modelo** mostrando `fato_vendas` ligada às `8` dimensões `R1`..`R8` sem ambiguidade).
  2. Percorrer as **5 páginas principais** interagindo com os segmentadores sincronizados (`ano`, `mes_abrev`, `regiao`, `canal`):
     - **`Visão Executiva`:** Cartões `R$ 3,45M` (`R$ 3.448.872,47`), `8.291` pedidos, `R$ 415,98` ticket, `31,2%` margem, `7,6%` atrasos; pico de Black Friday em Nov/2025 (`+99,5%` MoM); passar o mouse sobre uma barra para exibir a tooltip **`Dica Resumo` (`320x240`)**.
     - **`Produtos e Categorias`:** Dispersão (Eletrônicos no quadrante alto faturamento / baixa margem `16,78%` vs Moda `48,32%`) e matriz hierárquica `Categoria → Marca → Produto`.
     - **`Clientes`:** `1.571` ativos, `71,4%` taxa de recompra, Pareto `VIP (8+ pedidos)` (`17,70%` dos clientes gerando `57,06%` da receita) e mix de pagamento.
     - **`Logística e Clima`:** Queda da nota média no atraso (**H1**), barras de chuva (`Open-Meteo`), mapa nacional destacando Recife (`24,43%`), Salvador (`20,59%`) e Fortaleza (`20,10%`), e **Drill-through** em `EntregaJá` → **`Detalhe Transportadora / Loja`** + botão **Voltar**.
     - **`Canais e Lojas`:** Online (`89,91%`) vs `7` Lojas Físicas (`10,09%`, `0%` cancelamento) e série temporal `Dólar PTAX Médio × Eletrônicos` propagada via `TREATAS`.
  3. Abrir a aba oculta **`Conferência`** exibindo os deltas zerados (`Conf Diferença Faturamento vs Coluna = 0,00`, `8` avaliações inválidas auditadas, `0` pedidos sem clima/câmbio).
- **Fala objetiva:**
  - *"Todos os scripts M usam cultura `en-US` via ODBC `dsn=varejomax`. As medidas externas usam `TREATAS` para propagar filtros de `dim_tempo` e `dim_cidade` sem criar relacionamentos ambíguos no Star Schema."*

---

### Bloco 5 — Recomendações Estratégicas & Conclusão (13:00–15:00 · 2 min · Diego + equipe)
- **O que mostrar na tela:**
  - Permanecer na página `Logística e Clima` ou `Visão Executiva` do Power BI como pano de fundo executivo.
- **Recomendações práticas fundamentadas nos dados:**
  1. **Logística & SLA (H1 + H2):** Renegociar SLA e redistribuir volume da `EntregaJá` (`15,75%` de atraso) para `Veloz Express` (`2,79%`) ou `RotaBR` (`7,04%`) nas praças críticas do Nordeste (`Recife 24,43%`, `Salvador 20,59%`, `Fortaleza 20,10%`), priorizando um **CD avançado ou loja física âncora no Nordeste** (hoje sem presença física entre o Top 10).
  2. **Margem & Câmbio (H3):** Implantar gatilho de repasse dinâmico / hedge de compras em `Eletrônicos` quando o dólar PTAX superar `R$ 5,70` (onde a margem realizada cai para `15,63%` e `14,37%` acima de `R$ 6,00`), combinado com *cross-sell* de categorias de alta margem (`Moda` `47,64%` e `Beleza e Saúde` `42,83%`).
  3. **Omnichannel & Fidelização (H4):** Expandir programa de fidelidade para converter clientes `Recorrentes (2–4 pedidos)` rumo ao perfil `VIP (8+ pedidos)` (que já sustenta `57,06%` da receita) e adotar *buffer* automático de prazo (+1/+2 dias previstos) em rotas sob alerta de chuva forte (`>30 mm`) ou véspera de feriados/Black Friday.
- **Fechamento:** *"Toda a trilha é 100% auditável e reproduzível do dado bruto ao Power BI com diferença zero entre Excel, VBA, SQL e DAX. Estamos abertos às perguntas da banca."*

---

## 3. Banco de Perguntas & Respostas Técnicas para a Banca (Prof. Veríssimo)

| # | Possível Pergunta da Banca | Resposta Técnica Direta (com Evidências Reais) |
|---|---|---|
| **1** | *Por que existem **358 pedidos** com divergência entre itens do mesmo `ID_Pedido` e como vocês trataram sem distorcer os números?* | Em [`d1.csv`](02-bancodedados-sql/evidencias/d1.csv) mapeamos exatamente `358` pedidos com variação entre itens: `231` em `CEP`, `85` em `Transportadora`, `40` em `Forma_Pagamento`, `15` em `Data_Entrega_Real` e `4` em `Avaliacao_Cliente`. Na 3FN (`pedido` vs `item_pedido`), atributos variáveis por item ficaram em `item_pedido` para evitar perda de informação; na consolidação logística (`C5` e VBA [`modAtrasosLogistica.bas`](01-vba/macros/modAtrasosLogistica.bas)), os `85` casos de `Transportadora` ocorrem apenas quando um item traz `"Não informado"` e outro traz a transportadora efetiva — priorizamos a transportadora informada, reduzindo `"Não informado"` de `80` para `49` pedidos sem duplicar `ID_Pedido`. |
| **2** | *Como foram tratadas as **8 avaliações fora de 1–5** e as **18 linhas `Inconsistente`** (`data_entrega_real < data_envio`)?* | Preservamos todos os registros no banco (`13.613` linhas, sem exclusão silenciosa) e aplicamos filtragem explícita nas métricas ([`d2.csv`](02-bancodedados-sql/evidencias/d2.csv)): **(a)** As `8` notas inválidas (`0`, `6`, `7`, `8`, `9`, `10`) são filtradas por `1 <= nota <= 5` em VBA, SQL e DAX `[Nota Média Cliente]` e monitoradas por `[Conf Avaliações Inválidas] = 8`; **(b)** As `18` linhas `Inconsistente` (`17` entregues + `1` devolvido) são excluídas do `Prazo Real Médio` e não são rotuladas como `Atrasado`. |
| **3** | *Por que o projeto adotou **SQLite 3 + ODBC 64-bit** em vez de **PostgreSQL**?* | O **SQLite 3 (`3.50.4`)** oferece suporte completo a 3FN (`PK`, `FK`, `CHECK`, `UNIQUE`, índices), CTEs recursivas e Window Functions em um único arquivo versionável (`varejomax.db`, `journal_mode=DELETE`), reconstruído do zero em segundos por [`construir_banco.py`](02-bancodedados-sql/scripts_build/construir_banco.py) (apenas Python stdlib) e conectado ao Power BI Desktop via `SQLite3 ODBC Driver 64-bit` (`dsn=varejomax`). Isso garante **100% de portabilidade offline** na apresentação ao vivo, sem risco de falha de rede ou credenciais de nuvem. |
| **4** | *Por que não ligar `dim_cambio` diretamente a `dim_tempo`, ou `dim_clima` a `dim_cidade` e `dim_tempo` no Power BI?* | Porque `fato_vendas` já se relaciona com `dim_tempo` (`R6`), `dim_cidade` (`R3`), `dim_clima` (`R7`) e `dim_cambio` (`R8`). Ligar dimensões entre si criaria **caminhos de filtro ambíguos** no grafo do Power BI. Conforme documentado em [`modelo_powerbi.md`](04-powerbi-dashboard/docs/modelo_powerbi.md) e [`medidas_finais.dax`](04-powerbi-dashboard/dax/medidas_finais.dax), mantemos todos os `8` relacionamentos `*:1` em direção única e usamos **`TREATAS`** nas medidas `[Dólar PTAX Médio]`, `[Temp Máx Média]`, `[Temp Mín Média]` e `[Precipitação Média mm]` para propagar filtros do usuário com segurança. |
| **5** | *Por que a **Nota Média do Cliente** aparece como `4,11` no Power BI, `4,18` em `k5.csv` e `4,21` no VBA / `C5`?* | Trata-se de 3 granularidades complementares auditadas em [`k5.csv`](02-bancodedados-sql/evidencias/k5.csv) e [`resumo_validacao.md`](02-bancodedados-sql/evidencias/resumo_validacao.md), todas filtrando `1..5`: **`4,11`** é a média por linha avaliada considerando todos os status (padrão Etapa 3 e DAX `[Nota Média Cliente]`); **`4,18`** é a média por linha avaliada excluindo `Cancelado`/`Devolvido`; e **`4,21`** é a média consolidada por pedido `Entregue` (`3.517` pedidos avaliados — padrão VBA e consulta `C5`). |
| **6** | *Por que `C10` soma **`8.296`** pedidos válidos em vez de `8.291`, e por que `Frete` e `% Atraso` têm duas leituras?* | Em [`c10.csv`](02-bancodedados-sql/evidencias/c10.csv), a soma por `(faixa_precipitacao, tipo_dia_pedido)` dá `8.296` (`+5`) porque `5` pedidos tiveram itens entregues em dias distintos que caíram em faixas climáticas/tipos de dia diferentes na `chave_clima` (`COALESCE(data_entrega_real, data_pedido)`). Já no **Frete** (lançado apenas no `Item = 1`, `0` ocorrências em `Item <> 1` conforme [`d3.csv`](02-bancodedados-sql/evidencias/d3.csv)), `R$ 97.922,36` inclui todos os `8.996` pedidos (`[Frete Total]`) enquanto `R$ 93.707,67` considera só os `8.291` válidos (`K1`/`K5` e `[Frete Pedidos Válidos]`). E no **% Atraso**, `628 / 8.180 = 7,68%` sobre entregues (VBA/`C5`) vs `628 / 8.291 = 7,57%` (`7,6%`) no DAX sobre todos os válidos. |
| **7** | *Por que `data_entrega_prevista` não tem chave estrangeira (`FK`) para a tabela `calendario`?* | Conforme [`d4.csv`](02-bancodedados-sql/evidencias/d4.csv), `data_pedido` vai de `2025-01-01` a `2026-08-31` (`608` dias), mas `151` pedidos feitos no fim de agosto/2026 possuem `data_entrega_prevista` em setembro/2026 (máximo `2026-09-11`), além da janela comercial de `calendario`. |

---

## 4. Plano de Commits Git Multi-Autor (Preservando `MA2026-entrega-parcial/`)

Para organizar o histórico do repositório de forma limpa e rastreável por frente técnica antes do *push* final (sem tocar em nenhum arquivo de `MA2026-entrega-parcial/`):

```powershell
# 0. Verificar que MA2026-entrega-parcial/ não possui alterações
git status MA2026-entrega-parcial/

# 1. Frente APIs Web (consumo das 3 APIs, scripts M e planilha auxiliar)
git add MA2026-entrega-final/03-api-web/
git commit -m "feat(api-web): integra BCB PTAX, Open-Meteo Archive e BrasilAPI com scripts M e evidencias"

# 2. Frente VBA (template, 3 macros ANSI/CRLF e conferência R$ 0,00)
git add MA2026-entrega-final/01-vba/
git commit -m "feat(vba): adiciona template e macros de resumo de vendas, atrasos e reconciliacao"

# 3. Frente Banco de Dados SQLite 3FN + Star Schema + Consultas C1-C10
git add MA2026-entrega-final/02-bancodedados-sql/
git commit -m "feat(sql): implementa banco SQLite 3FN, views Star Schema, K1-K5, D1-D4 e C1-C10"

# 4. Frente Power BI (queries M en-US, DAX, tema, diagrama e guia de montagem)
git add MA2026-entrega-final/04-powerbi-dashboard/
git commit -m "feat(powerbi): adiciona consultas M, medidas DAX com TREATAS, validacao e tema corporativo"

# 5. Documentação executiva, roteiro de defesa final e .gitignore (se aplicável)
git add .gitignore README.md MA2026-entrega-final/README.md MA2026-entrega-final/ROTEIRO_DEFESA_FINAL.md
git commit -m "docs: atualiza README raiz, indice N2 com reconciliacao e roteiro de defesa final"
```

> **Lembrete pós-compilação no Windows:** Após gerar os binários locais `MA2026-entrega-final/01-vba/VarejoMax_04_VBA.xlsm` (+ PDFs) e `MA2026-entrega-final/04-powerbi-dashboard/VarejoMax_Dashboard.pbix` (+ `.pdf`), faça um commit final de entrega binária (`chore(release): adiciona VarejoMax_04_VBA.xlsm e VarejoMax_Dashboard.pbix/.pdf compilados`) e marque os checkboxes pendentes em `README.md`, `MA2026-entrega-final/README.md` e `validacao_resultados.md`.
