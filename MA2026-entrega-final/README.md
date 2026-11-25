# VarejoMax Brasil — Entrega Final (N2)

> **Projeto Prático MA2026** · FATEC SP — Microinformática Aplicada · Orientação: **Prof. Veríssimo**  
> **Escopo:** Integração completa das quatro frentes finais — **Excel VBA**, **APIs REST**, **Banco de Dados Relacional (SQLite 3FN + Views Star Schema)** e **Dashboard Power BI** — sobre os dados tratados da Entrega Parcial (N1).  
> **Roteiro de apresentação:** [ROTEIRO_DEFESA_FINAL.md](ROTEIRO_DEFESA_FINAL.md) · **Visão geral do repositório:** [README.md](../README.md)

---

## 1. Arquitetura de Ponta a Ponta (N1 → N2)

```text
┌───────────────────────────┐     ┌────────────────────────────────────┐
│  N1 — Entrega Parcial     │     │  03-api-web/ (3 APIs REST)         │
│  Excel Avançado + PQ (M)  │     │  • BCB Olinda PTAX (608 dias)      │
│  13.613 itens / 8.996 ped │     │  • Open-Meteo Archive (14.592 reg) │
└─────────────┬─────────────┘     │  • BrasilAPI Feriados (22 reg)     │
              │                   └──────────────────┬─────────────────┘
              ▼                                      │
┌───────────────────────────┐                        │
│  01-vba/                  │                        │
│  Template .xlsx + 3 .bas  │                        ▼
│  Resumo Categorias/Canais │     ┌────────────────────────────────────┐
│  Atrasos & Ped. Críticos  │     │  02-bancodedados-sql/              │
│  Reconciliação Automática │     │  SQLite 3.50.4 (varejomax.db)      │
└───────────────────────────┘     │  • 10 tabelas 3FN + 9 tabelas stg  │
                                  │  • 9 Views Star Schema             │
                                  │  • Auditoria K1–K5, D1–D4, C1–C10  │
                                  └──────────────────┬─────────────────┘
                                                     │ ODBC 64-bit (dsn=varejomax)
                                                     ▼
                                  ┌────────────────────────────────────┐
                                  │  04-powerbi-dashboard/             │
                                  │  11 scripts M (en-US) + DAX        │
                                  │  8 relacionamentos + TREATAS       │
                                  │  5 págs. principais + 3 de apoio   │
                                  └────────────────────────────────────┘
```

### Papéis de cada camada na Entrega Final
1. **Base saneada (N1 → CSVs em `02-bancodedados-sql/csv/`):** `13.613` linhas de itens (`8.996` pedidos totais, `8.291` válidos após excluir `355` cancelados e `350` devolvidos), preservando integralmente as anomalias mapeadas na auditoria (`8` avaliações fora de `1..5`, `18` linhas `Inconsistente` com `data_entrega_real < data_envio` e `358` pedidos com variação entre itens).
2. **Automação Operacional ([`01-vba/`](01-vba/README.md)):** 3 módulos VBA em memória (`Scripting.Dictionary` + `QuickSort`) rodando em `< 1s`, com filtros dinâmicos em `Menu_Operacional`, geração de quadros por Categoria, Canal, Transportadora e `628` pedidos atrasados, além de batimento automático (`Diferença = R$ 0,00`) e exportação em PDF paisagem.
3. **Enriquecimento Externo ([`03-api-web/`](03-api-web/README.md)):** Coleta real sem chave de API via Python stdlib e funções Power Query M (`RelativePath`/`Query`), unificando câmbio diário PTAX (`608` dias: `418` oficiais + `190` carry-forward), clima histórico das `24` cidades (`14.592` registros, `0` nulos) e `22` feriados nacionais (`2025–2026`).
4. **Persistência & SQL Analítico ([`02-bancodedados-sql/`](02-bancodedados-sql/README.md)):** Banco **SQLite 3** (`varejomax.db`, `3.50.4`) em 3FN (`PRAGMA foreign_keys = ON`, `PRAGMA integrity_check = ok`, carga idempotente 2×), expondo `9` views em Star Schema e validando `K1`–`K5`, `D1`–`D4` e `C1`–`C10`.
5. **Painel Executivo ([`04-powerbi-dashboard/`](04-powerbi-dashboard/README.md)):** Modelo conectado via ODBC 64-bit (`dsn=varejomax`) com cultura `"en-US"`, `8` relacionamentos `*:1`, propagação externa via `TREATAS` (sem ambiguidade entre dimensões), tema corporativo `#FAF7F2` / `#1E3A5F` / `#D95D39` e bateria `validacao.dax`.

---

## 2. Tabela de Reconciliação Inter-Etapas (`Etapa 3` × `VBA` × `SQLite K1–K5` × `Power BI DAX`)

Todos os módulos foram conciliados contra os mesmos gabaritos oficiais, garantindo **diferença zero** nos totais financeiros e volumétricos e explicitando as diferenças intencionais de granularidade entre visões de item e de pedido.

| Métrica / Indicador | Referência (Etapa 3 / N1–N2) | Excel VBA (`01-vba`) | SQLite (`K1`–`K5` / `C1`–`C10`) | Power BI (`medidas_finais.dax`) | Regra de Granularidade / Observação |
|---|---|---|---|---|---|
| **Linhas na tabela fato** | `13.613` | `13.613` | `13.613` (`K1`) | `13.613` (`[Conf Linhas Fato]`) | 1 linha por `(id_pedido, item)` |
| **Pedidos totais distintos** | `8.996` | `8.996` | `8.996` (`K1`/`K5`) | `8.996` (`[Qtd Pedidos Total]`) | Inclui `355` cancelados (`3,95%`) e `350` devolvidos (`3,89%`) |
| **Pedidos válidos** | `8.291` | `8.291` | `8.291` (`K1`/`K5`) | `8.291` (`[Qtd Pedidos]`) | Exclui `Cancelado` e `Devolvido` |
| **Unidades vendidas (válidos)** | `16.581` | `16.581` | `16.581` (`K1`/`K5`) | `16.581` (`[Qtd Itens]`) | Soma de `quantidade` nos `8.291` válidos (`17.946` totais) |
| **Faturamento Bruto** | `R$ 3.797.167,40` | `R$ 3.797.167,40` | `R$ 3.797.167,40` (`K1`/`K5`) | `R$ 3.797.167,40` (`[Faturamento Bruto]`) | Exclui apenas `Cancelado` *(em `c2.csv`, `R$ 3.628.062,60` reflete bruto apenas dos válidos para comparar com o líquido)* |
| **Faturamento Líquido** | `R$ 3.448.872,47` | `R$ 3.448.872,47` | `R$ 3.448.872,47` (`K1`/`K5`) | `R$ 3.448.872,47` (`[Faturamento]`) | Exclui `Cancelado` e `Devolvido`; `Diferença = R$ 0,00` |
| **Ticket Médio** | `R$ 415,98` | `R$ 415,98` | `R$ 415,98` (`K5`) | `R$ 415,98` (`[Ticket Médio]`) | `Faturamento Líquido ÷ 8.291` |
| **Custo Ref. Total / Margem Ref.** | `R$ 2.372.700,48` (`31,2%`) | `R$ 2.372.700,48` (`31,20%`) | `R$ 2.372.700,48` (`31,20%`) | `R$ 2.372.700,48` (`31,2%`) | Lucro Bruto = `R$ 1.076.171,99` |
| **Custo Realizado / Margem Real.** | `R$ 2.402.823,74` (`30,3%`) | `R$ 2.402.823,74` (`30,33%`) | `R$ 2.402.823,74` (`30,33%`) | `R$ 2.402.823,74` (`30,3%`) | Custo unitário efetivo da venda na fato |
| **Frete Total (lançado no Item 1)** | `R$ 97.922,36` *(geral)* / `R$ 93.707,67` *(válidos)* | — | `R$ 93.707,67` (`K1`/`K5`, válidos) | `R$ 97.922,36` (`[Frete Total]`) / `R$ 93.707,67` (`[Frete Pedidos Válidos]`) | **`R$ 97.922,36`** soma todos os pedidos (`4.948` itens com frete `> 0` em `Item = 1`); **`R$ 93.707,67`** restringe aos `8.291` válidos |
| **Pedidos Atrasados & % Atraso** | `628` (`7,6%` s/ válidos; `7,68%` s/ entregues) | `628` / `8.180` entregues (**`7,68%`**) | `628` / `8.180` entregues (**`7,68%`**, `K5`/`C5`) | `628` / `8.291` válidos (**`7,57%` → `7,6%`**) | VBA e `C5` dividem pelos `8.180` `Entregue`; DAX `[% Pedidos Atrasados]` divide pelos `8.291` válidos (inclui `78` `Em trânsito` + `33` `Em separação`); `654` inclui `26` atrasos `Devolvido` |
| **Nota Média Cliente (1 a 5)** | `4,11` *(item geral)* | `4,21` *(pedido entregue)* | `4,11` *(item)* / `4,18` *(válido)* / `4,21` *(entregue)* (`K5`) | `4,11` (`[Nota Média Cliente]`) | Todas descartam `8` notas inválidas (`0` e `6..10`): **`4,11`** é média por item avaliado (todos status); **`4,18`** exclui `Cancelado`/`Devolvido`; **`4,21`** consolida 1 nota por pedido `Entregue` (`3.517` pedidos) |
| **Clientes Ativos / Novos / Recor.** | `1.571` / `449` / `1.122` (`71,4%`) | — | `1.571` (`449` + `612` + `232` + `278` em `C7`; `1.607` em `dim_cliente`) | `1.571` / `449` / `1.122` (`71,4%`) | `36` cadastrados sem pedido válido; `278` VIPs (`17,70%`) geram `57,06%` da receita |
| **Cobertura APIs & Clima (`C10`)** | `0` órfãos | — | `608` PTAX / `14.592` Clima / `22` Feriados (`K4` = `0` órfãos) | `0` sem clima / `0` sem câmbio (`V-K5.10`/`11`) | Em `C10`, a soma por `(faixa_precipitacao, tipo_dia_pedido)` totaliza **`8.296`** (`+5` sobre `8.291`) pois `5` pedidos possuem itens entregues em dias distintos na `chave_clima` |

---

## 3. Inventário Completo da Entrega Final (`MA2026-entrega-final/`)

### Documentação Executiva e Roteiro
- [`README.md`](README.md) — Este índice executivo, arquitetura e reconciliação inter-etapas.
- [`ROTEIRO_DEFESA_FINAL.md`](ROTEIRO_DEFESA_FINAL.md) — Roteiro cronometrado de 15 minutos, falas, telas, banco de Q&A da banca e plano de commits Git.

### `01-vba/` — Automação Operacional em Excel VBA
- [`01-vba/README.md`](01-vba/README.md) — Guia dos módulos, abas e instruções para gerar o `.xlsm`.
- [`01-vba/VarejoMax_04_VBA_Template.xlsx`](01-vba/VarejoMax_04_VBA_Template.xlsx) — Pasta de trabalho modelo (6 abas populadas e resultados pré-calculados).
- Macros (ANSI `Windows-1252`, CRLF): [`modResumoVendas.bas`](01-vba/macros/modResumoVendas.bas) · [`modAtrasosLogistica.bas`](01-vba/macros/modAtrasosLogistica.bas) · [`modUtilitariosExportacao.bas`](01-vba/macros/modUtilitariosExportacao.bas)
- Evidências e referências: [`conferencia_resultados_vba.md`](01-vba/evidencias/conferencia_resultados_vba.md) · [`resumo_vendas_por_categoria_referencia.csv`](01-vba/relatorios_gerados/resumo_vendas_por_categoria_referencia.csv) · [`relatorio_atrasos_por_transportadora_referencia.csv`](01-vba/relatorios_gerados/relatorio_atrasos_por_transportadora_referencia.csv)

### `02-bancodedados-sql/` — Banco Relacional SQLite 3FN & Views Star Schema
- [`02-bancodedados-sql/README.md`](02-bancodedados-sql/README.md) — Guia de reconstrução, conexão ODBC e resumo de `C1`–`C10`.
- [`02-bancodedados-sql/varejomax.db`](02-bancodedados-sql/varejomax.db) — Banco SQLite `3.50.4` validado (`journal_mode=DELETE`).
- Scripts SQL: [`varejomax_ddl.sql`](02-bancodedados-sql/sql/varejomax_ddl.sql) · [`varejomax_carga.sql`](02-bancodedados-sql/sql/varejomax_carga.sql) · [`varejomax_views.sql`](02-bancodedados-sql/sql/varejomax_views.sql) · [`varejomax_conferencia.sql`](02-bancodedados-sql/sql/varejomax_conferencia.sql) · [`varejomax_diagnostico.sql`](02-bancodedados-sql/sql/varejomax_diagnostico.sql) · [`varejomax_consultas.sql`](02-bancodedados-sql/sql/varejomax_consultas.sql)
- Builder & Docs: [`construir_banco.py`](02-bancodedados-sql/scripts_build/construir_banco.py) · [`modelo_relacional.md`](02-bancodedados-sql/docs/modelo_relacional.md) · [`diagrama_er.png`](02-bancodedados-sql/docs/diagrama_er.png)
- CSVs tratados e APIs (`9` arquivos em [`02-bancodedados-sql/csv/`](02-bancodedados-sql/csv/)): [`fVendas.csv`](02-bancodedados-sql/csv/fVendas.csv), [`dProduto.csv`](02-bancodedados-sql/csv/dProduto.csv), [`dCliente.csv`](02-bancodedados-sql/csv/dCliente.csv), [`dCidade.csv`](02-bancodedados-sql/csv/dCidade.csv), [`dCanalLoja.csv`](02-bancodedados-sql/csv/dCanalLoja.csv), [`dTransportadora.csv`](02-bancodedados-sql/csv/dTransportadora.csv), [`cambio_ptax_2025_2026.csv`](02-bancodedados-sql/csv/cambio_ptax_2025_2026.csv), [`clima_24_cidades_2025_2026.csv`](02-bancodedados-sql/csv/clima_24_cidades_2025_2026.csv), [`feriados_nacionais_2025_2026.csv`](02-bancodedados-sql/csv/feriados_nacionais_2025_2026.csv)
- Evidências ([`02-bancodedados-sql/evidencias/`](02-bancodedados-sql/evidencias/)): [`resumo_validacao.md`](02-bancodedados-sql/evidencias/resumo_validacao.md), `k1`–`k5.csv`, `d1`–`d4.csv`, `c1`–`c10.csv`, [`sqlite_master.csv`](02-bancodedados-sql/evidencias/sqlite_master.csv) e [`explain_query_plan.txt`](02-bancodedados-sql/evidencias/explain_query_plan.txt)

### `03-api-web/` — Consumo das 3 APIs Públicas REST
- [`03-api-web/README.md`](03-api-web/README.md) — Resumo das fontes, padrão CSV (`en-US`) e reprodução.
- [`03-api-web/VarejoMax_05_APIs_Web.xlsx`](03-api-web/VarejoMax_05_APIs_Web.xlsx) — Planilha consolidada com 4 abas gerada via OpenXML.
- Scripts Python & Power Query M: [`coletar_apis.py`](03-api-web/scripts_coleta/coletar_apis.py) · [`gerar_excel_apis.py`](03-api-web/scripts_coleta/gerar_excel_apis.py) · [`fnPTAX_BCB.pq`](03-api-web/scripts_m/fnPTAX_BCB.pq) · [`fnClima_OpenMeteo.pq`](03-api-web/scripts_m/fnClima_OpenMeteo.pq) · [`fnFeriados_BrasilAPI.pq`](03-api-web/scripts_m/fnFeriados_BrasilAPI.pq)
- Documentação & Evidências: [`especificacao_apis.md`](03-api-web/docs/especificacao_apis.md) · [`resumo_coleta_apis.md`](03-api-web/evidencias/resumo_coleta_apis.md) · [`sample_ptax.json`](03-api-web/evidencias/sample_ptax.json) · [`sample_openmeteo_sp.json`](03-api-web/evidencias/sample_openmeteo_sp.json) · [`sample_brasilapi_2025.json`](03-api-web/evidencias/sample_brasilapi_2025.json)
- CSVs coletados (`3` arquivos em [`03-api-web/dados_api/`](03-api-web/dados_api/)): [`cambio_ptax_2025_2026.csv`](03-api-web/dados_api/cambio_ptax_2025_2026.csv) · [`clima_24_cidades_2025_2026.csv`](03-api-web/dados_api/clima_24_cidades_2025_2026.csv) · [`feriados_nacionais_2025_2026.csv`](03-api-web/dados_api/feriados_nacionais_2025_2026.csv)

### `04-powerbi-dashboard/` — Modelo Semântico, DAX, Tema & Guia `.pbix`
- [`04-powerbi-dashboard/README.md`](04-powerbi-dashboard/README.md) — Visão geral, roteiro resumido e referências numéricas.
- Documentação & Diagrama: [`modelo_powerbi.md`](04-powerbi-dashboard/docs/modelo_powerbi.md) · [`guia_montagem_pbix.md`](04-powerbi-dashboard/docs/guia_montagem_pbix.md) · [`diagrama_modelo.png`](04-powerbi-dashboard/docs/diagrama_modelo.png)
- Tema & DAX: [`tema_varejomax.json`](04-powerbi-dashboard/tema/tema_varejomax.json) · [`medidas_finais.dax`](04-powerbi-dashboard/dax/medidas_finais.dax) · [`validacao.dax`](04-powerbi-dashboard/dax/validacao.dax)
- Consultas Power Query M (`11` scripts em [`04-powerbi-dashboard/m/`](04-powerbi-dashboard/m/)): [`pConexaoODBC.pq`](04-powerbi-dashboard/m/pConexaoODBC.pq), [`tst_Conexao.pq`](04-powerbi-dashboard/m/tst_Conexao.pq), [`dim_produto.pq`](04-powerbi-dashboard/m/dim_produto.pq), [`dim_cliente.pq`](04-powerbi-dashboard/m/dim_cliente.pq), [`dim_cidade.pq`](04-powerbi-dashboard/m/dim_cidade.pq), [`dim_canal_loja.pq`](04-powerbi-dashboard/m/dim_canal_loja.pq), [`dim_transportadora.pq`](04-powerbi-dashboard/m/dim_transportadora.pq), [`dim_tempo.pq`](04-powerbi-dashboard/m/dim_tempo.pq), [`dim_cambio.pq`](04-powerbi-dashboard/m/dim_cambio.pq), [`dim_clima.pq`](04-powerbi-dashboard/m/dim_clima.pq), [`fato_vendas.pq`](04-powerbi-dashboard/m/fato_vendas.pq)
- Evidências: [`validacao_resultados.md`](04-powerbi-dashboard/evidencias/validacao_resultados.md)

---

## 4. Checklist de Finalização Manual no Windows (Pré-Defesa)

Execute esta sequência em ambiente **Windows 64-bit** com **Microsoft Excel Desktop** e **Power BI Desktop** instalados antes da apresentação ao vivo:

### Passo 1 — Compilar o Excel VBA (`VarejoMax_04_VBA.xlsm`)
- [ ] Abrir [`01-vba/VarejoMax_04_VBA_Template.xlsx`](01-vba/VarejoMax_04_VBA_Template.xlsx) no Microsoft Excel.
- [ ] Pressionar `Alt + F11` (VBE) → `Arquivo → Importar Arquivo...` (`Ctrl + M`) e importar os 3 arquivos `.bas` de `01-vba/macros/` (`modResumoVendas.bas`, `modAtrasosLogistica.bas`, `modUtilitariosExportacao.bas`).
- [ ] *(Opcional)* Vincular botões de controle de formulário na aba `Menu_Operacional` às macros principais.
- [ ] Salvar como **`01-vba/VarejoMax_04_VBA.xlsm`** (*Pasta de Trabalho Habilitada para Macro do Excel*).
- [ ] Executar `AtualizarBasesERelatorios` (`Alt + F8`) e confirmar `MsgBox`: **`RECONCILIAÇÃO OFICIAL: APROVADA (Diferença = R$ 0,00)`**.
- [ ] Executar `ExportarRelatoriosPDF` e verificar a geração dos 2 PDFs em `01-vba/relatorios_gerados/`.

### Passo 2 — Configurar Driver SQLite ODBC 64-bit (`dsn=varejomax`)
- [ ] Instalar o **SQLite3 ODBC Driver (64-bit)** (`sqliteodbc_w64.exe`).
- [ ] Abrir **Administrador de Fonte de Dados ODBC (64 bits)** (`odbcad32.exe`) → adicionar DSN `varejomax` (exatamente em minúsculas) apontando para o caminho absoluto de [`02-bancodedados-sql/varejomax.db`](02-bancodedados-sql/varejomax.db) com **FK Support** habilitado.

### Passo 3 — Montar e Validar o Power BI (`VarejoMax_Dashboard.pbix` e `.pdf`)
- [ ] Seguir o roteiro completo de [`04-powerbi-dashboard/docs/guia_montagem_pbix.md`](04-powerbi-dashboard/docs/guia_montagem_pbix.md):
  1. Importar os `11` scripts `.pq` de `04-powerbi-dashboard/m/` (desabilitar carga em `tst_Conexao`).
  2. Configurar os `8` relacionamentos (`R1`..`R8`, `*:1`, direção única) e marcar `dim_tempo` como **Tabela de Datas** (com *Sort by Column* e categorias geográficas).
  3. Criar `_Medidas = { BLANK() }` e aplicar [`medidas_finais.dax`](04-powerbi-dashboard/dax/medidas_finais.dax) via **DAX Query View**.
  4. Importar [`tema_varejomax.json`](04-powerbi-dashboard/tema/tema_varejomax.json) e construir as **5 páginas principais** + **3 páginas de apoio** (`Detalhe Transportadora / Loja`, `Dica Resumo` `320x240` e `Conferência`).
  5. Sincronizar os segmentadores nas 5 páginas principais e ajustar interações/bloqueios de filtro.
- [ ] Executar [`validacao.dax`](04-powerbi-dashboard/dax/validacao.dax) (`V-K4`, `V-K5`, `V-C1`..`V-C10`, `V-CONF`) e preencher [`validacao_resultados.md`](04-powerbi-dashboard/evidencias/validacao_resultados.md).
- [ ] Salvar **`04-powerbi-dashboard/VarejoMax_Dashboard.pbix`** e exportar **`04-powerbi-dashboard/VarejoMax_Dashboard.pdf`**.
