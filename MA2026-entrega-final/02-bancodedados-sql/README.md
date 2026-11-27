# Etapa IV — Banco de Dados Relacional 3FN, Views Star Schema & SQL Analítico (`02-bancodedados-sql/`)

Camada de persistência oficial do projeto **VarejoMax Brasil** em **SQLite 3** (`varejomax.db`), unificando os cadastros e vendas saneados na Etapa 2 com as três APIs públicas integradas (**PTAX BCB**, **Open-Meteo** e **BrasilAPI**), servindo de fonte única da verdade para as consultas SQL (`C1`–`C10`) e para o dashboard Power BI via ODBC 64-bit.

---

## Estrutura de Arquivos

| Caminho | Descrição |
|---|---|
| [`varejomax.db`](varejomax.db) | Banco SQLite 3 populado, validado e fechado em `journal_mode=DELETE` |
| [`sql/varejomax_ddl.sql`](sql/varejomax_ddl.sql) | DDL 3FN (`10` tabelas, PK, FK, CHECK, UNIQUE), `9` índices e `9` tabelas `stg_*` |
| [`sql/varejomax_carga.sql`](sql/varejomax_carga.sql) | Carga idempotente (2× idênticas), CTE recursiva do calendário (`608` dias + feriados BrasilAPI) |
| [`sql/varejomax_views.sql`](sql/varejomax_views.sql) | Camada Star Schema com `9` views (`fato_vendas` + `8` `dim_*`) aderentes aos `.pq` |
| [`sql/varejomax_conferencia.sql`](sql/varejomax_conferencia.sql) | Batimento `K1`–`K5` (`diferenca = 0` e reprodução exata dos KPIs da Etapa 3) |
| [`sql/varejomax_diagnostico.sql`](sql/varejomax_diagnostico.sql) | Diagnósticos `D1`–`D4` (358 pedidos com variação entre itens, anomalias e nulos) |
| [`sql/varejomax_consultas.sql`](sql/varejomax_consultas.sql) | `10` consultas analíticas (`C1`–`C10`) com CTEs, Window Functions e APIs externas |
| [`scripts_build/construir_banco.py`](scripts_build/construir_banco.py) | Builder automatizado (Python stdlib + `sqlite3`) que reconstrói e audita todo o banco |
| [`csv/`](csv) | `6` CSVs tratados da Etapa 2 + `3` CSVs das APIs externas |
| [`docs/modelo_relacional.md`](docs/modelo_relacional.md) | Documentação 3FN, dicionário de dados, índices e contrato das views |
| [`docs/diagrama_er.png`](docs/diagrama_er.png) | Diagrama ER completo (3FN + APIs + camada Star Schema) |
| [`evidencias/`](evidencias) | Exportações `k1..k5.csv`, `d1..d4.csv`, `c1..c10.csv`, `sqlite_master.csv`, `explain_query_plan.txt` e [`resumo_validacao.md`](evidencias/resumo_validacao.md) |

---

## Como Reconstruir e Validar do Zero

Execute o builder automatizado na raiz do repositório (requer apenas Python 3 com stdlib):

```powershell
python "MA2026-entrega-final/02-bancodedados-sql/scripts_build/construir_banco.py"
```

O script:
1. Recria `varejomax.db` aplicando `sql/varejomax_ddl.sql` (`PRAGMA foreign_keys = ON`).
2. Carrega os 9 CSVs de `csv/` nas tabelas `stg_*`.
3. Executa `sql/varejomax_carga.sql` **duas vezes consecutivas**, comprovando idempotência.
4. Cria as 9 views em `sql/varejomax_views.sql`.
5. Valida `PRAGMA integrity_check = ok` e `PRAGMA foreign_key_check` (`0` violações).
6. Executa `K1`–`K5`, `D1`–`D4` e `C1`–`C10`, grava os `.csv` e o `EXPLAIN QUERY PLAN` em `evidencias/`, e fecha o banco sem arquivos `.db-journal`, `.db-wal` ou `.db-shm`.

---

## Conexão ODBC 64-bit para o Power BI (`varejomax`)

1. Instale o **SQLite3 ODBC Driver (64-bit)** (`ch-werner.de/sqliteodbc/`).
2. Abra **Fontes de Dados ODBC (64 bits)** (`odbcad32.exe`) → aba **DSN de Usuário** → **Adicionar** → selecione `SQLite3 ODBC Driver`.
3. Configure:
   - **Data Source Name:** `varejomax`
   - **Database Name:** caminho absoluto para `MA2026-entrega-final/02-bancodedados-sql/varejomax.db`
   - **FK Support:** marcado (`PRAGMA foreign_keys = ON`)
4. No Power BI Desktop, o parâmetro `pConexaoODBC = "dsn=varejomax"` alimentará diretamente as 9 views (`dim_*` e `fato_vendas`) conforme os scripts `.pq` em [`../04-powerbi-dashboard/m/`](../04-powerbi-dashboard/m/).

---

## Resumo dos Achados Analíticos (`C1`–`C10`)

| Consulta | Tema / Hipótese | Principal Evidência Apurada no Banco |
|---|---|---|
| **C1** | Evolução Mensal & Black Friday (H4) | Pico em **Nov/2025** (`819` pedidos, **R$ 332.762,15**, `+99,50%` MoM); acumulado válido de **R$ 3.448.872,47** |
| **C2** | Categorias & Margens | **Eletrônicos** concentra **42,69%** do faturamento (R$ 1,47M), mas tem a menor margem (**16,78%** ref. / **15,70%** real) vs **Moda** (**47,64%** real) |
| **C3** | Canais & Lojas | Canal **Online** soma **89,91%** (`Site` 46,07%, `App` 24,87%, `Marketplace` 18,97%) contra **10,09%** das 7 lojas físicas |
| **C4** | Top 10 Cidades & Lojas Físicas | **São Paulo** lidera (**26,14%**, atraso `3,27%`); capitais NE sem loja no Top 10 (`Recife` 24,43%, `Salvador` 20,59%, `Fortaleza` 20,10%) lideram atrasos |
| **C5** | SLA Transportadoras & Nota (H1, H3) | Consolidação estrita de `8.180` pedidos entregues e `628` atrasos (`7,68%`): **EntregaJá** (`15,75%`), **Não informado** (`14,29%`), **TransNorte** (`10,21%`), **RotaBR** (`7,04%`), **Veloz Express** (`2,79%`), **Retirada** (`0,00%`); atrasos derrubam a nota (ex.: `4,27 → 3,41`) |
| **C6** | Pedidos Críticos `>= 5` dias (H2) | Todos os **15** pedidos críticos (máx. `11` dias em Manaus/AM) estão no **Norte, Nordeste e Centro-Oeste** |
| **C7** | Segmentação RFV (Pareto) | **17,70%** dos clientes (**VIP 8+ pedidos**) geram **57,06%** do faturamento (**R$ 1.968.032,01**) |
| **C8** | Cancelamentos, Devoluções & Descontos | Cancelamentos concentram-se no **Online** (`0%` em lojas físicas); `Eletrônicos Online` perde **R$ 158.918,71** em cancelados/devolvidos |
| **C9** | API 1: Dólar PTAX × Eletrônicos | Margem realizada de **Eletrônicos** recua de **17,03%** (`< R$ 5,10`) para **14,37%** (`>= R$ 6,00`) |
| **C10** | API 2: Clima Open-Meteo & Feriados × SLA | Atraso sobe de **4,67%** em dias úteis sem chuva para **14,04%** com chuva forte (`>30 mm`) e **16,67%** em feriados com chuva forte *(soma `8.296` por incluir 5 pedidos multi-dia em faixas distintas)* |
