# Resumo de Validação Técnica — Banco SQLite `varejomax.db`

> **Executor:** Script automatizado [`construir_banco.py`](../scripts_build/construir_banco.py) (SQLite `3.50.4`)  
> **Integridade Estrutural:** `PRAGMA integrity_check = ok` | `PRAGMA foreign_key_check = 0 violações` | `journal_mode = DELETE` (zero arquivos `.db-journal`, `.db-wal` ou `.db-shm`).

---

## 1. Prova de Idempotência (`varejomax_carga.sql` executado 2× seguidas)

| Tabela | 1ª Execução | 2ª Execução | Status |
|---|---:|---:|---|
| `cidade` | 24 | 24 | Idêntico |
| `canal_loja` | 10 | 10 | Idêntico |
| `produto` | 49 | 49 | Idêntico |
| `cliente` | 1.607 | 1.607 | Idêntico |
| `transportadora` | 6 | 6 | Idêntico |
| `calendario` | 608 | 608 | Idêntico |
| `pedido` | 8.996 | 8.996 | Idêntico |
| `item_pedido` | 13.613 | 13.613 | Idêntico |
| `cambio_ptax` | 608 | 608 | Idêntico |
| `clima_cidade` | 14.592 | 14.592 | Idêntico |

---

## 2. Batimento `K1`–`K5` ([`varejomax_conferencia.sql`](../sql/varejomax_conferencia.sql))

### `K1` — Volumetria e Totais Financeiros (`diferenca = 0` em todas as 15 linhas)

| Indicador | Staging (`stg_*`) | Banco (`varejomax.db`) | Diferença |
|---|---:|---:|---:|
| Itens de pedido (linhas) | 13.613 | 13.613 | **0** |
| Pedidos totais distintos | 8.996 | 8.996 | **0** |
| Pedidos válidos (sem Cancelado/Devolvido) | 8.291 | 8.291 | **0** |
| Unidades vendidas (válidos) | 16.581 | 16.581 | **0** |
| Faturamento Líquido (válidos) | 3.448.872,47 | 3.448.872,47 | **0,00** |
| Faturamento Bruto (não Cancelado) | 3.797.167,40 | 3.797.167,40 | **0,00** |
| Frete Total (válidos) | 93.707,67 | 93.707,67 | **0,00** |
| Produtos (`dim_produto`) | 49 | 49 | **0** |
| Clientes (`dim_cliente`) | 1.607 | 1.607 | **0** |
| Cidades (`dim_cidade`) | 24 | 24 | **0** |
| Canais/Lojas (`dim_canal_loja`) | 10 | 10 | **0** |
| Transportadoras (`dim_transportadora`) | 6 | 6 | **0** |
| Calendário (`dim_tempo`) | 608 | 608 | **0** |
| Câmbio PTAX (`dim_cambio`) | 608 | 608 | **0** |
| Clima 24 Cidades (`dim_clima`) | 14.592 | 14.592 | **0** |

### `K2`, `K3` e `K4` — Integridade, Regras de Negócio e Cobertura de APIs
- **`K2` (12 verificações de órfãos FK):** `0` órfãos em todas as 12 relações.
- **`K3` (7 verificações de regras de negócio):** `0` violações (valores positivos, descontos `[0,1]`, frete exclusivo no `item = 1`, domínio `'Sim'`/`'Não'` no calendário, paridade centavo-a-centavo de `valor_bruto` e `valor_liquido`, zero duplicatas de PK).
- **`K4` (5 verificações de cobertura temporal/APIs):** `0` linhas sem match entre `fato_vendas` e `dim_tempo`, `dim_cambio` e `dim_clima` (`100%` de cobertura nas 13.613 vendas e nas 14.592 combinações `cidade × data`).

### `K5` — KPIs Executivos Oficiais

| KPI | Valor Apurado (`k5.csv`) |
|---|---:|
| Pedidos Totais / Pedidos Válidos | 8.996 / 8.291 |
| Unidades Vendidas (Válidos) | 16.581 |
| Faturamento Bruto (não Cancelado) | R$ 3.797.167,40 |
| Faturamento Líquido (Válidos) | R$ 3.448.872,47 |
| Ticket Médio | R$ 415,98 |
| Itens por Pedido | 2,00 |
| Custo Referência Total / Margem Bruta Ref. % | R$ 2.372.700,48 / 31,20% |
| Custo Realizado Total / Margem Realizada % | R$ 2.402.823,74 / 30,33% |
| Frete Total (Válidos) | R$ 93.707,67 |
| % Vendas Online | 89,91% |
| Pedidos Entregues / Atrasados / % Atraso | 8.180 / 628 / 7,68% |
| % Cancelados / % Devolvidos | 3,95% / 3,89% |
| Nota Média Cliente — Item (todas as linhas 1–5, padrão Etapa 3 / DAX) | 4,11 |
| Nota Média Cliente — Item válido (excl. Cancelado/Devolvido) | 4,18 |
| Nota Média Cliente — Pedido Entregue (padrão VBA / `C5`) | 4,21 |

---

## 3. Diagnósticos `D1`–`D4` ([`varejomax_diagnostico.sql`](../sql/varejomax_diagnostico.sql))

- **`D1` (Variação entre itens no mesmo pedido):** `358` pedidos distintos (`CEP`: 231, `Transportadora`: 85, `Forma_Pagamento`: 40, `Data_Entrega_Real`: 15, `Avaliacao_Cliente`: 4).
- **`D2` (Anomalias preservadas da Etapa 2):** `8` itens com avaliação fora de `1..5` (`0` e `6..10`); `18` itens com `status_atraso = 'Inconsistente'` (`data_entrega_real < data_envio`).
- **`D3` (Nulos e regra de frete):** `1` cliente sem cidade (`NULL`), `11` clientes sem e-mail (`NULL`), `408` itens sem CEP, `1.192` pedidos sem `data_envio`, `689` itens sem `data_entrega_real`, `7.949` itens sem avaliação; `4.948` itens com `Frete_Valor > 0` em `Item = 1` e **`0`** em `Item <> 1`.
- **`D4` (Cobertura de datas):** `data_pedido` entre `2025-01-01` e `2026-08-31`; `151` registros com `data_entrega_prevista` após `2026-08-31` (máximo `2026-09-11`), justificando a não imposição de FK em `data_entrega_prevista`.

---

## 4. Plano de Execução (`EXPLAIN QUERY PLAN`) e Principais Achados `C1`–`C10`

- **Otimização (`explain_query_plan.txt`):** Todas as consultas `C1`–`C10` utilizam `SEARCH ... USING INDEX` nas dimensões (`pedido`, `produto`, `canal_loja`, `cidade`, `transportadora`, `calendario`, `cambio_ptax`, `clima_cidade`), sem `AUTOMATIC COVERING INDEX`. Em `C5` e `C6`, o filtro `status_pedido = 'Entregue'` utiliza `ix_pedido_status`; em `C10`, o join por `(cidade, COALESCE(data_entrega_real, data_pedido))` utiliza a PK composta `sqlite_autoindex_clima_cidade_1`.
- **Principais achados:**
  - `C1`: Pico de Black Friday em `2025-11` (`819` pedidos, `R$ 332.762,15` líquidos, `+99,50%` MoM).
  - `C2`: `Eletrônicos` responde por `42,69%` da receita (`R$ 1.472.456,60`), porém com a menor margem (`16,78%` ref. / `15,70%` real), enquanto `Moda` alcança `47,64%` real.
  - `C3`: `Site` (`46,07%`), `App` (`24,87%`) e `Marketplace` (`18,97%`) somam `89,91%` vs `10,09%` das 7 lojas físicas.
  - `C4`: `São Paulo` lidera com `R$ 901.513,58` (`26,14%`, `3,27%` atraso), enquanto capitais do Nordeste no Top 10 sem loja física apresentam atrasos acima de `20%` (`Recife` `24,43%`, `Salvador` `20,59%`, `Fortaleza` `20,10%`).
  - `C5`: Consolidação estrita de `8.180` pedidos `Entregue` e `628` atrasos (`7,68%`), batendo 100% com o VBA: `EntregaJá` (`1.943` / `15,75%`), `Não informado` (`49` / `14,29%`), `TransNorte Logística` (`891` / `10,21%`), `RotaBR` (`2.315` / `7,04%`), `Veloz Express` (`2.187` / `2,79%`) e `Retirada em loja` (`795` / `0,00%`); entregas atrasadas derrubam a nota média em todas as transportadoras (ex.: `EntregaJá` cai de `4,27` para `3,41`).
  - `C6`: Os `15` pedidos com maior atraso crítico (`>= 5` dias, máx. `11` dias em `PED-005672` para Manaus/AM) concentram-se exclusivamente no Norte, Nordeste e Centro-Oeste.
  - `C7`: Clientes `VIP (8+ pedidos)` representam `17,70%` da base (`278` clientes) e geram `57,06%` do faturamento (`R$ 1.968.032,01`), contra `5,85%` dos `Novos (1 pedido)`.
  - `C8`: Cancelamentos (`0%` em lojas físicas) e devoluções concentram-se no canal Online (`Eletrônicos Online` soma `R$ 158.918,71` em receita perdida).
  - `C9`: Com a alta do dólar PTAX de `< R$ 5,10` para `>= R$ 6,00`, a margem realizada de `Eletrônicos` cai de `17,03%` para `14,37%`.
  - `C10`: O percentual de atraso sobe de `4,67%` em dias úteis sem chuva (`0 mm`) para `14,04%` em dias úteis com chuva forte (`>30 mm`) e `16,67%` em feriados com chuva forte. *(Nota de granularidade: a soma de `pedidos_validos` por `(faixa_precipitacao, tipo_dia_pedido)` resulta em `8.296` (+5 frente aos `8.291` pedidos válidos únicos) porque `5` pedidos possuem itens entregues em dias distintos com faixas de chuva diferentes na `chave_clima`.)*
