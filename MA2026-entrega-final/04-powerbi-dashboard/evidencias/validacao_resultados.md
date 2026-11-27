# Evidências de Validação — Power BI (`VarejoMax_Dashboard.pbix`)

> **Banco de Referência Verificado:** `../02-bancodedados-sql/varejomax.db` (SQLite `3.50.4`, `PRAGMA integrity_check = ok`, `0` violações de FK).  
> **Status do `.pbix`:** `_aguardando_execucao_pbix_` — os números do SQLite abaixo já estão verificados via `k1..k5` e `c1..c10`; a coluna do Power BI Desktop será preenchida após a montagem do `.pbix` no Windows e execução de [dax/validacao.dax](../dax/validacao.dax).  
> **Script de diagnóstico ODBC:** [m/tst_Conexao.pq](../m/tst_Conexao.pq)

---

## 1. Pré-validação da Conexão ODBC (`tst_Conexao.pq`)

| Objeto / View | SQLite (`varejomax.db` Verificado) | Power BI (`tst_Conexao.pq`) | Status |
|---|---|---|---|
| `sqlite_version` | `3.50.4` | `_aguardando_execucao_pbix_` | [ ] |
| `dim_produto` | `49` | `_aguardando_execucao_pbix_` | [ ] |
| `dim_cliente` | `1.607` | `_aguardando_execucao_pbix_` | [ ] |
| `dim_cidade` | `24` | `_aguardando_execucao_pbix_` | [ ] |
| `dim_canal_loja` | `10` | `_aguardando_execucao_pbix_` | [ ] |
| `dim_transportadora` | `6` | `_aguardando_execucao_pbix_` | [ ] |
| `dim_tempo` | `608` (`2025-01-01` a `2026-08-31`) | `_aguardando_execucao_pbix_` | [ ] |
| `dim_cambio` | `608` | `_aguardando_execucao_pbix_` | [ ] |
| `dim_clima` | `14.592` (`24 × 608`) | `_aguardando_execucao_pbix_` | [ ] |
| `fato_vendas` | `13.613` | `_aguardando_execucao_pbix_` | [ ] |

---

## 2. Validação `V-K4` — KPIs Oficiais vs. SQLite (`k1.csv` / `k5.csv`)

| Indicador (Medida DAX) | SQLite (`varejomax.db` Verificado) | Power BI (`V-K4`) | Diferença | Status |
|---|---|---|---|---|
| `[Faturamento Bruto]` | `R$ 3.797.167,40` | `_aguardando_execucao_pbix_` | `_aguardando_execucao_pbix_` | [ ] |
| `[Faturamento]` | `R$ 3.448.872,47` | `_aguardando_execucao_pbix_` | `_aguardando_execucao_pbix_` | [ ] |
| `[Qtd Pedidos]` (válidos) | `8.291` | `_aguardando_execucao_pbix_` | `_aguardando_execucao_pbix_` | [ ] |
| `[Qtd Pedidos Total]` | `8.996` | `_aguardando_execucao_pbix_` | `_aguardando_execucao_pbix_` | [ ] |
| `[Ticket Médio]` | `R$ 415,98` | `_aguardando_execucao_pbix_` | `_aguardando_execucao_pbix_` | [ ] |
| `[Qtd Itens]` (válidos) | `16.581` | `_aguardando_execucao_pbix_` | `_aguardando_execucao_pbix_` | [ ] |
| `[Itens por Pedido]` | `2,00` | `_aguardando_execucao_pbix_` | `_aguardando_execucao_pbix_` | [ ] |
| `[Frete Total]` *(todos status — nota 1)* | `R$ 97.922,36` | `_aguardando_execucao_pbix_` | `_aguardando_execucao_pbix_` | [ ] |
| `[Frete Pedidos Válidos]` *(nota 1)* | `R$ 93.707,67` | `_aguardando_execucao_pbix_` | `_aguardando_execucao_pbix_` | [ ] |
| `[Custo Total]` (ref.) | `R$ 2.372.700,48` | `_aguardando_execucao_pbix_` | `_aguardando_execucao_pbix_` | [ ] |
| `[Lucro Bruto]` (ref.) | `R$ 1.076.171,99` | `_aguardando_execucao_pbix_` | `_aguardando_execucao_pbix_` | [ ] |
| `[Margem Bruta %]` (ref.) | `31,20%` (`31,2%`) | `_aguardando_execucao_pbix_` | `_aguardando_execucao_pbix_` | [ ] |
| `[Custo Realizado]` | `R$ 2.402.823,74` | `_aguardando_execucao_pbix_` | `_aguardando_execucao_pbix_` | [ ] |
| `[Margem Realizada %]` | `30,33%` (`30,3%`) | `_aguardando_execucao_pbix_` | `_aguardando_execucao_pbix_` | [ ] |
| `[% Share Online]` | `89,91%` | `_aguardando_execucao_pbix_` | `_aguardando_execucao_pbix_` | [ ] |
| `[Qtd Pedidos Atrasados]` | `628` | `_aguardando_execucao_pbix_` | `_aguardando_execucao_pbix_` | [ ] |
| `[% Pedidos Atrasados]` *(s/ 8.291 válidos — nota 2)* | `7,57%` (`7,6%`) | `_aguardando_execucao_pbix_` | `_aguardando_execucao_pbix_` | [ ] |
| `[% Entregas no Prazo (OTD)]` *(nota 2)* | `92,32%` (`100% - 7,68%` entregues) | `_aguardando_execucao_pbix_` | `_aguardando_execucao_pbix_` | [ ] |
| `[Nota Média Cliente]` *(itens 1..5 — nota 3)* | `4,11` (`nota_media_item_todas`) | `_aguardando_execucao_pbix_` | `_aguardando_execucao_pbix_` | [ ] |
| `[Clientes Ativos]` | `1.571` | `_aguardando_execucao_pbix_` | `_aguardando_execucao_pbix_` | [ ] |
| `[Clientes Novos]` (`1` pedido) | `449` (`28,58%`) | `_aguardando_execucao_pbix_` | `_aguardando_execucao_pbix_` | [ ] |
| `[Clientes Recorrentes]` (`2+` pedidos) | `1.122` (`71,42%`) | `_aguardando_execucao_pbix_` | `_aguardando_execucao_pbix_` | [ ] |
| `[% Taxa Recompra]` | `71,42%` | `_aguardando_execucao_pbix_` | `_aguardando_execucao_pbix_` | [ ] |

---

## 3. Reconciliação Explícita das 3 Granularidades entre Camadas

1. **Frete Total (`R$ 97.922,36` vs. `R$ 93.707,67`):**
   - `[Frete Total]` (`SUM(fato_vendas[frete_valor])` = `R$ 97.922,36`) soma o frete de **todos os status** (já lançado exclusivamente no `item = 1` em `4.948` linhas, `0` em `item <> 1`).
   - `[Frete Pedidos Válidos]` (`R$ 93.707,67`, idêntico a `frete_total_valido` em `K1`/`K5`) exclui pedidos `Cancelado` e `Devolvido`.
2. **Taxa de Atraso (`7,57%` vs. `7,68%`):**
   - `[% Pedidos Atrasados]` no DAX (`628 / 8.291 = 7,57%` → `7,6%`) divide os `628` pedidos atrasados pelo universo de pedidos válidos (`8.291`).
   - Na consulta `K5` / `C5` e no VBA, `pct_atraso_entregues` (`628 / 8.180 = 7,68%`) usa como denominador apenas os `8.180` pedidos já com `status_pedido = 'Entregue'` (excluindo os `111` em trânsito/separação). Em `[% Entregas no Prazo (OTD)]`, o denominador soma apenas pedidos concluídos no prazo/adiantado + atrasados.
3. **Nota Média do Cliente (`4,11` vs. `4,18` vs. `4,21`):**
   - `[Nota Média Cliente]` (`4,11`) calcula `AVERAGE(avaliacao_cliente)` sobre os **`5.656` itens** com nota válida `1..5` (descartando os `8` registros inválidos `0` e `6..10`), seguindo o padrão oficial da Etapa 3 / DAX.
   - `K5` (`nota_media_item_valida = 4,18`) considera os **`5.415` itens válidos** (excluindo `Cancelado` e `Devolvido`).
   - VBA / `C5` (`nota_media_pedido_entregue = 4,21`) consolida **1 nota por pedido entregue** (`3.517` pedidos).

---

## 4. Validação `V-K5` e `V-CONF` — Integridade e Auditoria (`k1..k4.csv`)

| Verificação | SQLite (`varejomax.db` Verificado) | Power BI (`V-K5` / `V-CONF`) | Status |
|---|---|---|---|
| `V-K5.01 fato_vendas` linhas | `13.613` (`0` órfãos FK, `0` duplicatas PK) | `_aguardando_execucao_pbix_` | [ ] |
| `V-K5.10` Pedidos Cancelados | `355` (`3,95%`) | `_aguardando_execucao_pbix_` | [ ] |
| `V-K5.11` Pedidos Devolvidos | `350` (`3,89%`) | `_aguardando_execucao_pbix_` | [ ] |
| `V-K5.12` Avaliações Inválidas descartadas (`0`, `6..10`) | `8` itens (`D2`) | `_aguardando_execucao_pbix_` | [ ] |
| `V-K5.13` Pedidos sem correspondência em `dim_clima` | `0` (`K4` = 100% cobertura) | `_aguardando_execucao_pbix_` | [ ] |
| `V-K5.14` Pedidos sem correspondência em `dim_cambio` | `0` (`K4` = 100% cobertura) | `_aguardando_execucao_pbix_` | [ ] |
| `V-K5.15` Delta `[Faturamento]` vs `SUM(valor_liquido)` | `0,00` (`K3`: paridade centavo a centavo) | `_aguardando_execucao_pbix_` | [ ] |
| `V-K5.16` Delta `[Custo Total]` vs `SUM(custo_ref_total)` | `0,00` | `_aguardando_execucao_pbix_` | [ ] |
| `V-CONF` Delta `[Faturamento Online] + [Faturamento Loja Física]` vs `[Faturamento]` | `0,00` | `_aguardando_execucao_pbix_` | [ ] |
| `V-CONF` `[Clientes Novos] + [Clientes Recorrentes]` vs `[Clientes Ativos]` | `449 + 1.122 = 1.571` | `_aguardando_execucao_pbix_` | [ ] |

---

## 5. Conferência Cruzada `V-C1` a `V-C10` (Espelho 1-para-1 de `c1.csv`..`c10.csv`)

| Consulta | Escopo | SQLite (`varejomax.db` Verificado — Principais Achados) | Power BI (`V-C1..V-C10`) | Status |
|---|---|---|---|---|
| `V-C1` (`c1.csv`) | Evolução mensal (`2025-01`..`2026-08`) + MoM | Pico Black Friday `2025-11`: `819` pedidos, `R$ 332.762,15` líq., `+99,50%` MoM; `2025-12`: `583` ped., `R$ 249.516,94` | `_aguardando_execucao_pbix_` | [ ] |
| `V-C2` (`c2.csv`) | Categorias, Margem Ref. vs Real e `% Participação` | `Eletrônicos`: `R$ 1.472.456,60` (`42,69%`), Margem Ref `16,78%` / Real `15,70%`; `Moda`: `R$ 491.996,15` (`14,27%`), Real `47,64%` | `_aguardando_execucao_pbix_` | [ ] |
| `V-C3` (`c3.csv`) | Canais de Venda e Lojas | `Site`: `R$ 1.588.900,22` (`46,07%`); `App`: `R$ 857.691,62` (`24,87%`); `Marketplace`: `R$ 654.238,87` (`18,97%`); Lojas Físicas (`7` lojas): `10,09%` (líder `LJ01` `R$ 90.110,89`) | `_aguardando_execucao_pbix_` | [ ] |
| `V-C4` (`c4.csv`) | Regiões & Top 10 Cidades | `São Paulo`: `2.036` ped., `R$ 901.513,58` (`26,14%`, `3,27%` atraso); Nordeste s/ loja (`Recife` `24,43%`, `Salvador` `20,59%`, `Fortaleza` `20,10%` atraso) | `_aguardando_execucao_pbix_` | [ ] |
| `V-C5` (`c5.csv`) | Desempenho Logístico por Transportadora + Nota | `8.180` entregues / `628` atrasos (`7,68%`): `EntregaJá` `15,75%` (nota cai `4,27→3,41`), `TransNorte` `10,21%`, `RotaBR` `7,04%`, `Veloz Express` `2,79%`, `Retirada` `0,00%` *(nota: no Star Schema por item, 81 pedidos mistos com `Não informado` seguem grão de item)* | `_aguardando_execucao_pbix_` | [ ] |
| `V-C6` (`c6.csv`) | Top 15 Pedidos Críticos (`atraso_dias >= 5`) | Máx. `11` dias (`PED-005672`, Manaus/AM, `EntregaJá`); `9` dias (`PED-004552`, Belém/PA); `8` dias (`PED-003928`, Manaus/AM); todos os 15 no Norte, Nordeste ou Centro-Oeste | `_aguardando_execucao_pbix_` | [ ] |
| `V-C7` (`c7.csv`) | Clientes Ativos, RFV/Recorrência e Região | `1-Novo (1 ped)`: `449` cli (`28,58%`), `R$ 201.801,91` (`5,85%`); `2-Recorrente (2-4)`: `612` (`R$ 701.622,50`); `3-Fiel (5-7)`: `232` (`R$ 577.416,05`); `4-VIP (8+)`: `278` (`17,70%`), `R$ 1.968.032,01` (`57,06%`) | `_aguardando_execucao_pbix_` | [ ] |
| `V-C8` (`c8.csv`) | Descontos, Cancelamentos e Devoluções | Lojas físicas: `0,0%` cancelamento em todas as categorias; `Eletrônicos Online`: `93` cancel. (`4,02%`), `110` devol. (`4,75%`), `R$ 158.918,71` receita perdida | `_aguardando_execucao_pbix_` | [ ] |
| `V-C9` (`c9.csv`) | API 1 — Câmbio PTAX × `Eletrônicos` vs `Demais` | Em `Eletrônicos`, margem real cai de `17,03%` (`< R$ 5,10`) para `14,37%` (`>= R$ 6,00`); faixa `R$ 5,10–5,39` concentra `R$ 641.502,55` em Eletrônicos | `_aguardando_execucao_pbix_` | [ ] |
| `V-C10` (`c10.csv`) | API 2 — Clima & Calendário × Atrasos | Atraso sobe de `4,67%` (dia útil sem chuva `0 mm`) para `14,04%` (dia útil chuva forte `>30 mm`) e `16,67%` (feriado chuva forte) *(nota: 5 pedidos com itens entregues em faixas de chuva distintas)* | `_aguardando_execucao_pbix_` | [ ] |

---

## 6. Checklist Visual e Funcional do `.pbix` e `.pdf`

- [ ] DSN de Sistema ODBC 64-bit `varejomax` apontando para `../02-bancodedados-sql/varejomax.db` testado com sucesso.
- [ ] Tema `tema_varejomax.json` aplicado (fundo `#FAF7F2` em todas as páginas).
- [ ] `dim_tempo` marcada como Tabela de Datas e ordenação cronológica (`mes_abrev`, `nome_mes`, `ano_mes`, `nome_dia`) verificada.
- [ ] 5 páginas principais visíveis (`Visão Executiva`, `Produtos e Categorias`, `Clientes`, `Logística e Clima`, `Canais e Lojas`).
- [ ] Segmentações sincronizadas entre as 5 páginas principais.
- [ ] Página `Detalhe Transportadora / Loja` funcionando via Drill-through com botão Voltar.
- [ ] Página `Dica Resumo` (`320x240`) oculta e exibindo tooltip corretamente.
- [ ] Página `Conferência` oculta e com deltas zerados.
- [ ] Arquivo `VarejoMax_Dashboard.pbix` salvo e `VarejoMax_Dashboard.pdf` exportado sem visuais cortados.
