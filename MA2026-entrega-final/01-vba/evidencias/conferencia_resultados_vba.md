# Conferência de Resultados e Evidências de Reconciliação — Etapa VBA (`01-vba/`)

> **Projeto:** VarejoMax Brasil (FATEC SP — Microinformática Aplicada, Prof. Veríssimo)  
> **Fonte oficial conferida:** `02-bancodedados-sql/csv/` (`fVendas.csv`, `dProduto.csv`, `dCidade.csv`)  
> **Pasta de trabalho modelo:** [VarejoMax_04_VBA_Template.xlsx](../VarejoMax_04_VBA_Template.xlsx)

---

## 1. Batimento com o Gabarito Oficial de Reconciliação

| Indicador Oficial | Gabarito Esperado | Apurado VBA / Template | Diferença | Status |
|---|---|---|---|---|
| Linhas em `Base_Vendas` (`fVendas.csv`) | 13.613 | 13.613 | 0 | APROVADO |
| Pedidos Totais (todos os status) | 8.996 | 8.996 | 0 | APROVADO |
| Pedidos Válidos (excl. `Cancelado` e `Devolvido`) | 8.291 | 8.291 | 0 | APROVADO |
| Faturamento Bruto (excl. `Cancelado`) | R$ 3.797.167,40 | R$ 3.797.167,40 | R$ 0,00 | APROVADO |
| Faturamento Líquido (excl. `Cancelado` e `Devolvido`) | R$ 3.448.872,47 | R$ 3.448.872,47 | R$ 0,00 | APROVADO |
| Ticket Médio (Fat. Líquido ÷ Pedidos Válidos) | R$ 415,98 | R$ 415,98 | R$ 0,00 | APROVADO |
| Qtd. Itens/Unidades Válidas (excl. Canc/Dev) | 16.581 | 16.581 | 0 | APROVADO |
| Qtd. Itens/Unidades Totais (todos status) | 17.946 | 17.946 | 0 | APROVADO |
| Custo Total Ref. Válido (`Cad_Produtos.Custo_Ref`) | R$ 2.372.700,48 | R$ 2.372.700,48 | R$ 0,00 | APROVADO |
| Margem Bruta % (Ref.) | 31,2% | 31,20% | 0,0 p.p. | APROVADO |
| Custo Realizado Válido (`fVendas.Custo_Unitario`) | R$ 2.402.823,74 | R$ 2.402.823,74 | R$ 0,00 | APROVADO |
| Margem Realizada % | 30,3% | 30,33% | 0,0 p.p. | APROVADO |
| Pedidos Atrasados Válidos (`Entregue` + `Atrasado`) | 628 | 628 | 0 | APROVADO |
| Pedidos Atrasados Totais (incl. 26 `Devolvido`) | 654 | 654 (628 válidos + 26 devolvidos) | 0 | APROVADO |

---

## 2. Quadro 1 — Desempenho por Categoria (`Resumo_Categorias`)

Referência exportada em [resumo_vendas_por_categoria_referencia.csv](../relatorios_gerados/resumo_vendas_por_categoria_referencia.csv).

| Categoria | Pedidos Válidos* | Unidades Vendidas | Faturamento Bruto (R$) | Faturamento Líquido (R$) | Ticket Médio (R$) | Custo Ref. Total (R$) | Margem Bruta (%) | Participação (%) |
|---|---:|---:|---:|---:|---:|---:|---:|---:|
| **Eletrônicos** | 2.345 | 3.187 | R$ 1.636.914,60 | R$ 1.472.456,60 | R$ 627,91 | R$ 1.225.387,66 | **16,78%** *(alerta < 20%)* | 42,69% |
| Casa e Decoração | 2.008 | 2.746 | R$ 696.733,70 | R$ 638.476,72 | R$ 317,97 | R$ 392.184,67 | 38,57% | 18,51% |
| Moda | 2.555 | 3.567 | R$ 547.264,00 | R$ 491.996,15 | R$ 192,56 | R$ 254.240,65 | 48,32% | 14,27% |
| Esporte e Lazer | 937 | 1.326 | R$ 413.069,30 | R$ 383.317,49 | R$ 409,09 | R$ 226.858,74 | 40,82% | 11,11% |
| Beleza e Saúde | 1.814 | 2.647 | R$ 288.922,00 | R$ 267.746,23 | R$ 147,60 | R$ 151.575,18 | 43,39% | 7,76% |
| Livros e Papelaria | 1.969 | 3.108 | R$ 214.263,80 | R$ 194.879,28 | R$ 98,97 | R$ 122.453,58 | 37,16% | 5,65% |
| **TOTAL GERAL** | **8.291** | **16.581** | **R$ 3.797.167,40** | **R$ 3.448.872,47** | **R$ 415,98** | **R$ 2.372.700,48** | **31,20%** | **100,00%** |

*\*Nota:* Pedidos Válidos por categoria contabiliza `ID_Pedido` distintos contendo ao menos um item da categoria (um pedido multicategoria é contado em cada categoria aplicável). O `TOTAL GERAL` consolida os **8.291 pedidos únicos válidos** da base.

---

## 3. Quadro 2 — Desempenho por Canal de Venda (`Resumo_Categorias`)

| Canal de Venda | Pedidos Válidos | Unidades Vendidas | Faturamento Bruto (R$) | Faturamento Líquido (R$) | Ticket Médio (R$) | Custo Ref. Total (R$) | Margem Bruta (%) | Participação (%) |
|---|---:|---:|---:|---:|---:|---:|---:|---:|
| Site | 3.763 | 7.429 | R$ 1.755.375,90 | R$ 1.588.900,22 | R$ 422,24 | R$ 1.059.258,85 | 33,33% | 46,07% |
| App | 2.161 | 4.394 | R$ 947.113,10 | R$ 857.691,62 | R$ 396,90 | R$ 600.773,86 | 29,95% | 24,87% |
| Marketplace | 1.572 | 3.201 | R$ 735.572,90 | R$ 654.238,87 | R$ 416,18 | R$ 471.649,71 | 27,91% | 18,97% |
| Loja Física | 795 | 1.557 | R$ 359.105,50 | R$ 348.041,76 | R$ 437,79 | R$ 241.018,06 | 30,75% | 10,09% |
| **TOTAL GERAL** | **8.291** | **16.581** | **R$ 3.797.167,40** | **R$ 3.448.872,47** | **R$ 415,98** | **R$ 2.372.700,48** | **31,20%** | **100,00%** |

---

## 4. Quadro 1 & 2 — Desempenho por Transportadora e Pedidos Críticos (`Relatorio_Atrasos`)

Referência exportada em [relatorio_atrasos_por_transportadora_referencia.csv](../relatorios_gerados/relatorio_atrasos_por_transportadora_referencia.csv).

### Regras de Saneamento Aplicadas por Pedido (`ID_Pedido`)

1. **Escopo de Entregas Concluídas:** dos **8.291** pedidos válidos, **8.180** possuem `Status_Pedido = "Entregue"` (os demais **111** estão em `Em trânsito` = 78 ou `Em separação` = 33, ainda sem data de entrega real). Os **26** atrasos ocorridos em pedidos `Devolvido` ficam isolados fora do quadro de pedidos válidos (`628 + 26 = 654`).
2. **Variação entre itens do mesmo pedido (`358` pedidos na base):**
   - **Transportadora (`85` pedidos com divergência):** ocorre exclusivamente quando um item traz `"Não informado"` e outro item do mesmo `ID_Pedido` traz a transportadora efetiva. A macro e o gabarito consolidam o pedido pela **transportadora informada** (evitando subestimar transportadoras reais e reduzindo `"Não informado"` de 80 para 49 pedidos).
   - **Entregas `Inconsistente` (`17` pedidos entregues, `18` linhas no total com `Devolvido`):** `Data_Entrega_Real < Data_Envio` gera `Prazo_Real_Dias` negativo; essas linhas não entram como `Atrasado` e são desconsideradas no cálculo de `Prazo Real Médio`.
   - **Avaliação fora de 1–5 (`8` linhas com notas `0` e `6..10`):** filtradas estritamente com `1 <= nota <= 5`, consolidando 1 avaliação válida por `ID_Pedido` (`3.517` pedidos avaliados válidos).

### Tabela Consolidada por Transportadora

| Transportadora | Pedidos Entregues | Pedidos Atrasados | % Atraso | Prazo Previsto Médio (dias) | Prazo Real Médio (dias) | Atraso Médio nos Atrasados (dias) | Nota Média (1 a 5) |
|---|---:|---:|---:|---:|---:|---:|---:|
| EntregaJá | 1.943 | 306 | 15,75% | 6,62 | 5,78 | 1,73 | 4,13 |
| RotaBR | 2.315 | 163 | 7,04% | 6,56 | 5,08 | 1,65 | 4,23 |
| TransNorte Logística | 891 | 91 | 10,21% | 9,49 | 8,02 | 1,59 | 4,19 |
| Veloz Express | 2.187 | 61 | 2,79% | 5,30 | 3,31 | 1,34 | 4,25 |
| Não informado | 49 | 7 | 14,29% | 7,00 | 5,53 | 1,57 | 4,35 |
| Retirada em loja | 795 | 0 | 0,00% | — | — | — | 4,29 |
| **TOTAL GERAL** | **8.180** | **628** | **7,68%** | **6,56** | **5,10** | **1,65** | **4,21** |

### Resumo de Pedidos Críticos Atrasados (`Quadro 2`)

- Total de linhas no Quadro 2: **628** pedidos atrasados entregues, ordenados por `Atraso_Dias DESC, ID_Pedido ASC`.
- Pedidos críticos destacados em vermelho claro (`Atraso_Dias >= 5` dias): **15 pedidos** (máximo de **11 dias** de atraso no pedido `PED-005672` via `EntregaJá`).

---

## 5. Checklist de Validação das 3 Macros VBA

| # | Módulo / Macro | Critério de Validação | Resultado Esperado |
|---|---|---|---|
| 1 | `modResumoVendas.GerarResumoVendasCategoria` | Execução sem filtros (`Menu_Operacional!C6:C8` vazios) em `< 1s` | Quadro 1 (6 categorias) + Quadro 2 (4 canais) somando `8.291` pedidos e `R$ 3.448.872,47` líquido; destaque em `Eletrônicos` (`16,78% < 20%`). |
| 2 | `modResumoVendas.GerarResumoVendasCategoria` | Execução com filtro de período (`C6`/`C7`) e/ou canal (`C8`) | Atualiza cabeçalho `A3`, filtra linhas e recalcula `% Participação` sobre o total filtrado. |
| 3 | `modAtrasosLogistica.GerarRelatorioAtrasosTransportadora` | Execução sem filtros (`Menu_Operacional!C9:C10` vazios) | Quadro 1 soma `8.180` entregues e `628` atrasados (`7,68%`); Quadro 2 lista os `628` atrasados ordenados por `Atraso_Dias DESC`, destacando os `15` pedidos com `>= 5` dias. |
| 4 | `modAtrasosLogistica.GerarRelatorioAtrasosTransportadora` | Execução com filtro `C9 = "EntregaJá"` ou `C10 = "Nordeste"` / `"SP"` | Filtra por transportadora e/ou UF/Região via `Cad_Cidades` sem duplicar `ID_Pedido`. |
| 5 | `modUtilitariosExportacao.AtualizarBasesERelatorios` | Execução completa | Chama `RefreshAll`, executa as duas macros principais, valida batimento com `3448872.47` (`Diferença = R$ 0,00`) e exibe `MsgBox`. |
| 6 | `modUtilitariosExportacao.ExportarRelatoriosPDF` | Planilha salva como `.xlsm` | Cria pasta `relatorios_gerados/` se ausente, configura `xlLandscape` + `FitToPagesWide = 1` e exporta 2 PDFs com timestamp `YYYYMMDD_HHMM`. |
| 7 | `modUtilitariosExportacao.LimparFiltrosERelatorios` | Limpeza operacional | Limpa `Menu_Operacional!C6:C10` e os quadros gerados preservando banners executivos (`A1:I3` e `A1:J4`). |
| 8 | Todas as macros (`Option Explicit` + `CleanExit`) | Tratamento de erro e performance | Desativa `ScreenUpdating`, `EnableEvents` e usa `xlCalculationManual`, restaurando obrigatoriamente em `CleanExit:`. |
