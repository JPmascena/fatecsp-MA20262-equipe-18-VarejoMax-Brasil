# Etapa 1 / Fase VBA — Automação Operacional em Excel VBA (`01-vba/`)

> **Projeto:** VarejoMax Brasil · FATEC SP — Microinformática Aplicada (Prof. Veríssimo)  
> **Escopo:** Automação de relatórios gerenciais e auditoria logística sobre os dados tratados da Etapa 2 (`13.613` itens / `8.996` pedidos totais / `8.291` pedidos válidos).

---

## 1. Objetivo

Substituir a consolidação manual de planilhas por um fluxo automatizado em **VBA (Visual Basic for Applications)** capaz de:

1. Processar os **13.613 registros** de `Base_Vendas` em menos de **1 segundo** utilizando arrays em memória e `Scripting.Dictionary`.
2. Aplicar integralmente as **regras oficiais de negócio e reconciliação** (tratamento de `Cancelado`, `Devolvido`, `Inconsistente`, notas fora de 1–5 e variação entre itens do mesmo `ID_Pedido`).
3. Gerar relatórios executivos formatados com alertas condicionais e exportação automatizada em **PDF paisagem**.

---

## 2. Estrutura de Arquivos Entregues

```text
01-vba/
├── README.md
├── VarejoMax_04_VBA_Template.xlsx              # Modelo completo com 6 abas e resultados pré-calculados
├── macros/
│   ├── modResumoVendas.bas                     # Macro de resumo por Categoria e Canal
│   ├── modAtrasosLogistica.bas                 # Macro de desempenho por Transportadora e Pedidos Críticos
│   └── modUtilitariosExportacao.bas            # Orquestração, conferência de batimento, PDF e limpeza
├── relatorios_gerados/
│   ├── resumo_vendas_por_categoria_referencia.csv
│   └── relatorio_atrasos_por_transportadora_referencia.csv
└── evidencias/
    └── conferencia_resultados_vba.md           # Gabarito completo apurado e checklist de testes
```

### Abas de [VarejoMax_04_VBA_Template.xlsx](VarejoMax_04_VBA_Template.xlsx)

| Aba | Finalidade | Volume |
|---|---|---|
| `Menu_Operacional` | Painel de filtros opcionais (`C6:C10`), validação por lista, KPIs oficiais e guia de macros | Painel executivo |
| `Resumo_Categorias` | Quadro 1 (6 Categorias) e Quadro 2 (4 Canais de Venda), com destaque em margem `< 20%` | 2 quadros + totais |
| `Relatorio_Atrasos` | Quadro 1 (6 Transportadoras) e Quadro 2 (`628` pedidos atrasados ordenados por `Atraso_Dias DESC`) | 6 + 628 linhas |
| `Cad_Produtos` | Tabela de referência de produtos (`ID_Produto`, `Categoria`, `Custo_Ref`, etc.) | 49 linhas |
| `Cad_Cidades` | Tabela de referência geográfica (`Cidade`, `UF`, `Regiao`, coordenadas) | 24 linhas |
| `Base_Vendas` | Fato transacional completa com datas e valores tipados | 13.613 linhas |

---

## 3. Descrição dos Módulos VBA (`macros/`)

### 3.1 [modResumoVendas.bas](macros/modResumoVendas.bas)
- **Procedimento:** `Public Sub GerarResumoVendasCategoria()`
- **Filtros lidos de `Menu_Operacional`:** `C6` (Data Inicial), `C7` (Data Final) e `C8` (Canal).
- **Regras implementadas:**
  - Carrega `Cad_Produtos` e `Base_Vendas` em arrays Variant (`Scripting.Dictionary` com contagem distinta de `ID_Pedido`).
  - **Faturamento Bruto:** exclui apenas `Status_Pedido = "Cancelado"`.
  - **Faturamento Líquido, Pedidos Válidos, Unidades e Custo Ref. Total:** excluem `Cancelado` e `Devolvido`.
  - Gera **Quadro 1** por Categoria (ordenado por Faturamento Líquido DESC, destacando `Eletrônicos` com margem `16,78% < 20%`) e **Quadro 2** por Canal (`Site`, `App`, `Marketplace`, `Loja Física`).

### 3.2 [modAtrasosLogistica.bas](macros/modAtrasosLogistica.bas)
- **Procedimento:** `Public Sub GerarRelatorioAtrasosTransportadora()`
- **Filtros lidos de `Menu_Operacional`:** `C9` (Transportadora), `C10` (UF ou Região, resolvido via `Cad_Cidades`) e `C6:C7` (Data, se preenchidos).
- **Regras implementadas:**
  - Agrupa por `ID_Pedido` considerando pedidos com `Status_Pedido = "Entregue"` (`8.180` pedidos), evitando dupla contagem nos **358 pedidos** que variam entre itens.
  - Consolida a `Transportadora` do pedido priorizando o valor informado quando algum item traz `"Não informado"`.
  - Exclui entregas `Inconsistente` (`Data_Entrega_Real < Data_Envio`) do cálculo de `Prazo Real Médio`.
  - Filtra avaliações estritamente no intervalo `1 <= nota <= 5` (descartando notas `0` e `6..10`).
  - Gera **Quadro 1** (Desempenho por Transportadora) e **Quadro 2** (628 pedidos críticos atrasados ordenados via `QuickSort` por `Atraso_Dias DESC`, destacando em vermelho claro os **15 pedidos** com `Atraso_Dias >= 5`).

### 3.3 [modUtilitariosExportacao.bas](macros/modUtilitariosExportacao.bas)
- `Public Sub AtualizarBasesERelatorios()`: executa `ThisWorkbook.RefreshAll`, dispara as duas macros de relatório, valida o batimento automático contra **R$ 3.448.872,47** (`8.291` pedidos válidos e `628` atrasados) quando sem filtros, e exibe `MsgBox` com o resultado e tempo de execução.
- `Public Sub ExportarRelatoriosPDF()`: configura `xlLandscape` + `FitToPagesWide = 1` e exporta `Resumo_Categorias` e `Relatorio_Atrasos` via `ExportAsFixedFormat Type:=xlTypePDF` para `ThisWorkbook.Path & "\relatorios_gerados\"` com carimbo `YYYYMMDD_HHMM`.
- `Public Sub LimparFiltrosERelatorios()`: limpa `Menu_Operacional!C6:C10` e os quadros gerados preservando cabeçalhos e formatação base.

> **Boas práticas obrigatórias aplicadas:** todas as `Sub`s públicas usam `Option Explicit`, `On Error GoTo ErrHandler`, `Application.ScreenUpdating = False`, `Application.EnableEvents = False` e `Application.Calculation = xlCalculationManual`, com restauração garantida no bloco `CleanExit:`.

---

## 4. Como Importar os Módulos `.bas` e Gerar o `.xlsm` para Demonstração ao Vivo

Como arquivos `.xlsx` não armazenam projeto VBA por padrão OpenXML, siga os passos abaixo no Microsoft Excel antes da apresentação ao vivo:

> **Codificação garantida para VBE:** todos os 3 arquivos `.bas` em `macros/` estão gravados em **ANSI (`Windows-1252` / `cp1252`)** com quebras de linha **CRLF**, evitando *mojibake* (ex.: `Ã§`, `Ã£`) ao importar no Visual Basic Editor do Excel no Windows.

1. Abra **`VarejoMax_04_VBA_Template.xlsx`** no Microsoft Excel (Desktop).
2. Pressione **`Alt + F11`** para abrir o **Visual Basic Editor (VBE)**.
3. No menu superior do VBE, clique em **Arquivo → Importar Arquivo...** (`Ctrl + M`) e importe, um a um, os 3 módulos da pasta `01-vba/macros/`:
   - `modResumoVendas.bas`
   - `modAtrasosLogistica.bas`
   - `modUtilitariosExportacao.bas`
4. *(Opcional — Botões no Menu)* Na aba `Menu_Operacional`, ative a guia **Desenvolvedor → Inserir → Botão (Controle de Formulário)** para criar botões ligados a `AtualizarBasesERelatorios`, `GerarResumoVendasCategoria`, `GerarRelatorioAtrasosTransportadora`, `ExportarRelatoriosPDF` e `LimparFiltrosERelatorios`.
5. No Excel, vá em **Arquivo → Salvar Como** e salve na mesma pasta `01-vba/` com o tipo **Pasta de Trabalho Habilitada para Macro do Excel (`*.xlsm`)** e nome **`VarejoMax_04_VBA.xlsm`**.
6. Pressione **`Alt + F8`**, selecione **`AtualizarBasesERelatorios`** e clique em **Executar** — o `MsgBox` confirmará **RECONCILIAÇÃO OFICIAL: APROVADA (Diferença = R$ 0,00)**.

---

## 5. Resumo dos Números Apurados

- **Pedidos Totais:** 8.996 | **Pedidos Válidos:** 8.291 | **Unidades Válidas:** 16.581
- **Faturamento Bruto (excl. Cancelado):** R$ 3.797.167,40
- **Faturamento Líquido (excl. Cancelado/Devolvido):** R$ 3.448.872,47
- **Ticket Médio:** R$ 415,98
- **Custo Ref. Total:** R$ 2.372.700,48 (**Margem Bruta:** 31,20%)
- **Pedidos Entregues:** 8.180 | **Pedidos Atrasados Válidos:** 628 (7,68%)

Consulte a conferência completa linha a linha em [conferencia_resultados_vba.md](evidencias/conferencia_resultados_vba.md).
