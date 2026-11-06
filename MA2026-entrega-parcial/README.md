# Projeto Evolutivo de Dados — VarejoMax Brasil (N1: Etapas 1–3)

> **Microinformática Aplicada (ILP590)** — Fatec São Paulo — Prof. Veríssimo — 2º semestre/2026
> Tema: **Varejo e e-commerce** — Inteligência de Vendas, Logística e Clientes da
> VarejoMax Brasil (empresa fictícia, rede omnichannel)
> Pergunta central: *"Como os dados de vendas, logística e clientes podem apoiar
> decisões no varejo e no e-commerce?"*

Pipeline do projeto (cada etapa consome a anterior e nada é redigitado):

```
Dado bruto → Excel Avançado (1) → Power Query (2) → Power Pivot/DAX (3) → VBA → SQL → API/Web → Power BI
```

## Estrutura do repositório

```
MA2026-entrega-parcial/
├── 01-excel/          Etapa 1 — Preparação, padronização e auditoria (Excel avançado)
│   ├── VarejoMax_01_Excel.xlsx
│   └── README.md
├── 02-powerquery/     Etapa 2 — Importação, limpeza e transformação (Power Query/M)
│   ├── VarejoMax_02_PowerQuery_VIVO.xlsx
│   └── README.md
├── 03-powerpivot/     Etapa 3 — Modelagem analítica e medidas (Power Pivot/DAX)
│   ├── VarejoMax_03_PowerPivot.xlsx
│   ├── dax/medidas_preliminares.dax
│   ├── docs/modelo_relacional.md · docs/diagrama_modelo.png
│   ├── scripts_m/*.pq (11 consultas)
│   ├── evidencias/checklist_conferencia.md
│   └── README.md
└── 04-documentacao/   Relatório de execução e orientações de reprodução
    ├── Relatorio_Execucao_VarejoMax_N1.docx / .pdf
    └── COMO_REPRODUZIR.md
```

A pasta `MA2026-entrega-final/` (VBA, banco de dados/SQL, APIs e dashboard Power BI —
N2, entrega em 27/11/2026) será acrescentada nas próximas etapas.

## As três etapas da N1 em resumo

| Etapa | Tecnologia | O que entrega | Números-chave |
|---|---|---|---|
| 1 | Excel avançado | Base auditável: Bruto preservado, Tratado com 85 colunas (PROCX, validações, formatação condicional), Cadastros oficiais, Auditoria com 4.733 ocorrências | 13.699 linhas brutas |
| 2 | Power Query | Limpeza automatizada e repetível de toda a pasta de brutos: fVendas consolidada, 6 dimensões, Rejeitados com motivo, conferências | 13.613 aceitas + 86 rejeitadas = 13.699 (dif 0) |
| 3 | Power Pivot & DAX | Modelo estrela (fVendas + 6 dimensões + dCalendario), 6 relacionamentos N:1, dCalendario marcada como Tabela de Datas, 20 medidas DAX, validação e reconciliação | [Faturamento] R$ 3,448,872.47 — dif 0 vs Etapa 2 |

## Regras respeitadas nesta entrega

1. **Nenhum arquivo das etapas 1 e 2 foi alterado.** A Etapa 3 lê a última versão
   salva em disco da Etapa 2 (parâmetro `pArquivoEtapa2`); correções são feitas na
   origem e propagadas com Atualizar Tudo.
2. Nenhum dado é digitado na Etapa 3: tudo vem da Etapa 2 ou é gerado (dCalendario).
3. A regra de ouro da Etapa 3 vale: o modelo não cria nem perde dinheiro — todas as
   diferenças da aba Conferencia são 0.

## Como reproduzir a Etapa 3 no Excel

Ver `03-powerpivot/README.md` (passo a passo com o parâmetro `pArquivoEtapa2`,
Atualizar Tudo, Gerenciar Relações, Marcar como Tabela de Datas e criação das
medidas). O ajuste do caminho do parâmetro é o único necessário por máquina.

## Pendências registradas (transparência)

- fMetas não existe na massa de dados do VarejoMax: o % Meta do modelo CaféData foi
  transposto para [% Vendas Online] e [% Pedidos Atrasados] (mesma técnica
  CALCULATE + DIVIDE). O lugar da fMetas na estrela está reservado.
- lst_Feriados é provisória (11 feriados nacionais fixados) — será substituída pela
  BrasilAPI na etapa de APIs, junto com dClima (Open-Meteo) e dCambio (PTAX).
- Avaliações fora da escala (notas 0 e 6–10, 8 casos) e 18 entregas "Inconsistente"
  (entrega antes do envio): sinalizadas na auditoria da Etapa 1, sem tratamento —
  decisão de negócio pendente.
- Custo praticado em 2026 difere do custo de referência do catálogo (reajuste):
  [Margem Bruta %] 31,2% (referência) × [Margem Realizada %] 30,3% (praticado).
