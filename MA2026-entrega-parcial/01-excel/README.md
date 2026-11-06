# Etapa 1 — Excel Avançado (VarejoMax)

> Projeto Evolutivo de Dados — Microinformática Aplicada (ILP590) — Fatec São Paulo
> Tema: Varejo e e-commerce — VarejoMax Brasil (rede omnichannel fictícia)

## Objetivo da etapa

Transformar a planilha bruta de vendas (aba `Vendas_Bruto` do arquivo
`VarejoMax_massa_de_dados.xlsx`, 13.699 linhas de dados) em uma base confiável e
auditável dentro do Excel, seguindo o procedimento do guia-modelo (Etapa 1 do
CaféData): dados originais preservados, colunas tratadas calculadas ao lado dos
originais, buscas contra cadastros oficiais, validações, formatação condicional e
registro de auditoria.

## Arquivo

`VarejoMax_01_Excel.xlsx` — abas:

| Aba | Função |
|---|---|
| LEIA-ME | Apresentação do arquivo e instruções de uso |
| Bruto | Cópia fiel dos dados originais (nunca alterada — rastreabilidade) |
| Tratado | Base de trabalho: 85 colunas com os campos originais + colunas tratadas (datas, textos, valores, chaves) + buscas PROCX contra os cadastros |
| Cadastros | Tabelas de referência: Tab_Produtos (49), Tab_Clientes (1.607), Tab_Cidades (24), listas oficiais (canais, pagamentos, transportadoras) e de-paras (categoria, cidade, status) |
| Auditoria | Registro das inconsistências encontradas (4.733 ocorrências): linha de origem, campo, valor encontrado, tipo de problema, ação necessária, situação e observação |

## O que foi feito (procedimento do guia, adaptado ao tema)

- Inspeção dos dados de origem: 30 colunas misturando pedido, cliente, produto,
  pagamento e logística; datas como texto, moedas como "R$ 1.439,91", CEPs sem zero
  à esquerda, grafias divergentes de cidade/status/categoria, pedidos de teste,
  duplicatas e quantidades inválidas — tudo registrado na aba Auditoria.
- Padronização **sem sobrescrever os originais**: colunas tratadas calculadas ao lado
  (ARRUMAR/PRI.MAIÚCULA, conversão de datas, CEP de 8 dígitos com zero à esquerda,
  desconto normalizado para 0–1, valores monetários).
- Buscas **PROCX** contra Tab_Produtos/Tab_Cidades/Tab_Clientes: categoria, preço de
  referência e custo por produto; cidade/UF/região; cadastro do cliente — sem
  correspondência fica "Não cadastrado"/"Sem correspondência" (visível, não oculto).
- Colunas calculadas de negócio: Valor_Bruto, Valor_Liquido, prazos previsto/real,
  atraso em dias e Status_Atraso (Adiantado / No prazo / Atrasado / Inconsistente / -).
- Validação de dados em lista (categoria, status, canal, pagamento) e formatação
  condicional para campos vazios, quantidades inválidas e preços com desvio.
- Painel de conferência na aba Auditoria (linhas na aba Bruto × linhas em Tab_Vendas).

## Entregas seguintes

- A Etapa 2 automatiza esta limpeza no Power Query para todos os arquivos brutos:
  ver `../02-powerquery/`.
- Os cadastros oficiais desta etapa são a fonte das dimensões da Etapa 3.
