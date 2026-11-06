# Etapa 2 — Power Query aplicado (VarejoMax)

> Projeto Evolutivo de Dados — Microinformática Aplicada (ILP590) — Fatec São Paulo

## Objetivo da etapa

Automatizar a limpeza da Etapa 1 de forma repetível: o Power Query lê a pasta de
arquivos brutos, aplica as regras de limpeza e entrega tabelas tratadas prontas para
o modelo analítico da Etapa 3. Arquivo novo ou corrigido na pasta? Basta
**Atualizar Tudo**.

## Arquivo

`VarejoMax_02_PowerQuery_VIVO.xlsx` — abas: fVendas (13.613 linhas), Conferencia,
Rejeitados (86), dTransportadora (6), dCanalLoja (10), dCidade (24), dCliente (1.607),
dProduto (49).

## Consultas (grupos no Editor do Power Query)

| Grupo | Consultas |
|---|---|
| 01_Parametros | pPastaBrutos, pArquivoEtapa1, pDataInicio (01/01/2025), pDataFim (31/08/2026) |
| 02_Origem | cad_Etapa1 (lista o arquivo da Etapa 1), cad_Canais, cad_DeParaCategoria, cad_DeParaCidade, cad_DeParaStatus, cad_Pagamentos, cad_Transportadoras, ColunasOrigem (dicionário das 30 colunas) |
| Funções | fnData (datas texto/número → data), fnNumero (moeda texto → número), fnDePara (mesclagem com deduplicação), fnLerVendas (lê a aba Vendas_Bruto numerando as linhas) |
| Staging | stg_Vendas_Bruto (anexa todos os arquivos .xlsx da pasta, guardando Arquivo_Origem e Linha_Origem) |
| Saída | fVendas, Rejeitados, dProduto, dCliente, dCidade, dCanalLoja, dTransportadora |
| Conferência | Conf_Reconciliacao, Conf_Motivos, Conf_Colunas, Conf_FK |

## Regras de limpeza (rastreáveis)

1. Linhas totalmente em branco → Rejeitados ("Linha em branco").
2. Pedidos de teste (TESTE-xxxx) → Rejeitados.
3. Duplicatas exatas (id_pedido + item) → mantém a 1ª, demais → Rejeitados (70).
4. Quantidade zero ou negativa (estorno) → Rejeitados com motivo específico.
5. Data do pedido inválida/ausente ou fora de 01/01/2025–31/08/2026 → Rejeitados.
6. Padronização via de-paras da Etapa 1 (cidade, status, canal, pagamento,
   transportadora); sem correspondência fica explícito ("Sem correspondência",
   "Fora da lista", "Não informado").
7. Loja vazia em canal online = "Online (sem loja)" + chave Chave_Canal_Loja.
8. Datas de logística (envio, prevista, real) convertidas; frete vazio = 0;
   desconto em pontos (5 = 5%) normalizado; CEP de 8 dígitos com zero à esquerda.
9. Colunas calculadas: Valor_Bruto/Valor_Liquido (arredondados), prazos, atraso e
   Status_Atraso.

**Regra de ouro:** linhas lidas − aceitas − rejeitadas = 0, comprovada na aba
Conferencia: **13.699 lidas = 13.613 aceitas + 86 rejeitadas (diferença 0)**.

## Conferencia (números de referência — usados pela Etapa 3)

| Verificação | Esperado | Apurado | Diferença |
|---|---|---|---|
| Linhas Lidas | 13.699 | 13.699 | 0 |
| Linhas Aceitas | 13.613 | 13.613 | 0 |
| Rejeitadas | 86 | 86 | 0 |
| Soma Valor_Liquido (sem Cancelado/Devolvido) | 3.448.872,47 | 3.448.872,47 | 0 |
| Soma Valor_Bruto (sem Cancelado) | 3.797.167,40 | 3.797.167,40 | 0 |
| Duplicatas / Testes / Qtd zero / Qtd negativa | 70 / 6 / 5 / 5 | 70 / 6 / 5 / 5 | 0 |

## Entrega seguinte

A Etapa 3 lê estas tabelas e monta o modelo estrela com medidas DAX:
ver `../03-powerpivot/`.
