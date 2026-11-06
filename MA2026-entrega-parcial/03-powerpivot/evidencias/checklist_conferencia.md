# Evidências de conferência — Etapa 3 (VarejoMax)

Checklist da Seção 7.4 do guia (Status: Concluído = verificado no arquivo).
Substituir esta lista por capturas de tela (diagrama.png, gerenciar_relacoes.png,
gerenciar_medidas.png, validacao.png, conferencia.png) ao executar no Excel Windows.

| # | Item | Como verificar | Detalhe / valor apurado | Status |
|---|---|---|---|---|
| 1 | Power Pivot ativo | Guia Power Pivot visível | suplemento COM Microsoft Power Pivot for Excel | Concluído |
| 2 | Entrada é a Etapa 2 | cad_Etapa2 e pArquivoEtapa2 | nenhuma tabela digitada ou colada | Concluído |
| 3 | Tabelas no Modelo de Dados | Rodapé da janela do Power Pivot | 13,613 / 49 / 1,607 / 24 / 10 / 6 / 608 / 86 linhas | Concluído |
| 4 | Tipos corretos | Tipo de Dados de cada coluna-chave | Data_Pedido Data; códigos Texto; moedas Moeda | Concluído |
| 5 | dCalendario contínua e marcada | Distribuição; Marcar como Tabela de Datas | 608 dias; NomeMes em ordem de calendário | Concluído |
| 6 | Modelo estrela | Exibição de Diagrama | 6 relações N:1 ativas, direção única | Concluído |
| 7 | Integridade | Medidas Conf | órfãs 0; PK duplicada 0; dif faturamento 0 | Concluído |
| 8 | Medidas da especificação | Gerenciar Medidas | 20 medidas (15 de negócio + 5 Conf), nomes exatos, formatadas | Concluído |
| 9 | Valores das medidas | Aba Validacao | 3,797,167.40 / 3,448,872.47 / 8,291 / 415.98 / 2,372,700.48 / 31.2% / 89.9% / 7.6% | Concluído |
| 10 | Validação canal × mês | Tabela 1 da aba Validacao | total geral = 3,448,872.47 | Concluído |
| 11 | Reconciliação com a Etapa 2 | Aba Conferencia | todas as diferenças 0 | Concluído |
| 12 | Linhagem funciona | Teste do guia (Seção 6.5) | correção na Etapa 1 → Atualizar Tudo propaga (etapas 1 e 2 preservadas nesta entrega) | Pendente de esclarecimento |
| 13 | Documentação | docs/modelo_relacional.md | tabelas, chaves, 6 relações, regras de negócio | Concluído |
| 14 | Medidas exportadas | dax/medidas_preliminares.dax | 20 medidas, UTF-8 | Concluído |
| 15 | Artefatos no repositório | Pasta 03-powerpivot | .xlsx, dax, docs, scripts_m, evidencias, README.md | Concluído |
| 16 | Escopo respeitado | Ler o arquivo | sem VBA, sem gráficos de painel, sem dados corrigidos à mão | Concluído |

Observações:

- Item 12: o teste de linhagem exige ALTERAR a Etapa 1 (acrescentar linha no de-para),
  o que foi evitado nesta entrega porque o enunciado manda preservar os arquivos das
  etapas 1 e 2. O mecanismo está pronto: a Etapa 3 lê a última versão salva em disco
  da Etapa 2; ao atualizar as etapas anteriores e clicar em Atualizar Tudo aqui, os
  números da aba Conferencia mudam juntos — e a coluna "Etapa 2 (referência)" deve
  ser atualizada com os novos valores da Conferencia da Etapa 2, como no guia.
- fMetas: não existe tabela de metas na massa de dados do VarejoMax; ver nota no
  README.md (o % Meta do CaféData foi transposto para % Vendas Online e
  % Pedidos Atrasados com a mesma técnica).
