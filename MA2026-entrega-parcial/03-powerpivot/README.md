# Etapa 3 — Power Pivot & DAX (VarejoMax)

> Projeto Evolutivo de Dados — Microinformática Aplicada (ILP590) — Fatec São Paulo
> Tema: Varejo e e-commerce — Inteligência de Vendas, Logística e Clientes da VarejoMax Brasil

## Objetivo da etapa

Organizar as tabelas tratadas da Etapa 2 em um **modelo de dados** (as tabelas passam a
"conversar" por meio de relacionamentos) e criar **medidas DAX** — cálculos escritos uma
única vez que respondem corretamente em qualquer recorte (canal, mês, categoria,
região, transportadora). Este modelo será reaproveitado no Power BI (Etapa final, N2).

**Entrada obrigatória:** `../02-powerquery/VarejoMax_02_PowerQuery_VIVO.xlsx`
(ler pelo parâmetro `pArquivoEtapa2` — nenhum dado é digitado nesta etapa).
**Os arquivos das Etapas 1 e 2 não foram alterados** (regra do projeto: correção
sempre na origem, propagada com Atualizar Tudo).

## Conteúdo da pasta

| Caminho | Conteúdo |
|---|---|
| `VarejoMax_03_PowerPivot.xlsx` | Pasta de trabalho da etapa: abas LEIA-ME, Modelo, dCalendario, MedidasDAX, Validacao e Conferencia + conexão Power Query embutida (11 consultas: pArquivoEtapa2, cad_Etapa2, lst_Feriados, fVendas, dProduto, dCliente, dCidade, dCanalLoja, dTransportadora, Rejeitados, dCalendario) |
| `dax/medidas_preliminares.dax` | Código de todas as medidas, copiado da Área de Cálculo / Gerenciar Medidas (UTF-8) |
| `docs/modelo_relacional.md` | Documentação do modelo: tabelas, grão, chaves, 6 relacionamentos, cardinalidades e regras de negócio |
| `docs/diagrama_modelo.png` | Diagrama do modelo estrela (equivalente à captura da Exibição de Diagrama) |
| `scripts_m/*.pq` | Código M de cada consulta deste arquivo (Editor Avançado → copiar) |
| `evidencias/` | Checklist de conferência da entrega (Seção 7.4 do guia) e lugar para as capturas de tela |

## Como abrir

1. Abra no **Excel para Windows** (o Power Pivot não existe no Mac nem no Excel Web),
   com o suplemento ativado: Arquivo → Opções → Suplementos → Suplementos de COM →
   *Microsoft Power Pivot for Excel*.
2. Em `Dados → Obter Dados → Iniciar Editor do Power Query → Gerenciar Parâmetros`,
   ajuste `pArquivoEtapa2` para o caminho completo de
   `VarejoMax_02_PowerQuery_VIVO.xlsx` **no seu computador** (é o único ajuste necessário).
3. `Dados → Atualizar Tudo`. As consultas recarregam o arquivo da Etapa 2.
4. `Power Pivot → Gerenciar`: confira as guias das tabelas, `Design → Gerenciar Relações`
   (6 relações ativas) e `Design → Marcar como Tabela de Datas` na guia dCalendario
   (coluna Data).
5. Recrie as medidas no Power Pivot colando de `dax/medidas_preliminares.dax`
   (no Windows pt-BR, trocar vírgulas por ponto e vírgula) e as tabelas dinâmicas da aba
   Validacao a partir do Modelo de Dados. A aba Conferencia deve mostrar diferença 0.

## Tabelas e relacionamentos

Modelo estrela com **fVendas** no centro (13.613 linhas — um item de pedido aceito) e
6 dimensões: dProduto (49), dCliente (1.607), dCidade (24), dCanalLoja (10),
dTransportadora (6) e dCalendario (608 dias, 01/01/2025 a 31/08/2026, marcada como
Tabela de Datas). Rejeitados (86 linhas) fica fora da estrela, para auditoria.
Detalhes completos em `docs/modelo_relacional.md`.

## Medidas e a pergunta que cada uma responde

| Medida | Pergunta | Valor (base atual) |
|---|---|---|
| [Faturamento Bruto] | Quanto vendemos antes dos descontos? | R$ 3,797,167.40 |
| [Faturamento] | Quanto entrou (líquido, sem cancelados/devolvidos)? | R$ 3,448,872.47 |
| [Qtd Pedidos] | Quantas vendas (pedidos)? | 8,291 |
| [Ticket Médio] | Quanto o cliente gasta por pedido? | R$ 415.98 |
| [Qtd Itens] / [Itens por Pedido] | Tamanho típico do pedido | 16,581 itens / 2.00 por pedido |
| [Custo Total] | Custo do que foi vendido (cadastro) | R$ 2,372,700.48 |
| [Margem Bruta %] | Margem sobre o faturamento | 31.2% |
| [Custo Realizado] / [Margem Realizada %] | Com o custo praticado (2026 reajustado) | R$ 2,402,823.74 / 30.3% |
| [Faturamento Online] | Quanto vende o e-commerce? | R$ 3,100,830.71 |
| [% Vendas Online] | Participação do e-commerce | 89.9% |
| [Qtd Pedidos Atrasados] / [% Pedidos Atrasados] | Entregas fora do prazo | 628 / 7.6% |
| [Cresc. MoM %] | Crescimento mês a mês (inteligência de tempo) | ex.: 2025-02 −19,4%; 2025-03 +58,7% |
| Conf ... (5 medidas) | Conferência: linhas, órfãs, PKs, diferença | 13.613 / 86 / 13.699; 0; 0; 0 |

## Resultados da conferência

- Linhas: 13.699 lidas = 13.613 aceitas + 86 rejeitadas (igual à Conf_Reconciliacao da Etapa 2).
- [Faturamento] = **R$ 3.448.872,47** = soma de Valor_Liquido da Etapa 2 (excluindo
  Cancelados e Devolvidos) — diferença **0**.
- [Faturamento Bruto] = **R$ 3.797.167,40** = soma de Valor_Bruto sem cancelados (Etapa 2) — diferença **0**.
- Integridade: 0 FKs órfãs, 0 PKs duplicadas, 0 vendas fora do calendário.

## Adaptações ao tema (transposição CaféData → VarejoMax)

- **dFilial** → desdobrada em **dCidade** (destino) e **dCanalLoja** (canal × loja,
  com chave composta "Canal | Loja").
- **fMetas / [% Meta]**: a massa de dados do VarejoMax não traz tabela de metas. O papel
  de "realizado ÷ esperado" foi transposto para **[% Vendas Online]** e
  **[% Pedidos Atrasados]** (mesma técnica: CALCULATE + DIVIDE). Se a diretoria publicar
  metas mensais, a tabela fMetas (Cod, AnoMes, Meta, Mes_Inicio) entra na estrela ligada
  a dCanalLoja/dCidade e dCalendario sem tocar nas medidas existentes.
- **[% Bebidas Frias]** → **[% Vendas Online]** (participação de um subconjunto do
  negócio — aqui, o e-commerce — sobre o total).
- **[Faturamento]** herdou a regra de negócio da Etapa 2: exclui pedidos Cancelados e
  Devolvidos (CALCULATE com filtros de Status_Pedido) — no CaféData todos os cupons
  valem.
- **[Custo Total]**: usa dProduto[Custo_Ref] (equivalente ao Custo do CaféData);
  [Custo Realizado] complementa com o custo praticado (reajuste de 2026).

## Pendências herdadas e próximos passos

- Avaliação de clientes com notas fora da escala (0 e 6–10 aparecem 8 vezes) e 18 casos
  "Inconsistente" em Status_Atraso (entrega antes do envio): sinalizados na auditoria da
  Etapa 1, sem tratamento — decisão de negócio pendente.
- Etapa de APIs: substituir lst_Feriados pela BrasilAPI e acrescentar dClima/dCambio
  (lugares já reservados no diagrama).
- N2: banco de dados (SQL) e Power BI consumirão este mesmo modelo estrela.
