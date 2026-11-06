# Como reproduzir a Etapa 3 no Excel (guia rápido)

Requisitos: Excel para **Windows** (Microsoft 365 / 2021 / 2024, preferência 64 bits),
Power Pivot ativado, arquivos das Etapas 1 e 2 salvos e fechados.

1. **Ativar o Power Pivot** (uma vez por usuário): Arquivo → Opções → Suplementos →
   Gerenciar: *Suplementos de COM* → Ir → marcar *Microsoft Power Pivot for Excel* → OK.
2. **Abrir** `VarejoMax_03_PowerPivot.xlsx` (pasta 03-powerpivot).
3. **Ajustar o parâmetro**: Dados → Obter Dados → Iniciar Editor do Power Query →
   Gerenciar Parâmetros → `pArquivoEtapa2` → caminho completo de
   `VarejoMax_02_PowerQuery_VIVO.xlsx` no seu computador. (Único ajuste por máquina;
   as demais consultas dependem dele.)
4. **Atualizar Tudo** (Dados → Atualizar Tudo). As 11 consultas recarregam: fVendas,
   dProduto, dCliente, dCidade, dCanalLoja, dTransportadora, Rejeitados (da Etapa 2) e
   dCalendario (gerada). Conferência rápida: aba dCalendario com 608 dias
   (01/01/2025 a 31/08/2026) sem lacunas.
5. **Carregar no Modelo de Dados**: para cada consulta do grupo 03_Modelo, botão
   direito → Carregar Para… → *Apenas Criar Conexão* + **Adicionar estes dados ao
   Modelo de Dados** (sem criar planilhas). pArquivoEtapa2, cad_Etapa2 e lst_Feriados
   ficam somente como conexão.
6. **Relacionamentos**: Power Pivot → Gerenciar → Exibição de Diagrama → arrastar as
   colunas (ou Dados → Relações → Novo): R1 fVendas[ID_Produto]→dProduto[ID_Produto];
   R2 fVendas[ID_Cliente]→dCliente[ID_Cliente]; R3 fVendas[Cidade]→dCidade[Cidade];
   R4 fVendas[Chave_Canal_Loja]→dCanalLoja[Chave_Canal_Loja];
   R5 fVendas[Transportadora]→dTransportadora[Transportadora];
   R6 fVendas[Data_Pedido]→dCalendario[Data]. Conferir em Design → Gerenciar Relações:
   6 relações, todas Ativo = Sim, *:1, direção única.
7. **Marcar como Tabela de Datas**: guia dCalendario → Design → Marcar como Tabela de
   Datas → coluna `Data`. Também Classificar por Coluna: NomeMes por Mes, MesAbrev por
   Mes, NomeDia por DiaSemana.
8. **Criar as medidas**: Power Pivot → Medidas → Nova Medida, tabela-base fVendas,
   colando cada fórmula de `../03-powerpivot/dax/medidas_preliminares.dax`
   (Windows pt-BR: trocar vírgula por ponto e vírgula) com o formato indicado.
   São 20 medidas: 15 de negócio + 5 Conf.
9. **Validar**: recriar as tabelas dinâmicas da aba Validacao a partir de
   *Do Modelo de Dados* (Tabela 1: dCanalLoja[Canal] × dCalendario[NomeMes] ×
   [Faturamento]; Tabela 2: dProduto[Categoria] × [Faturamento]/[Custo Total]/
   [Margem Bruta %]; Tabela 3 e 4 conforme a aba; Tabela 5: [Cresc. MoM %] por mês) e
   conferir com os valores de referência impressos.
10. **Conferir**: na aba Conferencia, coluna "Modelo (Etapa 3)" com
    `=VALORCUBO("ThisWorkbookDataModel";"[Measures].[<medida>]")` e Diferença = C−B:
    todas 0. Capturar telas (diagrama, Gerenciar Relações, Gerenciar Medidas,
    Validacao, Conferencia) para `03-powerpivot/evidencias/`.

Teste de linhagem (quando autorizado a mexer na Etapa 1): acrescente uma linha ao
de-para na Etapa 1 → salve → Atualizar Tudo na Etapa 2 → salve → Atualizar Tudo na
Etapa 3: os números mudam sem digitar nada aqui; depois desfaça e atualize de novo
(guia, Seção 6.5). Nesta entrega as etapas 1 e 2 foram preservadas sem alteração.
