# Guia Prático de Montagem do `VarejoMax_Dashboard.pbix`

Este roteiro descreve passo a passo como montar, validar e exportar o relatório final **`VarejoMax_Dashboard.pbix`** e **`VarejoMax_Dashboard.pdf`** no Power BI Desktop (Windows 64-bit) a partir dos artefatos prontos em `04-powerbi-dashboard/`.

---

## Etapa 1 — Configuração do DSN ODBC 64-bit (`varejomax`)

1. Certifique-se de que o banco `varejomax.db` já foi gerado pela etapa SQL/API e contém as 8 views dimensionais + `fato_vendas`.
2. Instale o **SQLite3 ODBC Driver (64-bit)** (Christian Werner — `sqliteodbc_w64.exe`).
3. Abra **Administrador de Fonte de Dados ODBC (64-bit)** (`odbcad32.exe`):
   - Aba **DSN de Sistema** → **Adicionar...** → selecione **SQLite3 ODBC Driver**.
   - **Data Source Name:** `varejomax` *(exatamente em minúsculas)*.
   - **Database Name:** caminho completo para o arquivo `varejomax.db`.
   - Clique em **OK** para salvar.
4. No Power BI Desktop, vá em **Arquivo → Opções e configurações → Opções → Configurações Regionais** e confirme a localidade de importação ou mantenha os scripts M com cultura `"en-US"` (já configurada em todos os arquivos `.pq`).

---

## Etapa 2 — Importação das Consultas M (`m/`)

1. No Power BI Desktop, clique em **Obter dados → Consulta Nula** para abrir o **Editor do Power Query**.
2. Crie o parâmetro **`pConexaoODBC`**:
   - **Gerenciar Parâmetros → Novo Parâmetro** (ou cole o conteúdo de [pConexaoODBC.pq](../m/pConexaoODBC.pq) no Editor Avançado):
     - Nome: `pConexaoODBC`
     - Tipo: `Texto`
     - Valor atual: `dsn=varejomax`
3. Crie uma consulta nula chamada **`tst_Conexao`**, abra **Editor Avançado**, cole [tst_Conexao.pq](../m/tst_Conexao.pq), confirme a credencial ODBC (se solicitada, escolha *Padrão ou Personalizado* sem usuário/senha) e verifique se todas as 9 views retornam suas contagens sem erro.
   - Clique com o botão direito em `tst_Conexao` → desmarque **Habilitar carga**.
4. Crie uma consulta nula para cada tabela abaixo, colando o respectivo script no **Editor Avançado**:
   - `dim_produto` → [dim_produto.pq](../m/dim_produto.pq)
   - `dim_cliente` → [dim_cliente.pq](../m/dim_cliente.pq)
   - `dim_cidade` → [dim_cidade.pq](../m/dim_cidade.pq)
   - `dim_canal_loja` → [dim_canal_loja.pq](../m/dim_canal_loja.pq)
   - `dim_transportadora` → [dim_transportadora.pq](../m/dim_transportadora.pq)
   - `dim_tempo` → [dim_tempo.pq](../m/dim_tempo.pq)
   - `dim_cambio` → [dim_cambio.pq](../m/dim_cambio.pq)
   - `dim_clima` → [dim_clima.pq](../m/dim_clima.pq)
   - `fato_vendas` → [fato_vendas.pq](../m/fato_vendas.pq)
5. Verifique na barra de status de cada consulta se as colunas decimais (`preco_unitario`, `desconto_pct`, `cotacao_venda`, `latitude`, `precipitacao`) foram convertidas corretamente (ex.: `6.5` → `6,5`, nunca `65`).
6. Clique em **Fechar e Aplicar**.

---

## Etapa 3 — Modelagem, Relacionamentos e Propriedades

1. Acesse a exibição **Modelo (Model View)** e configure os 8 relacionamentos (`*:1`, **Ativo**, **Direção de filtro Única**):
   - **R1:** `fato_vendas[id_produto]` → `dim_produto[id_produto]`
   - **R2:** `fato_vendas[id_cliente]` → `dim_cliente[id_cliente]`
   - **R3:** `fato_vendas[cidade]` → `dim_cidade[cidade]`
   - **R4:** `fato_vendas[chave_canal_loja]` → `dim_canal_loja[chave_canal_loja]`
   - **R5:** `fato_vendas[transportadora]` → `dim_transportadora[transportadora]`
   - **R6:** `fato_vendas[data_pedido]` → `dim_tempo[data]`
   - **R7:** `fato_vendas[chave_clima]` → `dim_clima[chave_clima]`
   - **R8:** `fato_vendas[data_pedido]` → `dim_cambio[data]`
   - *Atenção:* **não** crie relacionamento entre `dim_tempo` e `dim_cambio`, nem entre `dim_cidade`/`dim_tempo` e `dim_clima`.
2. **Marcar Tabela de Datas:** clique em `dim_tempo` → **Marcar como tabela de datas** → coluna `data`.
3. **Classificar por coluna (Sort by Column)** em `dim_tempo`:
   - `nome_mes` → `mes`
   - `mes_abrev` → `mes`
   - `ano_mes` → `mes_inicio`
   - `nome_dia` → `dia_semana`
4. **Categoria de dados** (para mapas):
   - `dim_cidade[cidade]` → *Cidade*
   - `dim_cidade[uf]` → *Estado ou Província*
   - `dim_cidade[latitude]` → *Latitude* (*Não resumir*)
   - `dim_cidade[longitude]` → *Longitude* (*Não resumir*)
   - `fato_vendas[cep]` → *Código Postal*
5. **Ocultar chaves técnicas** na visão de relatório: `fato_vendas[id_produto]`, `fato_vendas[id_cliente]`, `fato_vendas[chave_canal_loja]`, `fato_vendas[chave_clima]`, `fato_vendas[canal]`, `fato_vendas[loja]`.

---

## Etapa 4 — Criação da Tabela `_Medidas` e Aplicação do DAX

1. Na guia **Modelagem → Nova Tabela**, crie:
   ```dax
   _Medidas = { BLANK() }
   ```
2. Oculte a coluna `_Medidas[Value]`.
3. Abra a **DAX Query View** (ícone no menu esquerdo), cole todo o conteúdo de [medidas_finais.dax](../dax/medidas_finais.dax) e execute:
   - Clique em **Update model with changes** (ou *Atualizar modelo com alterações*) para gravar todas as medidas em `_Medidas`.
   - Execute a consulta `EVALUATE ROW(...)` ao final do arquivo para conferir rapidamente os KPIs principais.
4. Na exibição **Modelo**, selecione os grupos de medidas e preencha o campo **Pasta de exibição** conforme o cabeçalho de cada bloco (`01 Vendas`, `02 Custos & Margem`, `03 Logística & Satisfação`, `04 Clientes`, `05 Canais & Mix`, `06 Tempo, Câmbio & Clima`, `07 Rótulos`, `99 Conferência`).
5. Aplique os formatos numéricos:
   - Monetários (`Faturamento*`, `Ticket*`, `Custo*`, `Lucro*`, `Frete*`, `Desconto Concedido`): `R$ #,##0.00`
   - Percentuais (`*%*`, `*Margem*`): `0.0%`
   - Contagens (`Qtd*`, `Clientes*`, `Ranking*`): `#,##0`
   - Índices/Médias (`Itens por Pedido`, `Nota*`, `Prazo*`, `Atraso Médio*`, `Dólar*`, `Temp*`, `Precipitação*`, `Frequência*`, `Qtd Parcelas Média`): `0.00`

---

## Etapa 5 — Aplicação do Tema Corporativo

1. Vá em **Exibição → Temas → Procurar temas**.
2. Selecione [tema_varejomax.json](../tema/tema_varejomax.json).
3. Confirme que o fundo das páginas assumiu o tom off-white `#FAF7F2` e os títulos/cartões utilizam a paleta `#1E3A5F` (azul petróleo) e `#D95D39` (terracota).

---

## Etapa 6 — Construção das 5 Páginas Principais (16:9 — `1280x720`)

### Cabeçalho padrão (presente nas 5 páginas principais)
- Faixa superior (`#1E3A5F`, altura `68px`): Título da página à esquerda + Cartão ou Caixa de texto vinculada a `[Título Dinâmico Período]` e `[Subtítulo Filtros Ativos]` à direita.
- Barra de **Segmentadores (Slicers)** no topo/esquerda: `dim_tempo[ano]`, `dim_tempo[mes_abrev]`, `dim_cidade[regiao]`, `dim_canal_loja[canal]`.

### Página 1 — `Visão Executiva`
- **Linha de 5 Cartões KPI (topo):** `[Faturamento]`, `[Qtd Pedidos]`, `[Ticket Médio]`, `[Margem Bruta %]` (cor condicional via `[Cor Margem Bruta]`), `[% Pedidos Atrasados]` (cor condicional via `[Cor Status Atraso]`) + subtítulo de `[Rótulo MoM Faturamento]`.
- **Gráfico de Linhas e Colunas Empilhadas/Agrupadas:** Eixo X `dim_tempo[ano_mes]`, Colunas `[Faturamento]`, Linha `[Margem Bruta %]` (ou `[Faturamento Mês Anterior (MoM)]`).
- **Gráfico de Barras Horizontais:** `[Faturamento]` por `dim_produto[categoria]` com tooltip vinculado à página `Dica Resumo`.
- **Gráfico de Rosca:** `[Faturamento]` por `dim_canal_loja[canal]`.
- **Tabela Resumo Regional:** `dim_cidade[regiao]`, `[Faturamento]`, `[Qtd Pedidos]`, `[Margem Bruta %]`, `[% Pedidos Atrasados]`.

### Página 2 — `Produtos e Categorias`
- **4 Cartões KPI:** `[Faturamento Bruto]`, `[Desconto Concedido]`, `[Lucro Bruto]`, `[Qtd Itens]`.
- **Gráfico de Dispersão (Scatter):** Eixo X `[Faturamento]`, Eixo Y `[Margem Bruta %]`, Tamanho `[Qtd Itens]`, Detalhes `dim_produto[produto]`, Legenda `dim_produto[categoria]`.
- **Gráfico de Barras (Top 10 Produtos):** `dim_produto[produto]` com filtro Top N = 10 por `[Faturamento]`, tooltip `Dica Resumo`.
- **Matriz Hierárquica:** Linhas `dim_produto[categoria]` → `dim_produto[marca]` → `dim_produto[produto]`; Valores `[Faturamento]`, `[% Share Categoria]`, `[Custo Total]`, `[Custo Realizado]`, `[Margem Bruta %]`, `[Margem Realizada %]`, `[Ranking Produto Faturamento]`.

### Página 3 — `Clientes`
- **5 Cartões KPI:** `[Clientes Ativos]`, `[Clientes Novos]`, `[Clientes Recorrentes]`, `[% Taxa Recompra]`, `[Frequência Média de Compra]`.
- **Gráfico de Barras Agrupadas por Região:** Eixo `dim_cliente[regiao]`, Valores `[Clientes Novos]` e `[Clientes Recorrentes]`.
- **Gráfico de Rosca:** `[Faturamento Clientes Recorrentes]` vs `[Faturamento] - [Faturamento Clientes Recorrentes]` (ou `Faturamento` por perfil).
- **Gráfico de Colunas:** `[Faturamento]` e `[Ticket Médio]` por `fato_vendas[forma_pagamento]`, além de cartão `[% Share Pix]` e `[Qtd Parcelas Média]`.
- **Mapa / Barras por Cidade:** `dim_cidade[cidade]` (`latitude`/`longitude`), tamanho `[Faturamento por Cliente]`, cor `[% Taxa Recompra]`.

### Página 4 — `Logística e Clima`
- **5 Cartões KPI:** `[% Entregas no Prazo (OTD)]`, `[Qtd Pedidos Atrasados]`, `[Prazo Previsto Médio (dias)]`, `[Prazo Real Médio (dias)]`, `[Nota Média Cliente]`.
- **Tabela de Transportadoras (com Drill-through ativo):** `dim_transportadora[transportadora]`, `[Qtd Pedidos]`, `[Qtd Pedidos Atrasados]`, `[% Pedidos Atrasados]` (formatação condicional), `[Atraso Médio (dias)]`, `[Nota Média Cliente]`.
- **Gráfico de Colunas — Hipótese 1 (Atraso x Avaliação):** Eixo `fato_vendas[status_atraso]`, Colunas `[Nota Média Cliente]`, Linha `[% Taxa Devolução]`.
- **Gráfico Combinado — Hipótese Clima:** Eixo `dim_clima[faixa_precipitacao]`, Colunas `[% Pedidos Atrasados]`, Linha `[Precipitação Média mm]` (e cartões comparativos `[% Pedidos Atrasados Dias Chuvosos]` vs `[% Pedidos Atrasados Dias Secos]`).
- **Mapa Geográfico:** `dim_cidade[latitude]` e `dim_cidade[longitude]`, bolhas por `[Qtd Pedidos Atrasados]`, saturação por `[% Pedidos Atrasados]`.

### Página 5 — `Canais e Lojas`
- **4 Cartões KPI:** `[Faturamento Online]`, `[Faturamento Loja Física]`, `[% Share Online]`, `[Dólar PTAX Médio]`.
- **Gráfico de Colunas / Barras por Loja:** `dim_canal_loja[loja]` (com Drill-through para `Detalhe Transportadora / Loja`), Valores `[Faturamento]`, `[Ticket Médio]`, `[Margem Bruta %]`.
- **Gráfico de Linhas (Câmbio x Eletrônicos):** Eixo `dim_tempo[ano_mes]`, Valores `[Dólar PTAX Médio]` (eixo secundário) e `Faturamento` filtrado ou comparado por `dim_produto[categoria]` = `"Eletrônicos"`.
- **Matriz Canal × Forma de Pagamento:** Linhas `dim_canal_loja[tipo_canal]` → `dim_canal_loja[canal]`, Colunas `fato_vendas[forma_pagamento]`, Valores `[Faturamento]` e `[% Share Canal]`.

---

## Etapa 7 — Construção das 3 Páginas de Apoio

### Página 6 — `Detalhe Transportadora / Loja` (Drill-through)
- Em **Informações da página → Tipo de página:** selecione **Detalhamento (Drill-through)**.
- Adicione como campos de detalhamento: `dim_transportadora[transportadora]` e `dim_canal_loja[loja]` (opção *Manter todos os filtros* ativada).
- Insira botão **Voltar** automático no canto superior esquerdo.
- Visuais: Cartões de resumo do item detalhado (`[Faturamento]`, `[Qtd Pedidos]`, `[% Pedidos Atrasados]`, `[Nota Média Cliente]`), gráfico mensal (`dim_tempo[ano_mes]`) e tabela detalhada de pedidos (`fato_vendas[id_pedido]`, `dim_tempo[data]`, `dim_cidade[cidade]`, `fato_vendas[status_pedido]`, `fato_vendas[status_atraso]`, `[Faturamento]`, `[Nota Média Cliente]`).

### Página 7 — `Dica Resumo` (Tooltip Oculta `320x240`)
- Clique com o botão direito na aba → **Ocultar página**.
- Em **Configurações da tela → Tipo:** selecione **Dica de ferramenta (Tooltip)** e ajuste as dimensões para **`320 px × 240 px`**.
- Em **Informações da página**, ative **Permitir uso como dica de ferramenta**.
- Adicione 4 mini cartões compactos + 1 mini barra:
  - `[Faturamento]`, `[Margem Bruta %]`, `[Qtd Pedidos]`, `[% Pedidos Atrasados]`, `[Nota Média Cliente]`.
- Nos visuais das páginas 1, 2, 4 e 5, em **Dicas de ferramenta → Tipo: Página de relatório**, aponte para `Dica Resumo`.

### Página 8 — `Conferência` (Oculta)
- Clique com o botão direito na aba → **Ocultar página**.
- Insira cartões e matrizes com **todas** as medidas da pasta `99 Conferência` + os números de referência (`13.613` linhas fato, `8.291` pedidos válidos, `8.996` pedidos totais, `R$ 3.797.167,40` bruto, `R$ 3.448.872,47` líquido, `628` atrasos, `8` avaliações inválidas, `0` pedidos sem clima).

---

## Etapa 8 — Sincronização, Bloqueio de Filtros, Interações e Validação Final

1. **Sincronizar Segmentações (Exibição → Sincronizar segmentações):**
   - Sincronize `dim_tempo[ano]`, `dim_tempo[mes_abrev]`, `dim_cidade[regiao]` e `dim_canal_loja[canal]` entre as **5 páginas principais** (visível e sincronizado). Não sincronize com `Dica Resumo` nem com `Conferência`.
2. **Bloqueio de Filtros (Painel de Filtros):**
   - Bloqueie (ícone de cadeado) e oculte (ícone de olho) filtros técnicos fixos aplicados a visuais específicos (ex.: Top 10 produtos ou filtro `Eletrônicos`), evitando que o usuário final os altere por engano.
3. **Editar Interações (Formato → Editar interações):**
   - Nas páginas com gráficos comparativos de série temporal (`ano_mes`), configure para que cliques em gráficos de categoria/canal **filtrem** (ícone funil, não apenas realcem) os cartões e tabelas, e garanta que cartões de contexto fixo (quando aplicável) respeitem apenas os segmentadores globais.
4. **Validação Completa DAX:**
   - Abra a **DAX Query View**, execute todos os blocos de [validacao.dax](../dax/validacao.dax) (`V-K4`, `V-K5`, `V-C1`..`V-C10`, `V-CONF`) e registre os valores obtidos em [validacao_resultados.md](../evidencias/validacao_resultados.md).
5. **Salvar e Exportar:**
   - Salve o arquivo em `04-powerbi-dashboard/VarejoMax_Dashboard.pbix`.
   - Exporte em **Arquivo → Exportar → Exportar para PDF** como `04-powerbi-dashboard/VarejoMax_Dashboard.pdf` (garantindo que as 5 páginas principais estejam visíveis e sem visuais cortados).
