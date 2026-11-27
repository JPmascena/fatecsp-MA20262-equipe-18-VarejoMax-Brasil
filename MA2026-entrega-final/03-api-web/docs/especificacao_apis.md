# Especificação Técnica das APIs Públicas REST — VarejoMax Brasil

Este documento detalha os contratos, parâmetros, regras de transformação e integração relacional/analítica das três APIs REST consumidas na camada **`03-api-web/`**.

---

## 1. Padrão Geral de Saída (CSVs)

Todos os arquivos gerados em `dados_api/` seguem um padrão único para compatibilidade direta com **SQLite** e **Power BI** (cultura `en-US` na importação numérica):

- **Codificação:** `UTF-8` (sem BOM)
- **Delimitador de colunas:** `;` (ponto e vírgula)
- **Separador decimal:** `.` (**ponto decimal** — 4 casas decimais em câmbio PTAX; 1 casa decimal em variáveis climáticas)
- **Formato de datas:** ISO 8601 (`yyyy-MM-dd`)
- **Recorte temporal oficial:** `2025-01-01` a `2026-08-31` (**608 dias corridos**)

---

## 2. API 1 — Câmbio Dólar PTAX (Banco Central do Brasil — Olinda)

| Item | Especificação |
|---|---|
| **Provedor** | Banco Central do Brasil (Portal Dados Abertos / Serviço Olinda OData v1) |
| **Método / Autenticação** | `GET` público, sem chave de API |
| **Endpoint** | `https://olinda.bcb.gov.br/olinda/servico/PTAX/versao/v1/odata/CotacaoDolarPeriodo(dataInicial=@dataInicial,dataFinalCotacao=@dataFinalCotacao)` |
| **Query Parameters** | `@dataInicial='12-20-2024'`, `@dataFinalCotacao='08-31-2026'`, `$format=json`, `$select=cotacaoCompra,cotacaoVenda,dataHoraCotacao` |
| **Formato de data na URL** | `MM-DD-YYYY` (obrigatório na API Olinda BCB) |
| **Saída CSV** | `dados_api/cambio_ptax_2025_2026.csv` (**608 linhas**) |
| **Script Power Query M** | `scripts_m/fnPTAX_BCB.pq` |

### Regras de Transformação (PTAX)
1. **Janela de lookback (`12-20-2024`):** iniciada 12 dias antes de `2025-01-01` para assegurar uma cotação útil válida em dezembro de 2024, pois `2025-01-01` (Confraternização Universal) não possui boletim oficial.
2. **Deduplicação intradiária:** quando há mais de um boletim no mesmo dia útil, mantém-se exclusivamente o registro com maior `dataHoraCotacao` (boletim de fechamento).
3. **Série contínua + Carry-forward:** gera-se o calendário completo de `2025-01-01` a `2026-08-31` (608 dias). Sábados, domingos e feriados recebem a cotação do último dia útil anterior (`Table.FillDown` em M / carry-forward em Python).
4. **Rastreabilidade (`boletim_origem`):**
   - `Oficial BCB` (418 dias no período)
   - `Carry-forward` (190 dias no período)
5. **Faixas analíticas (`faixa_dolar`, base `cotacao_venda`):**
   Considerando a amplitude real observada no período (`R$ 4.8973` a `R$ 6.2086`), adotaram-se 5 faixas equilibradas:
   - `1-Abaixo de R$ 5,10`
   - `2-R$ 5,10 a R$ 5,39`
   - `3-R$ 5,40 a R$ 5,69`
   - `4-R$ 5,70 a R$ 5,99`
   - `5-R$ 6,00 ou mais`

### Esquema (`cambio_ptax_2025_2026.csv`)
- `data` (`TEXT` / `DATE`, `yyyy-MM-dd`, PK)
- `cotacao_compra` (`REAL`, 4 casas decimais, `.` decimal)
- `cotacao_venda` (`REAL`, 4 casas decimais, `.` decimal)
- `boletim_origem` (`TEXT`: `Oficial BCB` | `Carry-forward`)
- `faixa_dolar` (`TEXT`)

---

## 3. API 2 — Clima Histórico Diário (Open-Meteo Archive API)

| Item | Especificação |
|---|---|
| **Provedor** | Open-Meteo Historical Weather API (ERA5 / reanálise meteorológica) |
| **Método / Autenticação** | `GET` público, sem chave de API |
| **Endpoint** | `https://archive-api.open-meteo.com/v1/archive` |
| **Query Parameters** | `latitude={lat}`, `longitude={lon}`, `start_date=2025-01-01`, `end_date=2026-08-31`, `daily=temperature_2m_max,temperature_2m_min,precipitation_sum`, `timezone=America/Sao_Paulo` |
| **Origem das coordenadas** | `02-bancodedados-sql/csv/dCidade.csv` (24 cidades, vírgula convertida para ponto nos parâmetros) |
| **Controle de taxa** | Intervalo de `0.3s` entre cidades + retry com backoff exponencial para HTTP `429`/`5xx` |
| **Saída CSV** | `dados_api/clima_24_cidades_2025_2026.csv` (**14.592 linhas** = `24 × 608`) |
| **Script Power Query M** | `scripts_m/fnClima_OpenMeteo.pq` |

### Regras de Transformação (Clima)
1. **Fuso horário:** `timezone=America/Sao_Paulo` garante alinhamento dos dias civis (`daily.time`) com o calendário comercial da VarejoMax Brasil.
2. **Integridade da chave `cidade`:** a grafia exata da coluna `Cidade` em `dCidade.csv` (incluindo acentos e espaços, como `São Paulo`, `Brasília`, `Ribeirão Preto`, `Florianópolis`, `São Luís`, `João Pessoa`, `Maceió`, `Goiânia`, `Belém`, `Vitória`, `Cuiabá`) é preservada 100% idêntica para junção exata no SQLite e Power BI.
3. **Tratamento de nulos:** caso a API retorne `null` em dias recentes, aplica-se fallback controlado (carry-forward para temperaturas e `0.0` para precipitação). Na coleta oficial deste ciclo ocorreram **0 valores nulos**.
4. **Arredondamento e classificação (`faixa_precipitacao`):**
   Os valores são arredondados para 1 casa decimal antes da classificação:
   - `0-Sem chuva (0 mm)` se `precipitacao == 0`
   - `1-Chuva fraca (0,1-10 mm)` se `0 < precipitacao <= 10`
   - `2-Chuva moderada (10-30 mm)` se `10 < precipitacao <= 30`
   - `3-Chuva forte (>30 mm)` se `precipitacao > 30`

### Esquema (`clima_24_cidades_2025_2026.csv`)
- `cidade` (`TEXT`, FK para `dCidade.Cidade`)
- `data_ref` (`TEXT` / `DATE`, `yyyy-MM-dd`, FK para calendário)
- `temp_max` (`REAL`, °C, 1 casa decimal, `.` decimal)
- `temp_min` (`REAL`, °C, 1 casa decimal, `.` decimal)
- `precipitacao` (`REAL`, mm, 1 casa decimal, `.` decimal)
- `faixa_precipitacao` (`TEXT`)

---

## 4. API 3 — Feriados Nacionais (BrasilAPI)

| Item | Especificação |
|---|---|
| **Provedor** | BrasilAPI — Feriados Nacionais |
| **Método / Autenticação** | `GET` público, sem chave de API |
| **Endpoints** | `https://brasilapi.com.br/api/feriados/v1/2025` e `.../2026` |
| **Recorte** | Datas filtradas entre `2025-01-01` e `2026-08-31` |
| **Saída CSV** | `dados_api/feriados_nacionais_2025_2026.csv` (**22 linhas**) |
| **Script Power Query M** | `scripts_m/fnFeriados_BrasilAPI.pq` |

### Regras de Transformação (Feriados)
- Combinação dos endpoints de `2025` (14 feriados) e `2026` (8 feriados até `2026-08-31`), totalizando **22 feriados nacionais**.
- Padronização do campo `type` (`"national"` -> `"Nacional"`).

### Esquema (`feriados_nacionais_2025_2026.csv`)
- `data` (`TEXT` / `DATE`, `yyyy-MM-dd`, PK)
- `nome_feriado` (`TEXT`)
- `tipo` (`TEXT`: `Nacional`)

---

## 5. Integração com Banco SQLite e Power BI

- **Chaves de junção no SQLite (`02-bancodedados-sql`):**
  - `cambio_ptax_2025_2026.data` -> tabela calendário (`dCalendario.data`) / `fVendas.data_pedido` (`yyyy-MM-dd`).
  - `feriados_nacionais_2025_2026.data` -> tabela calendário (`dCalendario.data`) para marcação de dia útil/feriado e explicação de dias `Carry-forward` no câmbio.
  - `clima_24_cidades_2025_2026.(cidade, data_ref)` -> junção composta com `dCidade.Cidade` (grafia idêntica) e data do pedido/entrega para correlacionar precipitação (`faixa_precipitacao`) com atrasos logísticos.
- **Importação no SQLite / Power BI:**
  - Como o separador decimal é `.` (ponto), o SQLite realiza conversão numérica nativa (`CAST(... AS REAL)`), e no Power BI/Power Query recomenda-se usar `Culture = "en-US"` ao tipar colunas decimais a partir dos CSVs ou consumir diretamente as consultas M de `scripts_m/`.
