# Resumo da Coleta Real de APIs REST — VarejoMax Brasil

- **Data de execução:** 2026-11-25 14:15:46
- **Período analisado:** `2025-01-01` a `2026-08-31` (**608 dias corridos**)
- **Padrão dos CSVs:** UTF-8, delimitador `;`, separador decimal `.` (ponto — padrão `en-US` para compatibilidade direta com SQLite e Power BI), datas `yyyy-MM-dd`.

---

## 1. Câmbio Dólar PTAX (Banco Central do Brasil — Olinda)

- **Arquivo gerado:** `dados_api/cambio_ptax_2025_2026.csv`
- **Total de linhas:** **608**
- **Registros brutos retornados na janela (`12-20-2024` a `08-31-2026`):** 426 boletins
- **Dias no período (`2025-01-01` a `2026-08-31`):**
  - `Oficial BCB`: **418** dias úteis
  - `Carry-forward` (fins de semana/feriados): **190** dias
- **Estatísticas reais (`cotacao_compra`):**
  - Mínimo: **R$ 4.8967**
  - Máximo: **R$ 6.2080**
  - Média: **R$ 5.4145**
- **Estatísticas reais (`cotacao_venda`):**
  - Mínimo: **R$ 4.8973**
  - Máximo: **R$ 6.2086**
  - Média: **R$ 5.4151**

---

## 2. Clima Histórico Diário (Open-Meteo Archive API)

- **Arquivo gerado:** `dados_api/clima_24_cidades_2025_2026.csv`
- **Cidades atendidas:** **24** (grafia 100% preservada de `dCidade.csv`)
- **Total de linhas:** **14592** (`24 × 608`)
- **Valores nulos na API tratados via fallback:** **0**
- **Estatísticas reais no conjunto (24 cidades × 608 dias):**
  - `temp_max` (°C): mín **8.3** | máx **40.4** | média **28.3**
  - `temp_min` (°C): mín **0.5** | máx **28.7** | média **20.7**
  - `precipitacao` (mm): mín **0.0** | máx **105.4** | média **4.1**
- **Distribuição por `faixa_precipitacao`:**
  - `0-Sem chuva (0 mm)`: **3669** registros
  - `1-Chuva fraca (0,1-10 mm)`: **9017** registros
  - `2-Chuva moderada (10-30 mm)`: **1679** registros
  - `3-Chuva forte (>30 mm)`: **227** registros

---

## 3. Feriados Nacionais (BrasilAPI)

- **Arquivo gerado:** `dados_api/feriados_nacionais_2025_2026.csv`
- **Total de feriados no recorte (`2025-01-01` a `2026-08-31`):** **22**
  - **2025:** 14 feriados
  - **2026 (jan–ago):** 8 feriados
