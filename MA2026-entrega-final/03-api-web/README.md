# Fase 2 — Consumo de APIs Públicas REST (`03-api-web/`)

Camada de enriquecimento externo do projeto **VarejoMax Brasil** (Fatec SP — Microinformática Aplicada, Prof. Veríssimo), conectando dados reais de **câmbio (PTAX BCB)**, **clima histórico diário (Open-Meteo)** e **feriados nacionais (BrasilAPI)** ao modelo analítico de vendas e logística (`2025-01-01` a `2026-08-31`, **608 dias**).

---

## Estrutura da Pasta

```text
03-api-web/
├── README.md
├── VarejoMax_05_APIs_Web.xlsx               # Planilha consolidada (4 abas, OpenXML validado)
├── dados_api/
│   ├── cambio_ptax_2025_2026.csv            # 608 linhas (série contínua c/ carry-forward)
│   ├── clima_24_cidades_2025_2026.csv       # 14.592 linhas (24 cidades × 608 dias)
│   └── feriados_nacionais_2025_2026.csv     # 22 feriados nacionais (jan/2025–ago/2026)
├── docs/
│   └── especificacao_apis.md                # Endpoints, parâmetros, regras e integração SQLite
├── evidencias/
│   ├── resumo_coleta_apis.md                # Estatísticas reais da execução
│   ├── sample_ptax.json                     # Payload bruto retornado pela API PTAX BCB
│   ├── sample_openmeteo_sp.json             # Payload bruto retornado para São Paulo (Open-Meteo)
│   └── sample_brasilapi_2025.json           # Payload bruto retornado pela BrasilAPI (2025)
├── scripts_coleta/
│   ├── coletar_apis.py                      # Coleta real das 3 APIs (apenas Python stdlib)
│   └── gerar_excel_apis.py                  # Geração reprodutível do .xlsx via OpenXML stdlib
└── scripts_m/
    ├── fnPTAX_BCB.pq                        # Power Query M com RelativePath/Query + Table.FillDown
    ├── fnClima_OpenMeteo.pq                 # Função M (cidade, lat, lon) => table
    └── fnFeriados_BrasilAPI.pq              # Power Query M combinando 2025 e 2026
```

---

## Resumo dos Dados Reais Coletados

| Fonte | Endpoint / Serviço | Volume | Destaques Reais (`2025-01-01` a `2026-08-31`) |
|---|---|---|---|
| **BCB Olinda PTAX** | `CotacaoDolarPeriodo` | **608** dias | `418` dias `Oficial BCB` + `190` `Carry-forward`; Venda mín **R$ 4.8973**, máx **R$ 6.2086**, média **R$ 5.4151** |
| **Open-Meteo Archive** | `/v1/archive` (`America/Sao_Paulo`) | **14.592** reg. | **24 cidades** de `dCidade.csv` × 608 dias; `0` nulos; Temp máx `8.3`–`40.4 °C`; Precip máx `105.4 mm` |
| **BrasilAPI Feriados** | `/api/feriados/v1/{ano}` | **22** feriados | `14` feriados em 2025 + `8` feriados em 2026 (jan–ago) |

---

## Padrão dos Arquivos CSV

- **Encoding:** `UTF-8`
- **Delimitador:** `;`
- **Separador decimal:** **`.` (ponto)** — 4 casas decimais em câmbio (`0.0000`) e 1 casa decimal em clima (`0.0`), garantindo compatibilidade direta com SQLite e Power BI (`"en-US"`).
- **Cidades:** grafia preservada **100% idêntica** à coluna `Cidade` de [`02-bancodedados-sql/csv/dCidade.csv`](../02-bancodedados-sql/csv/dCidade.csv).

---

## Como Reproduzir (Python stdlib)

Sem dependências externas (utiliza apenas `urllib`, `json`, `csv`, `datetime`, `time`, `zipfile`, `xml.etree`):

```powershell
python .\scripts_coleta\coletar_apis.py
python .\scripts_coleta\gerar_excel_apis.py
```

Para uso direto no **Power BI Desktop / Excel Power Query**, utilize os códigos M em `scripts_m/` (todos parametrizados com `Web.Contents(..., [RelativePath=..., Query=...])`).
