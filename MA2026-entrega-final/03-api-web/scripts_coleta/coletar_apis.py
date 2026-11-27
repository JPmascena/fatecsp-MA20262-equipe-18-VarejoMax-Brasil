"""Coleta reprodutivel das 3 APIs publicas REST para o projeto VarejoMax Brasil.

Fontes:
1. Banco Central do Brasil (Olinda PTAX) - cambio diario USD/BRL (carry-forward em dias nao uteis)
2. Open-Meteo Archive API - clima historico diario para as 24 cidades de dCidade.csv
3. BrasilAPI Feriados - feriados nacionais de 2025 e 2026 (recorte 2025-01-01 a 2026-08-31)

Padrao dos CSVs gerados:
- Encoding: UTF-8
- Separador de colunas: ';'
- Separador decimal: '.' (4 casas em cambio, 1 casa em clima) para compatibilidade SQLite / Power BI en-US
- Datas: yyyy-MM-dd
"""

from __future__ import annotations

import csv
import datetime as dt
import json
import statistics
import time
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parents[1]
REPO_FINAL_DIR = BASE_DIR.parent
CIDADE_CSV = REPO_FINAL_DIR / "02-bancodedados-sql" / "csv" / "dCidade.csv"

DADOS_DIR = BASE_DIR / "dados_api"
EVIDENCIAS_DIR = BASE_DIR / "evidencias"

DATA_INICIO = dt.date(2025, 1, 1)
DATA_FIM = dt.date(2026, 8, 31)
TOTAL_DIAS = (DATA_FIM - DATA_INICIO).days + 1  # 608 dias

USER_AGENT = "VarejoMax-Brasil-Coleta/1.0 (Fatec-SP-MA2026)"


def fetch_json(url: str, max_retries: int = 5, base_wait: float = 1.5) -> tuple[object, bytes]:
    """Executa GET HTTP e retorna (objeto_json, bytes_brutos) com backoff exponencial."""
    last_err: Exception | None = None
    for attempt in range(1, max_retries + 1):
        req = urllib.request.Request(
            url,
            headers={
                "User-Agent": USER_AGENT,
                "Accept": "application/json",
            },
        )
        try:
            with urllib.request.urlopen(req, timeout=45) as resp:
                raw = resp.read()
                return json.loads(raw.decode("utf-8")), raw
        except urllib.error.HTTPError as exc:
            last_err = exc
            if exc.code in (429, 500, 502, 503, 504) and attempt < max_retries:
                wait = base_wait * (2 ** (attempt - 1))
                print(f"[WARN] HTTP {exc.code} em {url} (tentativa {attempt}/{max_retries}); aguardando {wait:.1f}s...")
                time.sleep(wait)
                continue
            raise
        except (urllib.error.URLError, TimeoutError) as exc:
            last_err = exc
            if attempt < max_retries:
                wait = base_wait * (2 ** (attempt - 1))
                print(f"[WARN] Erro de rede em {url}: {exc} (tentativa {attempt}/{max_retries}); aguardando {wait:.1f}s...")
                time.sleep(wait)
                continue
            raise
    assert last_err is not None
    raise last_err


def load_cidades(csv_path: Path) -> list[dict[str, str]]:
    """Le dCidade.csv preservando exatamente a grafia da coluna Cidade."""
    cidades: list[dict[str, str]] = []
    with csv_path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f, delimiter=";")
        for row in reader:
            cidades.append(
                {
                    "cidade": row["Cidade"],
                    "uf": row["UF"],
                    "regiao": row["Regiao"],
                    "lat": row["Latitude"].replace(",", ".").strip(),
                    "lon": row["Longitude"].replace(",", ".").strip(),
                }
            )
    if len(cidades) != 24:
        raise ValueError(f"Esperadas 24 cidades em {csv_path}, encontradas {len(cidades)}")
    return cidades


def classify_faixa_dolar(cotacao_venda: float) -> str:
    """Classifica cotacao de venda do dolar em faixas analiticas coerentes com 2025-2026."""
    if cotacao_venda < 5.10:
        return "1-Abaixo de R$ 5,10"
    if cotacao_venda < 5.40:
        return "2-R$ 5,10 a R$ 5,39"
    if cotacao_venda < 5.70:
        return "3-R$ 5,40 a R$ 5,69"
    if cotacao_venda < 6.00:
        return "4-R$ 5,70 a R$ 5,99"
    return "5-R$ 6,00 ou mais"


def classify_faixa_precipitacao(precip_mm: float) -> str:
    """Classifica precipitacao diaria (ja arredondada para 1 casa decimal)."""
    if precip_mm == 0.0:
        return "0-Sem chuva (0 mm)"
    if precip_mm <= 10.0:
        return "1-Chuva fraca (0,1-10 mm)"
    if precip_mm <= 30.0:
        return "2-Chuva moderada (10-30 mm)"
    return "3-Chuva forte (>30 mm)"


def parse_bcb_timestamp(ts_str: str) -> dt.datetime:
    """Converte dataHoraCotacao do BCB (com ou sem fracao de segundos)."""
    for fmt in ("%Y-%m-%d %H:%M:%S.%f", "%Y-%m-%d %H:%M:%S"):
        try:
            return dt.datetime.strptime(ts_str, fmt)
        except ValueError:
            continue
    return dt.datetime.fromisoformat(ts_str)


def collect_ptax() -> dict[str, object]:
    """Coleta PTAX diaria do BCB (12-20-2024 a 08-31-2026) e gera serie continua de 608 dias."""
    url = (
        "https://olinda.bcb.gov.br/olinda/servico/PTAX/versao/v1/odata/"
        "CotacaoDolarPeriodo(dataInicial=@dataInicial,dataFinalCotacao=@dataFinalCotacao)"
        "?@dataInicial='12-20-2024'&@dataFinalCotacao='08-31-2026'"
        "&$format=json&$select=cotacaoCompra,cotacaoVenda,dataHoraCotacao"
    )
    print("[1/3] Coletando PTAX BCB...")
    payload, raw_bytes = fetch_json(url)
    assert isinstance(payload, dict)

    (EVIDENCIAS_DIR / "sample_ptax.json").write_bytes(raw_bytes)

    latest_by_date: dict[dt.date, tuple[dt.datetime, float, float]] = {}
    for item in payload.get("value", []):
        ts = parse_bcb_timestamp(item["dataHoraCotacao"])
        d = ts.date()
        compra = float(item["cotacaoCompra"])
        venda = float(item["cotacaoVenda"])
        prev = latest_by_date.get(d)
        if prev is None or ts > prev[0]:
            latest_by_date[d] = (ts, compra, venda)

    # Garante cotacao util antes de 2025-01-01 para carry-forward inicial
    seed_dates = [d for d in latest_by_date if d < DATA_INICIO]
    if not seed_dates:
        raise RuntimeError("Nenhuma cotacao PTAX encontrada antes de 2025-01-01 para carry-forward.")
    last_date = max(seed_dates)
    _, last_compra, last_venda = latest_by_date[last_date]

    rows: list[list[str]] = []
    compras: list[float] = []
    vendas: list[float] = []
    oficiais = 0
    carry_forwards = 0

    for offset in range(TOTAL_DIAS):
        cur_date = DATA_INICIO + dt.timedelta(days=offset)
        if cur_date in latest_by_date:
            _, compra, venda = latest_by_date[cur_date]
            last_compra, last_venda = compra, venda
            origem = "Oficial BCB"
            oficiais += 1
        else:
            compra, venda = last_compra, last_venda
            origem = "Carry-forward"
            carry_forwards += 1

        compras.append(compra)
        vendas.append(venda)
        rows.append(
            [
                cur_date.isoformat(),
                f"{compra:.4f}",
                f"{venda:.4f}",
                origem,
                classify_faixa_dolar(venda),
            ]
        )

    out_csv = DADOS_DIR / "cambio_ptax_2025_2026.csv"
    with out_csv.open("w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f, delimiter=";", lineterminator="\n")
        writer.writerow(["data", "cotacao_compra", "cotacao_venda", "boletim_origem", "faixa_dolar"])
        writer.writerows(rows)

    return {
        "arquivo": str(out_csv),
        "linhas": len(rows),
        "oficiais": oficiais,
        "carry_forward": carry_forwards,
        "compra_min": min(compras),
        "compra_max": max(compras),
        "compra_media": statistics.mean(compras),
        "venda_min": min(vendas),
        "venda_max": max(vendas),
        "venda_media": statistics.mean(vendas),
        "total_boletins_api": len(payload.get("value", [])),
    }


def collect_openmeteo(cidades: list[dict[str, str]]) -> dict[str, object]:
    """Coleta clima historico diario (Open-Meteo Archive) para todas as 24 cidades."""
    print("[2/3] Coletando Clima Open-Meteo para 24 cidades...")
    rows: list[list[str]] = []
    null_replacements = 0
    all_tmax: list[float] = []
    all_tmin: list[float] = []
    all_prec: list[float] = []
    faixa_counts: dict[str, int] = {}

    expected_dates = [(DATA_INICIO + dt.timedelta(days=i)).isoformat() for i in range(TOTAL_DIAS)]

    for idx, c in enumerate(cidades, start=1):
        params = urllib.parse.urlencode(
            {
                "latitude": c["lat"],
                "longitude": c["lon"],
                "start_date": DATA_INICIO.isoformat(),
                "end_date": DATA_FIM.isoformat(),
                "daily": "temperature_2m_max,temperature_2m_min,precipitation_sum",
                "timezone": "America/Sao_Paulo",
            }
        )
        url = f"https://archive-api.open-meteo.com/v1/archive?{params}"
        payload, raw_bytes = fetch_json(url)
        assert isinstance(payload, dict)

        if c["cidade"] == "São Paulo":
            (EVIDENCIAS_DIR / "sample_openmeteo_sp.json").write_bytes(raw_bytes)

        daily = payload["daily"]
        times = daily["time"]
        tmax_list = daily["temperature_2m_max"]
        tmin_list = daily["temperature_2m_min"]
        prec_list = daily["precipitation_sum"]

        if times != expected_dates:
            raise RuntimeError(f"Serie de datas divergente para {c['cidade']}: {len(times)} dias")

        prev_tmax: float | None = None
        prev_tmin: float | None = None

        for d_str, raw_tmax, raw_tmin, raw_prec in zip(times, tmax_list, tmin_list, prec_list):
            if raw_tmax is None:
                null_replacements += 1
                tmax_val = prev_tmax if prev_tmax is not None else 25.0
            else:
                tmax_val = round(float(raw_tmax), 1)
                prev_tmax = tmax_val

            if raw_tmin is None:
                null_replacements += 1
                tmin_val = prev_tmin if prev_tmin is not None else 18.0
            else:
                tmin_val = round(float(raw_tmin), 1)
                prev_tmin = tmin_val

            if raw_prec is None:
                null_replacements += 1
                prec_val = 0.0
            else:
                prec_val = round(float(raw_prec), 1)

            # Evita '-0.0'
            if tmax_val == 0.0:
                tmax_val = 0.0
            if tmin_val == 0.0:
                tmin_val = 0.0
            if prec_val == 0.0:
                prec_val = 0.0

            faixa = classify_faixa_precipitacao(prec_val)
            faixa_counts[faixa] = faixa_counts.get(faixa, 0) + 1

            all_tmax.append(tmax_val)
            all_tmin.append(tmin_val)
            all_prec.append(prec_val)

            rows.append(
                [
                    c["cidade"],
                    d_str,
                    f"{tmax_val:.1f}",
                    f"{tmin_val:.1f}",
                    f"{prec_val:.1f}",
                    faixa,
                ]
            )

        print(f"  -> ({idx:02d}/24) {c['cidade']}: {len(times)} dias ok")
        if idx < len(cidades):
            time.sleep(0.3)

    out_csv = DADOS_DIR / "clima_24_cidades_2025_2026.csv"
    with out_csv.open("w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f, delimiter=";", lineterminator="\n")
        writer.writerow(["cidade", "data_ref", "temp_max", "temp_min", "precipitacao", "faixa_precipitacao"])
        writer.writerows(rows)

    return {
        "arquivo": str(out_csv),
        "linhas": len(rows),
        "cidades": len(cidades),
        "nulls_tratados": null_replacements,
        "tmax_min": min(all_tmax),
        "tmax_max": max(all_tmax),
        "tmax_media": statistics.mean(all_tmax),
        "tmin_min": min(all_tmin),
        "tmin_max": max(all_tmin),
        "tmin_media": statistics.mean(all_tmin),
        "prec_min": min(all_prec),
        "prec_max": max(all_prec),
        "prec_media": statistics.mean(all_prec),
        "faixa_counts": faixa_counts,
    }


def collect_feriados() -> dict[str, object]:
    """Coleta feriados nacionais de 2025 e 2026 na BrasilAPI e filtra o periodo."""
    print("[3/3] Coletando Feriados Nacionais BrasilAPI...")
    tipo_map = {"national": "Nacional"}
    registros: list[tuple[str, str, str]] = []
    counts_by_year: dict[int, int] = {}

    for year in (2025, 2026):
        url = f"https://brasilapi.com.br/api/feriados/v1/{year}"
        payload, raw_bytes = fetch_json(url)
        assert isinstance(payload, list)
        if year == 2025:
            (EVIDENCIAS_DIR / "sample_brasilapi_2025.json").write_bytes(raw_bytes)

        in_range = 0
        for item in payload:
            d_str = item["date"]
            d_obj = dt.date.fromisoformat(d_str)
            if DATA_INICIO <= d_obj <= DATA_FIM:
                raw_type = str(item.get("type", "national"))
                tipo_label = tipo_map.get(raw_type, raw_type)
                registros.append((d_str, str(item["name"]), tipo_label))
                in_range += 1
        counts_by_year[year] = in_range

    registros.sort(key=lambda r: r[0])

    out_csv = DADOS_DIR / "feriados_nacionais_2025_2026.csv"
    with out_csv.open("w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f, delimiter=";", lineterminator="\n")
        writer.writerow(["data", "nome_feriado", "tipo"])
        writer.writerows(registros)

    return {
        "arquivo": str(out_csv),
        "linhas": len(registros),
        "por_ano": counts_by_year,
    }


def write_summary_md(ptax_stats: dict[str, object], clima_stats: dict[str, object], fer_stats: dict[str, object]) -> None:
    """Gera evidencias/resumo_coleta_apis.md com os numeros reais coletados."""
    faixas_clima = clima_stats["faixa_counts"]
    assert isinstance(faixas_clima, dict)
    por_ano = fer_stats["por_ano"]
    assert isinstance(por_ano, dict)

    md = f"""# Resumo da Coleta Real de APIs REST — VarejoMax Brasil

- **Data de execução:** {dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
- **Período analisado:** `2025-01-01` a `2026-08-31` (**{TOTAL_DIAS} dias corridos**)
- **Padrão dos CSVs:** UTF-8, delimitador `;`, separador decimal `.` (ponto — padrão `en-US` para compatibilidade direta com SQLite e Power BI), datas `yyyy-MM-dd`.

---

## 1. Câmbio Dólar PTAX (Banco Central do Brasil — Olinda)

- **Arquivo gerado:** `dados_api/cambio_ptax_2025_2026.csv`
- **Total de linhas:** **{ptax_stats['linhas']}**
- **Registros brutos retornados na janela (`12-20-2024` a `08-31-2026`):** {ptax_stats['total_boletins_api']} boletins
- **Dias no período (`2025-01-01` a `2026-08-31`):**
  - `Oficial BCB`: **{ptax_stats['oficiais']}** dias úteis
  - `Carry-forward` (fins de semana/feriados): **{ptax_stats['carry_forward']}** dias
- **Estatísticas reais (`cotacao_compra`):**
  - Mínimo: **R$ {ptax_stats['compra_min']:.4f}**
  - Máximo: **R$ {ptax_stats['compra_max']:.4f}**
  - Média: **R$ {ptax_stats['compra_media']:.4f}**
- **Estatísticas reais (`cotacao_venda`):**
  - Mínimo: **R$ {ptax_stats['venda_min']:.4f}**
  - Máximo: **R$ {ptax_stats['venda_max']:.4f}**
  - Média: **R$ {ptax_stats['venda_media']:.4f}**

---

## 2. Clima Histórico Diário (Open-Meteo Archive API)

- **Arquivo gerado:** `dados_api/clima_24_cidades_2025_2026.csv`
- **Cidades atendidas:** **{clima_stats['cidades']}** (grafia 100% preservada de `dCidade.csv`)
- **Total de linhas:** **{clima_stats['linhas']}** (`24 × 608`)
- **Valores nulos na API tratados via fallback:** **{clima_stats['nulls_tratados']}**
- **Estatísticas reais no conjunto (24 cidades × 608 dias):**
  - `temp_max` (°C): mín **{clima_stats['tmax_min']:.1f}** | máx **{clima_stats['tmax_max']:.1f}** | média **{clima_stats['tmax_media']:.1f}**
  - `temp_min` (°C): mín **{clima_stats['tmin_min']:.1f}** | máx **{clima_stats['tmin_max']:.1f}** | média **{clima_stats['tmin_media']:.1f}**
  - `precipitacao` (mm): mín **{clima_stats['prec_min']:.1f}** | máx **{clima_stats['prec_max']:.1f}** | média **{clima_stats['prec_media']:.1f}**
- **Distribuição por `faixa_precipitacao`:**
"""
    for faixa in sorted(faixas_clima.keys()):
        md += f"  - `{faixa}`: **{faixas_clima[faixa]}** registros\n"

    md += f"""
---

## 3. Feriados Nacionais (BrasilAPI)

- **Arquivo gerado:** `dados_api/feriados_nacionais_2025_2026.csv`
- **Total de feriados no recorte (`2025-01-01` a `2026-08-31`):** **{fer_stats['linhas']}**
  - **2025:** {por_ano.get(2025, 0)} feriados
  - **2026 (jan–ago):** {por_ano.get(2026, 0)} feriados
"""
    (EVIDENCIAS_DIR / "resumo_coleta_apis.md").write_text(md, encoding="utf-8")


def main() -> None:
    DADOS_DIR.mkdir(parents=True, exist_ok=True)
    EVIDENCIAS_DIR.mkdir(parents=True, exist_ok=True)

    cidades = load_cidades(CIDADE_CSV)
    ptax_stats = collect_ptax()
    clima_stats = collect_openmeteo(cidades)
    fer_stats = collect_feriados()
    write_summary_md(ptax_stats, clima_stats, fer_stats)

    print("\n=== COLETA CONCLUIDA COM SUCESSO ===")
    print("PTAX:", ptax_stats)
    print("Clima:", {k: v for k, v in clima_stats.items() if k != "faixa_counts"}, clima_stats["faixa_counts"])
    print("Feriados:", fer_stats)


if __name__ == "__main__":
    main()
