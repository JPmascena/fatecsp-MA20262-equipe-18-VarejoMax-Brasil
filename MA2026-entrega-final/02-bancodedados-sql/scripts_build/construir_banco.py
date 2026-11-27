"""Builder automatizado do banco SQLite 3FN VarejoMax Brasil (Etapa IV)."""

from __future__ import annotations

import csv
import pathlib
import re
import sqlite3
import sys

BASE_DIR = pathlib.Path(__file__).resolve().parent.parent
SQL_DIR = BASE_DIR / "sql"
CSV_DIR = BASE_DIR / "csv"
EVIDENCIAS_DIR = BASE_DIR / "evidencias"
DB_PATH = BASE_DIR / "varejomax.db"

STAGING_FILES: list[tuple[str, str]] = [
    ("stg_cidade", "dCidade.csv"),
    ("stg_canal_loja", "dCanalLoja.csv"),
    ("stg_produto", "dProduto.csv"),
    ("stg_cliente", "dCliente.csv"),
    ("stg_transportadora", "dTransportadora.csv"),
    ("stg_vendas", "fVendas.csv"),
    ("stg_cambio_ptax", "cambio_ptax_2025_2026.csv"),
    ("stg_clima_cidade", "clima_24_cidades_2025_2026.csv"),
    ("stg_feriados", "feriados_nacionais_2025_2026.csv"),
]

CORE_TABLES = [
    "cidade",
    "canal_loja",
    "produto",
    "cliente",
    "transportadora",
    "calendario",
    "pedido",
    "item_pedido",
    "cambio_ptax",
    "clima_cidade",
]


def load_csv_to_staging(conn: sqlite3.Connection, table: str, csv_path: pathlib.Path) -> int:
    """Importa um CSV UTF-8 delimitado por ';' para a tabela stg_* correspondente."""
    with csv_path.open("r", encoding="utf-8-sig", newline="") as fh:
        reader = csv.reader(fh, delimiter=";")
        headers = next(reader)
        rows = [row for row in reader if row]

    placeholders = ",".join("?" for _ in headers)
    sql = f"INSERT INTO {table} ({','.join(headers)}) VALUES ({placeholders})"
    conn.executemany(sql, rows)
    return len(rows)


def parse_named_queries(sql_path: pathlib.Path) -> list[tuple[str, str]]:
    """Extrai blocos de consulta demarcados por '-- === NOME ==='."""
    text = sql_path.read_text(encoding="utf-8")
    pattern = re.compile(r"^--\s*===\s*([A-Za-z0-9_]+)\s*===\s*$", re.MULTILINE)
    matches = list(pattern.finditer(text))
    queries: list[tuple[str, str]] = []
    for idx, m in enumerate(matches):
        name = m.group(1).lower()
        start = m.end()
        end = matches[idx + 1].start() if idx + 1 < len(matches) else len(text)
        body = text[start:end].strip()
        queries.append((name, body))
    return queries


def write_query_csv(
    conn: sqlite3.Connection, query_sql: str, out_path: pathlib.Path
) -> tuple[list[str], list[tuple]]:
    """Executa uma consulta SELECT e grava o resultado em CSV UTF-8 com ';'."""
    cur = conn.execute(query_sql)
    headers = [desc[0] for desc in cur.description]
    rows = cur.fetchall()
    with out_path.open("w", encoding="utf-8", newline="") as fh:
        writer = csv.writer(fh, delimiter=";")
        writer.writerow(headers)
        writer.writerows(rows)
    return headers, rows


def table_counts(conn: sqlite3.Connection) -> dict[str, int]:
    """Retorna a contagem de registros de todas as tabelas 3FN e externas."""
    return {
        tbl: conn.execute(f"SELECT COUNT(*) FROM {tbl}").fetchone()[0]
        for tbl in CORE_TABLES
    }


def main() -> int:
    EVIDENCIAS_DIR.mkdir(parents=True, exist_ok=True)

    # Remove arquivo anterior e eventuais artefatos de journal/wal
    for suffix in ("", "-journal", "-wal", "-shm"):
        p = pathlib.Path(str(DB_PATH) + suffix)
        if p.exists():
            p.unlink()

    conn = sqlite3.connect(DB_PATH)
    try:
        conn.execute("PRAGMA journal_mode = DELETE;")
        conn.execute("PRAGMA foreign_keys = ON;")

        # 1. DDL
        print("[1/7] Executando varejomax_ddl.sql...")
        conn.executescript((SQL_DIR / "varejomax_ddl.sql").read_text(encoding="utf-8"))
        conn.execute("PRAGMA foreign_keys = ON;")

        # 2. Carga dos 9 CSVs em stg_*
        print("[2/7] Carregando 9 CSVs em tabelas stg_*...")
        for table, fname in STAGING_FILES:
            n = load_csv_to_staging(conn, table, CSV_DIR / fname)
            print(f"  - {table}: {n} linhas")
        conn.commit()

        # 3. Executar varejomax_carga.sql 2x para provar idempotencia
        carga_sql = (SQL_DIR / "varejomax_carga.sql").read_text(encoding="utf-8")
        print("[3/7] Executando varejomax_carga.sql (1a vez)...")
        conn.executescript(carga_sql)
        counts_1 = table_counts(conn)

        print("[3/7] Executando varejomax_carga.sql (2a vez - teste de idempotencia)...")
        conn.executescript(carga_sql)
        counts_2 = table_counts(conn)

        if counts_1 != counts_2:
            raise RuntimeError(f"Falha na idempotencia: {counts_1} != {counts_2}")
        print("  -> Idempotencia 100% confirmada! Contagens identicas nas 2 execucoes:")
        for k, v in counts_2.items():
            print(f"     {k}: {v}")

        # 4. Views Star Schema
        print("[4/7] Executando varejomax_views.sql...")
        conn.executescript((SQL_DIR / "varejomax_views.sql").read_text(encoding="utf-8"))
        conn.execute("PRAGMA foreign_keys = ON;")

        # 5. PRAGMA integrity_check & foreign_key_check
        print("[5/7] Validando PRAGMA integrity_check e PRAGMA foreign_key_check...")
        integ = conn.execute("PRAGMA integrity_check;").fetchall()
        fk_check = conn.execute("PRAGMA foreign_key_check;").fetchall()
        print(f"  -> integrity_check: {integ}")
        print(f"  -> foreign_key_check: {len(fk_check)} violacoes")
        if integ != [("ok",)] or fk_check:
            raise RuntimeError(f"Erro de integridade: integ={integ}, fk={fk_check}")

        # 6. Exportar sqlite_master + K1..K5 + D1..D4 + C1..C10
        print("[6/7] Executando e exportando evidencias (sqlite_master, K1-K5, D1-D4, C1-C10)...")
        write_query_csv(
            conn,
            "SELECT type, name, tbl_name FROM sqlite_master WHERE name NOT LIKE 'sqlite_%' ORDER BY type, name",
            EVIDENCIAS_DIR / "sqlite_master.csv",
        )

        for sql_file in ("varejomax_conferencia.sql", "varejomax_diagnostico.sql", "varejomax_consultas.sql"):
            for name, body in parse_named_queries(SQL_DIR / sql_file):
                headers, rows = write_query_csv(conn, body, EVIDENCIAS_DIR / f"{name}.csv")
                print(f"  -> {name}.csv: {len(rows)} linhas ({', '.join(headers[:4])}...)")

                # Validacao automatica de diferenca/orfaos/violacoes = 0 em K1..K4
                if name in ("k1", "k2", "k3", "k4"):
                    for row in rows:
                        val = row[-1]
                        if float(val) != 0.0:
                            raise RuntimeError(f"Falha em {name.upper()}: linha {row} tem diferenca != 0")

        # 7. EXPLAIN QUERY PLAN das consultas C1..C10
        print("[7/7] Gerando EXPLAIN QUERY PLAN para C1..C10...")
        eqp_lines: list[str] = [
            "=============================================================",
            "EXPLAIN QUERY PLAN -- Consultas Analiticas C1 a C10",
            "Banco: varejomax.db (SQLite " + sqlite3.sqlite_version + ")",
            "=============================================================",
            "",
        ]
        for name, body in parse_named_queries(SQL_DIR / "varejomax_consultas.sql"):
            eqp_lines.append(f"--- {name.upper()} ---")
            plan_rows = conn.execute(f"EXPLAIN QUERY PLAN {body}").fetchall()
            for r in plan_rows:
                eqp_lines.append(f"  {r}")
            eqp_lines.append("")
        (EVIDENCIAS_DIR / "explain_query_plan.txt").write_text(
            "\n".join(eqp_lines), encoding="utf-8"
        )

        # Garantir journal_mode=DELETE antes de fechar
        conn.execute("PRAGMA journal_mode = DELETE;")
        conn.commit()
    finally:
        conn.close()

    # Verificar ausencia de arquivos temporarios journal/wal/shm
    leftovers = [
        p.name
        for suffix in ("-journal", "-wal", "-shm")
        if (p := pathlib.Path(str(DB_PATH) + suffix)).exists()
    ]
    if leftovers:
        raise RuntimeError(f"Arquivos temporarios remanescentes: {leftovers}")

    print("CONCLUIDO! Banco varejomax.db construido e validado sem arquivos journal/wal/shm.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
