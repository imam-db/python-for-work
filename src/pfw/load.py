"""Load a tabular file into DuckDB (default) or Postgres."""

from __future__ import annotations

import argparse
import os
import re
from pathlib import Path

import pandas as pd

from pfw.common.io import read_table
from pfw.common.logging_setup import get_logger, setup_logging

log = get_logger("load")

DEFAULT_DUCKDB_PATH = Path("output") / "pfw.duckdb"
DEFAULT_POSTGRES_ENV = "PFW_DATABASE_URL"
MODES = ("replace", "append", "upsert")
TABLE_NAME_RE = re.compile(r"^[A-Za-z_][A-Za-z0-9_]*$")


def _quote_ident(name: str) -> str:
    return '"' + name.replace('"', '""') + '"'


def _validate_table_name(table: str) -> str:
    if not TABLE_NAME_RE.match(table):
        raise ValueError(
            f"Invalid table name '{table}'. Use letters, digits, and underscores; "
            "must start with a letter or underscore."
        )
    return table


def _validate_keys(df: pd.DataFrame, keys: list[str]) -> list[str]:
    if not keys:
        raise ValueError("upsert mode requires --keys (one or more key columns)")
    missing = [column for column in keys if column not in df.columns]
    if missing:
        raise ValueError(
            f"Key column(s) not found in the input: {missing}. "
            f"Available columns: {', '.join(map(str, df.columns))}"
        )
    return keys


def _parse_keys(value: str | None) -> list[str]:
    if not value:
        return []
    return [part.strip() for part in value.split(",") if part.strip()]


def _duckdb_connect(db_path: str | Path):
    try:
        import duckdb
    except ImportError as exc:  # pragma: no cover - installed via requirements
        raise ImportError(
            "duckdb is required for the DuckDB backend. Install with: pip install duckdb"
        ) from exc

    path = Path(db_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    return duckdb.connect(str(path))


def _postgres_connect(database_url: str):
    try:
        import psycopg
    except ImportError as exc:
        raise ImportError(
            "psycopg is required for the Postgres backend. "
            "Install with: pip install 'python-for-work[postgres]'"
        ) from exc
    return psycopg.connect(database_url)


def _resolve_postgres_url(env_var: str) -> str:
    url = os.environ.get(env_var, "").strip()
    if not url:
        raise ValueError(
            f"Postgres URL not set. Export {env_var} with a connection string "
            "(never pass the password on the CLI)."
        )
    return url


def _exec_duckdb(conn, query: str, params: list | None = None):
    """Run a DuckDB statement and fully consume the result set."""
    result = conn.execute(query, params) if params is not None else conn.execute(query)
    try:
        rows = result.fetchall()
    except Exception:
        rows = []
    return rows


def _table_exists_duckdb(conn, table: str) -> bool:
    rows = _exec_duckdb(
        conn,
        """
        SELECT 1
        FROM information_schema.tables
        WHERE table_schema = 'main' AND table_name = ?
        LIMIT 1
        """,
        [table],
    )
    return bool(rows)


def _table_exists_postgres(conn, table: str) -> bool:
    with conn.cursor() as cur:
        cur.execute(
            """
            SELECT 1
            FROM information_schema.tables
            WHERE table_schema = 'public' AND table_name = %s
            LIMIT 1
            """,
            (table,),
        )
        return cur.fetchone() is not None


def _create_table_duckdb(conn, table: str, df: pd.DataFrame, keys: list[str]) -> None:
    conn.register("_pfw_load_src", df)
    quoted = _quote_ident(table)
    if keys:
        # DuckDB primary key requires NOT NULL columns; create then add constraint.
        _exec_duckdb(
            conn, f"CREATE TABLE {quoted} AS SELECT * FROM _pfw_load_src WHERE 1=0"
        )
        key_list = ", ".join(_quote_ident(column) for column in keys)
        _exec_duckdb(conn, f"ALTER TABLE {quoted} ADD PRIMARY KEY ({key_list})")
    else:
        _exec_duckdb(
            conn, f"CREATE TABLE {quoted} AS SELECT * FROM _pfw_load_src WHERE 1=0"
        )
    conn.unregister("_pfw_load_src")


def _create_table_postgres(conn, table: str, df: pd.DataFrame, keys: list[str]) -> None:
    from psycopg import sql

    type_map = {
        "int64": "BIGINT",
        "Int64": "BIGINT",
        "float64": "DOUBLE PRECISION",
        "Float64": "DOUBLE PRECISION",
        "bool": "BOOLEAN",
        "boolean": "BOOLEAN",
        "datetime64[ns]": "TIMESTAMP",
    }
    columns = []
    for name, dtype in df.dtypes.items():
        pg_type = type_map.get(str(dtype), "TEXT")
        columns.append(
            sql.SQL("{} {}").format(sql.Identifier(str(name)), sql.SQL(pg_type))
        )
    if keys:
        columns.append(
            sql.SQL("PRIMARY KEY ({})").format(
                sql.SQL(", ").join(sql.Identifier(column) for column in keys)
            )
        )
    query = sql.SQL("CREATE TABLE {} ({})").format(
        sql.Identifier(table),
        sql.SQL(", ").join(columns),
    )
    with conn.cursor() as cur:
        cur.execute(query)
    conn.commit()


def _insert_duckdb(conn, table: str, df: pd.DataFrame) -> int:
    conn.register("_pfw_load_src", df)
    _exec_duckdb(
        conn,
        f"INSERT INTO {_quote_ident(table)} BY NAME SELECT * FROM _pfw_load_src",
    )
    conn.unregister("_pfw_load_src")
    return len(df)


def _insert_postgres(conn, table: str, df: pd.DataFrame) -> int:
    from psycopg import sql
    from psycopg.rows import tuple_row

    if df.empty:
        return 0

    columns = [str(column) for column in df.columns]
    placeholders = sql.SQL(", ").join(sql.Placeholder() * len(columns))
    query = sql.SQL("INSERT INTO {} ({}) VALUES ({})").format(
        sql.Identifier(table),
        sql.SQL(", ").join(sql.Identifier(column) for column in columns),
        placeholders,
    )
    rows = [
        tuple(None if pd.isna(value) else value for value in row)
        for row in df.itertuples(index=False, name=None)
    ]
    with conn.cursor(row_factory=tuple_row) as cur:
        cur.executemany(query, rows)
    conn.commit()
    return len(rows)


def _upsert_duckdb(conn, table: str, df: pd.DataFrame, keys: list[str]) -> int:
    conn.register("_pfw_load_src", df)
    quoted = _quote_ident(table)
    key_list = ", ".join(_quote_ident(column) for column in keys)
    non_keys = [column for column in df.columns if column not in keys]
    if non_keys:
        assignments = ", ".join(
            f"{_quote_ident(column)} = excluded.{_quote_ident(column)}"
            for column in non_keys
        )
        conflict = f"ON CONFLICT ({key_list}) DO UPDATE SET {assignments}"
    else:
        conflict = f"ON CONFLICT ({key_list}) DO NOTHING"
    _exec_duckdb(
        conn,
        f"INSERT INTO {quoted} BY NAME SELECT * FROM _pfw_load_src {conflict}",
    )
    conn.unregister("_pfw_load_src")
    return len(df)


def _upsert_postgres(conn, table: str, df: pd.DataFrame, keys: list[str]) -> int:
    from psycopg import sql

    if df.empty:
        return 0

    columns = [str(column) for column in df.columns]
    non_keys = [column for column in columns if column not in keys]
    insert = sql.SQL("INSERT INTO {} ({}) VALUES ({})").format(
        sql.Identifier(table),
        sql.SQL(", ").join(sql.Identifier(column) for column in columns),
        sql.SQL(", ").join(sql.Placeholder() * len(columns)),
    )
    if non_keys:
        conflict = sql.SQL(" ON CONFLICT ({}) DO UPDATE SET {}").format(
            sql.SQL(", ").join(sql.Identifier(column) for column in keys),
            sql.SQL(", ").join(
                sql.SQL("{} = EXCLUDED.{}").format(
                    sql.Identifier(column), sql.Identifier(column)
                )
                for column in non_keys
            ),
        )
    else:
        conflict = sql.SQL(" ON CONFLICT ({}) DO NOTHING").format(
            sql.SQL(", ").join(sql.Identifier(column) for column in keys)
        )
    query = insert + conflict
    rows = [
        tuple(None if pd.isna(value) else value for value in row)
        for row in df.itertuples(index=False, name=None)
    ]
    with conn.cursor() as cur:
        cur.executemany(query, rows)
    conn.commit()
    return len(rows)


def load_table(
    input_file: str | Path,
    table: str,
    *,
    backend: str = "duckdb",
    db_path: str | Path = DEFAULT_DUCKDB_PATH,
    mode: str = "replace",
    keys: list[str] | None = None,
    postgres_env: str = DEFAULT_POSTGRES_ENV,
) -> dict:
    """Load a tabular file into DuckDB or Postgres.

    Modes:
    - replace: drop and recreate the table, then insert
    - append: create the table if missing, then insert
    - upsert: create the table if missing (with primary key), then upsert on keys
    """
    table = _validate_table_name(table)
    mode = mode.lower()
    if mode not in MODES:
        raise ValueError(f"Unknown mode '{mode}'. Choose from: {', '.join(MODES)}")

    backend = backend.lower()
    if backend not in {"duckdb", "postgres"}:
        raise ValueError("backend must be 'duckdb' or 'postgres'")

    input_path = Path(input_file)
    if not input_path.exists():
        raise FileNotFoundError(f"File '{input_file}' not found")

    key_list = list(keys or [])
    log.info("Reading file: %s", input_path)
    df = read_table(input_path)
    log.info("Total rows: %s", len(df))
    log.info("Target: %s.%s (mode=%s)", backend, table, mode)

    if mode == "upsert":
        key_list = _validate_keys(df, key_list)

    if backend == "duckdb":
        return _load_duckdb(df, table, db_path=db_path, mode=mode, keys=key_list)

    database_url = _resolve_postgres_url(postgres_env)
    return _load_postgres(df, table, database_url=database_url, mode=mode, keys=key_list)


def _load_duckdb(
    df: pd.DataFrame,
    table: str,
    *,
    db_path: str | Path,
    mode: str,
    keys: list[str],
) -> dict:
    conn = _duckdb_connect(db_path)
    try:
        exists = _table_exists_duckdb(conn, table)
        quoted = _quote_ident(table)

        if mode == "replace" and exists:
            _exec_duckdb(conn, f"DROP TABLE {quoted}")
            exists = False

        if not exists:
            _create_table_duckdb(conn, table, df, keys if mode == "upsert" else [])
            log.info("Created table %s in %s", table, db_path)

        if mode == "upsert":
            written = _upsert_duckdb(conn, table, df, keys)
        else:
            written = _insert_duckdb(conn, table, df)

        count = _exec_duckdb(conn, f"SELECT COUNT(*) FROM {quoted}")[0][0]
        log.info("Done! Wrote %s row(s); table now has %s row(s) in %s", written, count, db_path)
        return {
            "backend": "duckdb",
            "table": table,
            "written": written,
            "total": int(count),
            "db_path": str(db_path),
        }
    finally:
        conn.close()


def _load_postgres(
    df: pd.DataFrame,
    table: str,
    *,
    database_url: str,
    mode: str,
    keys: list[str],
) -> dict:
    conn = _postgres_connect(database_url)
    try:
        exists = _table_exists_postgres(conn, table)

        if mode == "replace" and exists:
            with conn.cursor() as cur:
                from psycopg import sql

                cur.execute(sql.SQL("DROP TABLE {}").format(sql.Identifier(table)))
            conn.commit()
            exists = False

        if not exists:
            _create_table_postgres(conn, table, df, keys if mode == "upsert" else [])
            log.info("Created table %s in Postgres", table)

        if mode == "upsert":
            written = _upsert_postgres(conn, table, df, keys)
        else:
            written = _insert_postgres(conn, table, df)

        with conn.cursor() as cur:
            from psycopg import sql

            cur.execute(sql.SQL("SELECT COUNT(*) FROM {}").format(sql.Identifier(table)))
            count = cur.fetchone()[0]
        log.info("Done! Wrote %s row(s); table now has %s row(s) in Postgres", written, count)
        return {
            "backend": "postgres",
            "table": table,
            "written": written,
            "total": int(count),
        }
    finally:
        conn.close()


def add_arguments(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("input", help="Input tabular file (.xlsx, .csv, .parquet)")
    parser.add_argument(
        "--table",
        "-t",
        required=True,
        help="Destination table name",
    )
    parser.add_argument(
        "--backend",
        choices=["duckdb", "postgres"],
        default="duckdb",
        help="Database backend (default: duckdb)",
    )
    parser.add_argument(
        "--db",
        default=str(DEFAULT_DUCKDB_PATH),
        help=f"DuckDB file path (default: {DEFAULT_DUCKDB_PATH})",
    )
    parser.add_argument(
        "--mode",
        choices=list(MODES),
        default="replace",
        help="Load mode: replace, append, or upsert (default: replace)",
    )
    parser.add_argument(
        "--keys",
        help="Comma-separated key columns for upsert mode",
    )
    parser.add_argument(
        "--postgres-url-env",
        default=DEFAULT_POSTGRES_ENV,
        help=f"Env var holding the Postgres URL (default: {DEFAULT_POSTGRES_ENV})",
    )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="pfw load",
        description="Load a tabular file into DuckDB or Postgres.",
    )
    add_arguments(parser)
    return parser


def run(args: argparse.Namespace) -> int:
    setup_logging()
    try:
        load_table(
            args.input,
            args.table,
            backend=args.backend,
            db_path=args.db,
            mode=args.mode,
            keys=_parse_keys(args.keys),
            postgres_env=args.postgres_url_env,
        )
    except (FileNotFoundError, ValueError, ImportError) as exc:
        log.error("Error: %s", exc)
        return 1
    return 0


def main(argv: list[str] | None = None) -> int:
    return run(build_parser().parse_args(argv))
