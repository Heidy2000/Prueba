"""
Conexión a la base de datos.

- Sin URL: archivo SQLite local (presupuesto.db o la ruta de PANEL_DB).
- Con URL postgresql://... (variable DATABASE_URL o secrets [base_datos] url): PostgreSQL (Supabase, AWS RDS…).
  Con un esquema (DB_SCHEMA o secrets [base_datos] esquema) las tablas quedan en ese esquema, para compartir la
  base con otros aplicativos sin choques de nombres.

El resto del código escribe SQL al estilo SQLite (parámetros con ?, rowid). Aquí se traduce para PostgreSQL.
"""
import re
import sqlite3
import time

import numpy as np
import pandas as pd

# tablas con identificador de fila (en SQLite es el rowid implícito; en PostgreSQL una columna id)
CON_ID = ("ejecucion", "reglas_fijas", "cargas")
# tablas que no van en el respaldo .db: pesadas y se rehacen cargando de nuevo los archivos de origen
SIN_RESPALDO = ("cartera_archivo", "sispo_archivo")


class Base:
    """Envoltorio común: execute/executemany/commit/fetch, leer() -> DataFrame, anexar(df, tabla)."""

    motor = "sqlite"

    def __init__(self, ruta=None, url=None, esquema=None):
        self.ruta, self.url = ruta, url
        self.esquema = re.sub(r"[^a-z0-9_]", "", (esquema or "").lower()) or None
        self._cache = {}          # resultados de lecturas: se reutilizan hasta que alguien escriba
        self._pendiente = False   # hay escrituras sin confirmar (commit) en la transacción actual
        self._abrir()

    # ---------------------------------------------------------------- conexión
    def _abrir(self):
        if self.url:
            import psycopg2
            from psycopg2.extensions import AsIs, register_adapter
            for t in (np.int64, np.int32):
                register_adapter(t, lambda v: AsIs(int(v)))
            for t in (np.float64, np.float32):
                register_adapter(t, lambda v: AsIs(repr(float(v))) if np.isfinite(v) else AsIs("NULL"))
            register_adapter(np.bool_, lambda v: AsIs(bool(v)))
            self.motor = "postgres"
            self.con = psycopg2.connect(self.url, connect_timeout=15, keepalives=1, keepalives_idle=30)
            if self.esquema:      # tablas propias en su esquema (base compartida con otros aplicativos)
                cur = self.con.cursor()
                cur.execute(f"CREATE SCHEMA IF NOT EXISTS {self.esquema}")
                cur.execute(f"SET search_path TO {self.esquema}")
                self.con.commit()
        else:
            self.con = sqlite3.connect(self.ruta, check_same_thread=False)

    def _sql(self, sql: str) -> str:
        if self.motor == "sqlite":
            return sql
        sql = re.sub(r"\browid\s+AS\s+id\s*,\s*\*", "*", sql, flags=re.I)
        sql = re.sub(r"\browid\b", "id", sql, flags=re.I)
        sql = sql.replace("%", "%%").replace("?", "%s")
        return sql

    def _correr(self, fn):
        """Ejecuta; si la conexión se cayó (Supabase cierra las inactivas) reconecta y reintenta una vez."""
        try:
            return fn()
        except Exception as err:
            if self.motor == "postgres":
                import psycopg2
                try:
                    self.con.rollback()
                except Exception:
                    pass
                if isinstance(err, (psycopg2.OperationalError, psycopg2.InterfaceError)):
                    time.sleep(0.5)
                    self._abrir()
                    return fn()
            raise

    # ---------------------------------------------------------------- API usada por la app
    def execute(self, sql, params=()):
        if not sql.lstrip().upper().startswith(("SELECT", "PRAGMA")):
            self._cache.clear()
            self._pendiente = True
        def f():
            cur = self.con.cursor()
            cur.execute(self._sql(sql), tuple(params))
            return cur
        return self._correr(f)

    def executemany(self, sql, filas):
        filas = [tuple(r) for r in filas]
        if not filas:
            return
        self._cache.clear()
        self._pendiente = True
        def f():
            cur = self.con.cursor()
            if self.motor == "postgres":
                from psycopg2.extras import execute_batch
                execute_batch(cur, self._sql(sql), filas, page_size=1000)
            else:
                cur.executemany(sql, filas)
            return cur
        return self._correr(f)

    def commit(self):
        self.con.commit()
        self._pendiente = False
        self._cache.clear()

    def rollback(self):
        self.con.rollback()
        self._pendiente = False
        self._cache.clear()

    def leer(self, sql, params=()) -> pd.DataFrame:
        """Consulta -> DataFrame. Las lecturas repetidas salen de memoria (rápido en la nube, donde cada consulta
        viaja por internet hasta la base) hasta que haya una escritura."""
        clave = (sql, tuple(params))
        if not self._pendiente and clave in self._cache:
            return self._cache[clave].copy()
        def f():
            cur = self.con.cursor()
            cur.execute(self._sql(sql), tuple(params))
            cols = [d[0] for d in cur.description]
            filas = cur.fetchall()
            if self.motor == "postgres" and not self._pendiente:
                self.con.commit()          # cierra la transacción de lectura (no deja la conexión "ocupada")
            return pd.DataFrame(filas, columns=cols)
        df = self._correr(f)
        if not self._pendiente:
            self._cache[clave] = df
            return df.copy()
        return df

    def anexar(self, df: pd.DataFrame, tabla: str):
        """Inserta las filas de un DataFrame (equivale a df.to_sql(..., if_exists='append'))."""
        if df.empty:
            return
        d = df.astype(object).where(df.notna(), None)
        cols = list(d.columns)
        sql = f'INSERT INTO {tabla} ({", ".join(cols)}) VALUES ({", ".join("?" * len(cols))})'
        self.executemany(sql, d.itertuples(index=False, name=None))

    # ---------------------------------------------------------------- esquema
    def crear_tablas(self, ddl: str):
        for orden in [s.strip() for s in ddl.split(";") if s.strip()]:
            if self.motor == "postgres":
                orden = re.sub(r"\bREAL\b", "DOUBLE PRECISION", orden)
                m = re.match(r"CREATE TABLE IF NOT EXISTS (\w+)\((.*)\)$", orden, flags=re.S)
                if m and m.group(1) in CON_ID:
                    orden = orden[:-1] + ", id BIGSERIAL PRIMARY KEY)"
            self.execute(orden)
        self.commit()

    def columnas(self, tabla: str) -> set:
        if self.motor == "postgres":
            r = self.execute("SELECT column_name FROM information_schema.columns WHERE table_name=? "
                             "AND table_schema=current_schema()", (tabla,))
            return {x[0] for x in r.fetchall()}
        return {x[1] for x in self.execute(f"PRAGMA table_info({tabla})")}

    def tablas(self) -> list:
        if self.motor == "postgres":
            r = self.execute("SELECT table_name FROM information_schema.tables WHERE table_schema=current_schema()")
        else:
            r = self.execute("SELECT name FROM sqlite_master WHERE type='table'")
        return sorted(x[0] for x in r.fetchall())

    def reemplazar_por(self, tabla: str, claves: dict, fila: dict):
        """Upsert simple: borra la fila con esas claves y la inserta de nuevo."""
        where = " AND ".join(f"{k}=?" for k in claves)
        self.execute(f"DELETE FROM {tabla} WHERE {where}", tuple(claves.values()))
        self.execute(f"INSERT INTO {tabla} ({', '.join(fila)}) VALUES ({', '.join('?' * len(fila))})",
                     tuple(fila.values()))

    # ---------------------------------------------------------------- respaldo
    def exportar_sqlite(self) -> bytes:
        """Copia completa de la base en un archivo .db (SQLite), sirve como respaldo en cualquier modo."""
        import os
        import tempfile
        fd, ruta = tempfile.mkstemp(suffix=".db")
        os.close(fd)
        try:
            destino = sqlite3.connect(ruta)
            for t in self.tablas():
                if t in SIN_RESPALDO:
                    continue
                df = self.leer(f"SELECT * FROM {t}")
                if "id" in df.columns and t in CON_ID and self.motor == "postgres":
                    df = df.drop(columns="id")
                df.to_sql(t, destino, index=False, if_exists="replace")
            destino.commit()
            destino.close()
            with open(ruta, "rb") as f:
                return f.read()
        finally:
            os.remove(ruta)

    def importar_sqlite(self, contenido: bytes) -> dict:
        """Reemplaza todas las tablas con las de un respaldo .db. Devuelve filas por tabla."""
        import os
        import tempfile
        fd, ruta = tempfile.mkstemp(suffix=".db")
        with os.fdopen(fd, "wb") as f:
            f.write(contenido)
        resumen = {}
        try:
            origen = sqlite3.connect(ruta)
            nombres = [r[0] for r in origen.execute("SELECT name FROM sqlite_master WHERE type='table'")]
            propias = set(self.tablas())
            for t in nombres:
                if t not in propias:
                    continue
                df = pd.read_sql_query(f"SELECT * FROM {t}", origen)
                cols = [c for c in df.columns if c in self.columnas(t) and c != "id"]
                self.execute(f"DELETE FROM {t}")
                self.anexar(df[cols], t)
                resumen[t] = len(df)
            origen.close()
            self.commit()
        except Exception:
            self.rollback()
            raise
        finally:
            os.remove(ruta)
        return resumen
