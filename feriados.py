# feriados.py
import sqlite3, datetime
from typing import Tuple, Optional

DB_PATH = "reloj_control.db"

# ---- soporte de feriados de Chile (python-holidays) ----
try:
    import holidays  # pip install holidays
    _HOL_LIB_OK = True
except Exception:
    holidays = None
    _HOL_LIB_OK = False


def _ensure_schema():
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS feriados (
            fecha TEXT PRIMARY KEY,          -- 'YYYY-MM-DD'
            nombre TEXT NOT NULL,
            irrenunciable INTEGER NOT NULL DEFAULT 0
        );
    """)
    con.commit()
    con.close()


def marcar_feriado(fecha: datetime.date, nombre: str, irrenunciable: bool = False):
    """Inserta/actualiza un feriado manual en la BD."""
    _ensure_schema()
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    cur.execute("""
        INSERT INTO feriados (fecha, nombre, irrenunciable)
        VALUES (?, ?, ?)
        ON CONFLICT(fecha) DO UPDATE SET
            nombre=excluded.nombre,
            irrenunciable=excluded.irrenunciable
    """, (fecha.isoformat(), nombre.strip(), 1 if irrenunciable else 0))
    con.commit()
    con.close()


def borrar_feriado(fecha: datetime.date):
    _ensure_schema()
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    cur.execute("DELETE FROM feriados WHERE fecha=?", (fecha.isoformat(),))
    con.commit()
    con.close()


def es_feriado(fecha: datetime.date) -> Tuple[bool, Optional[str], bool]:
    """
    Devuelve (es_feriado, nombre, irrenunciable).

    Prioridad:
    1) Lo que esté en la tabla 'feriados' (manual).
    2) Si está instalado 'holidays', consulta feriados de Chile del año correspondiente.
    """
    _ensure_schema()
    iso = fecha.isoformat()

    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    row = cur.execute("SELECT nombre, irrenunciable FROM feriados WHERE fecha=?", (iso,)).fetchone()
    con.close()
    if row:
        return True, row[0], bool(row[1])

    if _HOL_LIB_OK:
        cl = holidays.CL(years=fecha.year)
        if iso in cl:
            return True, str(cl.get(iso)), False

    return False, None, False


def sincronizar_feriados_chile(anio: int):
    """Rellena en BD los feriados del año indicado (si está instalada la librería 'holidays')."""
    if not _HOL_LIB_OK:
        raise RuntimeError("Instala primero: pip install holidays")
    _ensure_schema()
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    cl = holidays.CL(years=anio)
    for d, nombre in cl.items():
        cur.execute("""
            INSERT OR IGNORE INTO feriados (fecha, nombre, irrenunciable) VALUES (?, ?, 0)
        """, (str(d), str(nombre)))
    con.commit()
    con.close()
