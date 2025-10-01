# reportes.py (actualizado con buscador por NOMBRE + colores/leyenda PDF + atrasos persistentes)
import customtkinter as ctk
import sqlite3
from datetime import datetime, timedelta, date
from tkcalendar import Calendar
import tkinter as tk
import pandas as pd
import os
from tkinter import messagebox
import math
from collections import Counter
import smtplib
import ssl
import mimetypes
from email.message import EmailMessage

from feriados import es_feriado  # tu módulo de feriados

# ======== PDF ========
try:
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, KeepTogether
    from reportlab.lib.units import mm
    from reportlab.lib.enums import TA_LEFT
    HAS_PDF = True
except Exception:
    HAS_PDF = False


# ========= Configuración de ajuste especial (43h/44h) =========
AJUSTE_COLACION_FIJO_43_44 = True
MINUTOS_AJUSTE_FIJO = 150
TOLERANCIA_MIN = 5  # tolerancia ingreso (min)

# ========= Política de salida =========
SALIDA_TOLERANCIA_MIN = 5           # ±5 min de tolerancia visual (salida)
LATE_AFTER_EXIT_MINUTES = 40        # 40+ min tarde → (en capturador se pide observación)

# ========= Fallback SMTP =========
SMTP_FALLBACK = {
    "host": "mail.bioaccess.cl",
    "port": 465,  # SSL directo
    "user": "documentos_bd@bioaccess.cl",
    "password": "documentos@2025",
    "use_tls": False,
    "use_ssl": True,
    "remitente": "documentos_bd@bioaccess.cl",
}

# ===================== Helpers de hora =====================

def parse_hora_flexible(hora_str):
    if not hora_str:
        return None
    for fmt in ("%H:%M:%S", "%H:%M"):
        try:
            return datetime.strptime(hora_str.strip(), fmt)
        except ValueError:
            continue
    return None

def _ph_any(s):
    if not s:
        return None
    for f in ("%H:%M:%S", "%H:%M"):
        try:
            return datetime.strptime(s.strip(), f)
        except Exception:
            pass
    return None

# ===================== BD: tablas de atrasos =====================

def crear_tablas_atrasos(db_path="reloj_control.db"):
    con = sqlite3.connect(db_path)
    cur = con.cursor()

    cur.execute("""
    CREATE TABLE IF NOT EXISTS atrasos_diarios (
      rut             TEXT    NOT NULL,
      fecha           DATE    NOT NULL,
      minutos_atraso  INTEGER NOT NULL,
      hora_esperada   TEXT,
      hora_ingreso    TEXT,
      tolerancia_min  INTEGER NOT NULL DEFAULT 5,
      calculado_en    TEXT    NOT NULL DEFAULT (datetime('now')),
      PRIMARY KEY (rut, fecha)
    )""")

    cur.execute("""
    CREATE TABLE IF NOT EXISTS atrasos_mensuales (
      rut                     TEXT    NOT NULL,
      anio                    INTEGER NOT NULL,
      mes                     INTEGER NOT NULL,
      minutos_atraso_total    INTEGER NOT NULL,
      tolerancia_min          INTEGER NOT NULL DEFAULT 5,
      cerrado_en              TEXT    NOT NULL DEFAULT (datetime('now')),
      PRIMARY KEY (rut, anio, mes)
    )""")

    cur.execute("CREATE INDEX IF NOT EXISTS idx_atrasos_m_rut ON atrasos_mensuales(rut)")
    con.commit()
    con.close()

# Asegura las tablas al importar el módulo
try:
    crear_tablas_atrasos()
except Exception as _e:
    print("Aviso: no se pudo crear/verificar tablas de atrasos:", _e)


# ====== ATRASOS: cálculo y persistencia ======

def _calcular_min_atraso(ingreso_str, esperado_str, fecha_iso, toler=None):
    """Devuelve minutos de atraso descontando tolerancia. 0 si no corresponde."""
    if toler is None:
        toler = TOLERANCIA_MIN
    if not ingreso_str or ingreso_str == "--" or not esperado_str:
        return 0
    ti = _ph_any(ingreso_str); te = _ph_any(esperado_str)
    if not ti or not te:
        return 0
    base = datetime.strptime(fecha_iso, "%Y-%m-%d").date()
    di = datetime.combine(base, ti.time())
    de = datetime.combine(base, te.time())
    delta = (di - de).total_seconds()
    if delta > 60 * toler:
        return int(math.ceil((delta - 60 * toler) / 60.0))
    return 0


def guardar_atraso_diario(rut, fecha_iso, minutos, hora_esperada, hora_ingreso, toler=None):
    """UPSERT de atraso diario para (rut, fecha)."""
    if toler is None:
        toler = TOLERANCIA_MIN
    con = sqlite3.connect("reloj_control.db")
    cur = con.cursor()
    cur.execute("""
        INSERT INTO atrasos_diarios (rut, fecha, minutos_atraso, hora_esperada, hora_ingreso, tolerancia_min, calculado_en)
        VALUES (?,?,?,?,?,?, datetime('now'))
        ON CONFLICT(rut, fecha) DO UPDATE SET
          minutos_atraso=excluded.minutos_atraso,
          hora_esperada=excluded.hora_esperada,
          hora_ingreso=excluded.hora_ingreso,
          tolerancia_min=excluded.tolerancia_min,
          calculado_en=datetime('now')
    """, (rut, fecha_iso, int(minutos), hora_esperada, hora_ingreso, int(toler)))
    con.commit(); con.close()


def consolidar_atrasos_mensuales(rut, anio, mes, toler=None):
    """Suma atrasos_diarios del mes y los guarda/actualiza en atrasos_mensuales."""
    if toler is None:
        toler = TOLERANCIA_MIN
    desde, hasta = _month_range(anio, mes)
    con = sqlite3.connect("reloj_control.db"); cur = con.cursor()
    cur.execute("""
        SELECT COALESCE(SUM(minutos_atraso),0)
        FROM atrasos_diarios
        WHERE rut=? AND fecha BETWEEN ? AND ?
    """, (rut, desde.isoformat(), hasta.isoformat()))
    total = int(cur.fetchone()[0] or 0)

    cur.execute("""
        INSERT INTO atrasos_mensuales (rut, anio, mes, minutos_atraso_total, tolerancia_min, cerrado_en)
        VALUES (?,?,?,?,?, datetime('now'))
        ON CONFLICT(rut, anio, mes) DO UPDATE SET
          minutos_atraso_total=excluded.minutos_atraso_total,
          tolerancia_min=excluded.tolerancia_min,
          cerrado_en=datetime('now')
    """, (rut, int(anio), int(mes), total, int(toler)))
    con.commit(); con.close()
    return total


def consolidar_atrasos_por_rango(rut, desde_dt, hasta_dt):
    """Consolida todos los meses que tocan el rango [desde_dt, hasta_dt]."""
    y, m = desde_dt.year, desde_dt.month
    end_y, end_m = hasta_dt.year, hasta_dt.month
    while (y < end_y) or (y == end_y and m <= end_m):
        consolidar_atrasos_mensuales(rut, y, m)
        if m == 12:
            y += 1; m = 1
        else:
            m += 1


def leer_total_atraso_mensual(rut, anio, mes):
    con = sqlite3.connect("reloj_control.db"); cur = con.cursor()
    cur.execute("""SELECT minutos_atraso_total
                   FROM atrasos_mensuales
                   WHERE rut=? AND anio=? AND mes=?""", (rut, int(anio), int(mes)))
    row = cur.fetchone(); con.close()
    return int(row[0]) if row and row[0] is not None else 0


# ===================== Datos del trabajador / horarios =====================

def obtener_horario_del_dia(rut, fecha_dt):
    if isinstance(fecha_dt, str):
        fecha_dt = datetime.strptime(fecha_dt, "%Y-%m-%d")

    dias_en = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
    dias_es = ['lunes', 'martes', 'miercoles', 'jueves', 'viernes', 'sabado', 'domingo']
    dia_norm = dias_es[dias_en.index(fecha_dt.strftime("%A"))]

    con = sqlite3.connect("reloj_control.db")
    cur = con.cursor()
    cur.execute("""
        SELECT hora_entrada, hora_salida
        FROM horarios
        WHERE rut = ?
          AND lower(replace(replace(replace(replace(replace(dia,'á','a'),'é','e'),'í','i'),'ó','o'),'ú','u')) = ?
        ORDER BY time(hora_entrada)
    """, (rut, dia_norm))
    filas = cur.fetchall()
    con.close()

    if not filas:
        return (None, None)

    he = filas[0][0].strip() if filas[0][0] else None
    hs = filas[-1][1].strip() if filas[-1][1] else None  # ÚLTIMA SALIDA (fin de jornada)
    return (he, hs)

def obtener_horario_base(rut):
    try:
        con = sqlite3.connect("reloj_control.db")
        cur = con.cursor()
        cur.execute("""
            SELECT MIN(time(hora_entrada)), MAX(time(hora_salida))
            FROM horarios
            WHERE rut=?
        """, (rut,))
        row = cur.fetchone()
        con.close()
        he = row[0] if row and row[0] else None
        hs = row[1] if row and row[1] else None
        return he, hs
    except Exception:
        return (None, None)

def get_info_trabajador(conn, rut):
    cur = conn.cursor()
    try:
        cur.execute("PRAGMA table_info(trabajadores)")
        cols = {c[1].lower() for c in cur.fetchall()}
    except Exception:
        cols = set()
    if "correo" in cols:
        cur.execute("""
            SELECT nombre, apellido, profesion, correo
            FROM trabajadores
            WHERE rut = ?
        """, (rut,))
    else:
        cur.execute("""
            SELECT nombre, apellido, profesion
            FROM trabajadores
            WHERE rut = ?
        """, (rut,))
    row = cur.fetchone()
    if not row:
        return None
    if "correo" in cols:
        return row
    return (row[0], row[1], row[2], None)

def cargar_trabajadores():
    try:
        con = sqlite3.connect("reloj_control.db")
        cur = con.cursor()
        cur.execute("SELECT rut, nombre, apellido FROM trabajadores ORDER BY apellido, nombre")
        lista = []
        for rut, nom, ape in cur.fetchall():
            nom = nom or ""
            ape = ape or ""
            display = f"{ape.strip()}, {nom.strip()} — {rut.strip()}"
            lista.append({"rut": rut.strip(), "nombre": nom.strip(), "apellido": ape.strip(), "display": display})
        con.close()
        return lista
    except Exception:
        return []

# ===================== Carga Horaria (semanal) =====================

def calcular_carga_horaria_semana(rut):
    con = sqlite3.connect("reloj_control.db")
    cur = con.cursor()
    cur.execute("""
        SELECT lower(replace(replace(replace(replace(replace(dia,'á','a'),'é','e'),'í','i'),'ó','o'),'ú','u')) AS d,
               hora_entrada, hora_salida
        FROM horarios
        WHERE rut = ?
        ORDER BY d, time(hora_entrada)
    """, (rut,))
    filas = cur.fetchall()
    con.close()

    bloques_por_dia = {}
    for d, he, hs in filas:
        bloques_por_dia.setdefault(d, []).append((he, hs))

    total_min = 0
    dias_con_2o_mas_bloques = 0

    for d, bloques in bloques_por_dia.items():
        bloques_validos = 0
        for he, hs in bloques:
            t1 = parse_hora_flexible(he)
            t2 = parse_hora_flexible(hs)
            if t1 and t2:
                mins = int((t2 - t1).total_seconds() // 60)
                if mins > 0:
                    total_min += mins
                    bloques_validos += 1
        if bloques_validos >= 2:
            dias_con_2o_mas_bloques += 1

    if AJUSTE_COLACION_FIJO_43_44 and dias_con_2o_mas_bloques >= 4:
        if abs(total_min - 2430) <= TOLERANCIA_MIN or abs(total_min - 2490) <= TOLERANCIA_MIN:
            total_min += MINUTOS_AJUSTE_FIJO

    return total_min

# ===================== Utilidades BD / SMTP =====================

def obtener_cupo_admin_para_rut(rut):
    try:
        con = sqlite3.connect("reloj_control.db")
        cur = con.cursor()
        cur.execute("""
            SELECT valor FROM parametros_trabajador
            WHERE rut=? AND clave='dias_admin_cupo'
        """, (rut,))
        row = cur.fetchone()
        con.close()
        if row and row[0]:
            return int(row[0])
    except Exception:
        pass
    return None

def contar_admin_anio(rut: str, anio: int | None = None) -> int:
    try:
        if anio is None:
            anio = datetime.now().year
        con = sqlite3.connect("reloj_control.db")
        cur = con.cursor()
        cur.execute("""
            SELECT COUNT(*)
            FROM dias_libres
            WHERE rut=? AND strftime('%Y', fecha)=?
              AND motivo = 'Día Administrativo'
        """, (rut, str(anio)))
        n = cur.fetchone()[0] or 0
        con.close()
        return int(n)
    except Exception:
        return 0

# ===================== Extras mensuales (detección flexible) =====================

def _detectar_tabla_extras(conn):
    cur = conn.cursor()
    try:
        cur.execute("SELECT name FROM sqlite_master WHERE type='table'")
        tablas = [r[0] for r in cur.fetchall()]
    except Exception:
        return None

    candidatos = [
        "extras_mensuales", "extras_mensual", "minutos_extra_mensual",
        "minutos_extras_mensuales", "saldo_extras_mensual", "acumulado_extras_mensual"
    ]
    for t in tablas:
        if t.lower() in candidatos:
            try:
                cur.execute(f"PRAGMA table_info({t})")
                cols = {c[1].lower() for c in cur.fetchall()}
            except Exception:
                cols = set()
            col_min = next((c for c in ["minutos_extra","minutos","min_extras","minutos_extras","total_minutos"] if c in cols), None)
            col_anio = "anio" if "anio" in cols else ("year" if "year" in cols else None)
            col_mes  = "mes"  if "mes"  in cols else ("month" if "month" in cols else None)
            if col_min and col_anio and col_mes and "rut" in cols:
                return {"tabla": t, "col_min": col_min, "col_anio": col_anio, "col_mes": col_mes}
    for t in tablas:
        try:
            cur.execute(f"PRAGMA table_info({t})")
            cols = {c[1].lower() for c in cur.fetchall()}
        except Exception:
            continue
        if "rut" in cols and (("anio" in cols) or ("year" in cols)) and (("mes" in cols) or ("month" in cols)):
            col_min = next((c for c in ["minutos_extra","minutos","min_extras","minutos_extras","total_minutos"] if c in cols), None)
            if col_min:
                return {
                    "tabla": t,
                    "col_min": col_min,
                    "col_anio": "anio" if "anio" in cols else "year",
                    "col_mes":  "mes"  if "mes"  in cols else "month"
                }
    return None

def leer_minutos_extra_mes(rut, anio, mes):
    try:
        con = sqlite3.connect("reloj_control.db")
        meta = _detectar_tabla_extras(con)
        if not meta:
            con.close()
            return None
        t, cmin, canio, cmes = meta["tabla"], meta["col_min"], meta["col_anio"], meta["col_mes"]
        cur = con.cursor()
        cur.execute(f"SELECT {cmin} FROM {t} WHERE rut=? AND {canio}=? AND {cmes}=?", (rut, int(anio), int(mes)))
        row = cur.fetchone()
        con.close()
        if row and row[0] is not None:
            try:
                return int(row[0])
            except Exception:
                val = str(row[0])
                if ":" in val:
                    hh, mm = val.split(":")
                    return int(hh) * 60 + int(mm)
        return None
    except Exception:
        return None

def sumar_minutos_extras_anio(rut, anio):
    total = 0
    for m in range(1, 13):
        mins = leer_minutos_extra_mes(rut, anio, m)
        if mins:
            total += int(mins)
    return total

def sumar_minutos_extras_hasta_mes(rut, anio, mes):
    total = 0
    for m in range(1, mes+1):
        mins = leer_minutos_extra_mes(rut, anio, m)
        if mins:
            total += int(mins)
    return total

# ===================== Email (HTML simple) =====================

def _smtp_load_config():
    cfg = {}
    try:
        con = sqlite3.connect("reloj_control.db")
        cur = con.cursor()
        try:
            cur.execute("SELECT host, port, user, password, use_tls, use_ssl, remitente FROM smtp_config LIMIT 1")
            row = cur.fetchone()
            if row:
                cfg = {
                    "host": row[0],
                    "port": int(row[1]) if row[1] is not None else 0,
                    "user": row[2],
                    "password": row[3],
                    "use_tls": str(row[4]).lower() in ("1", "true", "t", "yes", "y"),
                    "use_ssl": str(row[5]).lower() in ("1", "true", "t", "yes", "y"),
                    "remitente": row[6] or row[2]
                }
                con.close()
                return cfg
        except Exception:
            pass
        try:
            cur.execute("SELECT clave, valor FROM parametros_smtp")
            rows = cur.fetchall()
            if rows:
                mapa = {k: v for k, v in rows}
                cfg = {
                    "host": mapa.get("host"),
                    "port": int(mapa.get("port", "0")),
                    "user": mapa.get("user"),
                    "password": mapa.get("password"),
                    "use_tls": str(mapa.get("use_tls", "true")).lower() in ("1", "true", "t", "yes", "y"),
                    "use_ssl": str(mapa.get("use_ssl", "false")).lower() in ("1", "true", "t", "yes", "y"),
                    "remitente": mapa.get("remitente", mapa.get("user"))
                }
                con.close()
                return cfg
        except Exception:
            pass
        con.close()
    except Exception:
        pass
    return None

def _smtp_send(to_list, cc_list, subject, body_text, html_body, attachment_path=None):
    cfg = _smtp_load_config() or SMTP_FALLBACK

    msg = EmailMessage()
    remitente = cfg.get("remitente") or cfg.get("user")
    msg["From"] = remitente
    msg["To"] = ", ".join(to_list) if to_list else ""
    if cc_list:
        msg["Cc"] = ", ".join(cc_list)
    msg["Subject"] = subject or "Reporte de Asistencia"
    msg.set_content(body_text or "")

    if html_body:
        msg.add_alternative(html_body, subtype="html")

    if attachment_path and os.path.exists(attachment_path):
        with open(attachment_path, "rb") as f:
            data = f.read()
        mime, _ = mimetypes.guess_type(attachment_path)
        if mime:
            maintype, subtype = mime.split("/", 1)
        else:
            maintype, subtype = "application", "octet-stream"
        msg.add_attachment(data, maintype=maintype, subtype=subtype, filename=os.path.basename(attachment_path))

    host = cfg["host"]
    port = cfg.get("port") or (465 if cfg.get("use_ssl") else 587)
    user = cfg.get("user")
    password = cfg.get("password")
    use_ssl = cfg.get("use_ssl")
    use_tls = cfg.get("use_tls")

    if use_ssl:
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL(host, port, context=context, timeout=30) as server:
            if user and password:
                server.login(user, password)
            server.send_message(msg)
    else:
        with smtplib.SMTP(host, port, timeout=30) as server:
            server.ehlo()
            if use_tls:
                context = ssl.create_default_context()
                server.starttls(context=context)
                server.ehlo()
            if user and password:
                server.login(user, password)
            server.send_message(msg)

def _html_email_informe(periodo: str, identidad: str):
    gen = datetime.now().strftime("%d-%m-%Y %H:%M")
    PAGE_BG = "#f3f4f6"
    CARD_BG = "#ffffff"
    TEXT    = "#111827"
    MUTED   = "#6b7280"
    BORDER  = "#e5e7eb"
    ACCENT  = "#0b5ea8"
    return f"""<!doctype html>
<html lang="es"><meta charset="utf-8">
<body style="margin:0;padding:24px;background:{PAGE_BG};color:{TEXT};font-family:Arial,Helvetica,sans-serif">
  <div style="max-width:820px;margin:0 auto;background:{CARD_BG};border:1px solid {BORDER};border-radius:12px;padding:20px;box-sizing:border-box">
    <h1 style="margin:0 0 10px 0;font-size:20px;color:{TEXT}">Reporte de Asistencia</h1>
    <p style="margin:8px 0 14px 0;line-height:1.55;color:{TEXT}">
      Estimado(a), se adjunta el <strong>reporte de asistencia</strong> correspondiente al período indicado.
      Documento generado por <strong>BioAccess – Control de Horarios</strong>.
    </p>
    <div style="margin:0 0 12px 0;line-height:1.7">
      <span style="display:inline-block;background:{ACCENT};color:#ffffff;border-radius:8px;padding:4px 10px;font-size:12px;margin-right:6px">Periodo: {periodo}</span>
      <span style="display:inline-block;background:#eaf2fb;border:1px solid #bfd6f7;color:{ACCENT};border-radius:8px;padding:4px 10px;font-size:12px">Funcionario: {identidad}</span>
    </div>
    <p style="margin:14px 0 0 0;color:{MUTED};font-size:12px">Generado el {gen}.</p>
  </div>
</body>
</html>"""

# ===================== Construcción de RESUMEN MENSUAL =====================

def _month_range(anio, mes):
    first = date(anio, mes, 1)
    if mes == 12:
        last = date(anio+1, 1, 1) - timedelta(days=1)
    else:
        last = date(anio, mes+1, 1) - timedelta(days=1)
    return first, last

def _obtener_registros_mes(rut, anio, mes):
    con = sqlite3.connect("reloj_control.db")
    cur = con.cursor()
    desde, hasta = _month_range(anio, mes)
    cur.execute("""
        SELECT fecha, hora_ingreso, hora_salida, observacion
        FROM registros
        WHERE rut=? AND fecha BETWEEN ? AND ?
        ORDER BY fecha
    """, (rut, desde.isoformat(), hasta.isoformat()))
    rows = cur.fetchall()
    con.close()
    reg = {}
    for fecha, hi, hs, obs in rows:
        reg[fecha] = {"ingreso": hi or "--", "salida": hs or "--", "obs": obs or ""}
    dia = desde
    while dia <= hasta:
        fiso = dia.isoformat()
        if fiso not in reg:
            reg[fiso] = {"ingreso": "--", "salida": "--", "obs": ""}
        dia += timedelta(days=1)
    return reg

def _calcular_metricas_mes(rut, anio, mes):
    registros = _obtener_registros_mes(rut, anio, mes)
    total_min_atraso = 0
    ci_cumple = 0
    ci_total = 0
    cs_cumple = 0
    cs_total = 0
    filas_diarias = []

    for fiso in sorted(registros.keys()):
        info = registros[fiso]
        ingreso = info["ingreso"] or "--"
        salida = info["salida"] or "--"
        obs = info["obs"] or ""

        he, hs = obtener_horario_del_dia(rut, fiso)
        atraso_min = 0
        if ingreso != "--" and he:
            t_i = _ph_any(ingreso)
            t_e = _ph_any(he)
            if t_i and t_e:
                base = datetime.strptime(fiso, "%Y-%m-%d").date()
                di = datetime.combine(base, t_i.time())
                de = datetime.combine(base, t_e.time())
                delta = (di - de).total_seconds()
                if delta > 60 * TOLERANCIA_MIN:
                    atraso_min = math.ceil((delta - 60 * TOLERANCIA_MIN) / 60.0)
                    total_min_atraso += atraso_min

        cumpl_i = "—"
        if he and ingreso != "--":
            ci_total += 1
            try:
                t_i = _ph_any(ingreso); t_e = _ph_any(he)
                base = datetime.strptime(fiso, "%Y-%m-%d").date()
                di = datetime.combine(base, t_i.time())
                de = datetime.combine(base, t_e.time())
                cumpl = di <= (de + timedelta(minutes=TOLERANCIA_MIN))
                cumpl_i = "Sí" if cumpl else "No"
                if cumpl: ci_cumple += 1
            except Exception:
                pass

        cumpl_s = "—"
        if hs and salida != "--":
            cs_total += 1
            try:
                t_s = _ph_any(salida); t_e = _ph_any(hs)
                base = datetime.strptime(fiso, "%Y-%m-%d").date()
                ds = datetime.combine(base, t_s.time())
                de = datetime.combine(base, t_e.time())
                cumpl = ds >= (de - timedelta(minutes=SALIDA_TOLERANCIA_MIN))
                cumpl_s = "Sí" if cumpl else "No"
                if cumpl: cs_cumple += 1
            except Exception:
                pass

        filas_diarias.append([
            datetime.strptime(fiso, "%Y-%m-%d").strftime("%d/%m/%Y"),
            ingreso, salida, cumpl_i, cumpl_s, atraso_min, obs
        ])

    return {
        "total_min_atraso": int(total_min_atraso),
        "cumpl_ingreso": (int(ci_cumple), int(ci_total)),
        "cumpl_salida": (int(cs_cumple), int(cs_total)),
        "filas_diarias": filas_diarias
    }

# ======== Helpers de PDF: Leyenda + Colores condicionales ========

def _tabla_leyenda_pdf():
    data = [
        ["\u25A0", "Incumplimiento horario (Ingreso/Salida = No)", colors.HexColor("#ef4444")],   # rojo
        ["\u25A0", "Administrativo / Permiso / Feriado", colors.HexColor("#3b82f6")],             # azul
        ["\u25A0", "Cumplimiento horario (Ingreso/Salida = Sí)", colors.HexColor("#10b981")],     # verde
        ["\u25A0", "Normal", colors.black]                                                        # negro
    ]
    rows = [[data[i][0], data[i][1]] for i in range(len(data))]
    t = Table(rows, colWidths=[6*mm, 110*mm])
    style = [
        ("GRID", (0,0), (-1,-1), 0.25, colors.lightgrey),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("ALIGN", (0,0), (0,-1), "CENTER"),
        ("FONTNAME", (0,0), (-1,-1), "Helvetica"),
        ("FONTSIZE", (0,0), (-1,-1), 9),
    ]
    for i, row in enumerate(data):
        style.append(("TEXTCOLOR", (0,i), (0,i), row[2]))
        style.append(("TEXTCOLOR", (1,i), (1,i), row[2]))
    t.setStyle(TableStyle(style))
    return t

def _aplicar_colores_condicionales_pdf(tabla, datos_comienzo_fila=1, cols_map=None):
    if cols_map is None:
        return
    estilos = []
    n_rows = len(tabla._cellvalues)
    cum_i = cols_map.get("cumpl_ing")
    cum_s = cols_map.get("cumpl_sal")
    obs_c = cols_map.get("obs")

    for r in range(datos_comienzo_fila, n_rows):
        if cum_i is not None:
            val = str(tabla._cellvalues[r][cum_i])
            if val == "Sí":
                estilos.append(("TEXTCOLOR", (cum_i, r), (cum_i, r), colors.HexColor("#10b981")))
            elif val == "No":
                estilos.append(("TEXTCOLOR", (cum_i, r), (cum_i, r), colors.HexColor("#ef4444")))
        if cum_s is not None:
            val = str(tabla._cellvalues[r][cum_s])
            if val == "Sí":
                estilos.append(("TEXTCOLOR", (cum_s, r), (cum_s, r), colors.HexColor("#10b981")))
            elif val == "No":
                estilos.append(("TEXTCOLOR", (cum_s, r), (cum_s, r), colors.HexColor("#ef4444")))

        if obs_c is not None:
            texto = ""
            cell = tabla._cellvalues[r][obs_c]
            try:
                if hasattr(cell, 'text'):
                    texto = cell.text or ""
                else:
                    texto = str(cell)
            except Exception:
                texto = str(cell)

            low = texto.lower()
            if ("administr" in low) or ("permiso" in low) or ("feriado" in low):
                estilos.append(("TEXTCOLOR", (obs_c, r), (obs_c, r), colors.HexColor("#3b82f6")))

    tabla.setStyle(TableStyle(estilos))

def _pintar_ingreso_salida_por_horario(tabla, data, start_row=1,
                                       col_ing=1, col_esp_ing=2, col_sal=3, col_esp_sal=4):
    estilos = []
    for r in range(start_row, len(data)):
        ti = parse_hora_flexible(str(data[r][col_ing]))
        te = parse_hora_flexible(str(data[r][col_esp_ing]))
        if ti and te:
            if ti > (te + timedelta(minutes=TOLERANCIA_MIN)):
                estilos.append(("TEXTCOLOR", (col_ing, r), (col_ing, r), colors.HexColor("#ef4444")))
            else:
                estilos.append(("TEXTCOLOR", (col_ing, r), (col_ing, r), colors.HexColor("#10b981")))

        ts = parse_hora_flexible(str(data[r][col_sal]))
        to = parse_hora_flexible(str(data[r][col_esp_sal]))
        if ts and to:
            if ts < (to - timedelta(minutes=SALIDA_TOLERANCIA_MIN)):
                estilos.append(("TEXTCOLOR", (col_sal, r), (col_sal, r), colors.HexColor("#ef4444")))
            else:
                estilos.append(("TEXTCOLOR", (col_sal, r), (col_sal, r), colors.HexColor("#10b981")))

    tabla.setStyle(TableStyle(estilos))

# ===================== Exportadores: Excel (módulo) y PDF =====================

def exportar_excel(datos_tabla, resumen, ruta_sugerida, rut_para_extras):
    columnas = [
        "Fecha", "Ingreso", "Esp. Ing.", "Salida", "Esp. Sal.",
        "Horas Trabajadas Por Día", "Minutos Atraso Día", "Min. Extra Mes", "Min. Extra Año (YTD)", "Observación"
    ]
    df = pd.DataFrame(datos_tabla, columns=columnas)
    try:
        df.to_excel(ruta_sugerida, index=False)
        with pd.ExcelWriter(ruta_sugerida, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            rws = []
            tmin = resumen["total_min_trabajados"]
            h, m = divmod(tmin, 60)
            rws.append(["Horas trabajadas (período)", f"{h:02d}:{m:02d}"])
            rws.append(["Minutos acumulados de atraso", resumen["total_min_atraso"]])
            rws.append(["Días administrativos usados (período)", resumen["dias_admin_usados"]])
            rws.append(["Días administrativos pendientes (período)", "—" if resumen["dias_admin_pendientes"] is None else resumen["dias_admin_pendientes"]])
            rws.append(["Feriados del período", resumen["feriados"]])
            rws.append(["Días completos", resumen["dias_ok"]])
            rws.append(["Días incompletos", resumen["dias_incompletos"]])

            anio_actual = datetime.now().year
            limite_admin = obtener_cupo_admin_para_rut(rut_para_extras) or 6
            admin_anio = contar_admin_anio(rut_para_extras, anio_actual)
            admin_rest = max(0, limite_admin - admin_anio)
            rws.append([f"Días administrativos usados (año {anio_actual})", f"{admin_anio}/{limite_admin} (restan {admin_rest})"])

            extras_anio = sumar_minutos_extras_anio(rut_para_extras, anio_actual)
            eh, em = divmod(int(extras_anio), 60)
            rws.append([f"Extras acumulados (año {anio_actual})", f"{eh:02d}:{em:02d}"])

            rws.append([])
            rws.append(["Observaciones (Top 10)", "Veces"])
            for obs, cnt in resumen["obs_counter"].most_common(10):
                rws.append([obs, cnt])
            pd.DataFrame(rws).to_excel(writer, index=False, header=False, sheet_name="Resumen")
        messagebox.showinfo("Exportación exitosa", f"Excel guardado:\n{os.path.basename(ruta_sugerida)}")
        try:
            os.startfile(ruta_sugerida)
        except Exception:
            pass
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo exportar Excel:\n{e}")

def exportar_pdf(rut, nombre, apellido, cargo, periodo, datos_tabla, resumen, abrir=True):
    if not HAS_PDF:
        messagebox.showwarning(
            "PDF no disponible",
            "No se encontró reportlab.\nInstala con: pip install reportlab\nSe intentará exportar Excel como respaldo."
        )
        carpeta_descargas = os.path.join(os.path.expanduser("~"), "Downloads")
        nombre_xlsx = f"reporte_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        exportar_excel(datos_tabla, resumen, os.path.join(carpeta_descargas, nombre_xlsx), rut)
        return None

    carpeta_descargas = os.path.join(os.path.expanduser("~"), "Downloads")
    ruta = os.path.join(carpeta_descargas, f"reporte_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf")

    doc = SimpleDocTemplate(
        ruta,
        pagesize=landscape(A4),
        leftMargin=8*mm, rightMargin=8*mm, topMargin=8*mm, bottomMargin=8*mm
    )
    usable = doc.width

    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="Small", fontSize=9, leading=11))
    styles.add(ParagraphStyle(name="Bold", parent=styles["Normal"], fontSize=10, leading=12))
    styles.add(ParagraphStyle(name="Header", parent=styles["Title"], fontSize=19, leading=22, alignment=TA_LEFT))

    story = []
    story.append(Paragraph("Reporte de Asistencia", styles["Header"]))
    story.append(Paragraph(f"Período: {periodo} | Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}", styles["Normal"]))
    story.append(Spacer(1, 4))

    # Identificación
    tabla_ident = Table(
        [["Nombre:", f"{nombre} {apellido}"],
         ["RUT:", rut],
         ["Cargo:", cargo or ""]],
        colWidths=[0.14*usable, 0.86*usable]
    )
    tabla_ident.setStyle(TableStyle([
        ("GRID",(0,0),(-1,-1),0.3,colors.grey),
        ("BACKGROUND",(0,0),(0,-1),colors.HexColor("#f0f4f8")),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("FONTSIZE",(0,0),(-1,-1),10),
        ("LEFTPADDING",(0,0),(-1,-1),6),("RIGHTPADDING",(0,0),(-1,-1),6),
    ]))
    story.append(KeepTogether([tabla_ident])); story.append(Spacer(1,6))

    # Leyenda
    story.append(Paragraph("<b>Leyenda</b>", styles["Bold"]))
    story.append(Spacer(1,2))
    story.append(_tabla_leyenda_pdf())
    story.append(Spacer(1,8))

    # Resumen
    tmin = resumen["total_min_trabajados"]; h,m = divmod(tmin,60)
    anio_actual = datetime.now().year
    admin_anio_usados = contar_admin_anio(rut, anio_actual)
    limite_admin = obtener_cupo_admin_para_rut(rut) or 6
    admin_anio_rest = max(0, limite_admin - admin_anio_usados)
    extras_anio = sumar_minutos_extras_anio(rut, anio_actual); eh,em = divmod(int(extras_anio),60)

    filas_resumen = [
        ["Horas trabajadas en el período", f"{h:02d}:{m:02d}"],
        ["Minutos acumulados de atraso", f"{resumen['total_min_atraso']} min"],
        ["Días administrativos usados (período)", resumen["dias_admin_usados"]],
        ["Días administrativos pendientes (período)", "—" if resumen["dias_admin_pendientes"] is None else resumen["dias_admin_pendientes"]],
        ["Feriados en el período", resumen["feriados"]],
        ["Días con registro completo", resumen["dias_ok"]],
        ["Días incompletos", resumen["dias_incompletos"]],
        [f"Días administrativos usados (año {anio_actual})", f"{admin_anio_usados} / {limite_admin} (restan {admin_anio_rest})"],
        [f"Extras acumulados (año {anio_actual})", f"{eh:02d}:{em:02d}"],
    ]
    tabla_resumen = Table(filas_resumen, colWidths=[0.76*usable, 0.24*usable])
    tabla_resumen.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(1,0),colors.HexColor("#f0f4f8")),
        ("GRID",(0,0),(-1,-1),0.3,colors.grey),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
        ("ALIGN",(1,0),(1,-1),"RIGHT"),
        ("FONTSIZE",(0,0),(-1,-1),10),
        ("LEFTPADDING",(0,0),(-1,-1),6),("RIGHTPADDING",(0,0),(-1,-1),6),
    ]))
    story.append(KeepTogether([Paragraph("<b>Resumen</b>", styles["Bold"]), Spacer(1,2), tabla_resumen])); story.append(Spacer(1,6))

    # Detalle Diario
    encabezados = ["Fecha","Ingreso","Esp. Ing.","Salida","Esp. Sal.","Trabajado","Min. Atraso Día","Min. Extra Mes","Min. Extra Año (YTD)","Observación"]
    data = [encabezados]
    for fila in datos_tabla:
        data.append(list(fila[:-1]) + [Paragraph(str(fila[-1] or ""), styles["Small"])])

    fractions = [0.08,0.07,0.07,0.07,0.07,0.10,0.10,0.11,0.13,0.20]
    col_widths = [f*usable for f in fractions]

    tabla = Table(data, colWidths=col_widths, repeatRows=1)
    tabla.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#e2e8f0")),
        ("FONTNAME",   (0,0), (-1,0), "Helvetica-Bold"),
        ("ALIGN",      (1,1), (8,-1), "CENTER"),
        ("VALIGN",     (0,0), (-1,-1), "MIDDLE"),
        ("FONTSIZE",   (0,0), (-1,-1), 10),
        ("GRID",       (0,0), (-1,-1), 0.25, colors.grey),
        ("LEFTPADDING",(0,0), (-1,-1), 4),
        ("RIGHTPADDING",(0,0), (-1,-1), 4),
    ]))

    _pintar_ingreso_salida_por_horario(tabla, data, start_row=1, col_ing=1, col_esp_ing=2, col_sal=3, col_esp_sal=4)
    _aplicar_colores_condicionales_pdf(tabla, datos_comienzo_fila=1, cols_map={"obs":9})

    story.append(Paragraph("<b>Detalle Diario</b>", styles["Bold"]))
    story.append(Spacer(1,2))
    story.append(tabla)

    doc.build(story)

    try:
        frame = exportar_pdf.__frame
        frame._ultimo_pdf_path = ruta
    except Exception:
        pass

    if abrir:
        messagebox.showinfo("PDF generado", f"PDF guardado en Descargas:\n{os.path.basename(ruta)}")
        try: os.startfile(ruta)
        except Exception: pass
    return ruta


# ===================== Construir Reportes (UI + lógica) =====================

def construir_reportes(frame_padre):
    registros_por_dia = {}
    # Elementos UI que se usan en callbacks
    label_estado = None
    label_datos = None
    label_total = None
    label_total_semana = None
    label_pactadas = None
    label_completadas = None
    label_resumen = None
    frame_tabla = None
    entry_rut = None
    entry_desde = None
    entry_hasta = None
    entry_nombre = None
    popup_listbox = {"win": None, "list": None}
    _lista_trabajadores = cargar_trabajadores()

    # -------- calendario --------
    def seleccionar_fecha(entry_target):
        top = tk.Toplevel()
        try:
            top.transient(frame.winfo_toplevel())
            top.grab_set()
        except Exception:
            pass
        cal = Calendar(top, date_pattern='dd/mm/yyyy', locale='es_CL')
        cal.pack(padx=10, pady=10)

        def poner_fecha():
            fecha = cal.get_date()
            entry_target.delete(0, 'end')
            entry_target.insert(0, fecha)
            top.destroy()

        ctk.CTkButton(top, text="Seleccionar", command=poner_fecha).pack(pady=5)

    # -------- horas pactadas por día administrativo --------
    def obtener_horas_administrativo(rut, fecha_str_o_dt, como_minutos=False):
        try:
            if isinstance(fecha_str_o_dt, str):
                fecha_dt = datetime.strptime(fecha_str_o_dt, "%Y-%m-%d")
            else:
                fecha_dt = fecha_str_o_dt

            nombre_dia = fecha_dt.strftime("%A")
            dias_map = {
                'Monday': 'lunes', 'Tuesday': 'martes', 'Wednesday': 'miercoles',
                'Thursday': 'jueves', 'Friday': 'viernes', 'Saturday': 'sabado', 'Sunday': 'domingo'
            }
            dia = dias_map.get(nombre_dia, '').lower()
            dia_norm = (dia
                        .replace('á', 'a').replace('é', 'e')
                        .replace('í', 'i').replace('ó', 'o').replace('ú', 'u'))

            conexion = sqlite3.connect("reloj_control.db")
            cursor = conexion.cursor()
            cursor.execute("""
                SELECT hora_entrada, hora_salida FROM horarios
                WHERE rut = ?
                  AND lower(replace(replace(replace(replace(replace(dia,'á','a'),'é','e'),'í','i'),'ó','o'),'ú','u')) = ?
            """, (rut, dia_norm))
            filas = cursor.fetchall()
            conexion.close()

            total_minutos = 0
            for entrada_str, salida_str in filas:
                if not entrada_str or not salida_str:
                    continue
                entrada = parse_hora_flexible(entrada_str)
                salida = parse_hora_flexible(salida_str)
                if not entrada or not salida:
                    continue
                minutos = int((salida - entrada).total_seconds() // 60)
                total_minutos += int(minutos)

            if como_minutos:
                return int(total_minutos)

            h, m = divmod(int(total_minutos), 60)
            return f"{int(h)}h {int(m)}min"

        except Exception as e:
            print("Error al calcular horas administrativas:", e)
            return 0 if como_minutos else "00:00"

    # -------- mezcla administrativos --------
    def agregar_dias_administrativos(regs_por_dia, rut, desde_dt, hasta_dt):
        conexion = sqlite3.connect("reloj_control.db")
        cursor = conexion.cursor()
        desde_str = desde_dt.strftime('%Y-%m-%d')
        hasta_str = hasta_dt.strftime('%Y-%m-%d')

        cursor.execute("""
            SELECT fecha, motivo FROM dias_libres 
            WHERE rut = ? AND fecha BETWEEN ? AND ?
        """, (rut, desde_str, hasta_str))

        dias_admin = cursor.fetchall()
        conexion.close()

        for fecha, motivo in dias_admin:
            try:
                fecha_dt = datetime.strptime(fecha, '%Y-%m-%d').date()
                trabajado = obtener_horas_administrativo(rut, fecha_dt)
                if fecha not in regs_por_dia:
                    regs_por_dia[fecha] = {
                        "ingreso": "--",
                        "salida": "--",
                        "obs_ingreso": motivo,
                        "obs_salida": "",
                        "trabajado": trabajado,
                        "es_admin": True
                    }
                else:
                    regs_por_dia[fecha]["obs_ingreso"] = motivo
                    regs_por_dia[fecha]["trabajado"] = trabajado
                    regs_por_dia[fecha]["es_admin"] = True
            except Exception as e:
                print(f"Error al procesar día administrativo {fecha}: {str(e)}")
                continue

    # -------- estadísticas y resumen del período --------
    def calcular_estadisticas_periodo(rut, registros_dict, desde_dt, hasta_dt):
        total_min_trabajados = 0
        total_min_atraso = 0
        dias_ok = 0
        dias_incompletos = 0
        feriados = 0
        obs_counter = Counter()
        admin_usados = 0

        for fecha in sorted(registros_dict):
            info = registros_dict[fecha]
            ingreso = info.get("ingreso") or "--"
            salida = info.get("salida") or "--"
            es_admin = info.get("es_admin", False)
            es_fer = info.get("es_feriado", False)
            obs_ingreso = info.get("obs_ingreso", "") or ""
            obs_salida = info.get("obs_salida", "") or ""
            obs_txt = (obs_ingreso + " " + obs_salida).strip()
            if obs_txt:
                obs_counter[obs_txt] += 1

            if es_fer:
                feriados += 1

            if ingreso != "--":
                try:
                    fecha_dt = datetime.strptime(fecha, "%Y-%m-%d")
                    nombre_dia = fecha_dt.strftime("%A")
                    dias_map = {"Monday":"Lunes","Tuesday":"Martes","Wednesday":"Miércoles",
                                "Thursday":"Jueves","Friday":"Viernes","Saturday":"Sábado","Sunday":"Domingo"}
                    nombre_dia_es = dias_map.get(nombre_dia, "")
                    conexion = sqlite3.connect("reloj_control.db")
                    cursor = conexion.cursor()
                    cursor.execute("""
                        SELECT hora_entrada FROM horarios
                        WHERE rut = ? AND dia = ?
                    """, (rut, nombre_dia_es))
                    resultado = cursor.fetchone()
                    conexion.close()
                    if resultado:
                        hora_esperada_entrada = resultado[0]
                        t_ingreso = _ph_any(ingreso)
                        t_esperada = _ph_any(hora_esperada_entrada)
                        if t_ingreso and t_esperada:
                            base_date = fecha_dt.date()
                            t_i = datetime.combine(base_date, t_ingreso.time())
                            t_e = datetime.combine(base_date, t_esperada.time())
                            delta = (t_i - t_e).total_seconds()
                            if delta > 300:
                                total_min_atraso += math.ceil((delta - 300) / 60.0)
                except Exception:
                    pass

            trabajado_str = info.get("trabajado")
            if not trabajado_str or trabajado_str == "0h 0min":
                if ingreso != "--" and salida != "--":
                    t1 = parse_hora_flexible(ingreso)
                    t2 = parse_hora_flexible(salida)
                    if t1 and t2:
                        mins = int((t2 - t1).total_seconds() // 60)
                        total_min_trabajados += mins
                        info["trabajado"] = f"{mins//60}h {mins%60}min"
                elif es_admin:
                    mins = obtener_horas_administrativo(rut, fecha, como_minutos=True)
                    total_min_trabajados += int(mins)
            else:
                mins = 0
                if "h" in trabajado_str:
                    try:
                        h_part, m_part = trabajado_str.replace("min", "").split("h")
                        mins = int(h_part.strip()) * 60 + int(m_part.strip())
                    except Exception:
                        mins = 0
                elif ":" in trabajado_str:
                    try:
                        h_part, m_part = trabajado_str.split(":")
                        mins = int(h_part) * 60 + int(m_part)
                    except Exception:
                        mins = 0
                total_min_trabajados += mins

            if (ingreso != "--" and salida != "--") or es_admin or es_fer:
                dias_ok += 1
            else:
                dias_incompletos += 1

            if es_admin:
                admin_usados += 1

        cupo = obtener_cupo_admin_para_rut(rut)
        admin_pendientes = None
        if cupo is not None:
            admin_pendientes = max(0, int(cupo) - int(admin_usados))

        return {
            "total_min_trabajados": int(total_min_trabajados),
            "total_min_atraso": int(total_min_atraso),
            "dias_admin_usados": int(admin_usados),
            "dias_admin_pendientes": admin_pendientes,
            "feriados": int(feriados),
            "dias_ok": int(dias_ok),
            "dias_incompletos": int(dias_incompletos),
            "obs_counter": obs_counter
        }

    # -------- Construir filas para PDF (detalle período visible) --------
    def construir_datos_pdf(regs_por_dia, rut):
        datos = []
        extras_mes_cache = {}
        extras_ytd_cache = {}
        for fecha in sorted(regs_por_dia):
            info = regs_por_dia[fecha]
            ingreso = info.get("ingreso", "--") or "--"
            salida = info.get("salida", "--") or "--"
            obs_ingreso = info.get("obs_ingreso", "") or ""
            obs_salida = info.get("obs_salida", "") or ""
            obs_txt = (obs_ingreso + (" | " if obs_ingreso and obs_salida else "") + obs_salida).strip()

            he, hs = obtener_horario_del_dia(rut, fecha)
            esp_ing = he or "—"
            esp_sal = hs or "—"

            # atraso del día
            min_atraso = 0
            if ingreso != "--" and he:
                try:
                    fecha_dt = datetime.strptime(fecha, "%Y-%m-%d")
                    t_i = _ph_any(ingreso)
                    t_e = _ph_any(he)
                    if t_i and t_e:
                        base = fecha_dt.date()
                        t_i2 = datetime.combine(base, t_i.time())
                        t_e2 = datetime.combine(base, t_e.time())
                        delta = (t_i2 - t_e2).total_seconds()
                        if delta > 60 * TOLERANCIA_MIN:
                            min_atraso = math.ceil((delta - 60 * TOLERANCIA_MIN) / 60.0)
                except Exception:
                    pass

            # persistimos el atraso del día
            try:
                if ingreso != "--" and he:
                    guardar_atraso_diario(rut, fecha, int(min_atraso), he, ingreso, TOLERANCIA_MIN)
            except Exception as e:
                print("WARN guardar_atraso_diario:", e)

            # trabajado
            trabajado = info.get("trabajado", "")
            if not trabajado or trabajado == "0h 0min":
                if ingreso != "--" and salida != "--":
                    t1 = parse_hora_flexible(ingreso)
                    t2 = parse_hora_flexible(salida)
                    if t1 and t2:
                        mins = int((t2 - t1).total_seconds() // 60)
                        h_, m_ = divmod(mins, 60)
                        trabajado = f"{h_:02d}:{m_:02d}"
                    else:
                        trabajado = "--:--"
                else:
                    trabajado = "--:--"
            else:
                if "h" in trabajado:
                    try:
                        h__, m__ = trabajado.replace("min", "").split("h")
                        trabajado = f"{int(h__.strip()):02d}:{int(m__.strip()):02d}"
                    except Exception:
                        pass

            # Min. Extra Mes
            try:
                fdt = datetime.strptime(fecha, "%Y-%m-%d").date()
                key = (fdt.year, fdt.month)
                if key not in extras_mes_cache:
                    val = leer_minutos_extra_mes(rut, fdt.year, fdt.month)
                    extras_mes_cache[key] = int(val) if val is not None else 0
                min_extra_mes = extras_mes_cache[key]
            except Exception:
                min_extra_mes = 0

            # Min. Extra Año (YTD: enero..mes)
            try:
                key2 = (fdt.year, fdt.month)
                if key2 not in extras_ytd_cache:
                    extras_ytd_cache[key2] = int(sumar_minutos_extras_hasta_mes(rut, fdt.year, fdt.month))
                min_extra_ytd = extras_ytd_cache[key2]
            except Exception:
                min_extra_ytd = 0

            datos.append([
                datetime.strptime(fecha, "%Y-%m-%d").strftime("%d/%m/%Y"),
                ingreso, esp_ing, salida, esp_sal, trabajado, min_atraso, min_extra_mes, min_extra_ytd, obs_txt
            ])
        return datos

    # -------- Ventana "Enviar por Correo" --------
    def abrir_envio_correo():
        ctx = getattr(frame, "_ultimo_contexto_pdf", None)
        if not ctx:
            messagebox.showwarning("Sin datos", "Primero genera un reporte (Buscar Reporte).")
            return

        ultimo_pdf = getattr(frame, "_ultimo_pdf_path", None)
        if (not ultimo_pdf or not os.path.exists(ultimo_pdf)) and HAS_PDF:
            try:
                ultimo_pdf = exportar_pdf(
                    rut=ctx["rut"],
                    nombre=ctx["nombre"],
                    apellido=ctx["apellido"],
                    cargo=ctx["cargo"],
                    periodo=ctx["periodo"],
                    datos_tabla=ctx["datos_tabla"],
                    resumen=ctx["resumen"],
                    abrir=False
                )
            except Exception:
                ultimo_pdf = None

        TopLevelCls = getattr(ctk, "CTkToplevel", None) or tk.Toplevel
        master = frame.winfo_toplevel()
        win = TopLevelCls(master)
        win.title("✉️ Enviar por Correo")
        try:
            win.resizable(False, False)
            win.transient(master)
            win.grab_set()
        except Exception:
            pass

        cont = ctk.CTkFrame(win, corner_radius=12)
        cont.pack(fill="both", expand=True, padx=16, pady=16)

        ctk.CTkLabel(cont, text="Destinatarios", font=("Arial", 16, "bold")).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0,8))

        ctk.CTkLabel(cont, text="Para:").grid(row=1, column=0, sticky="e", padx=(0,8), pady=6)
        entry_to = ctk.CTkEntry(cont, width=420, placeholder_text="correo@dominio.cl; otro@dominio.cl")
        entry_to.grid(row=1, column=1, sticky="w", pady=6)

        ctk.CTkLabel(cont, text="CC:").grid(row=2, column=0, sticky="e", padx=(0,8), pady=(0,6))
        entry_cc = ctk.CTkEntry(cont, width=420, placeholder_text="(opcional)")
        entry_cc.grid(row=2, column=1, sticky="w", pady=(0,6))

        try:
            conx = sqlite3.connect("reloj_control.db")
            curx = conx.cursor()
            curx.execute("PRAGMA table_info(trabajadores)")
            cols = {c[1].lower() for c in curx.fetchall()}
            if "correo" in cols:
                curx.execute("SELECT correo FROM trabajadores WHERE rut=?", (ctx["rut"],))
                row = curx.fetchone()
                if row and row[0] and "@" in row[0]:
                    entry_to.insert(0, row[0].strip())
            conx.close()
        except Exception:
            pass

        ctk.CTkLabel(cont, text="Asunto:").grid(row=3, column=0, sticky="e", padx=(0,8), pady=(4,6))
        asunto_def = f"Reporte de asistencia – {ctx['identidad']} – {ctx['periodo']}"
        entry_subject = ctk.CTkEntry(cont, width=420)
        entry_subject.grid(row=3, column=1, sticky="w", pady=(4,6))
        entry_subject.insert(0, asunto_def)

        ctk.CTkLabel(cont, text="Mensaje:").grid(row=4, column=0, sticky="ne", padx=(0,8), pady=(4,6))
        Textbox = getattr(ctk, "CTkTextbox", None)
        msg_def = (
            f"Estimado(a):\n\nAdjunto el reporte de asistencia correspondiente al período {ctx['periodo']}.\n"
            f"Funcionario: {ctx['identidad']}\n\nSaludos cordiales."
        )
        if Textbox:
            txt_msg = Textbox(cont, width=420, height=150)
            txt_msg.grid(row=4, column=1, sticky="w", pady=(4,6))
            txt_msg.insert("1.0", msg_def)
        else:
            wrap = tk.Frame(cont); wrap.grid(row=4, column=1, sticky="w", pady=(4,6))
            txt_msg = tk.Text(wrap, width=56, height=8)
            txt_msg.pack()
            txt_msg.insert("1.0", msg_def)

        row5 = ctk.CTkFrame(cont, fg_color="transparent"); row5.grid(row=5, column=0, columnspan=2, sticky="w", pady=(6,8))
        var_adj = tk.BooleanVar(value=True)
        adj_label = os.path.basename(getattr(frame, "_ultimo_pdf_path", "") or "") or "(se generará PDF)"
        ctk.CTkCheckBox(row5, text=f"Adjuntar PDF: {adj_label}", variable=var_adj).pack(anchor="w")

        btns = ctk.CTkFrame(cont, fg_color="transparent"); btns.grid(row=6, column=0, columnspan=2, sticky="ew", pady=(10,0))
        ctk.CTkButton(btns, text="Cancelar", fg_color="#475569", command=win.destroy, width=120).pack(side="left", padx=(0,6))

        def _enviar():
            to_raw = entry_to.get().strip()
            cc_raw = entry_cc.get().strip()
            subject = entry_subject.get().strip() or "Reporte de asistencia"
            body = txt_msg.get("1.0", "end").strip()
            if not to_raw:
                messagebox.showwarning("Falta destinatario", "Ingresa al menos un correo en 'Para'.")
                return
            to_list = [p.strip() for p in to_raw.replace(",", ";").split(";") if p.strip()]
            cc_list = [c.strip() for c in cc_raw.replace(",", ";").split(";") if c.strip()]
            attach_path = None
            if var_adj.get():
                nonlocal_pdf = getattr(frame, "_ultimo_pdf_path", None)
                if not nonlocal_pdf or not os.path.exists(nonlocal_pdf):
                    try:
                        nonlocal_pdf = exportar_pdf(
                            rut=ctx["rut"],
                            nombre=ctx["nombre"],
                            apellido=ctx["apellido"],
                            cargo=ctx["cargo"],
                            periodo=ctx["periodo"],
                            datos_tabla=ctx["datos_tabla"],
                            resumen=ctx["resumen"],
                            abrir=False
                        )
                    except Exception as e:
                        messagebox.showerror("PDF", f"No se pudo generar el PDF para adjuntar:\n{e}")
                        return
                attach_path = nonlocal_pdf

            try:
                html = _html_email_informe(ctx["periodo"], ctx["identidad"])
                _smtp_send(
                    to_list=to_list,
                    cc_list=cc_list,
                    subject=subject,
                    body_text=body,
                    html_body=html,
                    attachment_path=attach_path
                )
                messagebox.showinfo("Envío exitoso", "Correo enviado correctamente.")
                try:
                    win.grab_release()
                except Exception:
                    pass
                win.destroy()
            except Exception as e:
                messagebox.showerror("Error al enviar", f"No fue posible enviar el correo:\n{e}")

        ctk.CTkButton(btns, text="Enviar", fg_color="#22c55e", command=_enviar, width=160).pack(side="right", padx=(6,0))
        cont.grid_columnconfigure(1, weight=1)
        master.update_idletasks()
        w, h = 640, 460
        x = master.winfo_x() + (master.winfo_width() // 2) - (w // 2)
        y = master.winfo_y() + (master.winfo_height() // 2) - (h // 2)
        win.geometry(f"{w}x{h}+{max(x,0)}+{max(y,0)}")
        win.bind("<Return>", lambda e: _enviar())
        try:
            entry_to.focus_set()
        except Exception:
            pass

    # -------- Buscador por NOMBRE con autocompletado --------
    def _cerrar_popup():
        if popup_listbox["win"] is not None:
            try:
                popup_listbox["win"].destroy()
            except Exception:
                pass
            popup_listbox["win"] = None
            popup_listbox["list"] = None

    def _abrir_popup(opciones):
        _cerrar_popup()
        if not opciones:
            return
        master = frame.winfo_toplevel()
        win = tk.Toplevel(master)
        win.overrideredirect(True)
        win.attributes("-topmost", True)
        lst = tk.Listbox(win, height=min(8, len(opciones)))
        for it in opciones:
            lst.insert("end", it["display"])
        lst.pack(fill="both", expand=True)
        popup_listbox["win"] = win
        popup_listbox["list"] = lst

        try:
            x = entry_nombre.winfo_rootx()
            y = entry_nombre.winfo_rooty() + entry_nombre.winfo_height()
            w = entry_nombre.winfo_width()
            win.geometry(f"{w}x{min(220, 28*len(opciones))}+{x}+{y}")
        except Exception:
            pass

        def _seleccionar(event=None):
            sel = lst.curselection()
            if not sel:
                return
            idx = sel[0]
            item = opciones[idx]
            entry_rut.delete(0, "end")
            entry_rut.insert(0, item["rut"])
            _cerrar_popup()
            try:
                con = sqlite3.connect("reloj_control.db")
                info = get_info_trabajador(con, item["rut"])
                con.close()
                if info:
                    nombre, apellido, profesion, _ = info
                    label_datos.configure(text=f"RUT: {item['rut']}\nNombre: {nombre} {apellido}\nProfesión: {profesion or ''}")
            except Exception:
                pass

        lst.bind("<Double-Button-1>", _seleccionar)
        lst.bind("<Return>", _seleccionar)

        def _tecla(e):
            if e.keysym == "Escape":
                _cerrar_popup()
        win.bind("<Escape>", _tecla)

    def _filtrar_nombres(event=None):
        query = entry_nombre.get().strip().lower()
        if not query:
            _cerrar_popup()
            return
        terms = [t for t in query.split() if t]
        def match(item):
            full = f"{item['apellido']} {item['nombre']}".lower()
            return all(t in full for t in terms)
        opciones = [it for it in _lista_trabajadores if match(it)]
        _abrir_popup(opciones[:20])

    def _enter_buscar(event=None):
        lst = popup_listbox["list"]
        if lst is not None and lst.size() > 0:
            try:
                if not lst.curselection():
                    lst.selection_set(0)
                lst.event_generate("<Return>")
                return "break"
            except Exception:
                pass
        texto = entry_nombre.get().strip().lower()
        if texto:
            candidatos = [it for it in _lista_trabajadores if texto in it["display"].lower()]
            if candidatos:
                entry_rut.delete(0, "end")
                entry_rut.insert(0, candidatos[0]["rut"])
        buscar_reportes()
        return "break"

    # -------- Lógica de búsqueda y render --------
    def buscar_reportes():
        nonlocal registros_por_dia, label_estado, label_datos, label_total, label_total_semana, label_pactadas, label_completadas, label_resumen, frame_tabla
        _cerrar_popup()

        rut = entry_rut.get().strip()
        fecha_desde = entry_desde.get().strip()
        fecha_hasta = entry_hasta.get().strip()

        for widget in frame_tabla.winfo_children():
            widget.destroy()
        registros_por_dia = {}

        if not rut:
            label_estado.configure(text="⚠️ Selecciona un funcionario (nombre) o ingresa un RUT", text_color="red")
            return

        conexion = sqlite3.connect("reloj_control.db")
        info_trabajador = get_info_trabajador(conexion, rut)

        if info_trabajador:
            nombre, apellido, profesion, correo = info_trabajador
            label_datos.configure(text=f"RUT: {rut}\nNombre: {nombre} {apellido}\nProfesión: {profesion}")
        else:
            label_datos.configure(text="RUT no encontrado")
            label_estado.configure(text="⚠️ RUT no registrado", text_color="orange")
            conexion.close()
            return

        if not fecha_desde or not fecha_hasta:
            hoy = datetime.today()
            desde_dt = hoy.replace(day=1)
            hasta_dt = (hoy + timedelta(days=45)).replace(day=1) - timedelta(days=1)
        else:
            try:
                desde_dt = datetime.strptime(fecha_desde, "%d/%m/%Y")
                hasta_dt = datetime.strptime(fecha_hasta, "%d/%m/%Y")
            except ValueError:
                label_estado.configure(text="⚠️ Formato de fecha incorrecto", text_color="red")
                conexion.close()
                return

        cursor = conexion.cursor()
        cursor.execute("""
            SELECT fecha, hora_ingreso, hora_salida, observacion
            FROM registros
            WHERE rut = ? AND fecha BETWEEN ? AND ?
            ORDER BY fecha
        """, (rut, desde_dt.strftime('%Y-%m-%d'), hasta_dt.strftime('%Y-%m-%d')))
        registros = cursor.fetchall()
        conexion.close()

        for fecha, hora_ingreso, hora_salida, observacion in registros:
            if fecha not in registros_por_dia:
                registros_por_dia[fecha] = {
                    "ingreso": "--", "salida": "--",
                    "obs_ingreso": "", "obs_salida": "", "motivo": ""
                }
            if hora_ingreso:
                registros_por_dia[fecha]["ingreso"] = hora_ingreso
                registros_por_dia[fecha]["obs_ingreso"] = observacion or registros_por_dia[fecha]["obs_ingreso"]
            if hora_salida:
                registros_por_dia[fecha]["salida"] = hora_salida
                registros_por_dia[fecha]["obs_salida"] = observacion or registros_por_dia[fecha]["obs_salida"]

        agregar_dias_administrativos(registros_por_dia, rut, desde_dt, hasta_dt)

        # feriados + días vacíos
        dia = desde_dt.date()
        while dia <= hasta_dt.date():
            fiso = dia.isoformat()
            es_f, nombre_f, _ = es_feriado(dia)
            if es_f and fiso not in registros_por_dia:
                registros_por_dia[fiso] = {
                    "ingreso": "--", "salida": "--",
                    "obs_ingreso": f"Feriado: {nombre_f}",
                    "obs_salida": "",
                    "trabajado": "0h 0min",
                    "es_admin": False,
                    "es_feriado": True
                }
            if fiso not in registros_por_dia:
                registros_por_dia[fiso] = {
                    "ingreso": "--", "salida": "--",
                    "obs_ingreso": "", "obs_salida": "",
                    "trabajado": "0h 0min",
                    "es_admin": False,
                    "es_feriado": False
                }
            dia += timedelta(days=1)

        # Autocompletar "Cometido" sin salida guardada
        for fecha, info in registros_por_dia.items():
            obs_txt = (info.get("obs_ingreso") or "") + " " + (info.get("obs_salida") or "") + " " + (info.get("motivo") or "")
            obs_txt = obs_txt.lower()

            if "cometid" in obs_txt and (not info.get("salida") or info.get("salida") == "--"):
                try:
                    fecha_dt = datetime.strptime(fecha, "%Y-%m-%d")
                except Exception:
                    continue

                _, salida_horario = obtener_horario_del_dia(rut, fecha_dt)
                if not salida_horario:
                    continue

                info["salida"] = salida_horario
                t1 = parse_hora_flexible(info.get("ingreso")) if info.get("ingreso") and info.get("ingreso") != "--" else None
                t2 = parse_hora_flexible(salida_horario)
                if t1 and t2:
                    mins = int((t2 - t1).total_seconds() // 60)
                    h_, m_ = divmod(mins, 60)
                    info["trabajado"] = f"{h_}h {m_}min"
                else:
                    info["trabajado"] = obtener_horas_administrativo(rut, fecha_dt)

                info["es_admin"] = True
                if not info.get("obs_salida"):
                    info["obs_salida"] = "Salida autocompletada por Cometido"

        if registros_por_dia:
            # ENCABEZADOS
            encabezados = ["Fecha", "Ingreso", "Esp. Ing.", "Salida", "Esp. Sal.", "Trabajado", "Obs. Ingreso", "Obs. Salida"]
            for col, texto in enumerate(encabezados):
                ctk.CTkLabel(frame_tabla, text=texto, font=("Arial", 13, "bold")).grid(row=0, column=col, padx=8, pady=4, sticky="nsew")

            frame_tabla.grid_columnconfigure(0, weight=0, minsize=90)
            frame_tabla.grid_columnconfigure(1, weight=0, minsize=80)
            frame_tabla.grid_columnconfigure(2, weight=0, minsize=80)
            frame_tabla.grid_columnconfigure(3, weight=0, minsize=80)
            frame_tabla.grid_columnconfigure(4, weight=0, minsize=80)
            frame_tabla.grid_columnconfigure(5, weight=0, minsize=110)
            frame_tabla.grid_columnconfigure(6, weight=1, minsize=250)
            frame_tabla.grid_columnconfigure(7, weight=1, minsize=250)

            fila = 1
            total_minutos = 0
            for fecha in sorted(registros_por_dia):
                ingreso = registros_por_dia[fecha].get("ingreso", "--")
                salida = registros_por_dia[fecha].get("salida", "--")
                obs_ingreso = registros_por_dia[fecha].get("obs_ingreso", "")
                obs_salida = registros_por_dia[fecha].get("obs_salida", "")
                es_admin = registros_por_dia[fecha].get("es_admin", False)
                es_fer = registros_por_dia[fecha].get("es_feriado", False)
                trabajado = registros_por_dia[fecha].get("trabajado", "0h 0min")

                try:
                    fecha_dt = datetime.strptime(fecha, "%Y-%m-%d")
                except Exception:
                    fecha_dt = None
                he, hs = obtener_horario_del_dia(rut, fecha_dt if fecha_dt else fecha)
                esp_ing = he or "—"
                esp_sal = hs or "—"

                # Calcular trabajado efectivo
                if not es_admin and ingreso not in (None, "--") and salida not in (None, "--"):
                    t1 = parse_hora_flexible(ingreso)
                    t2 = parse_hora_flexible(salida)
                    if t1 and t2:
                        duracion = (t2 - t1)
                        minutos = int(duracion.total_seconds() // 60)
                        h_, m_ = divmod(minutos, 60)
                        trabajado = f"{int(h_)}h {int(m_)}min"
                        registros_por_dia[fecha]["trabajado"] = trabajado
                        total_minutos += int(minutos)
                    else:
                        trabajado = "0h 0min"
                        registros_por_dia[fecha]["trabajado"] = trabajado
                elif ingreso in (None, "--") or salida in (None, "--"):
                    if not es_fer:
                        trabajado = registros_por_dia[fecha].get("trabajado", "0h 0min")
                        registros_por_dia[fecha]["trabajado"] = trabajado

                # colores
                ingreso_color = None
                salida_color = None
                try:
                    if ingreso and ingreso != "--" and esp_ing and esp_ing != "—":
                        ti = parse_hora_flexible(ingreso); te = parse_hora_flexible(esp_ing)
                        if ti and te:
                            if ti > (te + timedelta(minutes=TOLERANCIA_MIN)):
                                ingreso_color = "red"
                            else:
                                ingreso_color = "green"
                except Exception:
                    pass
                try:
                    if salida and salida != "--" and esp_sal and esp_sal != "—":
                        ts = parse_hora_flexible(salida); to = parse_hora_flexible(esp_sal)
                        if ts and to:
                            if ts < (to - timedelta(minutes=SALIDA_TOLERANCIA_MIN)):
                                salida_color = "red"
                            else:
                                salida_color = "green"
                except Exception:
                    pass

                fecha_legible = datetime.strptime(fecha, "%Y-%m-%d").strftime("%d/%m/%Y")
                valores = [fecha_legible, ingreso or "--", esp_ing, salida or "--", esp_sal, trabajado, obs_ingreso or "", obs_salida or ""]

                for col, val in enumerate(valores):
                    color = None
                    if es_fer or (("feriad" in (obs_ingreso or "").lower()) or ("feriad" in (obs_salida or "").lower())):
                        color = "#3b82f6"  # azul feriado
                    elif es_admin or ("administr" in (obs_ingreso or "").lower()) or ("permiso" in (obs_ingreso or "").lower()) or ("administr" in (obs_salida or "").lower()) or ("permiso" in (obs_salida or "").lower()):
                        color = "#3b82f6"  # azul admin/permiso
                    else:
                        if col == 1 and ingreso_color:
                            color = ingreso_color
                        if col == 3 and salida_color:
                            color = salida_color
                        if col == 5 and trabajado != "0h 0min":
                            color = color or "green"

                    wrap_len = 280 if col in (6, 7) else 0
                    lbl = ctk.CTkLabel(frame_tabla, text=val)
                    if color:
                        lbl.configure(text_color=color)
                    if wrap_len:
                        lbl.configure(wraplength=wrap_len, justify="left")
                    lbl.grid(row=fila, column=col, padx=8, pady=2, sticky="nsew")
                fila += 1

            h_tot, m_tot = divmod(int(total_minutos), 60)
            label_total.configure(text=f"Total trabajado en el período (efectivo en tabla): {int(h_tot)}h {int(m_tot)}min", text_color="white")

            hoy = datetime.today().date()
            inicio_semana = hoy - timedelta(days=hoy.weekday())
            fin_semana = inicio_semana + timedelta(days=6)
            minutos_semana = 0
            for f, info in registros_por_dia.items():
                fecha_dt = datetime.strptime(f, "%Y-%m-%d").date()
                if inicio_semana <= fecha_dt <= fin_semana:
                    trabajado_str = info.get("trabajado", "")
                    if isinstance(trabajado_str, str) and "h" in trabajado_str:
                        try:
                            hh, mm = trabajado_str.replace("min", "").split("h")
                            minutos_semana += int(hh.strip()) * 60 + int(mm.strip())
                        except Exception:
                            pass
                    elif isinstance(trabajado_str, str) and ":" in trabajado_str:
                        try:
                            hh, mm = trabajado_str.split(":")
                            minutos_semana += int(hh) * 60 + int(mm)
                        except Exception:
                            pass

            h_sem, m_sem = divmod(int(minutos_semana), 60)
            label_total_semana.configure(text=f"Total trabajado en la semana (efectivo): {h_sem}h {m_sem}min")

            minutos_carga = calcular_carga_horaria_semana(rut)
            h_pac, m_pac = divmod(int(minutos_carga), 60)
            cumple = minutos_semana >= minutos_carga
            label_pactadas.configure(text=f"Carga Horaria (semanal): {h_pac}h {m_pac:02d}min")
            label_completadas.configure(
                text=f"Horas completadas esta semana: {h_sem}h {m_sem:02d}min",
                text_color="green" if cumple else "orange"
            )

            resumen = calcular_estadisticas_periodo(rut, registros_por_dia, desde_dt, hasta_dt)
            h_all, m_all = divmod(int(resumen["total_min_trabajados"]), 60)

            anio_actual = datetime.now().year
            limite_admin = obtener_cupo_admin_para_rut(rut) or 6
            admin_anio = contar_admin_anio(rut, anio_actual)
            admin_rest = max(0, limite_admin - admin_anio)

            extras_anio = sumar_minutos_extras_anio(rut, anio_actual)
            eh, em = divmod(int(extras_anio), 60)

            label_resumen.configure(
                text=(
                    f"Resumen (período): Trabajado {h_all:02d}:{m_all:02d} | "
                    f"Atraso {resumen['total_min_atraso']} min | "
                    f"Admin usados {resumen['dias_admin_usados']} | "
                    f"Admin pendientes {'—' if resumen['dias_admin_pendientes'] is None else resumen['dias_admin_pendientes']} | "
                    f"Feriados {resumen['feriados']} | Completos {resumen['dias_ok']} | Incompletos {resumen['dias_incompletos']} | "
                    f"Admin año {anio_actual}: {admin_anio}/{limite_admin} (restan {admin_rest}) | "
                    f"Extras año {anio_actual}: {eh:02d}:{em:02d}"
                )
            )

            # --- NUEVO BLOQUE FINAL ORDENADO DENTRO DE if registros_por_dia ---
            # 1) Construir datos PDF (esto persiste atrasos_diarios día a día)
            datos_pdf = construir_datos_pdf(registros_por_dia, rut)

            # 2) Consolidar todos los meses que toca el rango
            try:
                consolidar_atrasos_por_rango(rut, desde_dt, hasta_dt)
            except Exception as e:
                print("WARN consolidar_atrasos_por_rango:", e)

            # 3) Si el período visible es un único mes, usar el total mensual desde BD
            if desde_dt.year == hasta_dt.year and desde_dt.month == hasta_dt.month:
                try:
                    total_mes_bd = leer_total_atraso_mensual(rut, desde_dt.year, desde_dt.month)
                    resumen["total_min_atraso"] = int(total_mes_bd)
                except Exception as _e:
                    pass

            # 4) Contexto para exportar y estado
            frame._ultimo_contexto_pdf = {
                "rut": rut,
                "nombre": nombre,
                "apellido": apellido,
                "cargo": profesion,
                "identidad": f"{nombre} {apellido}",
                "periodo": f"{desde_dt.strftime('%d/%m/%Y')} a {hasta_dt.strftime('%d/%m/%Y')}",
                "datos_tabla": datos_pdf,
                "resumen": resumen
            }
            frame._ultimo_pdf_path = None
            label_estado.configure(text="✅ Reporte generado", text_color="green")

        else:
            label_total.configure(text="Total trabajado en el período (efectivo en tabla): 0h 0min", text_color="white")
            label_total_semana.configure(text="Total trabajado en la semana (efectivo): 0h 0min")
            minutos_carga = calcular_carga_horaria_semana(rut)
            h_pac, m_pac = divmod(int(minutos_carga), 60)
            label_pactadas.configure(text=f"Carga Horaria (semanal): {h_pac}h {m_pac:02d}min")
            label_completadas.configure(text="Horas completadas esta semana: 0h 00min", text_color="orange")
            label_resumen.configure(text="Resumen (período): —")
            frame._ultimo_contexto_pdf = None
            frame._ultimo_pdf_path = None
            label_estado.configure(text="✅ Reporte generado (sin marcas)", text_color="green")

    def exportar_pdf_click():
        ctx = getattr(frame, "_ultimo_contexto_pdf", None)
        if not ctx:
            messagebox.showwarning("Sin datos", "Primero genera un reporte (Buscar Reporte).")
            return
        exportar_pdf.__frame = frame  # pequeño traspaso de contexto
        exportar_pdf(
            rut=ctx["rut"],
            nombre=ctx["nombre"],
            apellido=ctx["apellido"],
            cargo=ctx["cargo"],
            periodo=ctx["periodo"],
            datos_tabla=ctx["datos_tabla"],
            resumen=ctx["resumen"],
            abrir=True
        )

    # ---------------- UI ----------------
    frame = ctk.CTkFrame(frame_padre, corner_radius=12)
    frame.grid(row=0, column=0, sticky="nsew")
    try:
        frame_padre.grid_rowconfigure(0, weight=1)
        frame_padre.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(3, weight=1)   # tabla
        frame.grid_columnconfigure(0, weight=1)
    except Exception:
        pass

    # Filtros
    filtros = ctk.CTkFrame(frame, fg_color="transparent")
    filtros.grid(row=0, column=0, sticky="ew", padx=10, pady=(10,6))
    for c in range(0, 12):
        filtros.grid_columnconfigure(c, weight=0)
    filtros.grid_columnconfigure(11, weight=1)

    ctk.CTkLabel(filtros, text="Funcionario:").grid(row=0, column=0, padx=(0,6))
    entry_nombre = ctk.CTkEntry(filtros, width=260, placeholder_text="Escribe nombre o apellido")
    entry_nombre.grid(row=0, column=1, padx=(0,12))
    entry_nombre.bind("<KeyRelease>", _filtrar_nombres)
    entry_nombre.bind("<Return>", _enter_buscar)
    entry_nombre.bind("<Escape>", lambda e: _cerrar_popup())

    ctk.CTkLabel(filtros, text="RUT:").grid(row=0, column=2, padx=(0,6))
    entry_rut = ctk.CTkEntry(filtros, width=180, placeholder_text="11.111.111-1")
    entry_rut.grid(row=0, column=3, padx=(0,12))

    ctk.CTkLabel(filtros, text="Desde:").grid(row=0, column=4, padx=(0,6))
    entry_desde = ctk.CTkEntry(filtros, width=120, placeholder_text="dd/mm/aaaa")
    entry_desde.grid(row=0, column=5)
    ctk.CTkButton(filtros, text="📅", width=36, command=lambda: seleccionar_fecha(entry_desde)).grid(row=0, column=6, padx=(6,12))

    ctk.CTkLabel(filtros, text="Hasta:").grid(row=0, column=7, padx=(0,6))
    entry_hasta = ctk.CTkEntry(filtros, width=120, placeholder_text="dd/mm/aaaa")
    entry_hasta.grid(row=0, column=8)
    ctk.CTkButton(filtros, text="📅", width=36, command=lambda: seleccionar_fecha(entry_hasta)).grid(row=0, column=9, padx=(6,12))

    ctk.CTkButton(filtros, text="Buscar Reporte", fg_color="#0ea5e9", command=buscar_reportes).grid(row=0, column=10, padx=(0,8))
    ctk.CTkButton(filtros, text="Exportar PDF", fg_color="#22c55e", command=exportar_pdf_click).grid(row=0, column=11, padx=(0,8))
    ctk.CTkButton(filtros, text="Enviar por correo", fg_color="#6366f1", command=abrir_envio_correo).grid(row=0, column=12)

    # Datos trabajador + estado
    header = ctk.CTkFrame(frame, fg_color="transparent")
    header.grid(row=1, column=0, sticky="ew", padx=10)
    header.grid_columnconfigure(0, weight=1)
    header.grid_columnconfigure(1, weight=1)

    label_datos = ctk.CTkLabel(header, text="RUT / Nombre / Profesión", anchor="w", justify="left")
    label_datos.grid(row=0, column=0, sticky="w")

    label_estado = ctk.CTkLabel(header, text="Listo.", anchor="e", justify="right", text_color="#a3a3a3")
    label_estado.grid(row=0, column=1, sticky="e")

    # Resúmenes rápidos
    res = ctk.CTkFrame(frame, fg_color="transparent")
    res.grid(row=2, column=0, sticky="ew", padx=10, pady=(6,6))
    res.grid_columnconfigure((0,1,2,3), weight=1)

    label_total = ctk.CTkLabel(res, text="Total trabajado en el período (efectivo en tabla): 0h 0min")
    label_total.grid(row=0, column=0, sticky="w", padx=(0,12))

    label_total_semana = ctk.CTkLabel(res, text="Total trabajado en la semana (efectivo): 0h 0min")
    label_total_semana.grid(row=0, column=1, sticky="w", padx=(0,12))

    label_pactadas = ctk.CTkLabel(res, text="Carga Horaria (semanal): —")
    label_pactadas.grid(row=0, column=2, sticky="w", padx=(0,12))

    label_completadas = ctk.CTkLabel(res, text="Horas completadas esta semana: 0h 00min")
    label_completadas.grid(row=0, column=3, sticky="w")

    # Tabla resultados
    frame_tabla = ctk.CTkFrame(frame, corner_radius=8)
    frame_tabla.grid(row=3, column=0, sticky="nsew", padx=10, pady=(4,6))

    # Resumen general al pie
    label_resumen = ctk.CTkLabel(frame, text="Resumen (período): —", justify="left", anchor="w")
    label_resumen.grid(row=4, column=0, sticky="ew", padx=10, pady=(0,10))

    # Fechas por defecto (mes actual)
    try:
        hoy = datetime.today()
        entry_desde.insert(0, hoy.replace(day=1).strftime("%d/%m/%Y"))
        fin = (hoy + timedelta(days=45)).replace(day=1) - timedelta(days=1)
        entry_hasta.insert(0, fin.strftime("%d/%m/%Y"))
    except Exception:
        pass

    return frame


__all__ = ["construir_reportes"]


if __name__ == "__main__":
    try:
        crear_tablas_atrasos()
        print("Listo: tablas de atrasos creadas (si no existían).")
    except Exception as e:
        print("No se pudo crear/verificar tablas:", e)
