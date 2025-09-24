import customtkinter as ctk
import sqlite3
from datetime import datetime, timedelta
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
TOLERANCIA_MIN = 5  # tolerancia

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
    hs = filas[-1][1].strip() if filas[-1][1] else None
    return (he, hs)

# ===================== Datos del trabajador =====================

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

# === NUEVO: contador anual de días administrativos (reinicia cada año) ===
def contar_admin_anio(rut: str, anio: int | None = None) -> int:
    try:
        if anio is None:
            anio = datetime.now().year
        con = sqlite3.connect("reloj_control.db")
        cur = con.cursor()
        # Ajusta a LIKE si tus motivos no son exactos:
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

# ===================== Extras mensuales (nueva tabla) =====================

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

# === NUEVO: acumulado anual de extras (reinicia cada año) ===
def sumar_minutos_extras_anio(rut, anio):
    total = 0
    for m in range(1, 13):
        mins = leer_minutos_extra_mes(rut, anio, m)
        if mins:
            total += int(mins)
    return total

# === NUEVO: acumulado YTD hasta un mes dado (para tabla diaria) ===
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
        try:
            with open(attachment_path, "rb") as f:
                data = f.read()
            mime, _ = mimetypes.guess_type(attachment_path)
            if mime:
                maintype, subtype = mime.split("/", 1)
            else:
                maintype, subtype = "application", "octet-stream"
            msg.add_attachment(data, maintype=maintype, subtype=subtype, filename=os.path.basename(attachment_path))
        except Exception as e:
            raise RuntimeError(f"No se pudo adjuntar el archivo: {e}")

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
  <div style="max-width:720px;margin:0 auto;background:{CARD_BG};border:1px solid {BORDER};border-radius:12px;padding:20px;box-sizing:border-box">
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

# ===================== Construir Reportes (UI + lógica) =====================

def construir_reportes(frame_padre):
    registros_por_dia = {}
    label_estado = None
    label_datos = None
    label_total = None
    label_total_semana = None
    label_pactadas = None
    label_completadas = None
    label_resumen = None

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
                    dias_map = {
                        "Monday": "Lunes", "Tuesday": "Martes", "Wednesday": "Miércoles",
                        "Thursday": "Jueves", "Friday": "Viernes", "Saturday": "Sábado", "Sunday": "Domingo"
                    }
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

    # -------- export EXCEL (respaldo) --------
    def exportar_excel(datos_tabla, resumen, ruta_sugerida, rut_para_extras):
        columnas = ["Fecha", "Ingreso", "Salida", "Horas Trabajadas Por Día", "Minutos Atraso Día", "Min. Extra Mes", "Min. Extra Año (YTD)", "Observación"]
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

                # NUEVO: Admin año + Extras año
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

    # -------- export PDF --------
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
        nombre_pdf = f"reporte_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
        ruta = os.path.join(carpeta_descargas, nombre_pdf)

        doc = SimpleDocTemplate(
            ruta,
            pagesize=landscape(A4),
            rightMargin=14*mm, leftMargin=14*mm,
            topMargin=12*mm, bottomMargin=12*mm
        )

        styles = getSampleStyleSheet()
        styles.add(ParagraphStyle(name="Small", fontSize=9, leading=11))
        styles.add(ParagraphStyle(name="Bold", parent=styles["Normal"], fontSize=10, leading=12, spaceBefore=2, spaceAfter=2))
        styles.add(ParagraphStyle(name="Header", parent=styles["Title"], fontSize=18, leading=22, alignment=TA_LEFT))

        story = []

        story.append(Paragraph("Reporte de Asistencia", styles["Header"]))
        story.append(Paragraph(
            f"Período: {periodo} &nbsp;&nbsp;|&nbsp;&nbsp; Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}",
            styles["Normal"]
        ))
        story.append(Spacer(1, 4))

        tabla_ident = Table(
            [
                ["Nombre:", f"{nombre} {apellido}"],
                ["Cargo:", cargo or ""],
            ],
            colWidths=[25*mm, 120*mm]
        )
        tabla_ident.setStyle(TableStyle([
            ("GRID", (0,0), (-1,-1), 0.3, colors.grey),
            ("BACKGROUND", (0,0), (0,-1), colors.HexColor("#f0f4f8")),
            ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
            ("FONTNAME", (0,0), (-1,-1), "Helvetica"),
            ("FONTSIZE", (0,0), (-1,-1), 10),
            ("LEFTPADDING", (0,0), (-1,-1), 6),
            ("RIGHTPADDING", (0,0), (-1,-1), 6),
        ]))
        story.append(KeepTogether([tabla_ident]))
        story.append(Spacer(1, 8))

        # RESUMEN (incluye admin año y extras año)
        tmin = resumen["total_min_trabajados"]
        h, m = divmod(tmin, 60)

        anio_actual = datetime.now().year
        admin_anio_usados = contar_admin_anio(rut, anio_actual)
        limite_admin = obtener_cupo_admin_para_rut(rut) or 6
        admin_anio_rest = max(0, limite_admin - admin_anio_usados)

        extras_anio = sumar_minutos_extras_anio(rut, anio_actual)
        eh, em = divmod(int(extras_anio), 60)

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
        tabla_resumen = Table(filas_resumen, colWidths=[120*mm, 40*mm])
        tabla_resumen.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (1,0), colors.HexColor("#f0f4f8")),
            ("GRID", (0,0), (-1,-1), 0.3, colors.grey),
            ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
            ("ALIGN", (1,0), (1,-1), "RIGHT"),
            ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
            ("FONTSIZE", (0,0), (-1,-1), 10),
            ("LEFTPADDING", (0,0), (-1,-1), 6),
            ("RIGHTPADDING", (0,0), (-1,-1), 6),
        ]))
        story.append(KeepTogether([Paragraph("<b>Resumen</b>", styles["Bold"]), Spacer(1, 2), tabla_resumen]))
        story.append(Spacer(1, 8))

        top_obs = [["Observación", "Veces"]]
        for obs, cnt in resumen["obs_counter"].most_common(10):
            top_obs.append([obs, cnt if cnt else 0])
        if len(top_obs) == 1:
            top_obs.append(["(Sin observaciones registradas)", 0])

        tabla_obs = Table(top_obs, colWidths=[140*mm, 20*mm])
        tabla_obs.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#eef2ff")),
            ("GRID", (0,0), (-1,-1), 0.3, colors.grey),
            ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
            ("ALIGN", (1,1), (1,-1), "RIGHT"),
            ("VALIGN", (0,0), (-1,-1), "TOP"),
            ("FONTSIZE", (0,0), (-1,-1), 9),
            ("LEFTPADDING", (0,0), (-1,-1), 5),
            ("RIGHTPADDING", (0,0), (-1,-1), 5),
        ]))
        story.append(KeepTogether([Paragraph("<b>Observaciones de Ingreso y Salida (Top 10)</b>", styles["Bold"]), Spacer(1, 2), tabla_obs]))
        story.append(Spacer(1, 8))

        # DETALLE DIARIO: añadimos "Min. Extra Año (YTD)" ANTES de Observación
        encabezados = ["Fecha", "Ingreso", "Salida", "Trabajado", "Min. Atraso Día", "Min. Extra Mes", "Min. Extra Año (YTD)", "Observación"]
        tabla_data = [encabezados] + datos_tabla
        col_widths = [28*mm, 20*mm, 20*mm, 24*mm, 26*mm, 26*mm, 30*mm, 114*mm]  # ~288mm
        tabla = Table(tabla_data, colWidths=col_widths, repeatRows=1)
        tabla.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#e2e8f0")),
            ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
            ("ALIGN", (1,1), (3,-1), "CENTER"),
            ("ALIGN", (4,1), (6,-1), "RIGHT"),
            ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
            ("FONTSIZE", (0,0), (-1,-1), 9),
            ("GRID", (0,0), (-1,-1), 0.25, colors.grey),
            ("LEFTPADDING", (0,0), (-1,-1), 4),
            ("RIGHTPADDING", (0,0), (-1,-1), 4),
        ]))
        story.append(Paragraph("<b>Detalle Diario</b>", styles["Bold"]))
        story.append(Spacer(1, 2))
        story.append(tabla)

        doc.build(story)

        frame._ultimo_pdf_path = ruta
        if abrir:
            messagebox.showinfo("PDF generado", f"PDF guardado en Descargas:\n{os.path.basename(ruta)}")
            try:
                os.startfile(ruta)
            except Exception:
                pass
        return ruta

    # -------- Construir filas para PDF --------
    def construir_datos_pdf(regs_por_dia, rut):
        datos = []
        extras_mes_cache = {}   # (anio, mes) -> minutos
        extras_ytd_cache = {}   # (anio, mes) -> minutos (enero..mes)
        for fecha in sorted(regs_por_dia):
            info = regs_por_dia[fecha]
            ingreso = info.get("ingreso", "--") or "--"
            salida = info.get("salida", "--") or "--"
            obs_ingreso = info.get("obs_ingreso", "") or ""
            obs_salida = info.get("obs_salida", "") or ""
            obs_txt = (obs_ingreso + (" | " if obs_ingreso and obs_salida else "") + obs_salida).strip()

            # atraso día
            min_atraso = 0
            if ingreso != "--":
                try:
                    fecha_dt = datetime.strptime(fecha, "%Y-%m-%d")
                    nombre_dia = fecha_dt.strftime("%A")
                    dias_map = {
                        "Monday": "Lunes", "Tuesday": "Martes", "Wednesday": "Miércoles",
                        "Thursday": "Jueves", "Friday": "Viernes", "Saturday": "Sábado", "Sunday": "Domingo"
                    }
                    nombre_dia_es = dias_map.get(nombre_dia, "")
                    con = sqlite3.connect("reloj_control.db")
                    cur = con.cursor()
                    cur.execute("""
                        SELECT hora_entrada FROM horarios
                        WHERE rut = ? AND dia = ?
                    """, (rut, nombre_dia_es))
                    res = cur.fetchone()
                    con.close()
                    if res:
                        h_esp = res[0]
                        t_i = _ph_any(ingreso)
                        t_e = _ph_any(h_esp)
                        if t_i and t_e:
                            base = fecha_dt.date()
                            t_i2 = datetime.combine(base, t_i.time())
                            t_e2 = datetime.combine(base, t_e.time())
                            delta = (t_i2 - t_e2).total_seconds()
                            if delta > 300:
                                min_atraso = math.ceil((delta - 300) / 60.0)
                except Exception:
                    pass

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
                ingreso, salida, trabajado, min_atraso, min_extra_mes, min_extra_ytd, obs_txt
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
        if Textbox:
            txt_msg = Textbox(cont, width=420, height=150)
            txt_msg.grid(row=4, column=1, sticky="w", pady=(4,6))
            msg_def = (
                f"Estimado(a):\n\nAdjunto el reporte de asistencia correspondiente al período {ctx['periodo']}.\n"
                f"Funcionario: {ctx['identidad']}\n\nSaludos cordiales."
            )
            txt_msg.insert("1.0", msg_def)
        else:
            wrap = tk.Frame(cont); wrap.grid(row=4, column=1, sticky="w", pady=(4,6))
            txt_msg = tk.Text(wrap, width=56, height=8)
            txt_msg.pack()
            txt_msg.insert("1.0", msg_def)

        row5 = ctk.CTkFrame(cont, fg_color="transparent"); row5.grid(row=5, column=0, columnspan=2, sticky="w", pady=(6,8))
        var_adj = tk.BooleanVar(value=True)
        adj_label = os.path.basename(ultimo_pdf) if (ultimo_pdf and os.path.exists(ultimo_pdf)) else "(se generará PDF)"
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

    # -------- Lógica de búsqueda y render --------
    def buscar_reportes():
        nonlocal registros_por_dia, label_estado, label_datos, label_total, label_total_semana, label_pactadas, label_completadas, label_resumen
        rut = entry_rut.get().strip()
        fecha_desde = entry_desde.get().strip()
        fecha_hasta = entry_hasta.get().strip()

        for widget in frame_tabla.winfo_children():
            widget.destroy()
        registros_por_dia = {}

        if not rut:
            label_estado.configure(text="⚠️ Ingresa un RUT", text_color="red")
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
            hasta_dt = (hoy + timedelta(days=45)).replace(day=1)
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
            dia += timedelta(days=1)

        dia = desde_dt.date()
        while dia <= hasta_dt.date():
            fiso = dia.isoformat()
            if fiso not in registros_por_dia:
                registros_por_dia[fiso] = {
                    "ingreso": "--", "salida": "--",
                    "obs_ingreso": "", "obs_salida": "",
                    "trabajado": "0h 0min",
                    "es_admin": False,
                    "es_feriado": False
                }
            dia += timedelta(days=1)

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
            encabezados = ["Fecha", "Ingreso", "Salida", "Trabajado", "Obs. Ingreso", "Obs. Salida"]
            for col, texto in enumerate(encabezados):
                ctk.CTkLabel(frame_tabla, text=texto, font=("Arial", 13, "bold")).grid(row=0, column=col, padx=10, pady=4)

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

                fecha_legible = datetime.strptime(fecha, "%Y-%m-%d").strftime("%d/%m/%Y")
                fila_datos = [fecha_legible, ingreso or "--", salida or "--", trabajado, obs_ingreso or "", obs_salida or ""]
                for col, val in enumerate(fila_datos):
                    color = None
                    if es_fer:
                        color = "#00aaff"
                    elif es_admin:
                        color = "#00bfff"
                    elif col == 3 and trabajado != "0h 0min":
                        color = "green"
                    elif col == 3 and trabajado == "0h 0min":
                        color = "orange"

                    if color:
                        ctk.CTkLabel(frame_tabla, text=val, text_color=color).grid(row=fila, column=col, padx=10, pady=2)
                    else:
                        ctk.CTkLabel(frame_tabla, text=val).grid(row=fila, column=col, padx=10, pady=2)
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

            # NUEVO: Admin año + Extras año en la UI
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

            frame._ultimo_contexto_pdf = {
                "rut": rut,
                "nombre": nombre,
                "apellido": apellido,
                "cargo": profesion,
                "identidad": f"{nombre} {apellido}",
                "periodo": f"{desde_dt.strftime('%d/%m/%Y')} a {hasta_dt.strftime('%d/%m/%Y')}",
                "datos_tabla": construir_datos_pdf(registros_por_dia, rut),
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

    # === INTERFAZ ===
    frame = ctk.CTkFrame(frame_padre)
    frame.pack(fill="both", expand=True)

    ctk.CTkLabel(frame, text="Reportes por Funcionario", font=("Arial", 16)).pack(pady=10)

    # === BUSCADOR POR NOMBRE ===
    lista_nombres = []
    dict_nombre_rut = {}
    try:
        conexion = sqlite3.connect("reloj_control.db")
        cursor = conexion.cursor()
        cursor.execute("SELECT rut, nombre, apellido FROM trabajadores")
        for rut_db, nombre_db, apellido_db in cursor.fetchall():
            nombre_completo = f"{nombre_db} {apellido_db}"
            lista_nombres.append(nombre_completo)
            dict_nombre_rut[nombre_completo] = rut_db
        conexion.close()
    except Exception as e:
        print(f"Error al cargar funcionarios: {e}")

    def actualizar_opciones(valor_actual):
        filtro = (valor_actual or "").lower()
        coincidencias = [n for n in lista_nombres if filtro in n.lower()]
        if not coincidencias:
            coincidencias = ["No hay coincidencias"]
        combo_funcionarios.configure(values=coincidencias)
        combo_funcionarios.set(valor_actual)
        if valor_actual not in dict_nombre_rut:
            entry_rut.delete(0, "end")

    def seleccionar_funcionario(valor):
        if valor == "No hay coincidencias":
            return
        rut_sel = dict_nombre_rut.get(valor, "")
        entry_rut.delete(0, "end")
        entry_rut.insert(0, rut_sel)

    combo_frame = ctk.CTkFrame(frame, fg_color="transparent")
    combo_frame.pack(pady=5)

    def al_hacer_click(event):
        if combo_funcionarios.get() == "Buscar Usuario por nombre":
            combo_funcionarios.set("")

    combo_funcionarios = ctk.CTkComboBox(
        combo_frame,
        values=["Buscar Usuario por nombre"],
        width=400,
        height=35,
        font=("Arial", 14),
        corner_radius=8,
        command=seleccionar_funcionario
    )
    combo_funcionarios.set("Buscar Usuario por nombre")
    combo_funcionarios.pack(side="left", padx=(0, 10))
    combo_funcionarios.bind("<FocusIn>", al_hacer_click)
    combo_funcionarios.bind("<KeyRelease>", lambda event: actualizar_opciones(combo_funcionarios.get()))
    combo_funcionarios.bind("<Return>", lambda event: actualizar_opciones(combo_funcionarios.get()))

    def limpiar_combobox():
        combo_funcionarios.set("Buscar Usuario por nombre")
        combo_funcionarios.configure(values=lista_nombres)
        entry_rut.delete(0, 'end')

    ctk.CTkButton(combo_frame, text="Limpiar", width=80, height=35, command=limpiar_combobox).pack(side="left")

    entry_rut = ctk.CTkEntry(frame, placeholder_text="RUT (Ej: 12345678-9)")
    entry_rut.pack(pady=5)
    entry_rut.bind("<Return>", lambda event: buscar_reportes())

    cont_fecha_desde = ctk.CTkFrame(frame, fg_color="transparent")
    cont_fecha_desde.pack(pady=2)
    entry_desde = ctk.CTkEntry(cont_fecha_desde, placeholder_text="Desde (dd/mm/aaaa)", width=200)
    entry_desde.pack(side="left")
    ctk.CTkButton(cont_fecha_desde, text="📅", width=40, command=lambda: seleccionar_fecha(entry_desde)).pack(side="left", padx=5)

    cont_fecha_hasta = ctk.CTkFrame(frame, fg_color="transparent")
    cont_fecha_hasta.pack(pady=2)
    entry_hasta = ctk.CTkEntry(cont_fecha_hasta, placeholder_text="Hasta (dd/mm/aaaa)", width=200)
    entry_hasta.pack(side="left")
    ctk.CTkButton(cont_fecha_hasta, text="📅", width=40, command=lambda: seleccionar_fecha(entry_hasta)).pack(side="left", padx=5)

    ctk.CTkButton(frame, text="Buscar Reporte", command=buscar_reportes).pack(pady=5)
    ctk.CTkButton(frame, text="Exportar PDF", command=exportar_pdf_click).pack(pady=(5, 2))
    ctk.CTkButton(frame, text="Enviar por Correo", command=abrir_envio_correo).pack(pady=(2, 5))

    label_datos = ctk.CTkLabel(frame, text="RUT: ---\nNombre: ---\nProfesión: ---", font=("Arial", 13), justify="left")
    label_datos.pack(pady=10)

    frame_tabla = ctk.CTkScrollableFrame(frame)
    frame_tabla.pack(pady=10, fill="both", expand=True)

    label_total = ctk.CTkLabel(frame, text="", font=("Arial", 12))
    label_total.pack(pady=2)

    label_total_semana = ctk.CTkLabel(frame, text="", font=("Arial", 12))
    label_total_semana.pack(pady=2)

    label_pactadas = ctk.CTkLabel(frame, text="", font=("Arial", 12))
    label_pactadas.pack(pady=2)

    label_completadas = ctk.CTkLabel(frame, text="", font=("Arial", 12))
    label_completadas.pack(pady=2)

    label_resumen = ctk.CTkLabel(frame, text="", font=("Arial", 12), justify="left")
    label_resumen.pack(pady=4)

    label_estado = ctk.CTkLabel(frame, text="")
    label_estado.pack(pady=10)
