# solicitudes.py
import os, sys, sqlite3, datetime, mimetypes, smtplib, ssl, traceback
import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox
from email.message import EmailMessage

__all__ = ["construir_solicitudes"]

# ---- PDF opcional (rellenar formulario existente) ----
try:
    from PyPDF2 import PdfReader, PdfWriter
except Exception:
    PdfReader = None
    PdfWriter = None

# ---- PDF opcional (fallback PDF simple) ----
try:
    from reportlab.lib.pagesizes import A4
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib import colors
    HAS_REPORTLAB = True
except Exception:
    HAS_REPORTLAB = False

# ---- Calendario ----
try:
    from tkcalendar import DateEntry
    HAS_TKCAL = True
except Exception:
    HAS_TKCAL = False

# -------------------- Paths base --------------------
def app_path():
    return os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else os.path.dirname(os.path.abspath(__file__))

BASE = app_path()
DB_PATH = os.path.join(BASE, "reloj_control.db")

FORM_DIR = os.path.join(BASE, "formularios")
FORM_PATH = os.path.join(FORM_DIR, "SolicitudPermiso.pdf")  # opcional
SALIDA_PDF_DIR = os.path.join(BASE, "salidas_solicitudes")
os.makedirs(SALIDA_PDF_DIR, exist_ok=True)

# -------------------- Email (destinatarios fijos + fallback SMTP) --------------------
DESTINATARIOS = [
    "dennis.flores@slepllanquihue.cl",
    
]

# Fallback (si no hay config en BD)
SMTP_FALLBACK = {
    "host": "mail.bioaccess.cl",
    "port": 465,              # SSL directo
    "user": "solicitud_reloj_control@bioaccess.cl",
    "password": "@solicitud2026",
    "use_ssl": True,
    "use_tls": False,
    "remitente": "solicitud_reloj_control@bioaccess.cl",
}

USAR_SMTP = True  # True para envío real

# =========================================================
#                 ESQUEMA / UTILIDADES BD
# =========================================================
def _ensure_schema():
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS folios (
            id INTEGER PRIMARY KEY CHECK (id = 1),
            ultimo_folio INTEGER NOT NULL DEFAULT 0
        );
    """)
    cur.execute("INSERT OR IGNORE INTO folios (id, ultimo_folio) VALUES (1, 0);")

    cur.execute("""
        CREATE TABLE IF NOT EXISTS solicitudes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            folio INTEGER NOT NULL,
            rut TEXT NOT NULL,
            nombre TEXT NOT NULL,
            tipo_permiso TEXT NOT NULL,
            fecha_desde TEXT NOT NULL,    -- ISO YYYY-MM-DD
            fecha_hasta TEXT NOT NULL,    -- ISO YYYY-MM-DD
            observacion TEXT,
            pdf_path TEXT NOT NULL,
            created_at TEXT NOT NULL
        );
    """)
    cur.execute("CREATE UNIQUE INDEX IF NOT EXISTS idx_solicitudes_folio ON solicitudes(folio);")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_solicitudes_rut ON solicitudes(rut);")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_solicitudes_created_at ON solicitudes(created_at);")
    con.commit()
    con.close()

def get_next_folio():
    con = sqlite3.connect(DB_PATH)
    try:
        con.execute("BEGIN IMMEDIATE;")
        row = con.execute("SELECT ultimo_folio FROM folios WHERE id=1").fetchone()
        ultimo = row[0] if row else 0
        siguiente = ultimo + 1
        con.execute("UPDATE folios SET ultimo_folio=? WHERE id=1", (siguiente,))
        con.commit()
        return siguiente
    except Exception:
        con.rollback()
        raise
    finally:
        con.close()

def guardar_solicitud_en_bd(folio, rut, nombre, tipo, desde, hasta, obs, pdf_path):
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    cur.execute("""
        INSERT INTO solicitudes (folio, rut, nombre, tipo_permiso, fecha_desde, fecha_hasta, observacion, pdf_path, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        folio, rut, nombre, tipo, desde, hasta, obs,
        pdf_path,
        datetime.datetime.now().isoformat(timespec="seconds")
    ))
    con.commit()
    con.close()

# =========================================================
#          NOMBRES / RUT y datos extra (cargo/profesión)
# =========================================================
def _cols_trabajadores():
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    try:
        return [r[1] for r in cur.execute("PRAGMA table_info(trabajadores);").fetchall()]
    finally:
        con.close()

def _trabajadores_tiene(col_name: str) -> bool:
    return col_name in _cols_trabajadores()

def cargar_nombres_ruts():
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    try:
        if _trabajadores_tiene("apellido"):
            cur.execute("SELECT rut, nombre, apellido FROM trabajadores ORDER BY nombre, apellido")
            filas = cur.fetchall()
            lista = [f"{n} {a}".strip() for (_, n, a) in filas]
            mapping = {f"{n} {a}".strip(): r for (r, n, a) in filas}
        else:
            cur.execute("SELECT rut, nombre FROM trabajadores ORDER BY nombre")
            filas = cur.fetchall()
            lista = [n for (_, n) in filas]
            mapping = {n: r for (r, n) in filas}
        return lista, mapping
    finally:
        con.close()

def obtener_cargo_por_rut(rut: str) -> str:
    if not rut:
        return ""
    cols = _cols_trabajadores()
    cand = None
    for c in ["cargo", "profesion", "profesión", "puesto", "cargo_profesion"]:
        if c in cols:
            cand = c; break
    if not cand:
        return ""
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    try:
        cur.execute(f"SELECT {cand} FROM trabajadores WHERE rut=?", (rut,))
        row = cur.fetchone()
        return (row[0] or "") if row else ""
    finally:
        con.close()

def obtener_correo_por_rut(rut: str) -> str:
    """Retorna correo del trabajador si existe columna 'correo'."""
    if not rut:
        return ""
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    try:
        cur.execute("PRAGMA table_info(trabajadores)")
        cols = {c[1].lower() for c in cur.fetchall()}
        if "correo" not in cols:
            return ""
        cur.execute("SELECT correo FROM trabajadores WHERE rut=?", (rut,))
        row = cur.fetchone()
        return (row[0] or "").strip() if row else ""
    except Exception:
        return ""
    finally:
        con.close()

# =========================================================
#             TIPOS DE PERMISO + AUTOCÓMPUTO
# =========================================================
opciones_permiso = [
    "Elija tipo de Solicitud o Permiso",
    "Día Administrativo (Día Completo)",
    "Día Administrativo (Medio Día)",
    "Día Administrativo (Horas)",
    "Permiso por Matrimonio/Acuerdo Unión Civil (5 días hábiles)",
    "Permiso por Defunción de Hijo (10 días corridos)",
    "Permiso por Defunción de Cónyuge/Conviviente Civil (7 días corridos)",
    "Permiso por Defunción de Hijo en Gestación (7 días hábiles)",
    "Permiso por Defunción de Padre/Madre/Hermano(a) (4 días hábiles)",
    "Permiso de Nacimiento Paternal (5 días corridos)",
    "Permiso de Alimentación (1 hora diaria)",
    "Permiso sin Goce de Sueldo (máx 6 meses)",
    "Cometido de Servicio",
    "Otro (especificar)"
]
dias_por_permiso = {
    "Permiso por Matrimonio/Acuerdo Unión Civil (5 días hábiles)": 5,
    "Permiso por Defunción de Hijo (10 días corridos)": 10,
    "Permiso por Defunción de Cónyuge/Conviviente Civil (7 días corridos)": 7,
    "Permiso por Defunción de Hijo en Gestación (7 días hábiles)": 7,
    "Permiso por Defunción de Padre/Madre/Hermano(a) (4 días hábiles)": 4,
    "Permiso de Nacimiento Paternal (5 días corridos)": 5,
    "Permiso sin Goce de Sueldo (máx 6 meses)": 180,
}
def _is_habiles(tipo_texto: str) -> bool:
    t = tipo_texto.lower()
    return "hábil" in t or "habil" in t

def _add_days_inclusive(start_date: datetime.date, days: int, habiles: bool) -> datetime.date:
    if days <= 1:
        return start_date
    if not habiles:
        return start_date + datetime.timedelta(days=days - 1)
    d = start_date
    count = 0
    while count < days:
        if d.weekday() < 5:
            count += 1
            if count == days:
                break
        d += datetime.timedelta(days=1)
    return d

# =========================================================
#                    PDF helpers (rellenar o fallback)
# =========================================================
def completar_pdf_campos(entrada_pdf: str, salida_pdf: str, campos: dict):
    if PdfReader is None or PdfWriter is None:
        raise RuntimeError("PyPDF2 no está instalado. Ejecuta: pip install PyPDF2==3.0.1")
    reader = PdfReader(entrada_pdf)
    writer = PdfWriter()
    writer.append_pages_from_reader(reader)

    root = reader.trailer.get("/Root", {})
    if "/AcroForm" in root:
        writer._root_object.update({"/AcroForm": root["/AcroForm"]})
        writer._root_object["/AcroForm"].update({"/NeedAppearances": True})
        fields = writer.get_fields() or {}
        to_update = {k: str(v) for k, v in campos.items() if k in fields}
        if to_update:
            writer.update_page_form_field_values(writer.pages, to_update)

    with open(salida_pdf, "wb") as f:
        writer.write(f)

def _pdf_simple_fallback(path_out: str, datos: dict):
    """Crea un PDF simple si no hay formulario o PyPDF2/llenado falla."""
    if not HAS_REPORTLAB:
        return False
    styles = getSampleStyleSheet()
    title = styles["Title"]
    normal = styles["Normal"]
    elems = []
    elems.append(Paragraph("Solicitud de Permiso (Reserva)", title))
    elems.append(Paragraph("Documento informativo; sujeto a validación administrativa.", normal))
    elems.append(Spacer(0, 8))
    rows = [
        ["Folio", f"{datos.get('folio','')}"],
        ["Nombre", datos.get("nombre","")],
        ["RUT", datos.get("rut","")],
        ["Cargo/Profesión", datos.get("cargo","")],
        ["Tipo de Permiso", datos.get("tipo","")],
        ["Desde", datos.get("desde","")],
        ["Hasta", datos.get("hasta","")],
        ["Detalle/Obs.", datos.get("observacion","") or "-"],
        ["Generado", datos.get("generado","")],
    ]
    tbl = Table(rows, colWidths=[140, 380])
    tbl.setStyle(TableStyle([
        ("GRID", (0,0), (-1,-1), 0.3, colors.grey),
        ("BACKGROUND", (0,0), (0,-1), colors.whitesmoke),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("FONTSIZE", (0,0), (-1,-1), 10),
        ("LEFTPADDING", (0,0), (-1,-1), 6),
        ("RIGHTPADDING", (0,0), (-1,-1), 6),
    ]))
    elems.append(tbl)
    doc = SimpleDocTemplate(path_out, pagesize=A4, leftMargin=36, rightMargin=36, topMargin=36, bottomMargin=36)
    doc.build(elems)
    return True

def nombre_archivo_por_formato(fecha: datetime.date, folio: int, usar_ddmmyyyy=True):
    fecha_str = fecha.strftime("%d%m%Y") if usar_ddmmyyyy else fecha.strftime("%Y%m%d")
    return f"{fecha_str}_F{folio:06}.pdf"

def _build_pdf_solicitud(folio: int, rut: str, nombre: str, cargo: str, tipo: str,
                         desde: datetime.date, hasta: datetime.date, observacion: str) -> str | None:
    """
    Intenta:
      1) Llenar el formulario PDF (si existe y PyPDF2 está disponible).
      2) Si falla, genera PDF simple (si reportlab está disponible).
      3) Si todo falla, retorna None.
    """
    os.makedirs(SALIDA_PDF_DIR, exist_ok=True)
    file_out = os.path.join(SALIDA_PDF_DIR, nombre_archivo_por_formato(desde, folio, usar_ddmmyyyy=True))
    generado = datetime.datetime.now().strftime("%d-%m-%Y %H:%M")

    # Campos candidatos (solo se llenan si existen en el PDF)
    campos = {
        "folio": f"{folio:06d}",
        "Folio": f"{folio:06d}",
        "RUT": rut,
        "rut": rut,
        "Nombre": nombre,
        "nombre": nombre,
        "Cargo": cargo,
        "cargo": cargo,
        "Tipo": tipo,
        "tipo": tipo,
        "Desde": desde.strftime("%d-%m-%Y"),
        "desde": desde.strftime("%d-%m-%Y"),
        "Hasta": hasta.strftime("%d-%m-%Y"),
        "hasta": hasta.strftime("%d-%m-%Y"),
        "Observacion": observacion or "",
        "observacion": observacion or "",
        "FechaEmision": generado.split(" ")[0],
        "HoraEmision": generado.split(" ")[1],
    }

    # 1) Intentar llenar formulario
    try:
        if FORM_PATH and os.path.exists(FORM_PATH) and PdfReader is not None and PdfWriter is not None:
            completar_pdf_campos(FORM_PATH, file_out, campos)
            return file_out
    except Exception:
        traceback.print_exc()

    # 2) Fallback PDF simple
    try:
        ok = _pdf_simple_fallback(file_out, {
            "folio": f"{folio:06d}",
            "rut": rut, "nombre": nombre, "cargo": cargo,
            "tipo": tipo,
            "desde": desde.strftime("%d-%m-%Y"),
            "hasta": hasta.strftime("%d-%m-%Y"),
            "observacion": observacion or "",
            "generado": generado,
        })
        if ok:
            return file_out
    except Exception:
        traceback.print_exc()

    # 3) Nada
    return None

# =========================================================
#                    EMAIL (config + envío)
# =========================================================
def _smtp_load_config():
    """
    Intenta leer configuración SMTP desde BD:
      - smtp_config(host,port,user,password,use_tls,use_ssl,remitente) o
      - parametros_smtp(clave,valor)
    Si no existe, retorna None (usaremos fallback).
    """
    cfg = {}
    try:
        con = sqlite3.connect(DB_PATH)
        cur = con.cursor()
        # Opción 1
        try:
            cur.execute("SELECT host, port, user, password, use_tls, use_ssl, remitente FROM smtp_config LIMIT 1")
            row = cur.fetchone()
            if row:
                cfg = {
                    "host": row[0],
                    "port": int(row[1]) if row[1] else 0,
                    "user": row[2],
                    "password": row[3],
                    "use_tls": str(row[4]).lower() in ("1","true","t","yes","y"),
                    "use_ssl": str(row[5]).lower() in ("1","true","t","yes","y"),
                    "remitente": row[6] or row[2],
                }
                con.close()
                return cfg
        except Exception:
            pass
        # Opción 2
        try:
            cur.execute("SELECT clave, valor FROM parametros_smtp")
            rows = cur.fetchall()
            if rows:
                m = {k:v for k,v in rows}
                cfg = {
                    "host": m.get("host"),
                    "port": int(m.get("port", "0")),
                    "user": m.get("user"),
                    "password": m.get("password"),
                    "use_tls": str(m.get("use_tls", "true")).lower() in ("1","true","t","yes","y"),
                    "use_ssl": str(m.get("use_ssl", "false")).lower() in ("1","true","t","yes","y"),
                    "remitente": m.get("remitente", m.get("user")),
                }
                con.close()
                return cfg
        except Exception:
            pass
        con.close()
    except Exception:
        pass
    return None

def _smtp_send(to_list, cc_list, subject, body_text, html_body=None, attachment_path=None):
    if not USAR_SMTP:
        info = f"(SIMULADO)\nPara: {', '.join(to_list)}\nCC: {', '.join(cc_list or [])}\nAsunto: {subject}\n\n{body_text}"
        messagebox.showinfo("Envío no configurado", info)
        return

    cfg = _smtp_load_config() or SMTP_FALLBACK
    host = cfg["host"]
    port = cfg.get("port") or (465 if cfg.get("use_ssl") else 587)
    user = cfg.get("user")
    password = cfg.get("password")
    use_ssl = cfg.get("use_ssl")
    use_tls = cfg.get("use_tls")
    remitente = cfg.get("remitente") or user

    msg = EmailMessage()
    msg["From"] = remitente
    msg["To"] = ", ".join([t for t in to_list if t])
    if cc_list:
        msg["Cc"] = ", ".join([c for c in cc_list if c])
    msg["Subject"] = subject or "Solicitud de Permiso (Reserva)"
    msg.set_content(body_text or "")

    if html_body:
        msg.add_alternative(html_body, subtype="html")

    if attachment_path and os.path.exists(attachment_path):
        ctype, _ = mimetypes.guess_type(attachment_path)
        maintype, subtype = (ctype.split("/", 1) if ctype else ("application", "octet-stream"))
        with open(attachment_path, "rb") as f:
            msg.add_attachment(f.read(), maintype=maintype, subtype=subtype, filename=os.path.basename(attachment_path))

    if use_ssl:
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL(host, port, context=context, timeout=30) as s:
            if user and password:
                s.login(user, password)
            s.send_message(msg)
    else:
        with smtplib.SMTP(host, port, timeout=30) as s:
            s.ehlo()
            if use_tls:
                context = ssl.create_default_context()
                s.starttls(context=context)
                s.ehlo()
            if user and password:
                s.login(user, password)
            s.send_message(msg)

def _html_email_solicitud(folio:int, nombre:str, rut:str, cargo:str, tipo:str, desde:str, hasta:str, obs:str):
    gen = datetime.datetime.now().strftime("%d-%m-%Y %H:%M")
    PAGE_BG = "#f3f4f6"; CARD_BG = "#ffffff"; TEXT = "#111827"; MUTED="#6b7280"; BORDER="#e5e7eb"; ACCENT="#0b5ea8"
    obs_html = (obs or "-").replace("\n", "<br>")
    return f"""<!doctype html>
<html lang="es"><meta charset="utf-8">
<body style="margin:0;padding:24px;background:{PAGE_BG};color:{TEXT};font-family:Arial,Helvetica,sans-serif">
  <div style="max-width:720px;margin:0 auto;background:{CARD_BG};border:1px solid {BORDER};border-radius:12px;padding:20px;box-sizing:border-box">
    <h1 style="margin:0 0 10px 0;font-size:20px;color:{TEXT}">Reserva de Solicitud</h1>
    <p style="margin:8px 0 10px 0;line-height:1.55">
      Se informa nueva <strong>reserva de solicitud</strong> (sujeta a validación administrativa).
    </p>
    <div style="margin:8px 0 12px 0;line-height:1.7">
      <span style="display:inline-block;background:{ACCENT};color:#fff;border-radius:8px;padding:4px 10px;font-size:12px;margin-right:6px">Folio: {folio:06d}</span>
      <span style="display:inline-block;background:#eaf2fb;border:1px solid #bfd6f7;color:{ACCENT};border-radius:8px;padding:4px 10px;font-size:12px;margin-right:6px">Tipo: {tipo}</span>
      <span style="display:inline-block;background:#eaf2fb;border:1px solid #bfd6f7;color:{ACCENT};border-radius:8px;padding:4px 10px;font-size:12px">Periodo: {desde} → {hasta}</span>
    </div>
    <table style="width:100%;border-collapse:collapse;border:1px solid {BORDER}">
      <tr><td style="background:#f8fafc;padding:8px;width:160px">Nombre</td><td style="padding:8px">{nombre}</td></tr>
      <tr><td style="background:#f8fafc;padding:8px">RUT</td><td style="padding:8px">{rut}</td></tr>
      <tr><td style="background:#f8fafc;padding:8px">Cargo / Profesión</td><td style="padding:8px">{cargo or "-"}</td></tr>
      <tr><td style="background:#f8fafc;padding:8px">Observación</td><td style="padding:8px">{obs_html}</td></tr>
      <tr><td style="background:#f8fafc;padding:8px">Generado</td><td style="padding:8px">{gen}</td></tr>
    </table>
    <p style="margin:12px 0 0 0;color:{MUTED};font-size:12px">
      Nota: esta comunicación es informativa y no reemplaza el formulario oficial firmado.
    </p>
  </div>
</body>
</html>"""

# =========================================================
#                    UI PRINCIPAL
# =========================================================
def construir_solicitudes(frame_padre, on_volver=None):
    """
    Layout centrado: labels a la izquierda, campos a la derecha.
    Fechas dd-mm-yy con calendario y autocalc de 'Hasta' (en fila aparte).
    Sin PDF obligatorio; correo con aviso de “reserva, no oficial”.
    """
    _ensure_schema()

    # limpiar y contenedor base
    for w in frame_padre.winfo_children():
        w.destroy()

    cont = ctk.CTkFrame(frame_padre)
    cont.pack(fill="both", expand=True, padx=20, pady=20)
    cont.grid_columnconfigure(0, weight=1)
    cont.grid_columnconfigure(2, weight=1)

    box = ctk.CTkFrame(cont)
    box.grid(row=0, column=1, sticky="n", padx=10, pady=10)

    titulo = ctk.CTkLabel(box, text="Solicitud de Permiso / Día Administrativo", font=("Arial", 18))
    titulo.grid(row=0, column=0, columnspan=4, pady=(0, 12))

    # ---------- FORM ---------
    form = ctk.CTkFrame(box, fg_color="transparent")
    form.grid(row=1, column=0, sticky="n")

    # Más padding vertical entre filas
    for i in range(0, 12):
        form.grid_rowconfigure(i, pad=10)

    # Labels con más espacio a la derecha
    LABEL_PADX = (0, 12)
    LABEL_W = 180

    # ---- Nombre (Combo) + Buscar
    ctk.CTkLabel(form, text="Nombre:", width=LABEL_W, anchor="e").grid(row=0, column=0, sticky="e", padx=LABEL_PADX)
    lista_nombres, dict_nombre_rut = cargar_nombres_ruts()
    primeros_10 = lista_nombres[:10]
    combo_nombre = ctk.CTkComboBox(form, values=primeros_10, width=320)
    combo_nombre.set("Buscar por Nombre")
    combo_nombre.grid(row=0, column=1, sticky="w")
    btn_buscar = ctk.CTkButton(form, text="🔍 Buscar")
    btn_buscar.grid(row=0, column=2, padx=(8, 0), sticky="w")

    # ---- RUT
    ctk.CTkLabel(form, text="RUT:", width=LABEL_W, anchor="e").grid(row=1, column=0, sticky="e", padx=LABEL_PADX)
    entry_rut = ctk.CTkEntry(form, placeholder_text="RUT: (Ej: 12345678-9)", width=240)
    entry_rut.grid(row=1, column=1, sticky="w")

    # ---- Cargo/Profesión (auto)
    ctk.CTkLabel(form, text="Cargo / Profesión:", width=LABEL_W, anchor="e").grid(row=2, column=0, sticky="e", padx=LABEL_PADX)
    entry_cargo = ctk.CTkEntry(form, placeholder_text="(auto)", width=320)
    entry_cargo.grid(row=2, column=1, columnspan=2, sticky="w")
    entry_cargo.configure(state="disabled")

    # ---- Tipo de permiso
    ctk.CTkLabel(form, text="Tipo de Permiso:", width=LABEL_W, anchor="e").grid(row=3, column=0, sticky="e", padx=LABEL_PADX)
    cmb_tipo = ctk.CTkOptionMenu(form, values=opciones_permiso, width=320)
    cmb_tipo.set(opciones_permiso[0])
    cmb_tipo.grid(row=3, column=1, columnspan=2, sticky="w")

    # ---- Fechas (en filas separadas)
    ctk.CTkLabel(form, text="Desde (dd-mm-yy):", width=LABEL_W, anchor="e").grid(row=4, column=0, sticky="e", padx=LABEL_PADX)
    if HAS_TKCAL:
        entry_desde = DateEntry(form, date_pattern="dd-mm-yy", width=14)
    else:
        entry_desde = ctk.CTkEntry(form, placeholder_text="dd-mm-yy", width=140)
    entry_desde.grid(row=4, column=1, sticky="w")

    ctk.CTkLabel(form, text="Hasta (dd-mm-yy):", width=LABEL_W, anchor="e").grid(row=5, column=0, sticky="e", padx=LABEL_PADX)
    if HAS_TKCAL:
        entry_hasta = DateEntry(form, date_pattern="dd-mm-yy", width=14)
    else:
        entry_hasta = ctk.CTkEntry(form, placeholder_text="dd-mm-yy", width=140)
    entry_hasta.grid(row=5, column=1, sticky="w")

    # ================= BLOQUE: Día Administrativo (Horas) =================
    frame_horas = ctk.CTkFrame(form)
    frame_horas.grid(row=6, column=0, columnspan=3, sticky="w", padx=0, pady=(0, 0))

    lbl_cant = ctk.CTkLabel(frame_horas, text="Cantidad de horas:", width=LABEL_W, anchor="e")
    lbl_cant.grid(row=0, column=0, sticky="e", padx=LABEL_PADX)
    horas_values = [str(i) for i in range(1, 11)]
    cmb_cantidad_horas = ctk.CTkOptionMenu(frame_horas, values=horas_values, width=100)
    cmb_cantidad_horas.set(horas_values[0])
    cmb_cantidad_horas.grid(row=0, column=1, sticky="w")

    lbl_inicio = ctk.CTkLabel(frame_horas, text="Hora inicio (HH:MM):", width=LABEL_W, anchor="e")
    lbl_inicio.grid(row=1, column=0, sticky="e", padx=LABEL_PADX, pady=(8, 0))
    entry_hora_inicio = ctk.CTkEntry(frame_horas, placeholder_text="HH:MM", width=120)
    entry_hora_inicio.grid(row=1, column=1, sticky="w", pady=(8, 0))

    lbl_fin = ctk.CTkLabel(frame_horas, text="Hora término (HH:MM):", width=LABEL_W, anchor="e")
    lbl_fin.grid(row=2, column=0, sticky="e", padx=LABEL_PADX, pady=(8, 0))
    entry_hora_fin = ctk.CTkEntry(frame_horas, placeholder_text="HH:MM", width=120)
    entry_hora_fin.grid(row=2, column=1, sticky="w", pady=(8, 0))
    entry_hora_fin.configure(state="disabled")

    def _set_estado_horas(enabled: bool):
        estado = "normal" if enabled else "disabled"
        cmb_cantidad_horas.configure(state=estado)
        entry_hora_inicio.configure(state=estado)
        entry_hora_fin.configure(state="disabled")
        if not enabled:
            cmb_cantidad_horas.set(horas_values[0])
            entry_hora_inicio.delete(0, "end")
            entry_hora_fin.configure(state="normal")
            entry_hora_fin.delete(0, "end")
            entry_hora_fin.configure(state="disabled")

    _set_estado_horas(False)

    # ---- Observación
    ctk.CTkLabel(form, text="Observación:", width=LABEL_W, anchor="e").grid(row=7, column=0, sticky="e", padx=LABEL_PADX)
    entry_obs = ctk.CTkEntry(form, placeholder_text="Motivo u observación (opcional)", width=480)
    entry_obs.grid(row=7, column=1, columnspan=2, sticky="w")

    # ---- Nota aclaratoria
    nota = ctk.CTkLabel(
        box,
        text=("⚠️ Esta solicitud NO es oficial. Solo reserva la ausencia y "
              "queda sujeta a autorización. El trámite formal debe realizarse "
              "en el área administrativa completando el documento físico."),
        text_color="#ffb74d",
        font=("Arial", 13),
        wraplength=700,
        justify="center",
    )
    nota.grid(row=2, column=0, pady=(10, 6), sticky="n")

    # ---- Botones
    fila_botones = ctk.CTkFrame(box, fg_color="transparent")
    fila_botones.grid(row=3, column=0, pady=(6, 0))
    btn_enviar = ctk.CTkButton(fila_botones, text="Enviar Solicitud")
    btn_limpiar = ctk.CTkButton(fila_botones, text="Limpiar Todo")
    btn_cancelar = ctk.CTkButton(fila_botones, text="Cancelar (Volver)")
    btn_enviar.pack(side="left", padx=6)
    btn_limpiar.pack(side="left", padx=6)
    btn_cancelar.pack(side="left", padx=6)

    # ---------- Buscador ----------
    def limpiar_placeholder(_e=None):
        if combo_nombre.get() == "Buscar por Nombre":
            combo_nombre.set("")
    def restaurar_placeholder(_e=None):
        if combo_nombre.get() == "":
            combo_nombre.set("Buscar por Nombre")
    combo_nombre.bind("<FocusIn>", limpiar_placeholder)
    combo_nombre.bind("<FocusOut>", restaurar_placeholder)

    def mostrar_sugerencias(_e=None):
        t = combo_nombre.get().lower()
        combo_nombre.configure(values=(primeros_10 if not t or t == "buscar por nombre" else combo_nombre.cget("values")))
    combo_nombre.bind("<FocusIn>", mostrar_sugerencias)

    def autocompletar_nombres(_e=None):
        t = combo_nombre.get().lower()
        if not t or t == "buscar por nombre":
            combo_nombre.configure(values=primeros_10)
        else:
            filtrados = [n for n in lista_nombres if t in n.lower()]
            combo_nombre.configure(values=(filtrados[:10] if filtrados else ["No encontrado"]));
    combo_nombre.bind("<KeyRelease>", autocompletar_nombres)

    def _rellenar_por_rut(rut: str):
        if not rut:
            return
        entry_rut.delete(0, "end")
        entry_rut.insert(0, rut)
        cargo = obtener_cargo_por_rut(rut)
        entry_cargo.configure(state="normal")
        entry_cargo.delete(0, "end")
        if cargo:
            entry_cargo.insert(0, cargo)
        entry_cargo.configure(state="disabled")

    def buscar_por_nombre(_e=None):
        nombre = combo_nombre.get()
        rut = dict_nombre_rut.get(nombre, "")
        if rut:
            _rellenar_por_rut(rut)
        else:
            messagebox.showwarning("Nombre no válido", "Selecciona un nombre válido de la lista.")
    combo_nombre.bind("<Return>", buscar_por_nombre)
    btn_buscar.configure(command=buscar_por_nombre)

    # ---------- Fechas y autocalculado ----------
    def _get_date_from_widget(widget):
        if HAS_TKCAL:
            return widget.get_date()
        s = widget.get().strip()
        try:
            return datetime.datetime.strptime(s, "%d-%m-%y").date()
        except ValueError:
            return None

    def _set_widget_date(widget, date_obj: datetime.date):
        if HAS_TKCAL:
            widget.set_date(date_obj)
        else:
            s = date_obj.strftime("%d-%m-%y")
            widget.delete(0, "end")
            widget.insert(0, s)

    def _is_habiles_text(tipo_texto: str) -> bool:
        t = tipo_texto.lower()
        return "hábil" in t or "habil" in t

    def _add_days_inclusive_local(start_date: datetime.date, days: int, habiles: bool) -> datetime.date:
        if days <= 1:
            return start_date
        if not habiles:
            return start_date + datetime.timedelta(days=days - 1)
        d = start_date
        count = 0
        while count < days:
            if d.weekday() < 5:
                count += 1
                if count == days:
                    break
            d += datetime.timedelta(days=1)
        return d

    def _auto_hasta(*_):
        tipo = cmb_tipo.get().strip()
        d = _get_date_from_widget(entry_desde)
        if not tipo or tipo == opciones_permiso[0] or d is None:
            return
        if tipo in dias_por_permiso:
            n = dias_por_permiso[tipo]
            h = _is_habiles_text(tipo)
            end_date = _add_days_inclusive_local(d, n, h)
        else:
            end_date = d
        _set_widget_date(entry_hasta, end_date)

    if HAS_TKCAL:
        entry_desde.bind("<<DateEntrySelected>>", _auto_hasta)
    else:
        entry_desde.bind("<FocusOut>", _auto_hasta)

    # ================= Cálculo automático de hora término =================
    def _parse_hhmm(txt: str) -> datetime.time | None:
        txt = (txt or "").strip()
        try:
            return datetime.datetime.strptime(txt, "%H:%M").time()
        except Exception:
            return None

    def _calc_hora_fin():
        """Calcula y coloca la hora término según inicio + cantidad horas."""
        if entry_hora_inicio.cget("state") == "disabled":
            return
        h_ini = _parse_hhmm(entry_hora_inicio.get())
        if not h_ini:
            entry_hora_fin.configure(state="normal")
            entry_hora_fin.delete(0, "end")
            entry_hora_fin.configure(state="disabled")
            return
        try:
            cant = int(cmb_cantidad_horas.get())
        except Exception:
            cant = 1
        base_date = datetime.date(2000, 1, 1)
        dt_ini = datetime.datetime.combine(base_date, h_ini)
        dt_fin = dt_ini + datetime.timedelta(hours=cant)
        hhmm = dt_fin.strftime("%H:%M")
        entry_hora_fin.configure(state="normal")
        entry_hora_fin.delete(0, "end")
        entry_hora_fin.insert(0, hhmm)
        entry_hora_fin.configure(state="disabled")

    cmb_cantidad_horas.configure(command=lambda *_: _calc_hora_fin())
    entry_hora_inicio.bind("<FocusOut>", lambda *_: _calc_hora_fin())
    entry_hora_inicio.bind("<KeyRelease>", lambda *_: _calc_hora_fin())

    # ---------- Habilitar/Deshabilitar bloque horas según tipo (sin ocultar) ----------
    def _on_tipo_change(*_):
        tipo_sel = cmb_tipo.get().strip()
        _auto_hasta()
        if tipo_sel == "Día Administrativo (Horas)":
            entry_hora_inicio.configure(state="normal")
            cmb_cantidad_horas.configure(state="normal")
        else:
            _set_estado_horas(False)

    cmb_tipo.configure(command=lambda *_: _on_tipo_change())
    _on_tipo_change()

    # ---------- Validación ----------
    def validar_llenado():
        faltantes = []
        nombre_sel = combo_nombre.get()
        rut = entry_rut.get().strip()
        if not nombre_sel or nombre_sel in ("", "Buscar por Nombre", "No encontrado"):
            faltantes.append("Nombre")
        if not rut:
            faltantes.append("RUT")
        tipo = cmb_tipo.get().strip()
        if not tipo or tipo == opciones_permiso[0]:
            faltantes.append("Tipo de permiso")
        d = _get_date_from_widget(entry_desde)
        h = _get_date_from_widget(entry_hasta)
        if d is None:
            faltantes.append("Fecha Desde (dd-mm-yy)")
        if h is None:
            faltantes.append("Fecha Hasta (dd-mm-yy)")
        if d and h and d > h:
            faltantes.append("Rango de fechas inválido (Desde > Hasta)")
        if tipo == "Día Administrativo (Horas)":
            txt_ini = entry_hora_inicio.get().strip()
            if not _parse_hhmm(txt_ini):
                faltantes.append("Hora inicio (HH:MM)")
            if not entry_hora_fin.get().strip():
                faltantes.append("Hora término (auto)")
        return faltantes

    # ---------- Enviar ----------
    def do_enviar():
        faltantes = validar_llenado()
        if faltantes:
            messagebox.showerror("Faltan campos", "Debes completar/corregir: " + ", ".join(faltantes))
            return

        nombre_sel = combo_nombre.get()
        rut_sel = entry_rut.get().strip()
        cargo = entry_cargo.get().strip()
        tipo = cmb_tipo.get().strip()
        d = _get_date_from_widget(entry_desde)
        h = _get_date_from_widget(entry_hasta)
        obs = entry_obs.get().strip()

        # Si es por horas, anexo detalle
        if tipo == "Día Administrativo (Horas)":
            try:
                cant = int(cmb_cantidad_horas.get())
            except Exception:
                cant = 1
            hora_ini_txt = entry_hora_inicio.get().strip()
            _calc_hora_fin()
            hora_fin_txt = entry_hora_fin.get().strip()
            detalle_horas = f" [Horas: {cant}, Inicio: {hora_ini_txt}, Término: {hora_fin_txt}]"
            obs = (obs + " " + detalle_horas).strip() if obs else detalle_horas.strip()

        # Folio
        try:
            folio = get_next_folio()
        except Exception as e:
            messagebox.showerror("Folio", f"No fue posible generar el folio correlativo:\n{e}")
            return

        # Generar PDF (si se puede)
        pdf_path = None
        try:
            pdf_path = _build_pdf_solicitud(
                folio=folio, rut=rut_sel, nombre=nombre_sel, cargo=cargo,
                tipo=tipo, desde=d, hasta=h, observacion=obs
            )
        except Exception:
            traceback.print_exc()
            pdf_path = None

        # Guardar en BD (usa la ruta real o 'MANUAL' si no hubo PDF)
        try:
            guardar_solicitud_en_bd(
                folio, rut_sel, nombre_sel, tipo,
                d.strftime("%Y-%m-%d"), h.strftime("%Y-%m-%d"),
                obs, pdf_path or "MANUAL"
            )
        except Exception as e:
            messagebox.showerror("Base de Datos", f"No se pudo guardar la solicitud:\n{e}")
            return

        # Preparar correo
        aviso = ("IMPORTANTE: Esta solicitud NO es oficial. Solo reserva la ausencia y "
                 "queda sujeta a autorización. El trámite formal debe realizarse "
                 "en el área administrativa completando el documento físico.")
        asunto = f"Reserva de Solicitud - Folio {folio:06d} - {nombre_sel}"
        cuerpo_txt = (
            f"{aviso}\n\n"
            f"Folio: {folio:06d}\n"
            f"Nombre: {nombre_sel}\n"
            f"RUT: {rut_sel}\n"
            f"Cargo/Profesión: {cargo or '-'}\n"
            f"Tipo: {tipo}\n"
            f"Desde: {d.strftime('%d-%m-%y')}  Hasta: {h.strftime('%d-%m-%y')}\n"
            f"Observación: {obs or '-'}\n"
            f"Generado por BioAccess."
        )
        cuerpo_html = _html_email_solicitud(
            folio=folio,
            nombre=nombre_sel,
            rut=rut_sel,
            cargo=cargo or "-",
            tipo=tipo,
            desde=d.strftime('%d-%m-%Y'),
            hasta=h.strftime('%d-%m-%Y'),
            obs=obs or "-"
        )

        # Destinatarios: fijos + CC al funcionario si existe correo
        to_list = list(DESTINATARIOS)
        cc_list = []
        correo_func = obtener_correo_por_rut(rut_sel)
        if correo_func and "@" in correo_func:
            cc_list.append(correo_func)

        try:
            _smtp_send(
                to_list=to_list,
                cc_list=cc_list,
                subject=asunto,
                body_text=cuerpo_txt,
                html_body=cuerpo_html,
                attachment_path=pdf_path
            )
            msg_ok = f"Reserva registrada (Folio {folio:06d}).\nSe envió el aviso por correo."
            if pdf_path is None:
                msg_ok += "\n(Nota: No se pudo adjuntar PDF; ver consola para detalle)."
            messagebox.showinfo("Listo", msg_ok)
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("Envío", f"No se pudo enviar el correo:\n{e}")

    btn_enviar.configure(command=do_enviar)

    # ---------- Limpiar / Cancelar ----------
    def do_limpiar():
        combo_nombre.set("Buscar por Nombre")
        entry_rut.delete(0, "end")
        entry_cargo.configure(state="normal"); entry_cargo.delete(0, "end"); entry_cargo.configure(state="disabled")
        cmb_tipo.set(opciones_permiso[0])
        if HAS_TKCAL:
            entry_desde.set_date(datetime.date.today())
            entry_hasta.set_date(datetime.date.today())
        else:
            entry_desde.delete(0, "end")
            entry_hasta.delete(0, "end")
        entry_obs.delete(0, "end")
        _set_estado_horas(False)
    btn_limpiar.configure(command=do_limpiar)

    def do_cancelar():
        if on_volver:
            on_volver()
        else:
            for w in frame_padre.winfo_children():
                w.destroy()
    btn_cancelar.configure(command=do_cancelar)

    if not HAS_TKCAL:
        messagebox.showinfo("Calendario", "Para desplegar calendario en las fechas, instala tkcalendar:\n\npip install tkcalendar")
