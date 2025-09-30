# resumen_dia.py
import os
import ssl
import smtplib
import mimetypes
import tempfile
import customtkinter as ctk
import tkinter as tk
import tkinter.ttk as ttk
import sqlite3
from datetime import datetime, date, timedelta
from email.message import EmailMessage
from xml.sax.saxutils import escape

DB = "reloj_control.db"
DAYS_ES = ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado', 'Domingo']

# -------- Reglas ----------
TOL_INGRESO_MIN = 5        # Ingreso OK hasta +5'
TOL_SALIDA_MAX_MIN = 40    # Salida OK desde hora esperada hasta +40'

# ======== PDF (reportlab) ========
try:
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib import colors
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_LEFT
    HAS_PDF = True
except Exception:
    HAS_PDF = False

# ======== SMTP fallback (igual que panel) ========
SMTP_FALLBACK = {
    "host": "mail.bioaccess.cl",
    "port": 465,
    "user": "documentos_bd@bioaccess.cl",
    "password": "documentos@2025",
    "use_tls": False,
    "use_ssl": True,
    "remitente": "documentos_bd@bioaccess.cl",
}

def _smtp_load_config():
    cfg = {}
    try:
        con = sqlite3.connect(DB); cur = con.cursor()
        try:
            cur.execute("SELECT host, port, user, password, use_tls, use_ssl, remitente FROM smtp_config LIMIT 1")
            row = cur.fetchone()
            if row:
                cfg = {
                    "host": row[0],
                    "port": int(row[1]) if row[1] is not None else 0,
                    "user": row[2],
                    "password": row[3],
                    "use_tls": str(row[4]).lower() in ("1","true","t","yes","y"),
                    "use_ssl": str(row[5]).lower() in ("1","true","t","yes","y"),
                    "remitente": row[6] or row[2]
                }
                con.close(); return cfg
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
                    "use_tls": str(mapa.get("use_tls", "true")).lower() in ("1","true","t","yes","y"),
                    "use_ssl": str(mapa.get("use_ssl", "false")).lower() in ("1","true","t","yes","y"),
                    "remitente": mapa.get("remitente", mapa.get("user"))
                }
                con.close(); return cfg
        except Exception:
            pass
        con.close()
    except Exception:
        pass
    return None

def _smtp_send(to_list, cc_list, subject, body_text, html_body=None, attachment_path=None):
    cfg = _smtp_load_config() or SMTP_FALLBACK
    msg = EmailMessage()
    remitente = cfg.get("remitente") or cfg.get("user")
    msg["From"] = remitente
    msg["To"] = ", ".join(to_list) if to_list else ""
    if cc_list: msg["Cc"] = ", ".join(cc_list)
    msg["Subject"] = subject or "Resumen Diario"
    msg.set_content(body_text or "")
    if html_body: msg.add_alternative(html_body, subtype="html")
    if attachment_path and os.path.exists(attachment_path):
        with open(attachment_path, "rb") as f:
            data = f.read()
        mime, _ = mimetypes.guess_type(attachment_path)
        maintype, subtype = (mime.split("/",1) if mime else ("application","octet-stream"))
        msg.add_attachment(data, maintype=maintype, subtype=subtype, filename=os.path.basename(attachment_path))

    host = cfg["host"]; port = cfg.get("port") or (465 if cfg.get("use_ssl") else 587)
    user = cfg.get("user"); password = cfg.get("password")
    use_ssl = cfg.get("use_ssl"); use_tls = cfg.get("use_tls")

    if use_ssl:
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL(host, port, context=context, timeout=30) as server:
            if user and password: server.login(user, password)
            server.send_message(msg)
    else:
        with smtplib.SMTP(host, port, timeout=30) as server:
            server.ehlo()
            if use_tls:
                context = ssl.create_default_context()
                server.starttls(context=context); server.ehlo()
            if user and password: server.login(user, password)
            server.send_message(msg)

# ================= Helpers =================
def _date_to_iso(d: date) -> str: return d.strftime("%Y-%m-%d")
def _parse_ddmmyyyy(s: str) -> date | None:
    try: return datetime.strptime(s.strip(), "%d/%m/%Y").date()
    except Exception: return None
def _format_ddmmyyyy(d: date) -> str: return d.strftime("%d/%m/%Y")
def _parse_hora(h: str) -> datetime | None:
    if not h: return None
    for fmt in ("%H:%M:%S", "%H:%M"):
        try: return datetime.strptime(h.strip(), fmt)
        except ValueError: continue
    return None
def _diff_minutes(a: str, b: str) -> int | None:
    ta, tb = _parse_hora(a), _parse_hora(b)
    if not ta or not tb: return None
    return int((ta - tb).total_seconds() // 60)  # + si a>b

# ================= Consultas =================
def _fetch_ingresos(fecha_iso: str):
    con = sqlite3.connect(DB); cur = con.cursor()
    cur.execute("""
        WITH rows AS (
          SELECT
            COALESCE(NULLIF(hora_ingreso, ''),
                     CASE WHEN lower(IFNULL(tipo,''))='ingreso' THEN IFNULL(hora,'') ELSE '' END
            ) AS h_real,
            IFNULL(hora_ingreso,'') AS h_ing_col,
            IFNULL(rut,'')  AS rut,
            IFNULL(nombre,'') AS nombre
          FROM registros
          WHERE DATE(fecha)=?
        )
        SELECT h_real, rut, nombre, h_ing_col
        FROM rows
        WHERE TRIM(IFNULL(h_real,'')) <> ''
        ORDER BY time(h_real) ASC
    """, (fecha_iso,))
    rows = cur.fetchall(); con.close()
    return rows

def _fetch_salidas(fecha_iso: str):
    con = sqlite3.connect(DB); cur = con.cursor()
    cur.execute("""
        WITH rows AS (
          SELECT
            COALESCE(NULLIF(hora_salida, ''),
                     CASE WHEN lower(IFNULL(tipo,''))='salida' THEN IFNULL(hora,'') ELSE '' END
            ) AS h_real,
            IFNULL(hora_ingreso,'') AS h_ing_col,
            IFNULL(rut,'')  AS rut,
            IFNULL(nombre,'') AS nombre
          FROM registros
          WHERE DATE(fecha)=?
        )
        SELECT h_real, rut, nombre, h_ing_col
        FROM rows
        WHERE TRIM(IFNULL(h_real,'')) <> ''
        ORDER BY time(h_real) ASC
    """, (fecha_iso,))
    rows = cur.fetchall(); con.close()
    return rows

def _fetch_observaciones(fecha_iso: str):
    con = sqlite3.connect(DB); cur = con.cursor()
    cur.execute("""
        WITH base AS (
          SELECT
            IFNULL(observacion,'') AS obs,
            IFNULL(rut,'')        AS rut,
            IFNULL(nombre,'')     AS nom,
            IFNULL(hora_ingreso,'') AS hi,
            IFNULL(hora_salida,'')  AS hs,
            IFNULL(hora,'')         AS hx,
            IFNULL(tipo,'')         AS tp
          FROM registros
          WHERE DATE(fecha)=?
        ),
        rows AS (
          SELECT
            COALESCE(NULLIF(hi,''), NULLIF(hs,''), hx, '') AS h,
            CASE
              WHEN lower(obs) LIKE '%permiso%' OR lower(tp) LIKE '%permiso%' THEN 'Permiso'
              ELSE 'Observación'
            END AS tipo,
            rut, nom,
            CASE
              WHEN TRIM(hi)<>'' THEN 'Ingreso'
              WHEN TRIM(hs)<>'' THEN 'Salida'
              ELSE 'Otro'
            END AS evento,
            lower(obs) AS obs_l,
            obs AS obs_txt
          FROM base
          WHERE TRIM(obs) <> ''
        )
        SELECT h, tipo, rut, nom, evento, obs_l, obs_txt
        FROM rows
        ORDER BY CASE WHEN TRIM(h)<>'' THEN time(h) ELSE time('23:59:59') END ASC
    """, (fecha_iso,))
    rows = cur.fetchall(); con.close()
    return rows

# ================= Horarios esperados =================
def _fetch_horarios_dia(rut: str, dia_es: str):
    con = sqlite3.connect(DB); cur = con.cursor()
    cur.execute("""
        SELECT TRIM(IFNULL(hora_entrada,'')) AS he, TRIM(IFNULL(hora_salida,'')) AS hs
        FROM horarios
        WHERE rut=? AND dia=?
          AND TRIM(IFNULL(hora_entrada,'')) <> ''
          AND TRIM(IFNULL(hora_salida,''))  <> ''
        ORDER BY time(he) ASC
    """, (rut, dia_es))
    rows = cur.fetchall(); con.close()
    return rows

def _ultima_salida_programada(rut: str, fecha: date) -> str:
    dia = DAYS_ES[fecha.weekday()]
    bloques = _fetch_horarios_dia(rut, dia)
    if not bloques: return ""
    ult, best_dt = None, None
    for _he, hs in bloques:
        ths = _parse_hora(hs)
        if not ths: continue
        if best_dt is None or ths > best_dt:
            best_dt, ult = ths, hs
    return ult or ""

def _expected_ingreso(rut: str, fecha: date, hora_real: str) -> str:
    dia = DAYS_ES[fecha.weekday()]
    bloques = _fetch_horarios_dia(rut, dia)
    if not bloques: return ""
    t = _parse_hora(hora_real) or _parse_hora(bloques[0][0])
    best_he, best_diff = None, None
    for he, _ in bloques:
        th = _parse_hora(he)
        if not th: continue
        diff = abs((t - th).total_seconds())
        if best_diff is None or diff < best_diff:
            best_he, best_diff = he, diff
    return best_he or bloques[0][0]

def _expected_salida(rut: str, fecha: date, _hora_real: str, _hora_ingreso_col: str | None = None) -> str:
    return _ultima_salida_programada(rut, fecha)

# ================= Estilos Treeview =================
def _setup_tree_styles(parent, base_name="Resumen"):
    style = ttk.Style(parent)
    try: style.theme_use("clam")
    except tk.TclError: pass
    tv = f"{base_name}.Treeview"; hd = f"{base_name}.Treeview.Heading"; sb = f"{base_name}.Vertical.TScrollbar"
    style.configure(tv, background="#0b1220", fieldbackground="#0b1220", foreground="#E5E7EB",
                    rowheight=26, borderwidth=0, relief="flat")
    style.map(tv, background=[("selected", "#0ea5e9")], foreground=[("selected", "white")])
    style.configure(hd, background="#1f2937", foreground="#E5E7EB", relief="flat")
    style.map(hd, background=[("active", "#374151")])
    style.configure(sb, background="#1f2937", troughcolor="#0b1220", bordercolor="#0b1220",
                    lightcolor="#0b1220", darkcolor="#0b1220", arrowcolor="#E5E7EB")

# ================= PDF del día =================
def _pdf_resumen_dia(path, fecha_dt: date, ingresos_pdf, salidas_pdf, observs_pdf, resumen_dict):
    if not HAS_PDF: raise RuntimeError("ReportLab no está instalado.")
    doc = SimpleDocTemplate(path, pagesize=landscape(A4),
                            rightMargin=14, leftMargin=14, topMargin=12, bottomMargin=12)
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="Header", parent=styles["Title"], fontSize=18, leading=22, alignment=TA_LEFT))
    styles.add(ParagraphStyle(name="Small", parent=styles["Normal"], fontSize=10, leading=12))
    styles.add(ParagraphStyle(name="Cell", parent=styles["Normal"], fontSize=9, leading=11))
    # Estilo especial para Detalle con envoltura agresiva
    obs_cell_style = ParagraphStyle(
        name="ObsCell",
        parent=styles["Normal"],
        fontSize=9,
        leading=11,
        wordWrap="CJK",  # permite quebrar palabras largas
    )

    story = []
    story.append(Paragraph("Reporte de Asistencia – Resumen Diario", styles["Header"]))
    story.append(Paragraph(
        f"Período: {fecha_dt.strftime('%d/%m/%Y')} &nbsp;&nbsp;|&nbsp;&nbsp; "
        f"Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}", styles["Small"]))
    story.append(Spacer(1, 6))

    intro_rows = [
        ["Ingresos del día", resumen_dict.get("ing_total", 0)],
        ["Dentro de rango (+≤5')", resumen_dict.get("ing_ok", 0)],
        ["Fuera de rango", resumen_dict.get("ing_bad", 0)],
        ["Con observaciones", resumen_dict.get("obs_distinct", 0)],
        ["Con cometidos", resumen_dict.get("cometidos", 0)],
        ["Administrativos", resumen_dict.get("administrativos", 0)],
        ["Otros", resumen_dict.get("otros", 0)],
    ]
    tbl_intro = Table([["Resumen del día", ""]] + intro_rows, colWidths=[160, 60])
    tbl_intro.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#e2e8f0")),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("GRID", (0,0), (-1,-1), 0.3, colors.grey),
        ("ALIGN", (1,1), (1,-1), "RIGHT"),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("FONTSIZE", (0,0), (-1,-1), 10),
        ("LEFTPADDING", (0,0), (-1,-1), 6),
        ("RIGHTPADDING", (0,0), (-1,-1), 6),
    ]))
    story.append(tbl_intro); story.append(Spacer(1, 8))

    def _mk_table(title, headers, rows, status_idx, col_widths=None, wrap_cols=None, wrap_style=None):
        story.append(Paragraph(f"<b>{title}</b>", styles["Small"]))
        data = [list(headers)]
        # envolver columnas indicadas con Paragraph para permitir salto de línea
        for r in rows:
            rr = list(r)
            if wrap_cols:
                for ci in wrap_cols:
                    try:
                        txt = rr[ci]
                        # Escapar HTML y respetar saltos de línea
                        txt = escape(str(txt or "")).replace("\n", "<br/>")
                        rr[ci] = Paragraph(txt, wrap_style or styles["Cell"])
                    except Exception:
                        rr[ci] = Paragraph("", wrap_style or styles["Cell"])
            data.append(rr)
        tbl = Table(data, repeatRows=1, colWidths=col_widths)
        styl = [
            ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#e2e8f0")),
            ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
            ("GRID", (0,0), (-1,-1), 0.25, colors.grey),
            ("VALIGN", (0,1), (-1,-1), "TOP"),     # importante para celdas con múltiples líneas
            ("FONTSIZE", (0,0), (-1,-1), 9),
            ("LEFTPADDING", (0,0), (-1,-1), 4),
            ("RIGHTPADDING", (0,0), (-1,-1), 4),
        ]
        for i, row in enumerate(rows, start=1):
            val = (row[status_idx] or "").strip()
            if val == "✓":
                styl.append(("TEXTCOLOR", (status_idx, i), (status_idx, i), colors.HexColor("#16a34a")))
            elif val == "✗":
                styl.append(("TEXTCOLOR", (status_idx, i), (status_idx, i), colors.HexColor("#ef4444")))
        tbl.setStyle(TableStyle(styl))
        story.append(tbl); story.append(Spacer(1, 8))

    # Ingresos / Salidas
    _mk_table("Ingresos",
              ["Hora", "Hora esperada", "RUT", "Nombre", "Estado"],
              ingresos_pdf, status_idx=4)

    _mk_table("Salidas",
              ["Hora", "Hora esperada", "RUT", "Nombre", "Estado"],
              salidas_pdf, status_idx=4)

    # Observaciones con "Detalle" (envuelto) y ancho de columna grande
    _mk_table("Observaciones / Permisos",
              ["Hora", "Evento", "RUT", "Nombre", "Detalle", "Estado"],
              observs_pdf,
              status_idx=5,
              col_widths=[50, 50, 70, 230, 330, 50],
              wrap_cols=[4],
              wrap_style=obs_cell_style)

    doc.build(story)

# ================= Ventana detalle (persona) =================
def _open_person_detail(parent, fecha: date, rut: str, nombre: str):
    fecha_iso = _date_to_iso(fecha)
    con = sqlite3.connect(DB); cur = con.cursor()

    cur.execute("""
        SELECT COALESCE(NULLIF(hora_ingreso,''), CASE WHEN lower(IFNULL(tipo,''))='ingreso' THEN IFNULL(hora,'') ELSE '' END)
        FROM registros
        WHERE DATE(fecha)=? AND rut=? AND TRIM(COALESCE(hora_ingreso, CASE WHEN lower(IFNULL(tipo,''))='ingreso' THEN IFNULL(hora,'') ELSE '' END))<>''
        ORDER BY time(COALESCE(hora_ingreso, hora)) ASC LIMIT 1
    """, (fecha_iso, rut))
    row_in = cur.fetchone()
    h_ing_real = row_in[0] if row_in and row_in[0] else ""

    cur.execute("""
        SELECT COALESCE(NULLIF(hora_salida,''), CASE WHEN lower(IFNULL(tipo,''))='salida' THEN IFNULL(hora,'') ELSE '' END)
        FROM registros
        WHERE DATE(fecha)=? AND rut=? AND TRIM(COALESCE(hora_salida, CASE WHEN lower(IFNULL(tipo,''))='salida' THEN IFNULL(hora,'') ELSE '' END))<>''
        ORDER BY time(COALESCE(hora_salida, hora)) DESC LIMIT 1
    """, (fecha_iso, rut))
    row_out = cur.fetchone()
    h_sal_real = row_out[0] if row_out and row_out[0] else ""

    cur.execute("""
        SELECT COALESCE(NULLIF(hora_ingreso,''), NULLIF(hora_salida,''), hora, '') AS h,
               COALESCE(observacion, '') AS obs,
               CASE WHEN TRIM(hora_ingreso)<>'' THEN 'Ingreso'
                    WHEN TRIM(hora_salida)<>'' THEN 'Salida'
                    ELSE 'Otro' END AS tipo
        FROM registros
        WHERE DATE(fecha)=? AND rut=? AND TRIM(COALESCE(observacion,''))<>''
        ORDER BY CASE WHEN TRIM(h)<>'' THEN time(h) ELSE time('23:59:59') END
    """, (fecha_iso, rut))
    obs_rows = cur.fetchall(); con.close()

    h_ing_esp = _expected_ingreso(rut, fecha, h_ing_real or "08:00")
    h_sal_esp = _expected_salida(rut, fecha, h_sal_real or "18:00", None)
    d_ing = _diff_minutes(h_ing_real or "", h_ing_esp or "")
    d_sal = _diff_minutes(h_sal_real or "", h_sal_esp or "")

    def ok_ing(delta): return (delta is not None) and (delta <= TOL_INGRESO_MIN)
    def ok_sal(delta): return (delta is not None) and (0 <= delta <= TOL_SALIDA_MAX_MIN)

    def txt_ing(delta):
        if delta is None: return "Sin datos"
        if delta <= TOL_INGRESO_MIN:
            atraso = max(0, delta)
            return "En rango" if atraso == 0 else f"En rango (+{atraso} min)"
        extra = delta - TOL_INGRESO_MIN
        return f"Tardío (+{extra} min sobre tolerancia)"

    def txt_sal(delta):
        if delta is None: return "Sin datos"
        if delta < 0: return f"Salida antes de hora (−{abs(delta)} min)"
        if delta <= TOL_SALIDA_MAX_MIN:
            return "En rango" if delta == 0 else f"En rango (+{delta} min)"
        return f"Demora en marcar (+{delta} min > {TOL_SALIDA_MAX_MIN}')"

    win = ctk.CTkToplevel(parent); win.title(f"Detalle – {nombre}")
    try: win.resizable(False, False); win.transient(parent); win.grab_set()
    except Exception: pass

    head = ctk.CTkFrame(win, fg_color="transparent"); head.pack(fill="x", padx=16, pady=12)
    ctk.CTkLabel(head, text=f"{nombre}", font=("Segoe UI", 16, "bold")).pack(anchor="w")
    ctk.CTkLabel(head, text=f"{DAYS_ES[fecha.weekday()]} {fecha.strftime('%d/%m/%Y')}", text_color="#9ca3af").pack(anchor="w")

    grid = ctk.CTkFrame(win); grid.pack(fill="both", expand=True, padx=16, pady=(0,12))
    grid.grid_columnconfigure(0, weight=1); grid.grid_columnconfigure(1, weight=1)

    def _card(col, titulo, h_real, h_esp, ok, text):
        card = ctk.CTkFrame(grid); card.grid(row=0, column=col, sticky="nsew", padx=(0,8) if col==0 else (8,0))
        ctk.CTkLabel(card, text=titulo, font=("Segoe UI", 14, "bold")).pack(anchor="w", padx=12, pady=(10,4))
        row = ctk.CTkFrame(card, fg_color="transparent"); row.pack(fill="x", padx=12)
        ctk.CTkLabel(row, text=f"Real: {h_real or '--:--'}").pack(side="left", padx=(0,12))
        ctk.CTkLabel(row, text=f"Esperada: {h_esp or '--:--'}").pack(side="left")
        bar = ctk.CTkProgressBar(card, height=14); bar.pack(fill="x", padx=12, pady=(8,2))
        bar.configure(progress_color=("#16a34a" if ok else "#ef4444")); bar.set(1.0 if ok else 0.5)
        ctk.CTkLabel(card, text=text, text_color=("#16a34a" if ok else "#ef4444")).pack(anchor="w", padx=12, pady=(0,10))

    _card(0, "Ingreso", h_ing_real, h_ing_esp, ok_ing(d_ing), txt_ing(d_ing))
    _card(1, "Salida",  h_sal_real, h_sal_esp, ok_sal(d_sal), txt_sal(d_sal))

    ctk.CTkLabel(win, text="Observaciones", font=("Segoe UI", 14, "bold")).pack(anchor="w", padx=16)
    box = getattr(ctk, "CTkTextbox", None)
    if box:
        tb = box(win, width=600, height=140); tb.pack(fill="both", expand=True, padx=16, pady=(4,10))
        for h, obs, tipo in obs_rows:
            tb.insert("end", f"[{tipo:7}] {h or '--:--'}  {obs}\n")
        tb.configure(state="disabled")
    else:
        fr = tk.Frame(win); fr.pack(fill="both", expand=True, padx=16, pady=(4,10))
        tb = tk.Text(fr, width=82, height=8); tb.pack()
        for h, obs, tipo in obs_rows:
            tb.insert("end", f"[{tipo:7}] {h or '--:--'}  {obs}\n")
        tb.configure(state="disabled")

    btns = ctk.CTkFrame(win, fg_color="transparent"); btns.pack(pady=10)
    ctk.CTkButton(btns, text="Cerrar", width=140, fg_color="#64748b", command=win.destroy).pack()

    try:
        parent.update_idletasks()
        w, h = 720, 460
        x = parent.winfo_x() + (parent.winfo_width() // 2) - (w // 2)
        y = parent.winfo_y() + (parent.winfo_height() // 2) - (h // 2)
        win.geometry(f"{w}x{h}+{max(x,0)}+{max(y,0)}")
    except Exception:
        pass

# ================= Vista principal =================
def construir_resumen_dia(frame_padre):
    state = {"fecha": date.today(), "filtro": "", "after_id": None}
    root = ctk.CTkFrame(frame_padre); root.pack(fill="both", expand=True)
    _setup_tree_styles(root)

    # Título y fecha
    ctk.CTkLabel(root, text="Resumen Diario", font=("Arial", 16, "bold")).pack(pady=(10, 2))
    lbl_dia = ctk.CTkLabel(root, text="", text_color="#9ca3af"); lbl_dia.pack(pady=(0, 4))

    # NAV centrado
    nav = ctk.CTkFrame(root, fg_color="transparent"); nav.pack(pady=(0, 10))
    btn_prev = ctk.CTkButton(nav, text="◀", width=36); btn_prev.pack(side="left", padx=4)
    entry_fecha = ctk.CTkEntry(nav, width=120, justify="center", placeholder_text="dd/mm/aaaa")
    entry_fecha.pack(side="left", padx=4)
    btn_hoy = ctk.CTkButton(nav, text="Hoy", width=60); btn_hoy.pack(side="left", padx=4)
    btn_next = ctk.CTkButton(nav, text="▶", width=36); btn_next.pack(side="left", padx=4)

    # Topbar (buscador + contadores + acciones)
    topbar = ctk.CTkFrame(root, fg_color="transparent"); topbar.pack(fill="x", padx=12, pady=(0, 6))
    left = ctk.CTkFrame(topbar, fg_color="transparent"); left.pack(side="left")
    ctk.CTkLabel(left, text="🔎", font=("Arial", 16)).pack(side="left", padx=(0,6))
    entry_filter = ctk.CTkEntry(left, placeholder_text="Filtrar por RUT / Nombre", width=360); entry_filter.pack(side="left")

    right = ctk.CTkFrame(topbar, fg_color="transparent"); right.pack(side="right")
    lbl_ing = ctk.CTkLabel(right, text="Ingresos: 0", text_color="#93c5fd")
    lbl_sal = ctk.CTkLabel(right, text="Salidas: 0",  text_color="#34d399")
    lbl_obs = ctk.CTkLabel(right, text="Obs/Permisos: 0", text_color="#fbbf24")
    btn_refresh = ctk.CTkButton(right, text="Actualizar", width=110)
    btn_pdf = ctk.CTkButton(right, text="Exportar PDF", width=120)
    btn_email = ctk.CTkButton(right, text="Enviar", width=90)
    for w in (lbl_ing, lbl_sal, lbl_obs, btn_refresh, btn_pdf, btn_email):
        w.pack(side="left", padx=(6 if w!=lbl_ing else 0,0))

    # Cuerpo (3 columnas)
    body = ctk.CTkFrame(root, fg_color="transparent"); body.pack(fill="both", expand=True, padx=12, pady=8)
    for col in range(3): body.grid_columnconfigure(col, weight=1)
    body.grid_rowconfigure(0, weight=1)

    # ---- INGRESOS ----
    frame_ing = ctk.CTkFrame(body); frame_ing.grid(row=0, column=0, sticky="nsew", padx=(0,6))
    ctk.CTkLabel(frame_ing, text="Ingresos", font=("Arial", 14, "bold")).pack(pady=(6,4), anchor="w", padx=8)
    cols_ing = ("Hora", "Hora esperada", "RUT", "Nombre")
    tv_ing = ttk.Treeview(frame_ing, columns=cols_ing, show="headings", style="Resumen.Treeview",
                          selectmode="browse", height=14)
    tv_ing.column("Hora", width=90, anchor="center")
    tv_ing.column("Hora esperada", width=120, anchor="center")
    tv_ing.column("RUT", width=130, anchor="w")
    tv_ing.column("Nombre", width=200, anchor="w")
    for c in cols_ing: tv_ing.heading(c, text=c)
    y_ing = ttk.Scrollbar(frame_ing, orient="vertical", command=tv_ing.yview, style="Resumen.Vertical.TScrollbar")
    tv_ing.configure(yscrollcommand=y_ing.set); tv_ing.pack(side="left", fill="both", expand=True, padx=(8,0), pady=(0,8))
    y_ing.pack(side="left", fill="y", pady=(0,8))

    # ---- SALIDAS ----
    frame_sal = ctk.CTkFrame(body); frame_sal.grid(row=0, column=1, sticky="nsew", padx=6)
    ctk.CTkLabel(frame_sal, text="Salidas", font=("Arial", 14, "bold")).pack(pady=(6,4), anchor="w", padx=8)
    cols_sal = ("Hora", "Hora esperada", "RUT", "Nombre")
    tv_sal = ttk.Treeview(frame_sal, columns=cols_sal, show="headings", style="Resumen.Treeview",
                          selectmode="browse", height=14)
    tv_sal.column("Hora", width=90, anchor="center")
    tv_sal.column("Hora esperada", width=120, anchor="center")
    tv_sal.column("RUT", width=130, anchor="w")
    tv_sal.column("Nombre", width=200, anchor="w")
    for c in cols_sal: tv_sal.heading(c, text=c)
    y_sal = ttk.Scrollbar(frame_sal, orient="vertical", command=tv_sal.yview, style="Resumen.Vertical.TScrollbar")
    tv_sal.configure(yscrollcommand=y_sal.set); tv_sal.pack(side="left", fill="both", expand=True, padx=(8,0), pady=(0,8))
    y_sal.pack(side="left", fill="y", pady=(0,8))

    # ---- OBS / PERMISOS ----
    frame_obs = ctk.CTkFrame(body); frame_obs.grid(row=0, column=2, sticky="nsew", padx=(6,0))
    ctk.CTkLabel(frame_obs, text="Observaciones / Permisos", font=("Arial", 14, "bold")).pack(pady=(6,4), anchor="w", padx=8)
    cols_obs = ("Hora", "Evento", "RUT", "Nombre", "Detalle")
    tv_obs = ttk.Treeview(frame_obs, columns=cols_obs, show="headings", style="Resumen.Treeview",
                          selectmode="browse", height=14)
    widths = {"Hora": 90, "Evento":110, "RUT":130, "Nombre":180, "Detalle":320}
    aligns = {"Hora":"center","Evento":"center","RUT":"w","Nombre":"w","Detalle":"w"}
    for c in cols_obs:
        tv_obs.column(c, width=widths[c], anchor=aligns[c]); tv_obs.heading(c, text=c)
    y_obs = ttk.Scrollbar(frame_obs, orient="vertical", command=tv_obs.yview, style="Resumen.Vertical.TScrollbar")
    tv_obs.configure(yscrollcommand=y_obs.set); tv_obs.pack(side="left", fill="both", expand=True, padx=(8,0), pady=(0,8))
    y_obs.pack(side="left", fill="y", pady=(0,8))

    # Tags colores
    for tv in (tv_ing, tv_sal, tv_obs):
        tv.tag_configure("even", background="#0f172a")
        tv.tag_configure("odd", background="#111827")
    tv_ing.tag_configure("ok", foreground="#16a34a")
    tv_ing.tag_configure("bad", foreground="#ef4444")
    tv_sal.tag_configure("ok", foreground="#16a34a")
    tv_sal.tag_configure("bad", foreground="#ef4444")

    # ---------- lógica -----------
    cache_rows = {"ing": [], "sal": [], "obs": []}
    last_pdf_payload = {"ing": [], "sal": [], "obs": [], "resume": {}}

    def _set_header_date():
        d = state["fecha"]
        lbl_dia.configure(text=f"{DAYS_ES[d.weekday()]} {_format_ddmmyyyy(d)}")
        entry_fecha.delete(0, "end"); entry_fecha.insert(0, _format_ddmmyyyy(d))

    def _clear_all():
        for tv in (tv_ing, tv_sal, tv_obs):
            for iid in tv.get_children(): tv.delete(iid)

    def _apply_filter():
        term = (state["filtro"] or "").strip().lower()
        d = state["fecha"]

        ing_disp = []
        ing_ok = ing_bad = 0
        for h_real, rut, nombre, h_ing_col in cache_rows["ing"]:
            h_esp = _expected_ingreso(rut, d, h_real)
            dm = _diff_minutes(h_real, h_esp)
            is_ok = (dm is not None) and (dm <= TOL_INGRESO_MIN)
            tag = "ok" if is_ok else "bad"
            if is_ok: ing_ok += 1
            else: ing_bad += 1
            row = (h_real, h_esp or "--:--", rut, nombre)
            ing_disp.append((row, tag))

        sal_disp = []
        sal_ok = sal_bad = 0
        for h_real, rut, nombre, h_ing_col in cache_rows["sal"]:
            h_esp = _expected_salida(rut, d, h_real, h_ing_col)
            dm = _diff_minutes(h_real, h_esp)
            is_ok = (dm is not None) and (0 <= dm <= TOL_SALIDA_MAX_MIN)
            tag = "ok" if is_ok else "bad"
            if is_ok: sal_ok += 1
            else: sal_bad += 1
            row = (h_real, h_esp or "--:--", rut, nombre)
            sal_disp.append((row, tag))

        obs_disp = [((h or "--:--", ev, r, n, txt), None) for (h, _t, r, n, ev, _ol, txt) in cache_rows["obs"]]
        obs_count = len(obs_disp)

        def match(values): return (term in " ".join([str(v or "") for v in values]).lower()) if term else True
        ing_f = [x for x in ing_disp if match(x[0])]
        sal_f = [x for x in sal_disp if match(x[0])]
        obs_f = [x for x in obs_disp if match(x[0])]

        _clear_all()
        for i, (vals, tag) in enumerate(ing_f):
            tv_ing.insert("", "end", values=vals, tags=(("even" if i%2==0 else "odd"), tag))
        for i, (vals, tag) in enumerate(sal_f):
            tv_sal.insert("", "end", values=vals, tags=(("even" if i%2==0 else "odd"), tag))
        for i, (vals, _tag) in enumerate(obs_f):
            tv_obs.insert("", "end", values=vals, tags=("even" if i%2==0 else "odd",))

        lbl_ing.configure(text=f"Ingresos: {len(ing_f)}")
        lbl_sal.configure(text=f"Salidas: {len(sal_f)}")
        lbl_obs.configure(text=f"Obs/Permisos: {len(obs_f)}")

        # payload para PDF
        last_pdf_payload["ing"] = [(a,b,c,d, "✓" if t=="ok" else "✗") for (a,b,c,d), t in ing_disp]
        last_pdf_payload["sal"] = [(a,b,c,d, "✓" if t=="ok" else "✗") for (a,b,c,d), t in sal_disp]

        # obs estado + texto
        obs_pdf = []
        cometidos = administrativos = otros = 0
        for (h, tipo, rut, nom, evento, ol, txt) in cache_rows["obs"]:
            estado = "—"
            if evento == "Ingreso":
                esp = _expected_ingreso(rut, d, h or "08:00")
                dm = _diff_minutes(h or "", esp or "")
                estado = "✓" if (dm is not None and dm <= TOL_INGRESO_MIN) else "✗"
            elif evento == "Salida":
                esp = _expected_salida(rut, d, h or "18:00", None)
                dm = _diff_minutes(h or "", esp or "")
                estado = "✓" if (dm is not None and 0 <= dm <= TOL_SALIDA_MAX_MIN) else "✗"
            obs_pdf.append((h or "--:--", evento, rut, nom, txt or "", estado))

            if "cometid" in ol: cometidos += 1
            elif "administr" in ol: administrativos += 1
            else: otros += 1
        last_pdf_payload["obs"] = obs_pdf
        last_pdf_payload["resume"] = {
            "ing_total": len(ing_disp),
            "ing_ok": ing_ok, "ing_bad": ing_bad,
            "obs_distinct": obs_count,
            "cometidos": cometidos, "administrativos": administrativos, "otros": otros,
        }

    def _load_day():
        fecha_iso = _date_to_iso(state["fecha"])
        cache_rows["ing"] = _fetch_ingresos(fecha_iso)
        cache_rows["sal"] = _fetch_salidas(fecha_iso)
        cache_rows["obs"] = _fetch_observaciones(fecha_iso)
        _apply_filter()

    # Handlers
    def _debounced_filter(_=None):
        if state["after_id"]:
            try: root.after_cancel(state["after_id"])
            except Exception: pass
            state["after_id"] = None
        def run():
            state["filtro"] = entry_filter.get(); _apply_filter()
        state["after_id"] = root.after(250, run)

    def _go_prev(): state["fecha"]=state["fecha"]-timedelta(days=1); _set_header_date(); _load_day()
    def _go_next(): state["fecha"]=state["fecha"]+timedelta(days=1); _set_header_date(); _load_day()
    def _go_today(): state["fecha"]=date.today(); _set_header_date(); _load_day()

    def _enter_date(_=None):
        d = _parse_ddmmyyyy(entry_fecha.get())
        if not d:
            entry_fecha.delete(0, "end"); entry_fecha.insert(0, _format_ddmmyyyy(state["fecha"])); return
        state["fecha"]=d; _set_header_date(); _load_day()

    def _on_pdf():
        if not HAS_PDF:
            tk.messagebox.showwarning("PDF", "Instala reportlab:  pip install reportlab")
            return
        try:
            carpeta = os.path.join(os.path.expanduser("~"), "Downloads")
            os.makedirs(carpeta, exist_ok=True)
            out = os.path.join(carpeta, f"resumen_{state['fecha'].strftime('%Y%m%d')}.pdf")
            _pdf_resumen_dia(out, state["fecha"], last_pdf_payload["ing"], last_pdf_payload["sal"],
                             last_pdf_payload["obs"], last_pdf_payload["resume"])
            tk.messagebox.showinfo("PDF", f"PDF guardado en Descargas:\n{os.path.basename(out)}")
            try: os.startfile(out)
            except Exception: pass
        except Exception as e:
            tk.messagebox.showerror("PDF", f"No se pudo generar el PDF:\n{e}")

    def _on_email():
        if not HAS_PDF:
            tk.messagebox.showwarning("PDF", "Instala reportlab:  pip install reportlab")
            return
        win = ctk.CTkToplevel(root); win.title("Enviar Resumen Diario")
        try: win.resizable(False, False); win.transient(root); win.grab_set()
        except Exception: pass
        cont = ctk.CTkFrame(win); cont.pack(padx=14, pady=14)
        ctk.CTkLabel(cont, text="Para:").grid(row=0, column=0, sticky="e", padx=(0,6))
        ent_to = ctk.CTkEntry(cont, width=360); ent_to.grid(row=0, column=1, pady=4); ent_to.insert(0, "destinatario@empresa.cl")
        ctk.CTkLabel(cont, text="Asunto:").grid(row=1, column=0, sticky="e", padx=(0,6))
        ent_sub = ctk.CTkEntry(cont, width=360); ent_sub.grid(row=1, column=1, pady=4)
        ent_sub.insert(0, f"Resumen Diario – {state['fecha'].strftime('%d/%m/%Y')}")
        ctk.CTkLabel(cont, text="Mensaje:").grid(row=2, column=0, sticky="ne", padx=(0,6))
        txt = getattr(ctk, "CTkTextbox", None)
        if txt:
            tb = txt(cont, width=360, height=120); tb.grid(row=2, column=1, pady=4)
            tb.insert("1.0", "Estimado(a):\n\nAdjunto el resumen diario de asistencia.\n\nSaludos.")
        else:
            tb = tk.Text(cont, width=48, height=7); tb.grid(row=2, column=1, pady=4)
            tb.insert("1.0", "Estimado(a):\n\nAdjunto el resumen diario de asistencia.\n\nSaludos.")

        def enviar():
            to = [p.strip() for p in ent_to.get().replace(",", ";").split(";") if p.strip()]
            if not to:
                tk.messagebox.showwarning("Enviar", "Ingresa al menos un destinatario."); return
            subject = ent_sub.get().strip() or "Resumen Diario"
            body = tb.get("1.0", "end").strip()

            tmp = tempfile.NamedTemporaryFile(prefix="resumen_", suffix=".pdf", delete=False)
            tmp.close()
            try:
                _pdf_resumen_dia(tmp.name, state["fecha"], last_pdf_payload["ing"], last_pdf_payload["sal"],
                                 last_pdf_payload["obs"], last_pdf_payload["resume"])
                _smtp_send(to, [], subject, body, html_body=None, attachment_path=tmp.name)
                tk.messagebox.showinfo("Enviar", "Correo enviado correctamente.")
                win.destroy()
            except Exception as e:
                tk.messagebox.showerror("Enviar", f"No fue posible enviar el correo:\n{e}")
            finally:
                try: os.remove(tmp.name)
                except Exception: pass

        btns = ctk.CTkFrame(cont, fg_color="transparent"); btns.grid(row=3, column=0, columnspan=2, pady=(10,0))
        ctk.CTkButton(btns, text="Cancelar", fg_color="#6b7280", width=120, command=win.destroy).pack(side="left", padx=6)
        ctk.CTkButton(btns, text="Enviar", fg_color="#22c55e", width=140, command=enviar).pack(side="left", padx=6)

        try:
            root.update_idletasks()
            w, h = 520, 320
            x = root.winfo_x() + (root.winfo_width() // 2) - (w // 2)
            y = root.winfo_y() + (root.winfo_height() // 2) - (h // 2)
            win.geometry(f"{w}x{h}+{max(x,0)}+{max(y,0)}")
            ent_to.focus_set()
        except Exception:
            pass

    # binds
    entry_filter.bind("<KeyRelease>", _debounced_filter)
    btn_prev.configure(command=_go_prev)
    btn_next.configure(command=_go_next)
    btn_hoy.configure(command=_go_today)
    btn_refresh.configure(command=_load_day)
    entry_fecha.bind("<Return>", _enter_date)
    entry_fecha.bind("<FocusOut>", _enter_date)
    btn_pdf.configure(command=_on_pdf)
    btn_email.configure(command=_on_email)

    def _dbl_ing(_e):
        sel = tv_ing.selection()
        if not sel: return
        vals = tv_ing.item(sel[0], "values")
        if len(vals) >= 4:
            _open_person_detail(root, state["fecha"], vals[2], vals[3])
    def _dbl_sal(_e):
        sel = tv_sal.selection()
        if not sel: return
        vals = tv_sal.item(sel[0], "values")
        if len(vals) >= 4:
            _open_person_detail(root, state["fecha"], vals[2], vals[3])
    tv_ing.bind("<Double-1>", _dbl_ing)
    tv_sal.bind("<Double-1>", _dbl_sal)

    # Carga inicial
    def _set_header_date():  # redefine para cierre sobre outer
        d = state["fecha"]
        lbl_dia.configure(text=f"{DAYS_ES[d.weekday()]} {_format_ddmmyyyy(d)}")
        entry_fecha.delete(0, "end"); entry_fecha.insert(0, _format_ddmmyyyy(d))
    _set_header_date()
    _load_day()
