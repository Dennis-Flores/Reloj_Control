# nomina.py
import customtkinter as ctk
import sqlite3
import tkinter as tk
import tkinter.ttk as ttk
import threading
import os
import webbrowser
import smtplib
from email.message import EmailMessage
from datetime import datetime

DB = "reloj_control.db"

# ==================== Correo de envío (BioAccess) ====================
SMTP_HOST = "mail.bioaccess.cl"
SMTP_PORT = 465  # SSL directo
SMTP_USER = "documentos_bd@bioaccess.cl"
SMTP_PASS = "documentos@2025"
USAR_SMTP = True

# ==================== Configuración UI / Data ====================
CHUNK_SIZE = 200
DEBOUNCE_MS = 250
DAYS_ORDER = ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado', 'Domingo']
PERIODS = ["Mañana", "Tarde", "Nocturna"]  # columnas

# ---------- Utilidades DB ----------
def _ensure_indexes():
    """Crea índices para acelerar búsquedas/ordenamientos."""
    try:
        con = sqlite3.connect(DB)
        cur = con.cursor()
        cur.execute("CREATE INDEX IF NOT EXISTS idx_trab_rut ON trabajadores (rut)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_trab_ap_nom ON trabajadores (apellido, nombre)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_trab_nom_busq ON trabajadores (nombre, apellido, rut)")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_horarios_rut ON horarios (rut)")
        con.commit()
        con.close()
    except Exception as e:
        print("No se pudieron crear índices:", e)

def _like_param(s: str) -> str:
    return f"%{(s or '').lower()}%"

def _count_funcionarios(filtro: str) -> int:
    con = sqlite3.connect(DB)
    cur = con.cursor()
    if filtro:
        cur.execute("""
            SELECT COUNT(*) FROM trabajadores
            WHERE lower(nombre || ' ' || apellido || rut) LIKE ?
        """, (_like_param(filtro),))
    else:
        cur.execute("SELECT COUNT(*) FROM trabajadores")
    total = cur.fetchone()[0]
    con.close()
    return total

def _fetch_page(filtro: str, sort_col: str, sort_dir_asc: bool, limit: int, offset: int):
    """Devuelve filas: (nombre, apellido, rut, profesion, correo, cumpleanos)."""
    sort_map = {
        "Nombre": "apellido, nombre",
        "RUT": "rut",
        "Profesión": "profesion",
        "Correo": "correo",
        "Cumpleaños": "cumpleanos"
    }
    order_by = sort_map.get(sort_col, "apellido, nombre")
    direction = "ASC" if sort_dir_asc else "DESC"

    con = sqlite3.connect(DB)
    cur = con.cursor()
    base_sql = """
        SELECT nombre, apellido, rut, IFNULL(profesion,'-') AS profesion,
               IFNULL(correo,'-') AS correo, IFNULL(cumpleanos,'-') AS cumpleanos
        FROM trabajadores
    """
    params = []
    if filtro:
        base_sql += " WHERE lower(nombre || ' ' || apellido || rut) LIKE ?"
        params.append(_like_param(filtro))
    base_sql += f" ORDER BY {order_by} {direction} LIMIT ? OFFSET ?"
    params.extend([limit, offset])

    cur.execute(base_sql, params)
    rows = cur.fetchall()
    con.close()
    return rows

# ---------- HORARIO: Pivot directo en SQL ----------
def _fetch_matrix_por_rut_sql(rut: str):
    """
    Devuelve (matrix, ignored_total) ya pivotado:
    matrix = {dia: {'Mañana':(he,hs)|'-', 'Tarde':..., 'Nocturna':...}}
    Ignora filas con hora_entrada/salida vacías. Asigna por orden de inicio.
    Compatible con SQLite sin funciones ventana (ROW_NUMBER emulado).
    """
    con = sqlite3.connect(DB)
    cur = con.cursor()

    # total filas para este RUT
    cur.execute("SELECT COUNT(*) FROM horarios WHERE rut = ?", (rut,))
    total_rows = cur.fetchone()[0] or 0

    # filas válidas (no vacías)
    cur.execute("""
        WITH clean AS (
          SELECT dia,
                 TRIM(hora_entrada) AS he,
                 TRIM(hora_salida)  AS hs
          FROM horarios
          WHERE rut = ?
            AND TRIM(IFNULL(hora_entrada,'')) <> ''
            AND TRIM(IFNULL(hora_salida,'')) <> ''
        )
        SELECT COUNT(*) FROM clean
    """, (rut,))
    valid_rows = cur.fetchone()[0] or 0
    ignored_total = max(total_rows - valid_rows, 0)

    # Pivot: 1º bloque del día -> Mañana, 2º -> Tarde, 3º -> Nocturna
    cur.execute(f"""
        WITH clean AS (
          SELECT dia,
                 TRIM(hora_entrada) AS he,
                 TRIM(hora_salida)  AS hs
          FROM horarios
          WHERE rut = ?
            AND TRIM(IFNULL(hora_entrada,'')) <> ''
            AND TRIM(IFNULL(hora_salida,'')) <> ''
        ),
        ranked AS (
          SELECT c.*,
                 1 + (
                   SELECT COUNT(*)
                   FROM clean c2
                   WHERE c2.dia = c.dia AND time(c2.he) < time(c.he)
                 ) AS rn
          FROM clean c
        ),
        pivot AS (
          SELECT dia,
                 MAX(CASE WHEN rn=1 THEN he END) AS man_ent,
                 MAX(CASE WHEN rn=1 THEN hs END) AS man_sal,
                 MAX(CASE WHEN rn=2 THEN he END) AS tar_ent,
                 MAX(CASE WHEN rn=2 THEN hs END) AS tar_sal,
                 MAX(CASE WHEN rn=3 THEN he END) AS noc_ent,
                 MAX(CASE WHEN rn=3 THEN hs END) AS noc_sal
          FROM ranked
          GROUP BY dia
        )
        SELECT dia, man_ent, man_sal, tar_ent, tar_sal, noc_ent, noc_sal
        FROM pivot
        ORDER BY CASE dia
            WHEN 'Lunes' THEN 1 WHEN 'Martes' THEN 2 WHEN 'Miércoles' THEN 3
            WHEN 'Jueves' THEN 4 WHEN 'Viernes' THEN 5 WHEN 'Sábado' THEN 6
            WHEN 'Domingo' THEN 7 ELSE 99 END
    """, (rut,))
    rows = cur.fetchall()
    con.close()

    matrix = {d: {p: "-" for p in PERIODS} for d in DAYS_ORDER}
    for dia, me, ms, te, ts, ne, ns in rows:
        if dia not in matrix:
            continue
        if me and ms:
            matrix[dia]["Mañana"] = (me, ms)
        if te and ts:
            matrix[dia]["Tarde"] = (te, ts)
        if ne and ns:
            matrix[dia]["Nocturna"] = (ne, ns)

    return matrix, ignored_total

# ---------- Render helpers ----------
def _matrix_to_text(matrix, header=True):
    lines = []
    if header:
        lines.append("Día        Mañana(Ent-Sal)  Tarde(Ent-Sal)   Nocturna(Ent-Sal)")
        lines.append("-" * 66)
    for d in DAYS_ORDER:
        m = matrix[d]
        def s(val):
            if val == "-":
                return "--:-- - --:--"
            he, hs = val
            return f"{he} - {hs}"
        lines.append(f"{d:<10} {s(m['Mañana']):<17} {s(m['Tarde']):<17} {s(m['Nocturna']):<17}")
    return "\n".join(lines)

def _matrix_to_html(nombre, rut, profesion, correo, cumple, matrix, ignored=0):
    """HTML claro, alto contraste y estilos inline (robustos ante dark-mode)."""
    gen = datetime.now().strftime("%d-%m-%Y %H:%M")
    # Estilos base (inline-friendly)
    CARD_BG   = "#ffffff"
    PAGE_BG   = "#f3f4f6"
    TEXT      = "#111827"
    MUTED     = "#6b7280"
    BORDER    = "#e5e7eb"
    ACCENT    = "#0b5ea8"   # azul corporativo
    TH_BG     = ACCENT
    TH_TXT    = "#ffffff"
    SUB_BG    = "#eaf2fb"   # subtítulos en claro
    ROW_EVEN  = "#ffffff"
    ROW_ODD   = "#f9fafb"

    def td(val):
        return f'<td style="padding:8px;border:1px solid {BORDER};text-align:center;color:{TEXT}">{val}</td>'

    def c(val):
        if val == "-":
            return td("--:--") + td("--:--")
        he, hs = val
        return td(he) + td(hs)

    rows = []
    for i, d in enumerate(DAYS_ORDER):
        bg = ROW_EVEN if i % 2 == 0 else ROW_ODD
        m = matrix[d]
        rows.append(
            f"<tr style='background:{bg}'>"
            f"<th style='text-align:left;padding:8px;border:1px solid {BORDER};color:{MUTED};background:{bg};width:140px;font-weight:600'>{d}</th>"
            f"{c(m['Mañana'])}{c(m['Tarde'])}{c(m['Nocturna'])}"
            "</tr>"
        )

    # --- Badge especial para correo: fondo claro + link con color controlado ---
    if correo and correo != "-":
        email_badge = (
            "<span style=\"display:inline-block;background:#eaf2fb;border:1px solid #bfd6f7;"
            "color:{0};border-radius:8px;padding:4px 10px;font-size:12px;margin-right:6px\">"
            "Correo: <a href=\"mailto:{1}\" style=\"color:{0};text-decoration:underline\">{1}</a>"
            "</span>"
        ).format(ACCENT, correo)
    else:
        email_badge = (
            "<span style=\"display:inline-block;background:#eaf2fb;border:1px solid #bfd6f7;"
            "color:{0};border-radius:8px;padding:4px 10px;font-size:12px;margin-right:6px\">"
            "Correo: -</span>"
        ).format(ACCENT)

    warn = (f"<p style='margin:8px 0 0 0;color:#b45309'>"
            f"Se ignoraron {ignored} fila(s) vacía(s) del horario."
            f"</p>") if ignored else ""

    html = f"""<!doctype html>
<html lang="es"><meta charset="utf-8">
<body style="margin:0;padding:24px;background:{PAGE_BG};color:{TEXT};font-family:Arial,Helvetica,sans-serif">
  <div style="max-width:900px;margin:0 auto;background:{CARD_BG};border:1px solid {BORDER};border-radius:12px;padding:20px;box-sizing:border-box">
    <h1 style="margin:0 0 8px 0;font-size:22px;color:{TEXT}">Horario de {nombre}</h1>

    <!-- Mensaje profesional -->
    <p style="margin:8px 0 14px 0;line-height:1.55;color:{TEXT}">
      Estimado(a), junto con saludar, se remite el detalle de horario vigente del funcionario indicado más abajo.
      Este informe es generado automáticamente por <strong>BioAccess – Control de Horarios</strong> y tiene fines informativos internos.
      En caso de dudas o discrepancias, por favor responda a este mismo correo para su revisión.
    </p>

    <div style="margin:0 0 12px 0;line-height:1.7">
      <span style="display:inline-block;background:{ACCENT};color:{TH_TXT};border-radius:8px;padding:4px 10px;font-size:12px;margin-right:6px">RUT: {rut}</span>
      <span style="display:inline-block;background:{ACCENT};color:{TH_TXT};border-radius:8px;padding:4px 10px;font-size:12px;margin-right:6px">Cargo: {profesion}</span>
      {email_badge}
      <span style="display:inline-block;background:{ACCENT};color:{TH_TXT};border-radius:8px;padding:4px 10px;font-size:12px">Cumple: {cumple}</span>
    </div>

    {warn}

    <table role="table" style="border-collapse:collapse;width:100%;background:{CARD_BG};border:1px solid {BORDER}">
      <thead>
        <tr>
          <th style="background:{TH_BG};color:{TH_TXT};padding:10px;border:1px solid {BORDER};text-align:left;width:140px">Día</th>
          <th colspan="2" style="background:{TH_BG};color:{TH_TXT};padding:10px;border:1px solid {BORDER};text-align:center">MAÑANA</th>
          <th colspan="2" style="background:{TH_BG};color:{TH_TXT};padding:10px;border:1px solid {BORDER};text-align:center">TARDE</th>
          <th colspan="2" style="background:{TH_BG};color:{TH_TXT};padding:10px;border:1px solid {BORDER};text-align:center">NOCTURNA</th>
        </tr>
        <tr>
          <th style="background:{SUB_BG};color:{TEXT};padding:8px;border:1px solid {BORDER};text-align:left"></th>
          <th style="background:{SUB_BG};color:{TEXT};padding:8px;border:1px solid {BORDER};text-align:center">Entrada</th>
          <th style="background:{SUB_BG};color:{TEXT};padding:8px;border:1px solid {BORDER};text-align:center">Salida</th>
          <th style="background:{SUB_BG};color:{TEXT};padding:8px;border:1px solid {BORDER};text-align:center">Entrada</th>
          <th style="background:{SUB_BG};color:{TEXT};padding:8px;border:1px solid {BORDER};text-align:center">Salida</th>
          <th style="background:{SUB_BG};color:{TEXT};padding:8px;border:1px solid {BORDER};text-align:center">Entrada</th>
          <th style="background:{SUB_BG};color:{TEXT};padding:8px;border:1px solid {BORDER};text-align:center">Salida</th>
        </tr>
      </thead>
      <tbody>
        {''.join(rows)}
      </tbody>
    </table>

    <p style="margin:14px 0 0 0;color:{MUTED};font-size:12px">
      Generado el {gen}. Sistema BioAccess – www.bioaccess.cl
    </p>
  </div>
</body>
</html>"""
    return html



def _save_html_and_open(nombre, rut, profesion, correo, cumple, matrix, ignored=0, open_after=True):
    html = _matrix_to_html(nombre, rut, profesion, correo, cumple, matrix, ignored)
    out_dir = os.path.join(os.getcwd(), "exportes_horario")
    os.makedirs(out_dir, exist_ok=True)
    fname = f"horario_{rut.replace('.', '').replace('-', '')}.html"
    fpath = os.path.join(out_dir, fname)
    with open(fpath, "w", encoding="utf-8") as f:
        f.write(html)
    if open_after:
        webbrowser.open(f"file://{os.path.abspath(fpath)}")
    return fpath

def _save_pdf_report(nombre, rut, profesion, correo, cumple, matrix):
    """Genera PDF con reportlab si está disponible. Devuelve path o None."""
    try:
        from reportlab.lib.pagesizes import A4
        from reportlab.pdfgen import canvas
        from reportlab.lib import colors
        from reportlab.lib.units import cm
    except Exception:
        return None

    out_dir = os.path.join(os.getcwd(), "exportes_horario")
    os.makedirs(out_dir, exist_ok=True)
    fname = f"horario_{rut.replace('.', '').replace('-', '')}.pdf"
    fpath = os.path.join(out_dir, fname)

    c = canvas.Canvas(fpath, pagesize=A4)
    W, H = A4
    x0, y0 = 2*cm, H - 2*cm

    c.setFont("Helvetica-Bold", 14)
    c.drawString(x0, y0, f"Horario - {nombre} ({rut})")
    y0 -= 12
    c.setFont("Helvetica", 10)
    c.drawString(x0, y0, f"Cargo: {profesion}   Correo: {correo}   Cumple: {cumple}")
    y0 -= 18

    # Tabla
    col_w = [3.2*cm, 2.2*cm, 2.2*cm, 2.2*cm, 2.2*cm, 2.2*cm, 2.2*cm]
    headers1 = ["Día", "MAÑANA", "", "TARDE", "", "NOCTURNA", ""]
    headers2 = ["", "Entrada", "Salida", "Entrada", "Salida", "Entrada", "Salida"]

    def draw_row(y, vals, bold=False, fill=None):
        x = x0
        for i, v in enumerate(vals):
            c.rect(x, y-12, col_w[i], 14, fill=0, stroke=1)
            c.setFont("Helvetica-Bold" if bold else "Helvetica", 9 if bold else 9)
            c.drawString(x+4, y-9, v)
            x += col_w[i]

    draw_row(y0, headers1, bold=True)
    y0 -= 14
    draw_row(y0, headers2, bold=True)
    y0 -= 16

    for d in DAYS_ORDER:
        row = []
        row.append(d)
        for p in PERIODS:
            val = matrix[d][p]
            if val == "-":
                row.extend(["--:--", "--:--"])
            else:
                row.extend([val[0], val[1]])
        draw_row(y0, row)
        y0 -= 16
        if y0 < 2.5*cm:
            c.showPage()
            y0 = H - 2*cm

    c.setFont("Helvetica-Oblique", 8)
    c.drawString(x0, 1.8*cm, "Generado por BioAccess.")
    c.save()
    return fpath

def _send_email(remitente_user, remitente_pass, host, port, to_list, cc_list, subject, text_body, html_body):
    if not USAR_SMTP:
        tk.messagebox.showinfo("Envío no configurado",
                               f"Para: {', '.join(to_list)}\nCC: {', '.join(cc_list)}\nAsunto: {subject}\n\n{text_body[:500]}...")
        return True

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = remitente_user
    msg["To"] = ", ".join([t for t in to_list if t])
    if cc_list:
        msg["Cc"] = ", ".join([c for c in cc_list if c])
    msg.set_content(text_body)
    if html_body:
        msg.add_alternative(html_body, subtype="html")

    try:
        with smtplib.SMTP_SSL(host, port) as s:
            s.login(remitente_user, remitente_pass)
            s.send_message(msg)
        return True
    except Exception as e:
        tk.messagebox.showerror("Correo", f"No fue posible enviar el correo:\n{e}")
        return False

# ---------- UI ----------
def construir_nomina(frame_padre):
    _ensure_indexes()

    state = {
        "filtro": "",
        "sort_col": "Nombre",
        "sort_asc": True,
        "loaded": 0,
        "total": 0,
        "search_after_id": None,
        "selected": {  # datos del funcionario seleccionado
            "rut": None,
            "nombre": "",
            "apellido": "",
            "profesion": "",
            "correo": "",
            "cumple": ""
        }
    }

    root = ctk.CTkFrame(frame_padre)
    root.pack(fill="both", expand=True)

    ctk.CTkLabel(root, text="Nómina de Funcionarios", font=("Arial", 16, "bold")).pack(pady=(10, 6))

    # ---- Barra de búsqueda ----
    topbar = ctk.CTkFrame(root, fg_color="transparent")
    topbar.pack(fill="x", padx=16, pady=(0, 6))

    ctk.CTkLabel(topbar, text="🔍", font=("Arial", 18)).pack(side="left", padx=(0, 6))
    entry_search = ctk.CTkEntry(topbar, placeholder_text="Buscar por nombre o RUT", width=320)
    entry_search.pack(side="left")
    btn_clear = ctk.CTkButton(topbar, text="Limpiar", width=70)
    btn_clear.pack(side="left", padx=(6, 0))

    lbl_count = ctk.CTkLabel(topbar, text="0 resultados", text_color="gray")
    lbl_count.pack(side="right")

    # ---- Layout principal ----
    body = ctk.CTkFrame(root, fg_color="transparent")
    body.pack(fill="both", expand=True, padx=12, pady=8)
    body.grid_columnconfigure(0, weight=3)   # lista
    body.grid_columnconfigure(1, weight=2)   # detalle
    body.grid_rowconfigure(1, weight=1)

    # ---- LISTA (Treeview) ----
    list_frame = ctk.CTkFrame(body)
    list_frame.grid(row=0, column=0, rowspan=2, sticky="nsew", padx=(0, 6))

    # Estilo oscuro
    style = ttk.Style(list_frame)
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass

    style.configure(
        "Nomina.Treeview",
        background="#111827",
        fieldbackground="#111827",
        foreground="#E5E7EB",
        rowheight=26,
        borderwidth=0,
        relief="flat",
    )
    style.map(
        "Nomina.Treeview",
        background=[("selected", "#0ea5e9")],
        foreground=[("selected", "white")],
    )
    style.configure(
        "Nomina.Treeview.Heading",
        background="#1f2937",
        foreground="#E5E7EB",
        relief="flat",
    )
    style.map("Nomina.Treeview.Heading", background=[("active", "#374151")])

    style.configure(
        "Nomina.Vertical.TScrollbar",
        background="#1f2937",
        troughcolor="#111827",
        bordercolor="#111827",
        lightcolor="#111827",
        darkcolor="#111827",
        arrowcolor="#E5E7EB",
    )

    columns = ("Nombre", "RUT", "Profesión", "Correo", "Cumpleaños")
    tree = ttk.Treeview(
        list_frame,
        columns=columns,
        show="headings",
        height=18,
        selectmode="browse",
        style="Nomina.Treeview",
    )
    yscroll = ttk.Scrollbar(
        list_frame, orient="vertical", command=tree.yview, style="Nomina.Vertical.TScrollbar"
    )
    tree.configure(yscrollcommand=yscroll.set, takefocus=True)

    # Anchuras
    tree.column("Nombre", width=240, anchor="w")
    tree.column("RUT", width=130, anchor="w")
    tree.column("Profesión", width=160, anchor="w")
    tree.column("Correo", width=220, anchor="w")
    tree.column("Cumpleaños", width=110, anchor="w")

    for col in columns:
        tree.heading(col, text=col)

    tree.grid(row=0, column=0, sticky="nsew")
    yscroll.grid(row=0, column=1, sticky="ns")
    list_frame.grid_rowconfigure(0, weight=1)
    list_frame.grid_columnconfigure(0, weight=1)

    # Ordenar por encabezado
    def _on_heading_click(col):
        def handler(*_):
            if state["sort_col"] == col:
                state["sort_asc"] = not state["sort_asc"]
            else:
                state["sort_col"] = col
                state["sort_asc"] = True
            _reload_from_scratch()
        return handler
    for col in columns:
        tree.heading(col, text=col, command=_on_heading_click(col))

    # Footer de la lista
    footer = ctk.CTkFrame(list_frame, fg_color="transparent")
    footer.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(6, 8))
    btn_more = ctk.CTkButton(footer, text="Mostrar más", width=140)

    # ---- DETALLE ----
    detail = ctk.CTkFrame(body)
    detail.grid(row=0, column=1, sticky="nsew")
    detail.grid_rowconfigure(7, weight=1)
    detail.grid_columnconfigure(0, weight=1)

    ctk.CTkLabel(detail, text="Detalle del Funcionario", font=("Arial", 14, "bold")).grid(row=0, column=0, sticky="w", padx=10, pady=(10, 6))
    lbl_info = ctk.CTkLabel(detail, text="Selecciona un funcionario", justify="left")
    lbl_info.grid(row=1, column=0, sticky="w", padx=10)

    # --- Correo (para/enviar) ---
    email_bar = ctk.CTkFrame(detail, fg_color="transparent")
    email_bar.grid(row=2, column=0, sticky="ew", padx=10, pady=(4, 0))
    email_bar.grid_columnconfigure(1, weight=1)
    email_bar.grid_columnconfigure(3, weight=1)

    ctk.CTkLabel(email_bar, text="Para:").grid(row=0, column=0, sticky="e", padx=(0,6))
    entry_to = ctk.CTkEntry(email_bar, placeholder_text="correo@dominio.cl", width=260)
    entry_to.grid(row=0, column=1, sticky="ew", padx=(0,10))
    ctk.CTkLabel(email_bar, text="CC:").grid(row=0, column=2, sticky="e", padx=(0,6))
    entry_cc = ctk.CTkEntry(email_bar, placeholder_text="opcional (separar por coma)", width=200)
    entry_cc.grid(row=0, column=3, sticky="ew")

    # Scroll con contenedor fijo para horario (tabla)
    sched_frame = ctk.CTkScrollableFrame(detail, label_text="Horario", label_font=("Arial", 13, "bold"))
    sched_frame.grid(row=4, column=0, sticky="nsew", padx=10, pady=(6, 6))
    sched_inner = ctk.CTkFrame(sched_frame, fg_color="transparent")
    sched_inner.pack(fill="x", pady=(4, 6))

    # Aviso de filas ignoradas
    warn_label = ctk.CTkLabel(detail, text="", text_color="#fbbf24")
    warn_label.grid(row=3, column=0, sticky="w", padx=12, pady=(2, 0))

    # Botonera bajo el horario
    btns_frame = ctk.CTkFrame(detail, fg_color="transparent")
    btns_frame.grid(row=5, column=0, sticky="ew", padx=10, pady=(0, 10))
    btn_export = ctk.CTkButton(btns_frame, text="Exportar (PDF/HTML)")
    btn_print  = ctk.CTkButton(btns_frame, text="Imprimir")
    btn_send   = ctk.CTkButton(btns_frame, text="Enviar por correo")
    btn_export.pack(side="left", padx=(0,8))
    btn_print.pack(side="left", padx=(0,8))
    btn_send.pack(side="left", padx=(0,8))

    # ---------- Render horario en tabla ----------
    def _clear_sched():
        for w in sched_inner.winfo_children():
            w.destroy()

    def _render_table(matrix, ignored=0):
        _clear_sched()
        grid = ctk.CTkFrame(sched_inner, fg_color="transparent")
        grid.pack(fill="x")

        # Encabezados superiores
        headers1 = ["Día", "MAÑANA", "", "TARDE", "", "NOCTURNA", ""]
        headers2 = ["", "Entrada", "Salida", "Entrada", "Salida", "Entrada", "Salida"]

        # helper cell
        def cell(parent, text, r, c, w=1, bold=False, muted=False, pad=(4,2), anchor="center"):
            font = ("Arial", 12, "bold") if bold else ("Arial", 12)
            color = "#9ca3af" if muted else None
            lbl = ctk.CTkLabel(parent, text=text, font=font, text_color=color, width=80, anchor=anchor)
            lbl.grid(row=r, column=c, columnspan=w, padx=6, pady=pad, sticky="nsew")

        # headers
        for i in range(7):
            grid.grid_columnconfigure(i, weight=1)
        cell(grid, headers1[0], 0, 0, bold=True, anchor="w")
        cell(grid, headers1[1], 0, 1, w=2, bold=True)
        cell(grid, headers1[3], 0, 3, w=2, bold=True)
        cell(grid, headers1[5], 0, 5, w=2, bold=True)

        cell(grid, headers2[1], 1, 1, bold=True)
        cell(grid, headers2[2], 1, 2, bold=True)
        cell(grid, headers2[3], 1, 3, bold=True)
        cell(grid, headers2[4], 1, 4, bold=True)
        cell(grid, headers2[5], 1, 5, bold=True)
        cell(grid, headers2[6], 1, 6, bold=True)

        r = 2
        for d in DAYS_ORDER:
            m = matrix[d]
            cell(grid, d, r, 0, bold=False, muted=True, anchor="w")
            def pair(val, rr, cc):
                if val == "-":
                    ent, sal = "--:--", "--:--"
                else:
                    ent, sal = val
                cell(grid, ent, rr, cc)
                cell(grid, sal, rr, cc+1)
            pair(m["Mañana"], r, 1)
            pair(m["Tarde"], r, 3)
            pair(m["Nocturna"], r, 5)
            r += 1

        warn_label.configure(text=(f"Se ignoraron {ignored} fila(s) vacía(s) del horario." if ignored else ""))

    # ---- Helpers UI ----
    def _update_count_label():
        mostrados = state["loaded"]
        total = state["total"]
        lbl_count.configure(text=f"{mostrados} de {total} resultados")
        if mostrados < total:
            btn_more.pack(side="top", pady=(6, 2))
        else:
            btn_more.pack_forget()

    # Filas alternadas
    tree.tag_configure("even", background="#0f172a")
    tree.tag_configure("odd", background="#111827")

    def _fill_tree(rows, append=False):
        if not append:
            for iid in tree.get_children():
                tree.delete(iid)
        base_index = state["loaded"] if append else 0
        for i, (nombre, apellido, rut, profesion, correo, cumple) in enumerate(rows):
            nombre_comp = f"{nombre} {apellido}".strip()
            tag = "even" if (base_index + i) % 2 == 0 else "odd"
            try:
                tree.insert("", "end", iid=str(rut), values=(nombre_comp, rut, profesion, correo, cumple), tags=(tag,))
            except tk.TclError:
                tree.insert("", "end", values=(nombre_comp, rut, profesion, correo, cumple), tags=(tag,))

    def _reload_from_scratch():
        state["loaded"] = 0
        state["total"] = 0
        _fetch_and_fill(next_chunk=True, reset=True)

    def _fetch_and_fill(next_chunk: bool, reset: bool = False):
        filtro = state["filtro"]
        sort_col = state["sort_col"]
        sort_asc = state["sort_asc"]

        def work():
            try:
                total = _count_funcionarios(filtro)
                offset = 0 if reset else state["loaded"]
                rows = _fetch_page(filtro, sort_col, sort_asc, CHUNK_SIZE if next_chunk else offset, offset)
            except Exception as e:
                print("Error consultando nómina:", e)
                return

            def apply():
                state["total"] = total
                if reset:
                    _fill_tree(rows, append=False)
                    state["loaded"] = len(rows)
                else:
                    _fill_tree(rows, append=True)
                    state["loaded"] += len(rows)
                _update_count_label()

            root.after(0, apply)

        threading.Thread(target=work, daemon=True).start()

    # Búsqueda con debounce
    def _on_search_changed(_=None):
        if state["search_after_id"] is not None:
            try:
                root.after_cancel(state["search_after_id"])
            except Exception:
                pass
            state["search_after_id"] = None

        def do_search():
            state["filtro"] = entry_search.get().strip()
            _reload_from_scratch()

        state["search_after_id"] = root.after(DEBOUNCE_MS, do_search)

    entry_search.bind("<KeyRelease>", _on_search_changed)
    btn_clear.configure(command=lambda: (entry_search.delete(0, "end"), _on_search_changed()))

    # 🔧 Reemplazo seguro del bind global (Ctrl+F) — sin usar bind_all
    try:
        root.winfo_toplevel().bind("<Control-f>", lambda e: entry_search.focus_set())
    except Exception:
        pass

    btn_more.configure(command=lambda: _fetch_and_fill(next_chunk=True, reset=False))

    # --- Selección robusta ---
    def _get_selected_rut():
        iid = tree.focus()
        if not iid:
            sel = tree.selection()
            if sel:
                iid = sel[0]
        if iid:
            # iid es el RUT porque lo usamos como iid al insertar
            return str(iid)
        return None

    def _fetch_info_trabajador(rut: str):
        con = sqlite3.connect(DB)
        cur = con.cursor()
        cur.execute("""
            SELECT nombre, apellido, IFNULL(profesion,'-'), IFNULL(correo,'-'), IFNULL(cumpleanos,'-')
            FROM trabajadores WHERE rut = ?
        """, (rut,))
        row = cur.fetchone()
        con.close()
        return row

    def _on_select(_=None):
        rut = _get_selected_rut()
        if not rut:
            return
        info = _fetch_info_trabajador(rut)
        if not info:
            lbl_info.configure(text="Sin información de este RUT")
            _clear_sched()
            warn_label.configure(text="")
            return

        nombre, apellido, profesion, correo, cumple = info
        nombre_comp = f"{nombre} {apellido}".strip()
        state["selected"].update({
            "rut": rut, "nombre": nombre, "apellido": apellido,
            "profesion": profesion, "correo": correo, "cumple": cumple
        })
        lbl_info.configure(
            text=f"Nombre : {nombre_comp}\n"
                 f"RUT    : {rut}\n"
                 f"Cargo  : {profesion}\n"
                 f"Correo : {correo}\n"
                 f"Cumple : {cumple}",
            justify="left"
        )

        # Prefill de correo "Para"
        entry_to.delete(0, "end")
        if correo and correo != "-":
            entry_to.insert(0, correo)

        # --- matriz desde SQL (pivot)
        matrix, ignored = _fetch_matrix_por_rut_sql(rut)
        _render_table(matrix, ignored=ignored)

        # Acciones
        def do_export():
            pdf_path = _save_pdf_report(nombre_comp, rut, profesion, correo, cumple, matrix)
            if pdf_path:
                tk.messagebox.showinfo("Exportar", f"PDF generado:\n{pdf_path}")
                try:
                    os.startfile(pdf_path)  # Windows
                except Exception:
                    webbrowser.open(f"file://{os.path.abspath(pdf_path)}")
            else:
                html_path = _save_html_and_open(nombre_comp, rut, profesion, correo, cumple, matrix, ignored=ignored, open_after=True)
                tk.messagebox.showinfo("Exportar", f"HTML generado:\n{html_path}")

        def do_print():
            html_path = _save_html_and_open(nombre_comp, rut, profesion, correo, cumple, matrix, ignored=ignored, open_after=False)
            printed = False
            if os.name == "nt":
                try:
                    os.startfile(html_path, "print")
                    printed = True
                except Exception:
                    printed = False
            if not printed:
                webbrowser.open(f"file://{os.path.abspath(html_path)}")
                tk.messagebox.showinfo("Imprimir", "Se abrió el documento en el navegador. Usa Ctrl+P para imprimir.")

        def do_send():
            to_raw = entry_to.get().strip()
            cc_raw = entry_cc.get().strip()
            to_list = [x.strip() for x in to_raw.split(",") if x.strip()]
            cc_list = [x.strip() for x in cc_raw.split(",") if x.strip()]

            if not to_list:
                tk.messagebox.showwarning("Correo", "Indica al menos un destinatario en 'Para'.")
                return

            matrix_txt = _matrix_to_text(matrix)
            html_body = _matrix_to_html(nombre_comp, rut, profesion, correo, cumple, matrix, ignored=ignored)

            # ======= NUEVO: texto profesional + sello de generación =======
            gen = datetime.now().strftime("%d-%m-%Y %H:%M")
            intro = (
                "Estimado(a), junto con saludar, se remite el detalle de horario vigente.\n"
                "Este informe es generado automáticamente por BioAccess – Control de Horarios.\n"
                "Si detecta alguna discrepancia, por favor responda a este mismo correo.\n"
            )
            text_body = (
                f"{intro}\n"
                f"Funcionario: {nombre_comp} ({rut})\n"
                f"Cargo: {profesion} | Correo: {correo} | Cumple: {cumple}\n"
                f"Generado: {gen}\n\n"
                f"{matrix_txt}\n\n"
                f"BioAccess."
            )
            # (opcional) cambia el asunto si prefieres
            subject = f"Horario vigente – {nombre_comp} ({rut})"

            ok = _send_email(SMTP_USER, SMTP_PASS, SMTP_HOST, SMTP_PORT,
                            to_list, cc_list, subject, text_body, html_body)
            if ok:
                tk.messagebox.showinfo("Correo", "Horario enviado correctamente.")


        btn_export.configure(command=do_export)
        btn_print.configure(command=do_print)
        btn_send.configure(command=do_send)

    # Bindings para asegurar selección
    def _bind_selection_events():
        for ev in ("<<TreeviewSelect>>", "<ButtonRelease-1>", "<KeyRelease-Up>", "<KeyRelease-Down>", "<Return>", "<Double-1>"):
            tree.bind(ev, lambda e: root.after(1, _on_select))
    _bind_selection_events()

    # Carga inicial
    _reload_from_scratch()
