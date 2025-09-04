# asistencia_funcionarios.py
import os
import sqlite3
import calendar
import datetime as dt
import tkinter as tk
from tkinter import ttk, messagebox
import customtkinter as ctk

# Calendario opcional
try:
    from tkcalendar import DateEntry
    HAS_TKCAL = True
except Exception:
    HAS_TKCAL = False

# ===== PDF =====
try:
    from reportlab.lib.pagesizes import A4, landscape  # <- landscape
    from reportlab.lib import colors
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm                # <- cm para anchos exactos
    from reportlab.lib.enums import TA_LEFT
    PDF_OK = True
except Exception:
    PDF_OK = False

# SMTP BioAccess
import smtplib
from email.message import EmailMessage
import mimetypes
import traceback
import sys
import subprocess

SMTP_HOST = "mail.bioaccess.cl"
SMTP_PORT = 465  # SSL directo
SMTP_USER = "documentos_bd@bioaccess.cl"
SMTP_PASS = "documentos@2025"
USAR_SMTP = True

MESES_ES = {
    1:"Enero",2:"Febrero",3:"Marzo",4:"Abril",5:"Mayo",6:"Junio",
    7:"Julio",8:"Agosto",9:"Septiembre",10:"Octubre",11:"Noviembre",12:"Diciembre"
}

def _downloads_dir():
    home = os.path.expanduser("~")
    cand = os.path.join(home, "Downloads")
    if os.path.isdir(cand): return cand
    xdg = os.environ.get("XDG_DOWNLOAD_DIR")
    if xdg and os.path.isdir(xdg): return xdg
    return home

def _abrir_archivo(path):
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)  # type: ignore
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception as e:
        messagebox.showinfo("Archivo generado", f"Se guardó en:\n{path}\n\nNo se pudo abrir automáticamente: {e}")

def _mes_range(year: int, month: int):
    last_day = calendar.monthrange(year, month)[1]
    start = dt.date(year, month, 1)
    end = dt.date(year, month, last_day)
    return start, end, last_day

def _leer_funcionarios(con):
    cur = con.cursor()
    cur.execute("""
        SELECT rut, (nombre || ' ' || apellido) AS nombre_completo, IFNULL(correo,'')
        FROM trabajadores
        ORDER BY nombre_completo COLLATE NOCASE;
    """)
    return cur.fetchall()  # [(rut, nombre, correo), ...]

def _cargar_registros(con, rut: str | None, f_ini: dt.date, f_fin: dt.date):
    """
    Devuelve filas normalizadas: (rut, nombre, fecha, tipo, hora)
    combinando eventos (tipo/hora) y columnas (hora_ingreso/hora_salida).
    """
    cur = con.cursor()
    params = [f_ini.strftime("%Y-%m-%d"), f_fin.strftime("%Y-%m-%d")]
    filtro = ""
    if rut:
        filtro = " AND r.rut = ? "
        params.append(rut)
    cur.execute(f"""
        SELECT 
            r.rut,
            COALESCE(t.nombre || ' ' || t.apellido, r.nombre, r.rut) AS nombre,
            r.fecha,
            MAX(CASE WHEN LOWER(IFNULL(r.tipo,'')) LIKE 'ing%%' THEN r.hora END) AS evt_ingreso,
            MAX(CASE WHEN LOWER(IFNULL(r.tipo,'')) LIKE 'sal%%' THEN r.hora END) AS evt_salida,
            MAX(NULLIF(TRIM(IFNULL(r.hora_ingreso,'')), '')) AS col_ingreso,
            MAX(NULLIF(TRIM(IFNULL(r.hora_salida,'')), '')) AS col_salida
        FROM registros r
        LEFT JOIN trabajadores t ON t.rut = r.rut
        WHERE r.fecha BETWEEN ? AND ? {filtro}
        GROUP BY r.rut, r.fecha
        ORDER BY r.rut, r.fecha;
    """, params)
    agg = cur.fetchall()

    norm_rows = []
    for rutx, nombre, fecha, evt_ing, evt_sal, col_ing, col_sal in agg:
        hora_ing = col_ing or evt_ing
        hora_sal = col_sal or evt_sal
        if hora_ing:
            norm_rows.append((rutx, nombre, fecha, "ingreso", (hora_ing or "")[:5]))
        if hora_sal:
            norm_rows.append((rutx, nombre, fecha, "salida", (hora_sal or "")[:5]))
    return norm_rows

def _resumen_por_rut(rows):
    from collections import defaultdict
    data = defaultdict(lambda: {
        "nombre": "",
        "fechas": set(),
        "por_fecha": {},  # fecha -> {"ingreso": hh:mm, "salida": hh:mm}
    })
    for rut, nombre, fecha, tipo, hora in rows:
        d = data[rut]
        d["nombre"] = nombre
        d["fechas"].add(fecha)
        if fecha not in d["por_fecha"]:
            d["por_fecha"][fecha] = {"ingreso": None, "salida": None}
        if tipo.lower().startswith("ing"):
            d["por_fecha"][fecha]["ingreso"] = hora
        elif tipo.lower().startswith("sal"):
            d["por_fecha"][fecha]["salida"] = hora
    return data

def _style_dark_treeview():
    style = ttk.Style()
    try:
        style.theme_use('clam')
    except Exception:
        pass
    style.configure(
        "Dark.Treeview",
        background="#111827",
        foreground="#E5E7EB",
        fieldbackground="#111827",
        rowheight=26,
        bordercolor="#2b3440",
        borderwidth=0
    )
    style.configure("Dark.Treeview.Heading",
        background="#1f2937",
        foreground="#E5E7EB",
        relief="flat"
    )
    style.map("Dark.Treeview",
        background=[("selected","#0ea5e9")],
        foreground=[("selected","#ffffff")]
    )
    return "Dark.Treeview", "Dark.Treeview.Heading"

def _construir_tabla_preview(parent):
    tv_style, _ = _style_dark_treeview()
    cols = ("nombre", "rut", "rango", "total_dias_mes", "dias_trabajados", "observaciones")
    tree = ttk.Treeview(parent, columns=cols, show="headings", height=12, style=tv_style)
    headers = {
        "nombre": "Nombre",
        "rut": "RUT",
        "rango": "Periodo",
        "total_dias_mes": "Días del periodo",
        "dias_trabajados": "Días con ingreso",
        "observaciones": "Detalle (completo / incompleto)"
    }
    for c in cols:
        tree.heading(c, text=headers[c])
        tree.column(c, width=180 if c not in ("observaciones","rango") else (520 if c=="observaciones" else 200), anchor="w")
    tree.pack(fill="both", expand=True)
    return tree

# ====== helpers correo ======
def _emails_para_ruts(con, ruts):
    if not ruts: return []
    qmarks = ",".join(["?"]*len(ruts))
    cur = con.cursor()
    try:
        cur.execute(f"SELECT rut, IFNULL(correo,''), (nombre||' '||apellido) FROM trabajadores WHERE rut IN ({qmarks})", tuple(ruts))
        rows = cur.fetchall()
    except Exception:
        rows = []
    emails = []
    for _rut, correo, _nom in rows:
        c = (correo or "").strip()
        if c and "@" in c:
            emails.append(c)
    seen = set(); out = []
    for e in emails:
        k = e.lower()
        if k in seen: continue
        seen.add(k); out.append(e)
    return out

def _html_email_asistencia_general(period_text: str, listado: list[tuple[str,str]]) -> str:
    gen = dt.datetime.now().strftime("%d-%m-%Y %H:%M")
    CARD_BG = "#ffffff"
    PAGE_BG = "#f3f4f6"
    TEXT    = "#111827"
    MUTED   = "#6b7280"
    BORDER  = "#e5e7eb"
    ACCENT  = "#0b5ea8"

    total = len(listado)
    show = listado[:10]
    extra = max(0, total - len(show))
    lis = "".join(
        f"<li style='margin:2px 0'>{n} <span style='color:{MUTED}'>({r})</span></li>"
        for n, r in show
    )
    if extra:
        lis += f"<li style='margin:2px 0;color:{MUTED}'>+ {extra} más…</li>"

    return f"""<!doctype html>
<html lang="es"><meta charset="utf-8">
<body style="margin:0;padding:24px;background:{PAGE_BG};color:{TEXT};font-family:Arial,Helvetica,sans-serif">
  <div style="max-width:900px;margin:0 auto;background:{CARD_BG};border:1px solid {BORDER};border-radius:12px;padding:20px;box-sizing:border-box">
    <h1 style="margin:0 0 8px 0;font-size:22px;color:{TEXT}">Asistencia de Funcionarios</h1>

    <p style="margin:8px 0 14px 0;line-height:1.55;color:{TEXT}">
      Estimado(a), junto con saludar, se remite el reporte de <strong>Asistencia de Funcionarios</strong> correspondiente al periodo indicado.
      Este mensaje es generado automáticamente por <strong>BioAccess – Control de Horarios</strong>.
      Se adjunta el documento en formato <strong>PDF</strong> para su revisión.
    </p>

    <div style="margin:0 0 12px 0;line-height:1.7">
      <span style="display:inline-block;background:{ACCENT};color:#ffffff;border-radius:8px;padding:4px 10px;font-size:12px;margin-right:6px">Periodo: {period_text}</span>
      <span style="display:inline-block;background:{ACCENT};color:#ffffff;border-radius:8px;padding:4px 10px;font-size:12px">Funcionarios: {total}</span>
    </div>

    <div style="background:#f9fafb;border:1px solid {BORDER};border-radius:10px;padding:12px;margin-bottom:10px">
      <div style="color:{MUTED};font-size:13px;margin-bottom:6px">Funcionarios incluidos:</div>
      <ul style="margin:0 0 0 18px;padding:0">{lis}</ul>
    </div>

    <p style="margin:14px 0 0 0;color:{MUTED};font-size:12px">
      Generado el {gen}. Sistema BioAccess – www.bioaccess.cl
    </p>
  </div>
</body>
</html>"""

def _send_pdf_email(to_list, cc_list, subject, text_body, html_body, attach_path) -> bool:
    if not USAR_SMTP:
        messagebox.showinfo("Correo (simulado)", f"Para: {', '.join(to_list)}\nCC: {', '.join(cc_list)}\nAsunto: {subject}")
        return True
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = SMTP_USER
    msg["To"] = ", ".join([t for t in to_list if t])
    if cc_list:
        msg["Cc"] = ", ".join([c for c in cc_list if c])
    msg.set_content(text_body or "Adjunto reporte de asistencia (PDF).")
    if html_body:
        msg.add_alternative(html_body, subtype="html")
    if attach_path and os.path.exists(attach_path):
        ctype, _ = mimetypes.guess_type(attach_path)
        maintype, subtype = (ctype.split("/",1) if ctype else ("application","pdf"))
        with open(attach_path, "rb") as f:
            msg.add_attachment(
                f.read(), maintype=maintype, subtype=subtype,
                filename=os.path.basename(attach_path)
            )
    try:
        with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, timeout=20) as s:
            s.login(SMTP_USER, SMTP_PASS)
            s.send_message(msg)
        return True
    except Exception as e:
        traceback.print_exc()
        messagebox.showerror("Correo", f"No fue posible enviar el correo:\n{e}")
        return False

# ====== PDF builder (horizontal + encabezado profesional) ======
def _build_pdf_asistencia_general(path_pdf, rows_data, period_text):
    """
    rows_data: lista de dicts con
      {"nombre":..., "rut":..., "rango":..., "total_dias":int, "trabajados":int, "obs":str}
    """
    if not PDF_OK:
        raise RuntimeError("ReportLab no está instalado.")

    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_LEFT

    styles   = getSampleStyleSheet()
    normal   = styles["Normal"]
    title    = styles["Title"]; title.alignment = TA_LEFT
    sub      = ParagraphStyle("Sub", parent=normal, fontSize=11, textColor=colors.HexColor("#0b5ea8"))
    note     = ParagraphStyle("Note", parent=normal, fontSize=9,  textColor=colors.HexColor("#6b7280"))
    obs_style= ParagraphStyle("Obs",  parent=normal, fontSize=9,  leading=11, textColor=colors.HexColor("#111827"))

    elems = []

    # ---------- Encabezado profesional ----------
    gen = dt.datetime.now().strftime("%d/%m/%Y %H:%M")

    # Título
    elems.append(Paragraph("Informe de Asistencia de Funcionarios", title))

    # Banda azul separadora
    band = Table([[""]], colWidths=[27*cm], rowHeights=[0.25*cm])
    band.setStyle(TableStyle([("BACKGROUND", (0,0), (-1,-1), colors.HexColor("#0b5ea8"))]))
    elems.append(band)
    elems.append(Spacer(0, 6))

    # Bloque institucional (izq) + metadatos (der)
    left_block = Paragraph(
        "Sistema Control de Horarios <b>BioAccess</b><br/>"
        "Liceo Ignacio Carrera Pinto<br/>"
        "Frutillar",
        sub
    )
    right_block = Paragraph(
        f"<b>Periodo:</b> {period_text}<br/>"
        f"<b>Emitido el:</b> {gen}<br/>"
        f"<b>Elaborado por:</b> BioAccess<br/>"
        f"<b>Contacto:</b> documentos_bd@bioaccess.cl",
        normal
    )
    meta_tbl = Table([[left_block, right_block]], colWidths=[13.2*cm, 13.8*cm])
    meta_tbl.setStyle(TableStyle([
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("LEFTPADDING",  (0,0), (-1,-1), 0),
        ("RIGHTPADDING", (0,0), (-1,-1), 0),
        ("TOPPADDING",   (0,0), (-1,-1), 2),
        ("BOTTOMPADDING",(0,0), (-1,-1), 2),
    ]))
    elems.append(meta_tbl)
    elems.append(Spacer(0, 8))

    # Nota breve
    elems.append(Paragraph(
        "Documento generado automáticamente para fines administrativos. "
        "Revise observaciones y días completos/incompletos en la última columna.",
        note
    ))
    elems.append(Spacer(0, 8))

    # ---------- Tabla principal ----------
    headers = ["Nombre", "RUT", "Días del periodo", "Con ingreso", "Detalle"]
    data = [headers]
    for row in rows_data:
        data.append([
            Paragraph(row["nombre"], normal),
            Paragraph(row["rut"], normal),
            str(row["total_dias"]),
            str(row["trabajados"]),
            Paragraph(row["obs"] or "—", obs_style),
        ])

    # Anchos pensados para A4 landscape con márgenes 36pt (≈27 cm útiles)
    col_w = [8.5*cm, 4.0*cm, 3.0*cm, 3.0*cm, 8.5*cm]  # ≈27 cm total

    tbl = Table(data, colWidths=col_w, repeatRows=1)
    tbl.setStyle(TableStyle([
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE", (0,0), (-1,0), 10),
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#0b5ea8")),
        ("TEXTCOLOR",  (0,0), (-1,0), colors.white),
        ("ALIGN", (2,1), (3,-1), "CENTER"),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("GRID", (0,0), (-1,-1), 0.25, colors.lightgrey),
        ("LEFTPADDING",  (0,0), (-1,-1), 4),
        ("RIGHTPADDING", (0,0), (-1,-1), 4),
        ("TOPPADDING",   (0,0), (-1,-1), 2),
        ("BOTTOMPADDING",(0,0), (-1,-1), 2),
    ]))
    # Filas alternadas sutiles
    for i in range(1, len(data)):
        bg = "#f6f7fb" if i % 2 == 0 else "#ffffff"
        tbl.setStyle(TableStyle([("BACKGROUND", (0,i), (-1,i), colors.HexColor(bg))]))

    elems.append(tbl)
    elems.append(Spacer(0, 10))
    elems.append(Paragraph(f"Generado por BioAccess — {gen}", note))

    # Documento (A4 horizontal + márgenes amplios)
    doc = SimpleDocTemplate(
        path_pdf,
        pagesize=landscape(A4),
        leftMargin=36, rightMargin=36, topMargin=36, bottomMargin=36
    )
    doc.build(elems)


# ================================================================
# ===================   VENTANA PRINCIPAL   ======================
# ================================================================
def abrir_asistencia(app_root, db_path: str, *, default_todos: bool = False):
    TopLevelCls = getattr(ctk, "CTkToplevel", None) or tk.Toplevel
    win = TopLevelCls(app_root)
    win.title("Asistencia de Funcionarios")
    try:
        win.resizable(True, True); win.transient(app_root); win.grab_set()
    except Exception: pass

    cont = ctk.CTkFrame(win)
    cont.pack(fill="both", expand=True, padx=12, pady=12)

    # ---------- fila 1: buscador ----------
    fila1 = ctk.CTkFrame(cont)
    fila1.pack(fill="x", padx=6, pady=(6, 10))

    ctk.CTkLabel(fila1, text="Buscar").grid(row=0, column=0, padx=(0, 8), pady=6, sticky="w")
    entry_buscar = ctk.CTkEntry(fila1, placeholder_text="Nombre o RUT")
    entry_buscar.grid(row=0, column=1, padx=(0, 10), pady=6, sticky="ew")

    ctk.CTkLabel(fila1, text="Funcionario").grid(row=0, column=2, padx=(0, 8), pady=6, sticky="w")
    combo_func = ctk.CTkComboBox(fila1, values=["Cargando..."])
    combo_func.grid(row=0, column=3, padx=(0, 8), pady=6, sticky="ew")

    var_todos = tk.BooleanVar(value=bool(default_todos))
    chk_todos = ctk.CTkCheckBox(fila1, text="Todos", variable=var_todos)
    chk_todos.grid(row=0, column=4, padx=(4, 0), pady=6, sticky="w")

    fila1.grid_columnconfigure(1, weight=1)
    fila1.grid_columnconfigure(3, weight=1)

    con = sqlite3.connect(db_path)
    funcionarios = _leer_funcionarios(con)
    listado = [f"{n} | {r}" for (r, n, _c) in funcionarios]
    combo_func.configure(values=listado if listado else ["(Sin registros)"])
    if listado: combo_func.set(listado[0])

    def filtrar():
        q = entry_buscar.get().strip().lower()
        vals = [x for x in listado if q in x.lower()] if q else listado
        combo_func.configure(values=vals if vals else ["(Sin coincidencias)"])
        if vals: combo_func.set(vals[0])
    entry_buscar.bind("<KeyRelease>", lambda e: filtrar())

    # ---------- fila 2: rango / mes ----------
    fila2 = ctk.CTkFrame(cont)
    fila2.pack(fill="x", padx=6, pady=(0, 10))

    modo_var = tk.StringVar(value="rango")
    r1 = ctk.CTkRadioButton(fila2, text="Rango de fechas", variable=modo_var, value="rango")
    r2 = ctk.CTkRadioButton(fila2, text="Mes completo", variable=modo_var, value="mensual")
    r1.grid(row=0, column=0, padx=(0, 12), pady=6, sticky="w")
    r2.grid(row=0, column=1, padx=(0, 12), pady=6, sticky="w")

    ctk.CTkLabel(fila2, text="Desde").grid(row=1, column=0, padx=(0, 8), pady=6, sticky="e")
    if HAS_TKCAL:
        entry_desde = DateEntry(fila2, date_pattern="yyyy-mm-dd", locale="es_CL")
        entry_hasta = DateEntry(fila2, date_pattern="yyyy-mm-dd", locale="es_CL")
    else:
        entry_desde = tk.Entry(fila2, width=12)
        entry_hasta = tk.Entry(fila2, width=12)
        entry_desde.insert(0, dt.date.today().strftime("%Y-%m-%d"))
        entry_hasta.insert(0, dt.date.today().strftime("%Y-%m-%d"))
    entry_desde.grid(row=1, column=1, padx=(0, 16), pady=6, sticky="w")

    ctk.CTkLabel(fila2, text="Hasta").grid(row=1, column=2, padx=(0, 8), pady=6, sticky="e")
    entry_hasta.grid(row=1, column=3, padx=(0, 16), pady=6, sticky="w")

    ctk.CTkLabel(fila2, text="Año").grid(row=1, column=4, padx=(12, 8), pady=6, sticky="e")
    combo_anio = ctk.CTkComboBox(fila2, values=[str(y) for y in range(dt.date.today().year - 3, dt.date.today().year + 2)])
    combo_anio.set(str(dt.date.today().year))
    combo_anio.grid(row=1, column=5, padx=(0, 16), pady=6, sticky="w")

    ctk.CTkLabel(fila2, text="Mes").grid(row=1, column=6, padx=(0, 8), pady=6, sticky="e")
    meses = [f"{i:02d}" for i in range(1, 13)]
    combo_mes = ctk.CTkComboBox(fila2, values=meses)
    combo_mes.set(f"{dt.date.today().month:02d}")
    combo_mes.grid(row=1, column=7, padx=(0, 0), pady=6, sticky="w")

    btn_prev = ctk.CTkButton(fila2, text="Previsualizar")
    btn_prev.grid(row=1, column=8, padx=(16, 0), pady=6, sticky="w")

    # ---------- tabla ----------
    tabla_frame = ctk.CTkFrame(cont)
    tabla_frame.pack(fill="both", expand=True, padx=6, pady=(0, 10))
    tabla = _construir_tabla_preview(tabla_frame)

    # ---------- botones inferiores ----------
    fila3 = ctk.CTkFrame(cont); fila3.pack(fill="x", padx=6, pady=(0, 6))
    btn_descargar = ctk.CTkButton(fila3, text="Descargar (PDF)", fg_color="#0ea5e9")
    btn_enviar = ctk.CTkButton(fila3, text="Enviar (PDF)", fg_color="#22c55e")
    btn_cancelar = ctk.CTkButton(fila3, text="Cancelar", fg_color="#475569", command=win.destroy)
    btn_descargar.pack(side="left", padx=4); btn_enviar.pack(side="left", padx=4); btn_cancelar.pack(side="right", padx=4)

    # ---------- lógica ----------
    def _obtener_rut_sel():
        if var_todos.get(): return None
        val = combo_func.get().strip()
        if " | " in val:
            _, rut = val.split(" | ", 1)
            return rut.strip()
        for r, n, _c in funcionarios:
            if val == f"{n} | {r}":
                return r
        return None

    def _rango_fechas():
        if modo_var.get() == "rango":
            if HAS_TKCAL:
                f_ini = dt.datetime.strptime(entry_desde.get(), "%Y-%m-%d").date()
                f_fin = dt.datetime.strptime(entry_hasta.get(), "%Y-%m-%d").date()
            else:
                f_ini = dt.datetime.strptime(entry_desde.get().strip(), "%Y-%m-%d").date()
                f_fin = dt.datetime.strptime(entry_hasta.get().strip(), "%Y-%m-%d").date()
            return f_ini, f_fin, f"{f_ini} a {f_fin}", None, (f_fin - f_ini).days + 1
        else:
            y = int(combo_anio.get()); m = int(combo_mes.get())
            f_ini, f_fin, last_day = _mes_range(y, m)
            etiqueta = f"{y} - {MESES_ES[m]}"
            return f_ini, f_fin, etiqueta, (y, m), last_day

    def previsualizar():
        try:
            rut_sel = _obtener_rut_sel()
            f_ini, f_fin, etiqueta_rango, ym, total_dias = _rango_fechas()
            rows = _cargar_registros(con, rut_sel, f_ini, f_fin)
            tabla.delete(*tabla.get_children())

            agg = _resumen_por_rut(rows)

            # Sin movimientos (todos)
            if rut_sel is None and not agg:
                for r, n, _c in funcionarios:
                    tabla.insert("", "end", values=(n, r, etiqueta_rango, total_dias, 0, "Sin movimientos"))
                return

            for rutx, info in (agg.items() if agg else []):
                nombre = info["nombre"]
                dias_trab = len(info["fechas"])
                completos = []
                incompletos = []
                for f, p in sorted(info["por_fecha"].items()):
                    if p["ingreso"] and p["salida"]:
                        completos.append(f)
                    else:
                        falt = []
                        if not p["ingreso"]: falt.append("ingreso")
                        if not p["salida"]:  falt.append("salida")
                        incompletos.append(f"{f} (sin {', '.join(falt)})")
                obs = []
                if completos:   obs.append(f"Días completos: {len(completos)}")
                if incompletos: obs.append("Incompletos: " + "; ".join(incompletos[:6]) + (" ..." if len(incompletos) > 6 else ""))
                if not obs:
                    obs = ["Sin movimientos"] if dias_trab == 0 else ["Registros sin detalle"]
                tabla.insert("", "end", values=(nombre, rutx, etiqueta_rango, total_dias, dias_trab, " | ".join(obs)))

            if rut_sel and not tabla.get_children():
                f = next(((r, n, _c) for (r, n, _c) in funcionarios if r == rut_sel), None)
                if f:
                    _, n, _c = f
                    tabla.insert("", "end", values=(n, rut_sel, etiqueta_rango, total_dias, 0, "Sin movimientos"))

        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("Previsualizar", f"Error al calcular asistencia:\n{e}")

    btn_prev.configure(command=previsualizar)

    def _collect_rows_for_pdf():
        rows = tabla.get_children()
        if not rows:
            return None, ""
        period_text = str(tabla.item(rows[0], "values")[2])
        payload = []
        for iid in rows:
            vals = list(tabla.item(iid, "values"))
            payload.append({
                "nombre": vals[0],
                "rut": vals[1],
                "rango": vals[2],
                "total_dias": vals[3],
                "trabajados": vals[4],
                "obs": vals[5],
            })
        return payload, period_text

    def descargar():
        if not PDF_OK:
            return messagebox.showerror("PDF", "Instala reportlab: pip install reportlab")
        rows, period_text = _collect_rows_for_pdf()
        if not rows:
            return messagebox.showinfo("PDF", "No hay datos para exportar.")
        try:
            path = os.path.join(_downloads_dir(), f"Asistencia_Funcionarios_{dt.datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf")
            _build_pdf_asistencia_general(path, rows, period_text)
            _abrir_archivo(path)
        except Exception as e:
            traceback.print_exc(); messagebox.showerror("PDF", f"No se pudo generar/abrir:\n{e}")
    btn_descargar.configure(command=descargar)

    def enviar_correo():
        if not PDF_OK:
            return messagebox.showerror("PDF", "Instala reportlab: pip install reportlab")
        rows, period_text = _collect_rows_for_pdf()
        if not rows:
            return messagebox.showinfo("Correo", "No hay datos para enviar. Primero previsualiza el reporte.")

        # Genera PDF temporal
        try:
            path = os.path.join(_downloads_dir(), f"Asistencia_Funcionarios_{dt.datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf")
            _build_pdf_asistencia_general(path, rows, period_text)
        except Exception as e:
            traceback.print_exc(); return messagebox.showerror("Correo", f"No se pudo preparar el PDF:\n{e}")

        # Recolectar correos sugeridos (según filas mostradas)
        ruts_tabla = [r["rut"] for r in rows]
        to_sugeridos = _emails_para_ruts(con, ruts_tabla)

        # Diálogo CTk envío (misma estética)
        TopSend = getattr(ctk, "CTkToplevel", None) or tk.Toplevel
        dlg = TopSend(win); dlg.title("Enviar por correo (PDF)")
        try:
            dlg.transient(win); dlg.grab_set(); dlg.resizable(False, False)
        except Exception:
            pass

        box = ctk.CTkFrame(dlg, corner_radius=12)
        box.pack(fill="both", expand=True, padx=16, pady=16)

        ctk.CTkLabel(box, text="Enviar reporte — Asistencia de Funcionarios", font=("Arial", 16, "bold")).grid(row=0, column=0, columnspan=2, pady=(0,8), sticky="w")

        ctk.CTkLabel(box, text="Para:").grid(row=1, column=0, sticky="e", padx=(0,8), pady=6)
        entry_to = ctk.CTkEntry(box, width=480, placeholder_text="correo1@dominio.cl, correo2@dominio.cl")
        entry_to.grid(row=1, column=1, sticky="w", pady=6)
        if to_sugeridos:
            entry_to.insert(0, ", ".join(to_sugeridos))

        ctk.CTkLabel(box, text="CC:").grid(row=2, column=0, sticky="e", padx=(0,8), pady=(0,6))
        entry_cc = ctk.CTkEntry(box, width=480, placeholder_text="(opcional) separar por coma")
        entry_cc.grid(row=2, column=1, sticky="w", pady=(0,6))

        ctk.CTkLabel(
            box,
            text=f"Periodo: {period_text}   |   Funcionarios incluidos: {len(rows)}",
            text_color="#9ca3af"
        ).grid(row=3, column=0, columnspan=2, sticky="w", pady=(4,2))

        btns = ctk.CTkFrame(box, fg_color="transparent")
        btns.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(10,0))
        btn_cancel = ctk.CTkButton(btns, text="Cancelar", fg_color="#475569", command=dlg.destroy, width=120)
        btn_send2 = ctk.CTkButton(btns, text="Enviar", fg_color="#22c55e", width=160)
        btn_cancel.pack(side="left", padx=(0,6))
        btn_send2.pack(side="right", padx=(6,0))
        box.grid_columnconfigure(1, weight=1)

        def do_send_now():
            to_raw = entry_to.get().strip()
            cc_raw = entry_cc.get().strip()
            to_list = [x.strip() for x in to_raw.split(",") if x.strip()]
            cc_list = [x.strip() for x in cc_raw.split(",") if x.strip()]
            if not to_list:
                return messagebox.showerror("Correo", "Indica al menos un destinatario en 'Para'.")

            gen = dt.datetime.now().strftime("%d-%m-%Y %H:%M")
            intro = (
                "Estimado(a), junto con saludar, se remite el reporte de Asistencia de Funcionarios.\n"
                "Este informe es generado automáticamente por BioAccess – Control de Horarios.\n"
                "Se adjunta el documento PDF para su revisión.\n"
            )
            text_body = (
                f"{intro}\n"
                f"Periodo  : {period_text}\n"
                f"Funcionarios incluidos: {len(rows)}\n"
                f"Generado : {gen}\n\n"
                f"BioAccess."
            )
            listado_nombrerut = [(r["nombre"], r["rut"]) for r in rows]
            html_body = _html_email_asistencia_general(period_text, listado_nombrerut)
            subject   = f"Asistencia de Funcionarios — {period_text}"

            ok = _send_pdf_email(to_list, cc_list, subject, text_body, html_body, path)
            if ok:
                messagebox.showinfo("Correo", "Reporte enviado correctamente.")
                try:
                    dlg.destroy()
                except Exception:
                    pass

        btn_send2.configure(command=do_send_now)

        # Centrar diálogo
        win.update_idletasks()
        try:
            w = max(640, dlg.winfo_reqwidth() + 32)
            h = max(260, dlg.winfo_reqheight() + 16)
        except Exception:
            w, h = 640, 260
        x = win.winfo_rootx() + (win.winfo_width() // 2) - (w // 2)
        y = win.winfo_rooty()  + (win.winfo_height() // 2) - (h // 2)
        dlg.geometry(f"{w}x{h}+{max(x,0)}+{max(y,0)}")

    btn_enviar.configure(command=enviar_correo)

    # Centrar ventana principal
    app_root.update_idletasks()
    w,h = 980, 620
    x = app_root.winfo_x() + (app_root.winfo_width() // 2) - (w // 2)
    y = app_root.winfo_y() + (app_root.winfo_height() // 2) - (h // 2)
    win.geometry(f"{w}x{h}+{max(x,0)}+{max(y,0)}")
