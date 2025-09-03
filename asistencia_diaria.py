# asistencia_diaria.py
import os, sys, sqlite3, calendar, datetime as dt, tkinter as tk
from tkinter import ttk, messagebox
import customtkinter as ctk

# ===== PDF =====
try:
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib import colors
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.graphics.shapes import Drawing, Circle, Polygon, String
    PDF_OK = True
except Exception:
    PDF_OK = False

# ===== Email =====
import smtplib, mimetypes, subprocess, traceback
from email.message import EmailMessage

MESES_ES = {1:"Enero",2:"Febrero",3:"Marzo",4:"Abril",5:"Mayo",6:"Junio",7:"Julio",8:"Agosto",9:"Septiembre",10:"Octubre",11:"Noviembre",12:"Diciembre"}

def _downloads_dir():
    home = os.path.expanduser("~")
    cand = os.path.join(home, "Downloads")
    if os.path.isdir(cand): return cand
    xdg = os.environ.get("XDG_DOWNLOAD_DIR")
    if xdg and os.path.isdir(xdg): return xdg
    return home

def _abrir_archivo(path):
    try:
        if sys.platform.startswith("win"): os.startfile(path)  # type: ignore
        elif sys.platform == "darwin": subprocess.Popen(["open", path])
        else: subprocess.Popen(["xdg-open", path])
    except Exception as e:
        messagebox.showinfo("Archivo generado", f"Se guardó en:\n{path}\n\nNo se pudo abrir automáticamente: {e}")

def _mes_range(year, month):
    last_day = calendar.monthrange(year, month)[1]
    return dt.date(year, month, 1), dt.date(year, month, last_day), last_day

def _feriados_set_range(con, d1: dt.date, d2: dt.date):
    cur = con.cursor()
    cur.execute(
        "SELECT fecha FROM feriados WHERE fecha BETWEEN ? AND ?",
        (d1.strftime("%Y-%m-%d"), d2.strftime("%Y-%m-%d"))
    )
    return {row[0] for row in cur.fetchall()}

def _panel_salida_autorizada_set(con, d1: dt.date, d2: dt.date):
    cur = con.cursor()
    try:
        cur.execute(
            "SELECT fecha FROM panel_flags WHERE fecha BETWEEN ? AND ? AND salida_anticipada=1",
            (d1.strftime("%Y-%m-%d"), d2.strftime("%Y-%m-%d"))
        )
        return {row[0] for row in cur.fetchall()}
    except Exception:
        return set()

def _maximize_without_covering_taskbar(win):
    """
    Maximiza usando el área de trabajo (work area) para no quedar debajo de la barra de tareas (Windows).
    """
    ok = False
    if sys.platform.startswith("win"):
        try:
            import ctypes
            from ctypes import wintypes
            SPI_GETWORKAREA = 0x0030
            rect = wintypes.RECT()
            ctypes.windll.user32.SystemParametersInfoW(SPI_GETWORKAREA, 0, ctypes.byref(rect), 0)
            width = rect.right - rect.left
            height = rect.bottom - rect.top
            win.geometry(f"{width}x{height}+{rect.left}+{rect.top}")
            win.update_idletasks()
            ok = True
        except Exception:
            ok = False
    if not ok:
        try:
            win.state("zoomed"); ok = True
        except Exception:
            pass
    if not ok:
        try:
            win.attributes("-zoomed", True); ok = True
        except Exception:
            pass
    if not ok:
        try:
            sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
            win.geometry(f"{sw}x{sh}+0+0"); win.update_idletasks()
        except Exception:
            pass

def _style_dark_treeview():
    style = ttk.Style()
    try: style.theme_use('clam')
    except Exception: pass

    style.configure(
        "Dark.Treeview",
        background="#111418",
        foreground="#e8eef5",
        fieldbackground="#111418",
        rowheight=30,
        bordercolor="#2b3440",
        borderwidth=0,
        font=("Segoe UI", 11)
    )
    style.configure(
        "Dark.Treeview.Heading",
        background="#1e293b",
        foreground="#f1f5f9",
        relief="flat",
        font=("Segoe UI Semibold", 10)
    )
    style.map(
        "Dark.Treeview",
        background=[("selected", "#0b4a6e")],
        foreground=[("selected", "#ffffff")]
    )
    return "Dark.Treeview"

def _leer_funcionarios(con):
    cur = con.cursor()
    cur.execute("""
        SELECT rut, (nombre || ' ' || apellido) AS nombre_completo
        FROM trabajadores
        ORDER BY nombre_completo COLLATE NOCASE;
    """)
    return cur.fetchall()

def _estado_en_fecha(con, rut: str, fecha: dt.date, salida_autorizada_set):
    """
    Devuelve estado para esa fecha:
      'OK' (ingreso y salida),
      'SI' (solo ingreso),
      'PS' (pendiente salida con autorización activa),
      '-'  (ausente).
    Soporta esquema nuevo (hora_ingreso/hora_salida) y clásico (tipo/hora).
    """
    f_iso = fecha.strftime("%Y-%m-%d")
    cur = con.cursor()
    cur.execute("""
        SELECT
            MAX(CASE
                WHEN hora_ingreso IS NOT NULL AND TRIM(hora_ingreso) <> '' THEN 1
                WHEN ( (lower(IFNULL(tipo,'')) LIKE 'ing%%' OR lower(IFNULL(tipo,'')) LIKE 'ent%%')
                       AND hora IS NOT NULL AND TRIM(hora) <> '' ) THEN 1
                ELSE 0
            END) AS tiene_ing,
            MAX(CASE
                WHEN hora_salida IS NOT NULL AND TRIM(hora_salida) <> '' THEN 1
                WHEN ( (lower(IFNULL(tipo,'')) LIKE 'sal%%' OR lower(IFNULL(tipo,'')) LIKE 'ret%%')
                       AND hora IS NOT NULL AND TRIM(hora) <> '' ) THEN 1
                ELSE 0
            END) AS tiene_sal
        FROM registros
        WHERE rut = ? AND DATE(fecha) = DATE(?)
    """, (rut, f_iso))
    row = cur.fetchone()
    ing, sal = (row or (0, 0))
    ing = int(ing or 0); sal = int(sal or 0)
    if ing:
        if sal:
            return "OK"
        else:
            return "PS" if (f_iso in salida_autorizada_set) else "SI"
    return "-"

# ===== Iconos vectoriales (PDF) =====
from reportlab.lib import colors as _rl_colors
def _icon_check(size=9, color=_rl_colors.black):
    d = Drawing(size, size); d.add(String(0, 0, "✓", fontSize=size, fillColor=color)); return d
def _icon_dash(size=9, color=_rl_colors.black):
    d = Drawing(size, size); d.add(String(0, 0, "-", fontSize=size, fillColor=color)); return d

def abrir_asistencia_diaria(app_root, db_path: str):
    Top = getattr(ctk, "CTkToplevel", None) or tk.Toplevel
    win = Top(app_root); win.title("Asistencia Diaria (Matriz)")
    _maximize_without_covering_taskbar(win)
    try:
        win.resizable(True, True); win.transient(app_root); win.grab_set()
        win.minsize(1000, 580)
    except Exception:
        pass

    cont = ctk.CTkFrame(win); cont.pack(fill="both", expand=True, padx=12, pady=(12, 20))

    # --- Controles superiores ---
    fila = ctk.CTkFrame(cont); fila.pack(fill="x", padx=6, pady=(6,10))

    # Fila 0: Buscar / Lista / Todos
    ctk.CTkLabel(fila, text="Buscar").grid(row=0, column=0, padx=(0,8), pady=6, sticky="w")
    entry_buscar = ctk.CTkEntry(fila, placeholder_text="Nombre o RUT")
    entry_buscar.grid(row=0, column=1, padx=(0,10), pady=6, sticky="ew")

    ctk.CTkLabel(fila, text="Funcionario(s)").grid(row=0, column=2, padx=(0,8), pady=6, sticky="w")
    listbox = tk.Listbox(fila, selectmode="extended", height=6, bg="#1e1e1e", fg="#e5e5e5")
    listbox.grid(row=0, column=3, padx=(0,8), pady=6, sticky="nsew")

    var_todos = tk.BooleanVar(value=True)
    ctk.CTkCheckBox(fila, text="Todos", variable=var_todos).grid(row=0, column=4, padx=(4,0), pady=6, sticky="w")

    # Fila 1: Modo + Año/Mes o Rango + Previsualizar
    ctk.CTkLabel(fila, text="Por mes / Rango de Días").grid(row=1, column=0, padx=(0,8), pady=6, sticky="e")
    combo_modo = ctk.CTkComboBox(fila, values=["Mes", "Rango"])
    combo_modo.set("Mes"); combo_modo.grid(row=1, column=1, padx=(0,12), pady=6, sticky="w")

    # Controles de Mes
    ctk.CTkLabel(fila, text="Año").grid(row=1, column=2, padx=(0,8), pady=6, sticky="e")
    combo_anio = ctk.CTkComboBox(fila, values=[str(y) for y in range(dt.date.today().year-3, dt.date.today().year+2)])
    combo_anio.set(str(dt.date.today().year)); combo_anio.grid(row=1, column=3, padx=(0,16), pady=6, sticky="w")

    ctk.CTkLabel(fila, text="Mes").grid(row=1, column=4, padx=(0,8), pady=6, sticky="e")
    combo_mes = ctk.CTkComboBox(fila, values=[MESES_ES[i] for i in range(1,13)])
    combo_mes.set(MESES_ES[dt.date.today().month]); combo_mes.grid(row=1, column=5, padx=(0,8), pady=6, sticky="w")

    # Controles de Rango
    lbl_desde = ctk.CTkLabel(fila, text="Desde (YYYY-MM-DD)")
    entry_desde = ctk.CTkEntry(fila, placeholder_text="YYYY-MM-DD")
    lbl_hasta = ctk.CTkLabel(fila, text="Hasta (YYYY-MM-DD)")
    entry_hasta = ctk.CTkEntry(fila, placeholder_text="YYYY-MM-DD")

    # Botón previsualizar (misma fila)
    btn_prev = ctk.CTkButton(fila, text="Previsualizar")
    btn_prev.grid(row=1, column=6, padx=(12,0), pady=6, sticky="w")

    # Distribución de columnas
    fila.grid_columnconfigure(1, weight=1)  # buscador
    fila.grid_columnconfigure(3, weight=1)
    fila.grid_columnconfigure(5, weight=1)

    # Toggle de modo
    def toggle_modo(*_):
        modo = combo_modo.get().lower()
        if modo == "rango":
            # Oculta Año/Mes, muestra Desde/Hasta en las mismas columnas
            combo_anio.grid_remove(); combo_mes.grid_remove()
            lbl_desde.grid(row=1, column=2, padx=(0,8), pady=6, sticky="e")
            entry_desde.grid(row=1, column=3, padx=(0,16), pady=6, sticky="w")
            lbl_hasta.grid(row=1, column=4, padx=(0,8), pady=6, sticky="e")
            entry_hasta.grid(row=1, column=5, padx=(0,8), pady=6, sticky="w")
        else:
            # Muestra Año/Mes, oculta Desde/Hasta
            lbl_desde.grid_remove(); entry_desde.grid_remove()
            lbl_hasta.grid_remove(); entry_hasta.grid_remove()
            combo_anio.grid(row=1, column=3, padx=(0,16), pady=6, sticky="w")
            combo_mes.grid(row=1, column=5, padx=(0,8), pady=6, sticky="w")
    combo_modo.configure(command=toggle_modo)
    toggle_modo()

    # ---- Datos base ----
    con = sqlite3.connect(db_path)
    funcionarios = _leer_funcionarios(con)
    items = [f"{n} | {r}" for (r, n) in funcionarios]
    for it in items: listbox.insert("end", it)

    def filtrar_lista(*_):
        q = entry_buscar.get().strip().lower()
        listbox.delete(0, "end")
        for it in items:
            if q in it.lower(): listbox.insert("end", it)
    entry_buscar.bind("<KeyRelease>", filtrar_lista)

    # --- Tabla ---
    tabla = ttk.Treeview(cont, show="headings", height=16, style=_style_dark_treeview())
    tabla.pack(fill="both", expand=True, padx=10, pady=(4,10))

    # --- Pie: leyenda + botones ---
    bottom = ctk.CTkFrame(cont); bottom.pack(fill="x", padx=8, pady=(0,14))
    bottom.pack_propagate(False); bottom.configure(height=54)

    ctk.CTkLabel(
        bottom,
        text=("Leyenda: Celdas → ✓ Jornada cerrada · S/I Solo Ingreso · P/S Pendiente salida · - Ausente"),
        text_color="#cbd5e1"
    ).pack(side="left", padx=6)

    btn_pdf  = ctk.CTkButton(bottom, text="Descargar (PDF)", width=148, fg_color="#0ea5e9")
    btn_send = ctk.CTkButton(bottom, text="Enviar (PDF)",     width=148, fg_color="#22c55e")
    btn_close= ctk.CTkButton(bottom, text="Cancelar (salir)", width=148, fg_color="#475569", command=win.destroy)

    btn_close.pack(side="right", padx=4)
    btn_send.pack(side="right",  padx=4)
    btn_pdf.pack(side="right",   padx=4)

    def _seleccionados():
        if var_todos.get(): return [r for (r, _n) in funcionarios]
        sel = [listbox.get(i) for i in listbox.curselection()]
        ruts = []
        for v in sel:
            if " | " in v:
                _, r = v.split(" | ", 1); ruts.append(r.strip())
        return ruts

    # ===== PREVISUALIZAR =====
    def previsualizar():
        try:
            modo = combo_modo.get().lower()
            if modo == "rango":
                try:
                    d1 = dt.datetime.strptime(entry_desde.get().strip(), "%Y-%m-%d").date()
                    d2 = dt.datetime.strptime(entry_hasta.get().strip(), "%Y-%m-%d").date()
                except Exception:
                    return messagebox.showerror("Rango", "Indica fechas válidas (YYYY-MM-DD) en Desde y Hasta.")
                if d1 > d2:
                    return messagebox.showerror("Rango", "La fecha 'Desde' no puede ser mayor que 'Hasta'.")
                date_list = [d1 + dt.timedelta(days=i) for i in range((d2 - d1).days + 1)]
                period_text = f"{d1.strftime('%Y-%m-%d')} — {d2.strftime('%Y-%m-%d')}"
            else:
                y = int(combo_anio.get())
                m = next((k for k,v in MESES_ES.items() if v == combo_mes.get()), dt.date.today().month)
                s, e, last_day = _mes_range(y, m)
                date_list = [dt.date(y, m, d) for d in range(1, last_day+1)]
                period_text = f"{MESES_ES[m]} {y}"

            # Feriados y autorización de salida por rango
            set_fer = _feriados_set_range(con, date_list[0], date_list[-1])
            salida_aut_set = _panel_salida_autorizada_set(con, date_list[0], date_list[-1])

            # Tipos de día
            day_types = []
            for d in date_list:
                f_iso = d.strftime("%Y-%m-%d")
                if f_iso in set_fer: day_types.append("F")
                else:
                    wd = d.weekday()
                    day_types.append("S" if wd==5 else ("D" if wd==6 else "N"))

            # Columnas
            cols = ["nombre", "rut"] + [str(i+1) for i in range(len(date_list))]
            tabla["columns"] = cols
            tabla.delete(*tabla.get_children())

            # Encabezados (solo números)
            tabla.heading("nombre", text="Nombre")
            tabla.column("nombre", width=300, anchor="w")
            tabla.heading("rut", text="RUT")
            tabla.column("rut", width=160, anchor="w")
            for i in range(len(date_list)):
                col_id = str(i+1)
                tabla.heading(col_id, text=str(i+1))
                tabla.column(col_id, width=36, anchor="center")

            ruts = _seleccionados()
            if not ruts:
                messagebox.showinfo("Asistencia diaria","No hay funcionarios seleccionados."); return

            # Totales para el subtítulo
            sab = sum(1 for t in day_types if t=="S")
            dom = sum(1 for t in day_types if t=="D")
            fer = sum(1 for t in day_types if t=="F")
            tabla.heading("nombre", text=f"Nombre  *  Sábados: {sab}, Domingos: {dom}, Feriados: {fer}")

            # Filas con estados
            for r, n in funcionarios:
                if r not in ruts: continue
                vals = [n, r]
                for d in date_list:
                    estado = _estado_en_fecha(con, r, d, salida_aut_set)
                    if estado == "OK": vals.append("✓")
                    elif estado == "SI": vals.append("S/I")
                    elif estado == "PS": vals.append("P/S")
                    else: vals.append("-")
                tabla.insert("", "end", values=vals)

            # Guardar info para PDF
            tabla._date_list = date_list          # type: ignore
            tabla._period_text = period_text      # type: ignore
            tabla._day_types = day_types          # type: ignore

        except Exception as e:
            traceback.print_exc(); messagebox.showerror("Asistencia diaria", f"Error al generar matriz:\n{e}")

    btn_prev.configure(command=previsualizar)

    # ===== PDF =====
    def _build_pdf(path_pdf, data_rows, period_text, date_list, day_types):
        styles = getSampleStyleSheet()
        elems = []

        # Encabezado institucional
        elems.append(Paragraph(f"<b>Asistencia Diaria — {period_text}</b>", styles["Title"]))
        elems.append(Paragraph("Sistema Control de Horarios <b>BioAccess</b>", styles["Normal"]))
        elems.append(Paragraph("Liceo Ignacio Carrera Pinto", styles["Normal"]))
        elems.append(Paragraph("Frutillar", styles["Normal"]))
        elems.append(Spacer(0, 12))

        # Ficha cuando es un solo funcionario
        if len(data_rows) == 1:
            row = data_rows[0]
            nombre = row["nombre"]; rut = row["rut"]
            cargo = ""
            try:
                conx = sqlite3.connect(db_path)
                curx = conx.cursor()
                curx.execute("PRAGMA table_info(trabajadores)")
                cols = {c[1].lower() for c in curx.fetchall()}
                if "profesion" in cols:
                    curx.execute("SELECT profesion FROM trabajadores WHERE rut = ?", (rut,))
                    r = curx.fetchone()
                    if r and r[0]: cargo = str(r[0])
                conx.close()
            except Exception:
                pass
            elems.append(Paragraph(f"<b>Nombre:</b> {nombre}", styles["Normal"]))
            elems.append(Paragraph(f"<b>RUT:</b> {rut}", styles["Normal"]))
            elems.append(Paragraph(f"<b>Cargo:</b> {cargo or '—'}", styles["Normal"]))
            elems.append(Paragraph(f"<b>Periodo:</b> {period_text}", styles["Normal"]))
            elems.append(Spacer(0, 10))

        # Leyenda
        elems.append(Paragraph(
            "Leyenda: ✓ Jornada cerrada · I = S/I Solo Ingreso · P = P/S Pendiente salida · - Ausente "
            "· Encabezados en color: <font color='blue'>Sábado</font>, "
            "<font color='red'>Domingo</font>, <font color='gold'>Feriado</font>",
            styles["Normal"]
        ))
        elems.append(Spacer(0, 8))

        # Encabezados (sin '-'), usando índice 1..n
        day_headers = [str(i+1) for i in range(len(date_list))]
        table_data = [["Nombre", "RUT"] + day_headers]

        # Celdas: ✓, I, P, -
        def cell_for(status):
            if status == "OK": return _icon_check(9)
            if status == "SI": return "I"
            if status == "PS": return "P"
            return _icon_dash(9)

        for row in data_rows:
            f = [row["nombre"], row["rut"]]
            for status in row["flags"]:  # "OK"/"SI"/"PS"/"-"
                f.append(cell_for(status))
            table_data.append(f)

        name_w, rut_w = 220, 120
        day_w = 16
        tbl = Table(table_data, colWidths=[name_w, rut_w] + [day_w]*len(day_headers))

        base = [
            ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
            ("FONTSIZE", (0,0), (-1,0), 10),
            ("FONTSIZE", (0,1), (-1,-1), 9),
            ("ALIGN", (2,1), (-1,-1), "CENTER"),
            ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
            ("GRID", (0,0), (-1,-1), 0.25, colors.lightgrey),
            ("LEFTPADDING", (0,0), (-1,-1), 3),
            ("RIGHTPADDING", (0,0), (-1,-1), 3),
            ("TOPPADDING", (0,0), (-1,-1), 2),
            ("BOTTOMPADDING", (0,0), (-1,-1), 2),
        ]
        # Color SOLO en el texto del encabezado según S/D/F
        for i, t in enumerate(day_types, start=2):
            if t == "S":
                base.append(("TEXTCOLOR", (i,0), (i,0), colors.blue))
            elif t == "D":
                base.append(("TEXTCOLOR", (i,0), (i,0), colors.red))
            elif t == "F":
                base.append(("TEXTCOLOR", (i,0), (i,0), colors.gold))

        tbl.setStyle(TableStyle(base))
        elems.append(tbl)

        # ===== Resumen por funcionario =====
        elems.append(Spacer(0, 12))
        elems.append(Paragraph("<b>Resumen por funcionario</b>", styles["Heading3"]))
        elems.append(Paragraph("Con ingreso = ✓ + I + P", styles["Italic"]))
        elems.append(Spacer(0, 6))

        total_days = len(day_headers)
        sum_rows = [["Nombre", "RUT", "Días", "Con ingreso", "Jornada cerrada", "Solo ingreso", "Pend. salida", "Ausente"]]
        agg_ing = agg_ok = agg_si = agg_ps = agg_abs = 0

        for row in data_rows:
            flags = row["flags"]
            c_ok = sum(1 for s in flags if s == "OK")
            c_si = sum(1 for s in flags if s == "SI")
            c_ps = sum(1 for s in flags if s == "PS")
            c_abs = sum(1 for s in flags if s == "-")
            c_ing = total_days - c_abs  # OK + SI + PS

            agg_ok += c_ok; agg_si += c_si; agg_ps += c_ps; agg_abs += c_abs; agg_ing += c_ing
            sum_rows.append([row["nombre"], row["rut"], total_days, c_ing, c_ok, c_si, c_ps, c_abs])

        # Totales
        sum_rows.append([
            "Totales", "",
            total_days * len(data_rows),
            agg_ing, agg_ok, agg_si, agg_ps, agg_abs
        ])

        # Tabla de resumen
        sum_name_w, sum_rut_w = 220, 120
        sum_num_w = 70
        tbl_sum = Table(sum_rows, colWidths=[sum_name_w, sum_rut_w] + [sum_num_w]*6)
        tbl_sum.setStyle(TableStyle([
            ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
            ("BACKGROUND", (0,0), (-1,0), colors.whitesmoke),
            ("ALIGN", (2,1), (-1,-1), "CENTER"),
            ("GRID", (0,0), (-1,-1), 0.25, colors.lightgrey),
            ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
            ("FONTSIZE", (0,0), (-1,-1), 9),
            ("FONTNAME", (0,-1), (-1,-1), "Helvetica-Bold"),  # Totales en negrita
        ]))
        elems.append(tbl_sum)

        # Pie
        try:
            ahora = dt.datetime.now().strftime("%d/%m/%Y %H:%M")
            elems.append(Spacer(0, 6))
            elems.append(Paragraph(f"Generado por BioAccess — {ahora}", styles["Italic"]))
        except Exception:
            pass

        doc = SimpleDocTemplate(
            path_pdf,
            pagesize=landscape(A4),
            leftMargin=36, rightMargin=36, topMargin=36, bottomMargin=36
        )
        doc.build(elems)

    def _collect_for_pdf():
        rows = tabla.get_children()
        if not rows: return None
        period_text = getattr(tabla, "_period_text", None)
        date_list = getattr(tabla, "_date_list", None)
        day_types = getattr(tabla, "_day_types", None)

        if not (period_text and date_list and day_types):
            # Fallback a mes actual (no debería ocurrir si se previsualiza)
            y = int(dt.date.today().year); m = int(dt.date.today().month)
            s, e, last_day = _mes_range(y, m)
            date_list = [dt.date(y, m, d) for d in range(1, last_day+1)]
            period_text = f"{MESES_ES[m]} {y}"
            set_fer = _feriados_set_range(con, date_list[0], date_list[-1])
            day_types = []
            for d in date_list:
                f_iso = d.strftime("%Y-%m-%d")
                if f_iso in set_fer: day_types.append("F")
                else:
                    wd = d.weekday()
                    day_types.append("S" if wd==5 else ("D" if wd==6 else "N"))

        data = []
        for iid in rows:
            vals = list(tabla.item(iid, "values"))
            nombre, rut = vals[0], vals[1]
            flags = []
            for cell in vals[2:]:
                s = str(cell).strip().upper()
                if s in ("✅","✓"): flags.append("OK")
                elif s.startswith("S"): flags.append("SI")
                elif s.startswith("P"): flags.append("PS")
                else: flags.append("-")
            data.append({"nombre": nombre, "rut": rut, "flags": flags})
        return data, period_text, date_list, day_types

    def descargar_pdf():
        if not PDF_OK: return messagebox.showerror("PDF","Instala reportlab: pip install reportlab")
        pack = _collect_for_pdf()
        if not pack: return messagebox.showinfo("PDF","No hay datos para exportar.")
        data, period_text, date_list, day_types = pack
        path = os.path.join(_downloads_dir(), f"Asistencia_Diaria_{dt.datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf")
        try: _build_pdf(path, data, period_text, date_list, day_types); _abrir_archivo(path)
        except Exception as e: traceback.print_exc(); messagebox.showerror("PDF", f"No se pudo generar/abrir:\n{e}")

    btn_pdf.configure(command=descargar_pdf)

    def enviar_pdf():
        if not PDF_OK: return messagebox.showerror("PDF","Instala reportlab: pip install reportlab")
        pack = _collect_for_pdf()
        if not pack: return messagebox.showinfo("Enviar","No hay datos para enviar.")
        data, period_text, date_list, day_types = pack
        path = os.path.join(_downloads_dir(), f"Asistencia_Diaria_{dt.datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf")
        try: _build_pdf(path, data, period_text, date_list, day_types)
        except Exception as e: traceback.print_exc(); return messagebox.showerror("Enviar", f"No se pudo preparar PDF:\n{e}")

        top = tk.Toplevel(win); top.title("Enviar por correo (PDF)")
        tk.Label(top, text="Para:").grid(row=0, column=0, padx=8, pady=8, sticky="e")
        entry_to = tk.Entry(top, width=40); entry_to.grid(row=0, column=1, padx=8, pady=8, sticky="w")
        def _do_send():
            to = entry_to.get().strip()
            if not to: return messagebox.showerror("Enviar","Indica un destinatario.")
            try:
                SMTP_HOST = os.environ.get("SMTP_HOST",""); SMTP_PORT = int(os.environ.get("SMTP_PORT","587"))
                SMTP_USER = os.environ.get("SMTP_USER",""); SMTP_PASS = os.environ.get("SMTP_PASS","")
                if not (SMTP_HOST and SMTP_USER and SMTP_PASS):
                    return messagebox.showwarning("Enviar","Falta configuración SMTP (SMTP_HOST/SMTP_USER/SMTP_PASS).")
                msg = EmailMessage()
                msg["Subject"] = f"Asistencia Diaria — {period_text}"; msg["From"]=SMTP_USER; msg["To"]=to
                msg.set_content("Adjunto reporte de asistencia diaria (PDF) generado desde BioAccess.")
                ctype,_ = mimetypes.guess_type(path); maintype, subtype = (ctype.split("/",1) if ctype else ("application","pdf"))
                with open(path,"rb") as f: msg.add_attachment(f.read(), maintype=maintype, subtype=subtype, filename=os.path.basename(path))
                with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=20) as s:
                    s.starttls(); s.login(SMTP_USER, SMTP_PASS); s.send_message(msg)
                messagebox.showinfo("Enviar","Enviado correctamente."); top.destroy()
            except Exception as e:
                traceback.print_exc(); messagebox.showerror("Enviar", f"No se pudo enviar:\n{e}")
        tk.Button(top, text="Cancelar", command=top.destroy).grid(row=1, column=0, padx=8, pady=(0,8), sticky="ew")
        tk.Button(top, text="Enviar", command=_do_send).grid(row=1, column=1, padx=8, pady=(0,8), sticky="ew")
        top.grab_set(); top.transient(win); entry_to.focus_set()

    btn_send.configure(command=enviar_pdf)
