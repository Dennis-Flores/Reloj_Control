# solicitudes.py
import os, sys, sqlite3, datetime, mimetypes, smtplib
import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox
from email.message import EmailMessage

__all__ = ["construir_solicitudes"]

# ---- PDF opcional ----
try:
    from PyPDF2 import PdfReader, PdfWriter
except Exception:
    PdfReader = None
    PdfWriter = None

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

# -------------------- Email --------------------
DESTINATARIOS = [
    "dennis.flores@slepllanquihue.cl",
    
    
]

# Cuenta oficial BioAccess (SMTP SSL directo 465)
SMTP_HOST = "mail.bioaccess.cl"
SMTP_PORT = 465
SMTP_USER = "solicitud_reloj_control@bioaccess.cl"
SMTP_PASS = "@solicitud2026"
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
#                    PDF helpers (opcional)
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

def nombre_archivo_por_formato(fecha: datetime.date, folio: int, usar_ddmmyyyy=True):
    fecha_str = fecha.strftime("%d%m%Y") if usar_ddmmyyyy else fecha.strftime("%Y%m%d")
    return f"{fecha_str}_F{folio:06}.pdf"

# =========================================================
#                    EMAIL
# =========================================================
def enviar_correo(destinatarios, asunto: str, cuerpo: str, adjunto: str | None = None):
    if not USAR_SMTP:
        info = f"Se habría enviado a: {', '.join(destinatarios)}\n\nASUNTO:\n{asunto}\n\nCUERPO:\n{cuerpo}"
        if adjunto:
            info += f"\n\nAdjunto: {adjunto}"
        messagebox.showinfo("Envío no configurado", info)
        return

    msg = EmailMessage()
    msg["Subject"] = asunto
    msg["From"] = SMTP_USER
    msg["To"] = ", ".join(destinatarios)
    msg.set_content(cuerpo)

    if adjunto and os.path.exists(adjunto):
        ctype, _ = mimetypes.guess_type(adjunto)
        maintype, subtype = (ctype.split("/", 1) if ctype else ("application", "octet-stream"))
        with open(adjunto, "rb") as f:
            msg.add_attachment(f.read(), maintype=maintype, subtype=subtype, filename=os.path.basename(adjunto))

    # IMPORTANTE: SSL directo en puerto 465
    with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT) as s:
        s.login(SMTP_USER, SMTP_PASS)
        s.send_message(msg)

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
    # Siempre visible, pero editable solo cuando corresponda
    frame_horas = ctk.CTkFrame(form)
    frame_horas.grid(row=6, column=0, columnspan=3, sticky="w", padx=0, pady=(0, 0))

    # Cantidad de horas
    lbl_cant = ctk.CTkLabel(frame_horas, text="Cantidad de horas:", width=LABEL_W, anchor="e")
    lbl_cant.grid(row=0, column=0, sticky="e", padx=LABEL_PADX)
    horas_values = [str(i) for i in range(1, 11)]
    cmb_cantidad_horas = ctk.CTkOptionMenu(frame_horas, values=horas_values, width=100)
    cmb_cantidad_horas.set(horas_values[0])
    cmb_cantidad_horas.grid(row=0, column=1, sticky="w")

    # Hora inicio
    lbl_inicio = ctk.CTkLabel(frame_horas, text="Hora inicio (HH:MM):", width=LABEL_W, anchor="e")
    lbl_inicio.grid(row=1, column=0, sticky="e", padx=LABEL_PADX, pady=(8, 0))
    entry_hora_inicio = ctk.CTkEntry(frame_horas, placeholder_text="HH:MM", width=120)
    entry_hora_inicio.grid(row=1, column=1, sticky="w", pady=(8, 0))

    # Hora término (auto)
    lbl_fin = ctk.CTkLabel(frame_horas, text="Hora término (HH:MM):", width=LABEL_W, anchor="e")
    lbl_fin.grid(row=2, column=0, sticky="e", padx=LABEL_PADX, pady=(8, 0))
    entry_hora_fin = ctk.CTkEntry(frame_horas, placeholder_text="HH:MM", width=120)
    entry_hora_fin.grid(row=2, column=1, sticky="w", pady=(8, 0))
    entry_hora_fin.configure(state="disabled")

    def _set_estado_horas(enabled: bool):
        estado = "normal" if enabled else "disabled"
        cmb_cantidad_horas.configure(state=estado)
        entry_hora_inicio.configure(state=estado)
        entry_hora_fin.configure(state="disabled")  # siempre de solo lectura
        if not enabled:
            # limpiar campos si se deshabilita
            cmb_cantidad_horas.set(horas_values[0])
            entry_hora_inicio.delete(0, "end")
            entry_hora_fin.configure(state="normal")
            entry_hora_fin.delete(0, "end")
            entry_hora_fin.configure(state="disabled")

    # Arranca deshabilitado (visible pero no editable)
    _set_estado_horas(False)

    # ---- Observación
    ctk.CTkLabel(form, text="Observación:", width=LABEL_W, anchor="e").grid(row=7, column=0, sticky="e", padx=LABEL_PADX)
    entry_obs = ctk.CTkEntry(form, placeholder_text="Motivo u observación (opcional)", width=480)
    entry_obs.grid(row=7, column=1, columnspan=2, sticky="w")
    # ===========================================================================

    # ---- Nota aclaratoria (fuente +1 → 13)
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
    # (cmb_tipo command se conecta más abajo)

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
            # limpiar fin si inicio no válido
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

    # Eventos que recalculan
    cmb_cantidad_horas.configure(command=lambda *_: _calc_hora_fin())
    entry_hora_inicio.bind("<FocusOut>", lambda *_: _calc_hora_fin())
    entry_hora_inicio.bind("<KeyRelease>", lambda *_: _calc_hora_fin())

    # ---------- Habilitar/Deshabilitar bloque horas según tipo (sin ocultar) ----------
    def _on_tipo_change(*_):
        tipo_sel = cmb_tipo.get().strip()
        _auto_hasta()
        if tipo_sel == "Día Administrativo (Horas)":
            _set_estado_horas(True)   # habilitar edición
        else:
            _set_estado_horas(False)  # mantener visible pero no editable

    # Conectar cambio de tipo y aplicar estado inicial
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

        # Validaciones extra si es por horas
        if tipo == "Día Administrativo (Horas)":
            h_ini = _parse_hhmm(entry_hora_inicio.get())
            if h_ini is None:
                faltantes.append("Hora inicio (HH:MM)")
            # Asegurar que hora fin esté calculada
            val_fin = entry_hora_fin.get().strip()
            if not val_fin:
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

        # Si es por horas, anexar detalles a la observación y asegurar cálculo fin
        detalle_horas = ""
        if tipo == "Día Administrativo (Horas)":
            try:
                cant = int(cmb_cantidad_horas.get())
            except Exception:
                cant = 1
            hora_ini_txt = entry_hora_inicio.get().strip()
            _calc_hora_fin()
            hora_fin_txt = entry_hora_fin.get().strip()
            detalle_horas = f" [Horas: {cant}, Inicio: {hora_ini_txt}, Término: {hora_fin_txt}]"
            if obs:
                obs = obs + detalle_horas
            else:
                obs = detalle_horas.strip()

        try:
            folio = get_next_folio()
        except Exception as e:
            messagebox.showerror("Folio", f"No fue posible generar el folio correlativo:\n{e}")
            return

        # Guardar (pdf_path = 'MANUAL')
        try:
            guardar_solicitud_en_bd(folio, rut_sel, nombre_sel, tipo, d.strftime("%Y-%m-%d"), h.strftime("%Y-%m-%d"), obs, "MANUAL")
        except Exception as e:
            messagebox.showerror("Base de Datos", f"No se pudo guardar la solicitud:\n{e}")
            return

        aviso = ("IMPORTANTE: Esta solicitud NO es oficial. Solo reserva la ausencia y "
                 "queda sujeta a autorización. Debe formalizarse en el área administrativa "
                 "completando el documento físico.")
        cuerpo = (
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
        try:
            enviar_correo(DESTINATARIOS, asunto=f"Reserva de Solicitud - Folio {folio:06d} - {nombre_sel}", cuerpo=cuerpo)
            messagebox.showinfo("Listo", f"Reserva registrada (Folio {folio:06d}).\nSe envió el aviso por correo.")
        except Exception as e:
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
