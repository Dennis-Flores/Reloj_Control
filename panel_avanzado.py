import os
import sys
import ssl
import smtplib
import mimetypes
import tempfile
import customtkinter as ctk
import sqlite3
from tkinter import messagebox, ttk
import importlib.util
import importlib
from email.message import EmailMessage

# ====== Resumen Global: constantes / helpers ======
TOL_INGRESO_MIN = 5        # para tardanzas (misma regla que usas a diario)
_DAYS_ES = ['Lunes','Martes','Miércoles','Jueves','Viernes','Sábado','Domingo']

def _fmt_rut_norm(s: str) -> str:
    return (s or "").upper().replace(".","").replace("-","").strip()

def _parse_hhmm(h: str):
    from datetime import datetime
    if not h: return None
    h = h.strip()
    for fmt in ("%H:%M:%S","%H:%M"):
        try: return datetime.strptime(h, fmt)
        except Exception: pass
    return None

def _diff_min(a: str, b: str) -> int | None:
    ta, tb = _parse_hhmm(a), _parse_hhmm(b)
    if not ta or not tb: return None
    return int((ta - tb).total_seconds() // 60)

def _fetch_horarios_dia_glob(con, rut: str, dia_es: str):
    cur = con.cursor()
    cur.execute("""
        SELECT TRIM(IFNULL(hora_entrada,'')), TRIM(IFNULL(hora_salida,''))
        FROM horarios
        WHERE REPLACE(REPLACE(UPPER(IFNULL(rut,'')),'.',''),'-','')
              = REPLACE(REPLACE(UPPER(?),'.',''),'-','')
          AND LOWER(TRIM(IFNULL(dia,''))) = LOWER(TRIM(?))
          AND TRIM(IFNULL(hora_entrada,'')) <> ''
          AND TRIM(IFNULL(hora_salida ,'')) <> ''
        ORDER BY time(hora_entrada) ASC
    """, (rut, dia_es))
    return cur.fetchall()

def _expected_ingreso_glob(con, rut: str, fecha_iso: str, hora_real: str) -> str:
    import datetime as _dt
    d = _dt.datetime.strptime(fecha_iso, "%Y-%m-%d").date()
    dia = _DAYS_ES[d.weekday()]
    bloques = _fetch_horarios_dia_glob(con, rut, dia)
    if not bloques: return ""
    t = _parse_hhmm(hora_real) or _parse_hhmm(bloques[0][0])
    best_he, best_diff = None, None
    for he, _ in bloques:
        th = _parse_hhmm(he)
        if not th: continue
        diff = abs((t - th).total_seconds())
        if best_diff is None or diff < best_diff:
            best_he, best_diff = he, diff
    return best_he or bloques[0][0]

def _label_perm(motivo_raw: str) -> str:
    s = (motivo_raw or "").lower()
    if "licenc" in s: return "Licencia médica"
    if "cometid" in s: return "Cometido de servicio"
    if "admin" in s or "día administrativo" in s or "dia administrativo" in s: return "Permiso administrativo"
    if "defunci" in s: return "Permiso por defunción"
    if "permiso" in s: return "Permiso"
    return "Otro permiso"

# --- Rutas/BD ---
def _app_path():
    return os.path.dirname(sys.executable) if getattr(sys, "frozen", False) \
           else os.path.dirname(os.path.abspath(__file__))

DB_PATH = os.path.join(_app_path(), "reloj_control.db")

# ======== SMTP (igual a tu resumen diario) ========
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
        con = sqlite3.connect(DB_PATH); cur = con.cursor()
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
    msg["Subject"] = subject or "Resumen Global"
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

# --- Feriados opcional ---
try:
    from feriados import sincronizar_feriados_chile
except Exception:
    sincronizar_feriados_chile = None

# --- tkcalendar (versión ya utilizada) ---
try:
    from tkcalendar import Calendar
    TKCAL_OK = True
except Exception:
    TKCAL_OK = False

# ---------- Utilidades de centrar/traer al frente ----------
def _center_on_parent(win, parent=None):
    """Centra la ventana 'win' respecto al padre (si hay) o a la pantalla."""
    try:
        win.update_idletasks()
        w = win.winfo_width() or win.winfo_reqwidth()
        h = win.winfo_height() or win.winfo_reqheight()
        if parent is not None:
            parent.update_idletasks()
            px, py = parent.winfo_rootx(), parent.winfo_rooty()
            pw, ph = parent.winfo_width(), parent.winfo_height()
            x = px + (pw - w) // 2
            y = py + (ph - h) // 2
        else:
            sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
            x = (sw - w) // 2
            y = (sh - h) // 2
        win.geometry(f"+{x}+{y}")
    except Exception:
        pass

def _lift_and_focus(win, parent=None):
    """Asegura que la ventana quede al frente y con foco."""
    try:
        if parent is not None:
            win.transient(parent)
        win.grab_set()
    except Exception:
        pass
    for fn in (win.lift, win.focus_force, win.update):
        try: fn()
        except Exception: pass
    try:
        win.attributes("-topmost", True)
        win.after(250, lambda: win.attributes("-topmost", False))
    except Exception:
        pass

# ---------- Helpers de fechas/texto ----------
_MESES_ES = {
    1:"Enero",2:"Febrero",3:"Marzo",4:"Abril",5:"Mayo",6:"Junio",
    7:"Julio",8:"Agosto",9:"Septiembre",10:"Octubre",11:"Noviembre",12:"Diciembre"
}
_DIAS_ES = ["Lunes","Martes","Miércoles","Jueves","Viernes","Sábado","Domingo"]

def _hoy_iso():
    import datetime as _dt
    return _dt.date.today().strftime("%Y-%m-%d")

def _iso_to_humano(iso: str) -> str:
    import datetime as _dt
    try:
        d = _dt.datetime.strptime(iso, "%Y-%m-%d").date()
        return f"{_DIAS_ES[d.weekday()]} {d.day} de {_MESES_ES[d.month]} de {d.year}"
    except Exception:
        return ""

# ---------- Estilo oscuro para ttk.Treeview ----------
def _style_dark_treeview():
    style = ttk.Style()
    try:
        style.theme_use('clam')
    except Exception:
        pass
    style.configure(
        "Dark.Treeview",
        background="#111418",
        foreground="#e8eef5",
        fieldbackground="#111418",
        rowheight=28,
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

# --- Esquema panel/flags ---
def _ensure_panel_schema():
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS panel_flags (
            fecha TEXT PRIMARY KEY,
            salida_anticipada INTEGER NOT NULL DEFAULT 0,
            salida_anticipada_obs TEXT,
            cierre_forzado INTEGER NOT NULL DEFAULT 0,
            cierre_forzado_obs TEXT
        );
    """)
    cur.execute("CREATE INDEX IF NOT EXISTS idx_registros_fecha_tipo ON registros(fecha, tipo);")
    con.commit()
    con.close()

# --- Esquema feriados (asegurar) ---
def _ensure_feriados_schema():
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS feriados (
            fecha TEXT PRIMARY KEY,         -- YYYY-MM-DD
            nombre TEXT,
            irrenunciable INTEGER NOT NULL DEFAULT 0
        );
    """)
    con.commit()
    con.close()

# --- Acciones panel ---
def _set_flag_salida_anticipada_activa(obs: str):
    _ensure_panel_schema()
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    cur.execute("""
        INSERT INTO panel_flags (fecha, salida_anticipada, salida_anticipada_obs)
        VALUES (?, 1, ?)
        ON CONFLICT(fecha) DO UPDATE SET
          salida_anticipada=1,
          salida_anticipada_obs=excluded.salida_anticipada_obs
    """, (_hoy_iso(), obs))
    con.commit()
    con.close()

def habilitar_salida_anticipada_todos(observacion):
    try:
        _set_flag_salida_anticipada_activa(observacion)
        messagebox.showinfo(
            "Salida Anticipada Activada",
            "Desde ahora, al marcar SALIDA, se registrará con la hora oficial del turno\n"
            "y se agregará la observación de autorización para todos los pendientes de hoy."
        )
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo activar la salida anticipada:\n{e}")

def cerrar_dia_para_todos(observacion):
    import datetime as _dt
    try:
        con = sqlite3.connect(DB_PATH)
        cur = con.cursor()
        cur.execute("""
            SELECT id, rut, fecha, hora_ingreso, COALESCE(observacion, '')
            FROM registros
            WHERE DATE(fecha) = DATE('now')
              AND hora_ingreso IS NOT NULL AND TRIM(hora_ingreso) <> ''
              AND (hora_salida IS NULL OR TRIM(hora_salida) = '')
        """)
        pendientes = cur.fetchall()
        if not pendientes:
            messagebox.showinfo("Nada que cerrar", "No hay ingresos pendientes para cerrar hoy.")
            con.close()
            return

        def _dia_semana_es(fecha_iso: str) -> str:
            d = _dt.datetime.strptime(fecha_iso, "%Y-%m-%d").date()
            return _DIAS_ES[d.weekday()]

        def _calc_salida(cursor, rut: str, fecha_iso: str, hora_ingreso_hhmm: str) -> str:
            dia = _dia_semana_es(fecha_iso)
            cursor.execute("""
                SELECT hora_entrada, hora_salida FROM horarios
                WHERE rut = ? AND dia = ?
            """, (rut, dia))
            turnos = cursor.fetchall()
            if not turnos:
                return "17:30"
            h_ing = _dt.datetime.strptime((hora_ingreso_hhmm or "08:00")[:5], "%H:%M")
            for h_e, h_s in turnos:
                if not h_e or not h_s: continue
                h_ini = _dt.datetime.strptime(h_e, "%H:%M")
                h_fin = _dt.datetime.strptime(h_s, "%H:%M")
                if h_fin < h_ini:
                    h_fin += _dt.timedelta(days=1)
                    if h_ing < h_ini: h_ing += _dt.timedelta(days=1)
                if h_ini <= h_ing <= h_fin:
                    return h_s[:5]
            salidas = [t[1] for t in turnos if t[1]]
            return (max(salidas)[:5] if salidas else "17:30")

        for reg_id, rut, fecha, hora_ingreso, _ in pendientes:
            hsal = _calc_salida(cur, rut, fecha, hora_ingreso or "08:00")
            cur.execute("""
                UPDATE registros
                   SET hora_salida = ?,
                       observacion = CASE
                           WHEN observacion IS NULL OR TRIM(observacion) = '' THEN ?
                           ELSE observacion || ' | ' || ?
                       END
                 WHERE id = ?
            """, (hsal, observacion, observacion, reg_id))

        _ensure_panel_schema()
        cur.execute("""
            INSERT INTO panel_flags (fecha, cierre_forzado, cierre_forzado_obs)
            VALUES (DATE('now'), 1, ?)
            ON CONFLICT(fecha) DO UPDATE SET
              cierre_forzado=1,
              cierre_forzado_obs=excluded.cierre_forzado_obs
        """, (observacion,))
        con.commit(); con.close()
        messagebox.showinfo("Cierre de Jornada", f"Se cerró la jornada de {len(pendientes)} funcionario(s).")
    except Exception as e:
        messagebox.showerror("Error", f"Error al cerrar jornada:\n{e}")

def _sincronizar_feriados_anios(anios):
    if sincronizar_feriados_chile is None:
        return messagebox.showerror("Feriados","No se encontró el módulo de feriados.")
    try:
        for a in anios: sincronizar_feriados_chile(a)
        messagebox.showinfo("Feriados", "Feriados sincronizados: " + ", ".join(map(str, anios)))
    except Exception as e:
        messagebox.showerror("Error al sincronizar feriados", str(e))

# --- CRUD feriados ---
def _upsert_feriado_manual(fecha_iso: str, nombre: str, irrenunciable: bool):
    _ensure_feriados_schema()
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    cur.execute("""
        INSERT INTO feriados (fecha, nombre, irrenunciable)
        VALUES (DATE(?), ?, ?)
        ON CONFLICT(fecha) DO UPDATE SET
            nombre = excluded.nombre,
            irrenunciable = excluded.irrenunciable
    """, (fecha_iso, nombre, 1 if irrenunciable else 0))
    con.commit()
    con.close()

def _delete_feriado(fecha_iso: str):
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    cur.execute("DELETE FROM feriados WHERE DATE(fecha) = DATE(?)", (fecha_iso,))
    con.commit(); con.close()

def _fetch_feriados(filtro_texto=""):
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    if filtro_texto:
        like = f"%{filtro_texto.lower()}%"
        cur.execute("""
            SELECT fecha, COALESCE(nombre,''), COALESCE(irrenunciable,0)
            FROM feriados
            WHERE LOWER(fecha) LIKE ? OR LOWER(nombre) LIKE ?
            ORDER BY fecha
        """, (like, like))
    else:
        cur.execute("""
            SELECT fecha, COALESCE(nombre,''), COALESCE(irrenunciable,0)
            FROM feriados
            ORDER BY fecha
        """)
    rows = cur.fetchall()
    con.close()
    return rows

# ---------- Calendario emergente (tkcalendar) ----------
def _open_calendar_popup(parent, initial_iso, on_pick):
    """Abre un popup con tkcalendar.Calendar centrado y al frente."""
    if not TKCAL_OK:
        messagebox.showerror(
            "Calendario",
            "No se encontró tkcalendar. Instala con:\n  pip install tkcalendar"
        )
        return
    import datetime as _dt
    try:
        base = _dt.datetime.strptime((initial_iso or "").strip(), "%Y-%m-%d").date()
    except Exception:
        base = _dt.date.today()

    top = ctk.CTkToplevel(parent)
    top.title("Seleccionar fecha")
    top.resizable(False, False)

    cal = Calendar(
        top,
        selectmode="day",
        locale="es_CL",
        date_pattern="yyyy-mm-dd"
    )
    cal.pack(padx=12, pady=12)
    cal.selection_set(base)

    btns = ctk.CTkFrame(top, fg_color="transparent")
    btns.pack(pady=(0,12))
    def _ok():
        iso = cal.get_date()
        on_pick(iso)
        top.destroy()
    def _hoy():
        cal.selection_set(_dt.date.today())

    ctk.CTkButton(btns, text="Hoy", width=90, command=_hoy).pack(side="left", padx=6)
    ctk.CTkButton(btns, text="Aceptar", width=120, fg_color="#2e7d32", command=_ok).pack(side="left", padx=6)
    ctk.CTkButton(btns, text="Cancelar", width=120, fg_color="#757575", command=top.destroy).pack(side="left", padx=6)

    _center_on_parent(top, parent)
    _lift_and_focus(top, parent)

# --- Diálogo para feriado manual (crear/editar) ---
def _abrir_dialogo_feriado_manual(parent, registro=None):
    """
    registro: None para nuevo, o tupla (fecha_iso, nombre, irrenunciable)
    """
    import datetime as _dt

    win = ctk.CTkToplevel(parent)
    win.title("Asignar Feriado Manual" if not registro else "Editar Feriado Manual")
    win.resizable(False, False)

    frame = ctk.CTkFrame(win)
    frame.pack(padx=18, pady=16)

    ctk.CTkLabel(frame, text="Define un feriado para cualquier día del año.").grid(row=0, column=0, columnspan=3, pady=(0,8))

    # Fecha
    ctk.CTkLabel(frame, text="Fecha (YYYY-MM-DD):").grid(row=1, column=0, sticky="e", padx=(0,6))
    entry_fecha = ctk.CTkEntry(frame, width=180, placeholder_text="YYYY-MM-DD")
    entry_fecha.grid(row=1, column=1, sticky="w", pady=4)
    btn_cal = ctk.CTkButton(frame, text="📅", width=36)
    btn_cal.grid(row=1, column=2, padx=(6,0))

    # Nombre
    ctk.CTkLabel(frame, text="Nombre del feriado:").grid(row=2, column=0, sticky="e", padx=(0,6), pady=(8,0))
    entry_nombre = ctk.CTkEntry(frame, width=260)
    entry_nombre.grid(row=2, column=1, columnspan=2, sticky="w", pady=(8,0))

    # Tipo
    ctk.CTkLabel(frame, text="Tipo:").grid(row=3, column=0, sticky="e", padx=(0,6), pady=(8,0))
    combo_tipo = ctk.CTkComboBox(frame, values=["Normal", "Irrenunciable"])
    combo_tipo.grid(row=3, column=1, sticky="w", pady=(8,0))

    # Panel informativo
    info = ctk.CTkFrame(win)
    info.pack(padx=18, pady=(6, 6), fill="x")
    lbl_largo = ctk.CTkLabel(info, text="", font=("Segoe UI", 13, "bold"))
    lbl_largo.pack(pady=(6,2))
    ctk.CTkLabel(
        info,
        text="• Feriado irrenunciable: por ley, no laborable para comercio general.\n"
             "• Si alguien trabaja ese día, puede implicar recargo de pago/compensación.",
        justify="left"
    ).pack(pady=(0,6))

    def _sync_largo(*_):
        iso = entry_fecha.get().strip()
        lbl_largo.configure(text=_iso_to_humano(iso) or "Fecha inválida")

    # Valores por defecto / edición
    if not registro:
        sugerida = f"{_hoy_iso()[:4]}-09-17"
        entry_fecha.insert(0, sugerida)
        entry_nombre.insert(0, "Feriado Manual")
        combo_tipo.set("Normal")
    else:
        f0, nom0, irr0 = registro
        entry_fecha.insert(0, f0)
        entry_nombre.insert(0, nom0 or "")
        combo_tipo.set("Irrenunciable" if int(irr0) == 1 else "Normal")

    _sync_largo()
    entry_fecha.bind("<KeyRelease>", _sync_largo)

    # Abrir calendario
    def _set_fecha(iso):
        entry_fecha.delete(0, 'end')
        entry_fecha.insert(0, iso)
        _sync_largo()

    btn_cal.configure(command=lambda: _open_calendar_popup(win, entry_fecha.get().strip(), _set_fecha))
    entry_fecha.bind("<Button-1>", lambda e: _open_calendar_popup(win, entry_fecha.get().strip(), _set_fecha))

    # Botones de acción
    btns = ctk.CTkFrame(win, fg_color="transparent")
    btns.pack(pady=(4, 12))

    def _guardar():
        fecha = entry_fecha.get().strip()
        nombre = (entry_nombre.get() or "").strip()
        tipo = combo_tipo.get().strip()

        # Validar fecha
        import datetime as _dt
        try:
            _dt.datetime.strptime(fecha, "%Y-%m-%d")
        except Exception:
            return messagebox.showerror("Fecha inválida", "Usa el formato YYYY-MM-DD (ej: 2025-09-17).")

        if not nombre:
            return messagebox.showerror("Nombre requerido", "Ingresa un nombre para el feriado.")

        try:
            if registro and fecha != registro[0]:
                _delete_feriado(registro[0])
            _upsert_feriado_manual(fecha, nombre, irrenunciable=(tipo.lower().startswith("irren")))
            messagebox.showinfo("Feriado guardado", f"Se guardó {fecha} como feriado ({tipo}).")
            win.destroy()
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar el feriado:\n{e}")

    ctk.CTkButton(btns, text="Guardar", fg_color="#2e7d32", width=140, command=_guardar).pack(side="left", padx=10)
    ctk.CTkButton(btns, text="Cancelar", fg_color="#757575", width=140, command=win.destroy).pack(side="left", padx=10)

    _center_on_parent(win, parent)
    _lift_and_focus(win, parent)

# --- Gestor (listar/editar/quitar) ---
def _abrir_gestor_feriados(parent):
    style_name = _style_dark_treeview()

    win = ctk.CTkToplevel(parent)
    win.title("Gestor de Feriados")
    win.geometry("900x560")
    win.minsize(860, 520)

    top = ctk.CTkFrame(win)
    top.pack(fill="x", padx=12, pady=(12, 6))

    ctk.CTkLabel(top, text="Buscar:").pack(side="left")
    entry_buscar = ctk.CTkEntry(top, width=260, placeholder_text="fecha o nombre")
    entry_buscar.pack(side="left", padx=(6,10))

    btn_close = ctk.CTkButton(top, text="Cerrar", width=110, fg_color="#4b5563", command=win.destroy)
    btn_del   = ctk.CTkButton(top, text="Eliminar", width=110, fg_color="#e53935")
    btn_edit  = ctk.CTkButton(top, text="Editar", width=110, fg_color="#1e88e5")
    btn_new   = ctk.CTkButton(top, text="Nuevo", width=110, fg_color="#2e7d32",
                              command=lambda: (_abrir_dialogo_feriado_manual(win), _refresh_after()))
    btn_close.pack(side="right", padx=4)
    btn_new.pack(side="right", padx=4)
    btn_edit.pack(side="right", padx=4)
    btn_del.pack(side="right", padx=4)

    mid = ctk.CTkFrame(win)
    mid.pack(fill="both", expand=True, padx=12, pady=(0, 12))

    tree = ttk.Treeview(mid, columns=("fecha","nombre","irren"), show="headings", height=16, style=style_name)
    tree.pack(fill="both", expand=True)
    tree.heading("fecha", text="Fecha")
    tree.heading("nombre", text="Nombre")
    tree.heading("irren", text="Irrenunciable")
    tree.column("fecha", width=120, anchor="center")
    tree.column("nombre", width=540, anchor="w")
    tree.column("irren", width=120, anchor="center")

    vs = ttk.Scrollbar(mid, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=vs.set)
    vs.place(relx=1.0, rely=0, relheight=1.0, anchor="ne")

    def _load(filtro=""):
        for i in tree.get_children():
            tree.delete(i)
        for f, n, irr in _fetch_feriados(filtro):
            tree.insert("", "end", values=(f, n, "Sí" if int(irr)==1 else "No"))

    def _get_sel():
        sel = tree.selection()
        if not sel: return None
        vals = tree.item(sel[0], "values")
        if not vals or len(vals) < 3: return None
        f, n, irr_txt = vals
        return (f, n, 1 if str(irr_txt).strip().lower().startswith("s") else 0)

    def _refresh_after():
        _load(entry_buscar.get().strip())

    def _on_new():
        _abrir_dialogo_feriado_manual(win)
        win.after(200, _refresh_after)

    def _on_edit():
        reg = _get_sel()
        if not reg:
            return messagebox.showinfo("Editar", "Selecciona un feriado de la lista.")
        _abrir_dialogo_feriado_manual(win, registro=reg)
        win.after(200, _refresh_after)

    def _on_del():
        reg = _get_sel()
        if not reg:
            return messagebox.showinfo("Eliminar", "Selecciona un feriado de la lista.")
        f, n, _irr = reg
        if messagebox.askyesno("Eliminar feriado", f"¿Eliminar el feriado del {f} ({n})?"):
            try:
                _delete_feriado(f)
                _load(entry_buscar.get().strip())
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo eliminar:\n{e}")

    btn_new.configure(command=_on_new)
    btn_edit.configure(command=_on_edit)
    btn_del.configure(command=_on_del)

    entry_buscar.bind("<KeyRelease>", lambda e: _load(entry_buscar.get().strip()))
    tree.bind("<Double-1>", lambda e: _on_edit())

    _load()

    _center_on_parent(win, parent)
    _lift_and_focus(win, parent)

def _abrir_resumen_global(parent):
    """
    Resumen global con rango de fechas: tardanzas, sin salida, observaciones
    y permisos/licencias (desde dias_libres). Exporta PDF con gráficos si hay matplotlib.
    """
    import datetime as _dt
    # reportlab opcional
    try:
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.lib import colors
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
        from reportlab.lib.styles import getSampleStyleSheet
        HAS_PDF = True
    except Exception:
        HAS_PDF = False

    # matplotlib opcional (para gráficos en PDF)
    try:
        import matplotlib.pyplot as _plt
        HAS_MPL = True
    except Exception:
        HAS_MPL = False

    style_name = _style_dark_treeview()

    win = ctk.CTkToplevel(parent)
    win.title("Resumen Global")
    # abrir maximizada (Windows / Linux)
    try:
        win.state('zoomed')
    except Exception:
        try: win.attributes('-zoomed', True)
        except Exception: pass
    win.minsize(1000, 620)

    # ------- Top: rango de fechas y acciones
    top = ctk.CTkFrame(win); top.pack(fill="x", padx=12, pady=(12,6))

    ctk.CTkLabel(top, text="Desde (YYYY-MM-DD):").pack(side="left")
    e_desde = ctk.CTkEntry(top, width=140); e_desde.pack(side="left", padx=(6,4))
    b_cal_d = ctk.CTkButton(top, text="📅", width=36); b_cal_d.pack(side="left", padx=(0,10))

    ctk.CTkLabel(top, text="Hasta (YYYY-MM-DD):").pack(side="left")
    e_hasta = ctk.CTkEntry(top, width=140); e_hasta.pack(side="left", padx=(6,4))
    b_cal_h = ctk.CTkButton(top, text="📅", width=36); b_cal_h.pack(side="left", padx=(0,10))

    # Rápidos
    quick = ctk.CTkComboBox(top, width=180,
                             values=["Últimos 7 días","Últimos 30 días","Este mes","Mes pasado","Este año","Año pasado","Rango personalizado"])
    quick.pack(side="left", padx=(10,10)); quick.set("Últimos 30 días")

    b_calc = ctk.CTkButton(top, text="Calcular", width=110, fg_color="#1e88e5")
    b_pdf  = ctk.CTkButton(top, text="Exportar PDF", width=130, fg_color="#0f766e")
    b_send = ctk.CTkButton(top, text="Enviar", width=110, fg_color="#2563eb")
    b_close= ctk.CTkButton(top, text="Cerrar", width=110, fg_color="#64748b", command=win.destroy)
    b_close.pack(side="right", padx=4)
    b_send.pack(side="right", padx=4)
    b_pdf.pack(side="right", padx=4)
    b_calc.pack(side="right", padx=4)

    # Fechas por defecto (últimos 30 días)
    hoy = _dt.date.today()
    e_hasta.insert(0, hoy.strftime("%Y-%m-%d"))
    e_desde.insert(0, (hoy - _dt.timedelta(days=30)).strftime("%Y-%m-%d"))

    # Calendarios
    b_cal_d.configure(command=lambda: _open_calendar_popup(win, e_desde.get().strip(), lambda iso: (e_desde.delete(0,'end'), e_desde.insert(0, iso))))
    b_cal_h.configure(command=lambda: _open_calendar_popup(win, e_hasta.get().strip(), lambda iso: (e_hasta.delete(0,'end'), e_hasta.insert(0, iso))))

    # ------- Resumen numérico
    resume = ctk.CTkFrame(win); resume.pack(fill="x", padx=12, pady=(0,6))
    lbl_r = ctk.CTkLabel(resume, text="", font=("Segoe UI", 13))
    lbl_r.pack(anchor="w", padx=6, pady=6)

    # ------- Notebook con tablas
    nb = ttk.Notebook(win); nb.pack(fill="both", expand=True, padx=12, pady=(0,12))

    # Tab Tardanzas
    tab_t = ctk.CTkFrame(nb); nb.add(tab_t, text="Tardanzas")
    frame_top = ctk.CTkFrame(tab_t); frame_top.pack(fill="both", expand=True, padx=8, pady=8)
    tv_top = ttk.Treeview(frame_top, columns=("rut","nombre","casos","min_tot","min_prom"), show="headings", height=10, style=style_name)
    for c,t,a,w in (("rut","RUT","w",140),("nombre","Nombre","w",360),("casos","Casos","center",80),
                    ("min_tot","Min. acumulados","center",130),("min_prom","Promedio min.","center",120)):
        tv_top.heading(c, text=t); tv_top.column(c, anchor=a, width=w)
    vs1 = ttk.Scrollbar(frame_top, orient="vertical", command=tv_top.yview); tv_top.configure(yscrollcommand=vs1.set)
    tv_top.pack(side="left", fill="both", expand=True); vs1.pack(side="left", fill="y", padx=(2,0))

    frame_dia = ctk.CTkFrame(tab_t); frame_dia.pack(fill="both", expand=True, padx=8, pady=(0,8))
    tv_dia = ttk.Treeview(frame_dia, columns=("fecha","casos"), show="headings", height=8, style=style_name)
    tv_dia.heading("fecha", text="Fecha"); tv_dia.column("fecha", width=140, anchor="center")
    tv_dia.heading("casos", text="Tardanzas"); tv_dia.column("casos", width=120, anchor="center")
    vs2 = ttk.Scrollbar(frame_dia, orient="vertical", command=tv_dia.yview); tv_dia.configure(yscrollcommand=vs2.set)
    tv_dia.pack(side="left", fill="both", expand=True); vs2.pack(side="left", fill="y", padx=(2,0))

    # Tab Permisos/Licencias
    tab_p = ctk.CTkFrame(nb); nb.add(tab_p, text="Permisos / Licencias")
    tv_perm = ttk.Treeview(tab_p, columns=("fecha","rut","nombre","motivo"), show="headings", height=16, style=style_name)
    for c,t,a,w in (("fecha","Fecha","center",120),("rut","RUT","w",140),("nombre","Nombre","w",360),("motivo","Motivo","w",360)):
        tv_perm.heading(c, text=t); tv_perm.column(c, anchor=a, width=w)
    vs3 = ttk.Scrollbar(tab_p, orient="vertical", command=tv_perm.yview); tv_perm.configure(yscrollcommand=vs3.set)
    tv_perm.pack(side="left", fill="both", expand=True, padx=8, pady=8); vs3.pack(side="left", fill="y", padx=(2,0))
    lbl_perm = ctk.CTkLabel(tab_p, text=""); lbl_perm.pack(anchor="e", padx=12, pady=(0,8))

    # Tab Sin salida
    tab_s = ctk.CTkFrame(nb); nb.add(tab_s, text="Ingresos sin salida")
    tv_nos = ttk.Treeview(tab_s, columns=("fecha","rut","nombre"), show="headings", height=18, style=style_name)
    for c,t,a,w in (("fecha","Fecha","center",120),("rut","RUT","w",160),("nombre","Nombre","w",520)):
        tv_nos.heading(c, text=t); tv_nos.column(c, anchor=a, width=w)
    vs4 = ttk.Scrollbar(tab_s, orient="vertical", command=tv_nos.yview); tv_nos.configure(yscrollcommand=vs4.set)
    tv_nos.pack(side="left", fill="both", expand=True, padx=8, pady=8); vs4.pack(side="left", fill="y", padx=(2,0))

    # Tab Observaciones (diarias)
    tab_o = ctk.CTkFrame(nb); nb.add(tab_o, text="Observaciones (Registros)")
    tv_obs = ttk.Treeview(tab_o, columns=("fecha","rut","nombre","texto"), show="headings", height=16, style=style_name)
    for c,t,a,w in (("fecha","Fecha","center",120),("rut","RUT","w",160),("nombre","Nombre","w",320),("texto","Texto","w",420)):
        tv_obs.heading(c, text=t); tv_obs.column(c, anchor=a, width=w)
    vs5 = ttk.Scrollbar(tab_o, orient="vertical", command=tv_obs.yview); tv_obs.configure(yscrollcommand=vs5.set)
    tv_obs.pack(side="left", fill="both", expand=True, padx=8, pady=8); vs5.pack(side="left", fill="y", padx=(2,0))

    # ---------- lógica ----------
    stats_cache = {}

    def _aplicar_quick():
        base = _dt.date.today()
        sel = (quick.get() or "").lower()
        if "7" in sel:
            d = base - _dt.timedelta(days=7)
        elif "30" in sel:
            d = base - _dt.timedelta(days=30)
        elif sel.startswith("este mes"):
            d = base.replace(day=1)
        elif sel.startswith("mes pasado"):
            first_this = base.replace(day=1)
            last_prev = first_this - _dt.timedelta(days=1)
            d = last_prev.replace(day=1); base = last_prev
        elif sel.startswith("este año"):
            d = base.replace(month=1, day=1)
        elif sel.startswith("año pasado"):
            d = base.replace(year=base.year-1, month=1, day=1)
            base = base.replace(year=base.year-1, month=12, day=31)
        else:
            return
        e_desde.delete(0,'end'); e_desde.insert(0, d.strftime("%Y-%m-%d"))
        e_hasta.delete(0,'end'); e_hasta.insert(0, base.strftime("%Y-%m-%d"))

    quick.bind("<<ComboboxSelected>>", lambda e: _aplicar_quick())

    def _calc():
        for tv in (tv_top, tv_dia, tv_perm, tv_nos, tv_obs):
            for iid in tv.get_children(): tv.delete(iid)

        f1, f2 = e_desde.get().strip(), e_hasta.get().strip()
        try:
            d1 = _dt.datetime.strptime(f1, "%Y-%m-%d").date()
            d2 = _dt.datetime.strptime(f2, "%Y-%m-%d").date()
            if d1 > d2: d1, d2 = d2, d1
        except Exception:
            messagebox.showerror("Fechas", "Usa formato YYYY-MM-DD."); return

        con = sqlite3.connect(DB_PATH); cur = con.cursor()

        # ------- INGRESOS: tardanzas
        cur.execute("""
            SELECT DATE(fecha) AS f, IFNULL(rut,''), IFNULL(nombre,''),
                   COALESCE(NULLIF(hora_ingreso,''), CASE WHEN lower(IFNULL(tipo,''))='ingreso' THEN IFNULL(hora,'') ELSE '' END) AS h_real
            FROM registros
            WHERE DATE(fecha) BETWEEN DATE(?) AND DATE(?)
              AND TRIM(COALESCE(hora_ingreso, CASE WHEN lower(IFNULL(tipo,''))='ingreso' THEN IFNULL(hora,'') ELSE '' END))<>''
            ORDER BY DATE(fecha) ASC
        """, (f1, f2))
        ingresos = cur.fetchall()

        tard_by_person = {}
        tard_by_day = {}
        total_ing = ok_ing = 0
        sum_min_atraso = 0

        for f, rut, nom, h_real in ingresos:
            he = _expected_ingreso_glob(con, rut, f, h_real)
            dm = _diff_min(h_real or "", he or "")
            total_ing += 1
            if he and dm is not None and dm > TOL_INGRESO_MIN:
                info = tard_by_person.setdefault(rut, {"nombre":nom, "casos":0, "min_tot":0})
                info["casos"] += 1
                info["min_tot"] += max(0, dm)
                tard_by_day[f] = tard_by_day.get(f, 0) + 1
                sum_min_atraso += max(0, dm)
            else:
                ok_ing += 1

        top_list = []
        for rut, d in tard_by_person.items():
            prom = (d["min_tot"] / d["casos"]) if d["casos"] else 0
            top_list.append((rut, d["nombre"], d["casos"], d["min_tot"], round(prom,1)))
        top_list.sort(key=lambda x: (-x[2], -x[3], x[1]))

        # ------- SIN SALIDA
        cur.execute("""
            SELECT DATE(fecha), IFNULL(rut,''), IFNULL(nombre,'')
            FROM registros
            WHERE DATE(fecha) BETWEEN DATE(?) AND DATE(?)
              AND TRIM(IFNULL(hora_ingreso,'')) <> ''
              AND (hora_salida IS NULL OR TRIM(hora_salida)='')
        """, (f1, f2))
        sin_sal = cur.fetchall()

        # ------- OBSERVACIONES (de registros)
        cur.execute("""
            SELECT DATE(fecha), IFNULL(rut,''), IFNULL(nombre,''), IFNULL(observacion,'')
            FROM registros
            WHERE DATE(fecha) BETWEEN DATE(?) AND DATE(?)
              AND TRIM(IFNULL(observacion,'')) <> ''
            ORDER BY DATE(fecha) ASC
        """, (f1, f2))
        obs_rows = cur.fetchall()

        # ------- PERMISOS/LICENCIAS (dias_libres)
        cur.execute("SELECT name FROM sqlite_master WHERE type='table' AND lower(name)='dias_libres'")
        has_dl = cur.fetchone() is not None
        permisos = []
        per_by_mot = {}
        if has_dl:
            cur.execute("""
                SELECT DATE(dl.fecha), IFNULL(dl.rut,''), IFNULL(t.nombre,''), IFNULL(dl.motivo,'')
                FROM dias_libres dl
                LEFT JOIN trabajadores t
                  ON REPLACE(REPLACE(UPPER(IFNULL(t.rut,'')),'.',''),'-','')
                   = REPLACE(REPLACE(UPPER(IFNULL(dl.rut,'')),'.',''),'-','')
                WHERE DATE(dl.fecha) BETWEEN DATE(?) AND DATE(?)
                ORDER BY DATE(dl.fecha) ASC
            """, (f1, f2))
            for f, rut, nom, mot in cur.fetchall():
                motivo = _label_perm(mot)
                permisos.append((f, rut, nom, motivo))
                per_by_mot[motivo] = per_by_mot.get(motivo, 0) + 1

        con.close()

        # ----- Resumen numérico
        tard_total = sum(v[2] for v in top_list)
        prom_atraso = round((sum_min_atraso / tard_total), 1) if tard_total else 0
        txt = (f"Período: {f1} → {f2}   |   Ingresos: {total_ing}  "
               f"| En rango: {ok_ing}  | Tardanzas: {tard_total}  (prom. atraso {prom_atraso} min)   "
               f"| Obs.reg.: {len(obs_rows)}  | Sin salida: {len(sin_sal)}  | Permisos/licencias: {len(permisos)}")
        lbl_r.configure(text=txt)

        # ----- llenar tablas
        for item in top_list[:100]:
            tv_top.insert("", "end", values=item)
        for f, c in sorted(tard_by_day.items()):
            tv_dia.insert("", "end", values=(f, c))
        for f, r, n, m in permisos:
            tv_perm.insert("", "end", values=(f, r, n, m))
        lbl_perm.configure(text=" | ".join([f"{k}: {v}" for k,v in sorted(per_by_mot.items())]) or "—")
        for f, r, n in sin_sal:
            tv_nos.insert("", "end", values=(f, r, n))
        for f, r, n, t in obs_rows:
            tv_obs.insert("", "end", values=(f, r, n, t))

        # cache para PDF
        stats_cache.clear()
        stats_cache.update({
            "desde": f1, "hasta": f2,
            "resumen_txt": txt,
            "top_tard": top_list,
            "tard_day": sorted(tard_by_day.items()),
            "permisos": permisos,
            "permisos_mot": per_by_mot,
            "sin_salida": sin_sal,
            "obs_rows": obs_rows,
            "ing_total": total_ing, "ing_ok": ok_ing, "tard_total": tard_total, "prom_atraso": prom_atraso
        })

    def _build_pdf_to(out_path: str):
        if not stats_cache:
            raise RuntimeError("Primero calcula el resumen.")
        # gráficos opcionales (PNG)
        chart_paths = []
        try:
            if HAS_MPL and stats_cache["top_tard"]:
                top10 = stats_cache["top_tard"][:10]
                labels = [x[1][:20] for x in top10]
                vals   = [x[2] for x in top10]
                import matplotlib.pyplot as _plt
                _plt.figure()
                _plt.bar(labels, vals)
                _plt.xticks(rotation=45, ha="right")
                _plt.title("Top 10 – Tardanzas (casos)")
                _plt.tight_layout()
                p1 = tempfile.mkstemp(prefix="chart_top_tard_", suffix=".png")[1]
                _plt.savefig(p1, dpi=160); _plt.close(); chart_paths.append(p1)
            if HAS_MPL and stats_cache["permisos_mot"]:
                labels = list(stats_cache["permisos_mot"].keys())
                vals   = [stats_cache["permisos_mot"][k] for k in labels]
                import matplotlib.pyplot as _plt
                _plt.figure()
                _plt.bar(labels, vals)
                _plt.xticks(rotation=30, ha="right")
                _plt.title("Permisos/Licencias por motivo")
                _plt.tight_layout()
                p2 = tempfile.mkstemp(prefix="chart_permisos_", suffix=".png")[1]
                _plt.savefig(p2, dpi=160); _plt.close(); chart_paths.append(p2)
        except Exception:
            chart_paths = []

        # PDF
        from reportlab.lib.pagesizes import A4, landscape
        from reportlab.lib import colors
        from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
        from reportlab.lib.styles import getSampleStyleSheet

        doc = SimpleDocTemplate(out_path, pagesize=landscape(A4), rightMargin=18, leftMargin=18, topMargin=14, bottomMargin=14)
        styles = getSampleStyleSheet()
        h = styles["Heading1"]; h.fontSize=18
        small = styles["Normal"]; small.fontSize=9
        body = []

        body.append(Paragraph("Reporte de Asistencia – Resumen Global", h))
        body.append(Paragraph(stats_cache["resumen_txt"], small))
        body.append(Spacer(1, 8))

        resumen_rows = [
            ["Ingresos", stats_cache["ing_total"]],
            ["En rango (≤5')", stats_cache["ing_ok"]],
            ["Tardanzas (casos)", stats_cache["tard_total"]],
            ["Prom. atraso (min)", stats_cache["prom_atraso"]],
            ["Observaciones (registros)", len(stats_cache["obs_rows"])],
            ["Ingresos sin salida", len(stats_cache["sin_salida"])],
            ["Permisos/Licencias", len(stats_cache["permisos"])]
        ]
        tbl = Table([["Resumen",""]] + resumen_rows, colWidths=[180, 70])
        tbl.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,0), colors.HexColor("#e2e8f0")),
            ("FONTNAME",(0,0),(-1,0), "Helvetica-Bold"),
            ("GRID",(0,0),(-1,-1), 0.3, colors.grey),
            ("ALIGN",(1,1),(1,-1), "RIGHT"),
            ("FONTSIZE",(0,0),(-1,-1), 10),
            ("VALIGN",(0,0),(-1,-1), "MIDDLE"),
        ]))
        body.append(tbl); body.append(Spacer(1,10))

        top = stats_cache["top_tard"][:20]
        if top:
            data = [["RUT","Nombre","Casos","Min. acumulados","Promedio min."]] + [list(x) for x in top]
            t = Table(data, colWidths=[120, 240, 60, 90, 90])
            t.setStyle(TableStyle([
                ("BACKGROUND",(0,0),(-1,0), colors.HexColor("#e2e8f0")),
                ("FONTNAME",(0,0),(-1,0), "Helvetica-Bold"),
                ("GRID",(0,0),(-1,-1), 0.25, colors.grey),
                ("ALIGN",(2,1),(4,-1), "CENTER"),
                ("FONTSIZE",(0,0),(-1,-1), 9),
            ]))
            body.append(Paragraph("<b>Top Tardanzas (primeros 20)</b>", small)); body.append(t); body.append(Spacer(1,8))

        per = stats_cache["permisos"][:60]
        if per:
            data = [["Fecha","RUT","Nombre","Motivo"]] + [list(x) for x in per]
            t = Table(data, colWidths=[70, 120, 220, 220])
            t.setStyle(TableStyle([
                ("BACKGROUND",(0,0),(-1,0), colors.HexColor("#e2e8f0")),
                ("FONTNAME",(0,0),(-1,0), "Helvetica-Bold"),
                ("GRID",(0,0),(-1,-1), 0.25, colors.grey),
                ("FONTSIZE",(0,0),(-1,-1), 9),
                ("TEXTCOLOR",(0,1),(-1,-1), colors.HexColor("#3b82f6")),
            ]))
            body.append(Paragraph("<b>Permisos / Licencias</b>", small)); body.append(t); body.append(Spacer(1,8))

        for p in chart_paths:
            try:
                body.append(Image(p, width=420, height=260)); body.append(Spacer(1,8))
            except Exception:
                pass

        doc.build(body)

        # limpia temporales
        for p in chart_paths:
            try: os.remove(p)
            except Exception: pass

        return out_path

    def _export_pdf():
        if not HAS_PDF:
            messagebox.showwarning("PDF", "No está disponible ReportLab. Instálalo con: pip install reportlab")
            return
        if not stats_cache:
            messagebox.showinfo("PDF", "Primero calcula el resumen."); return
        try:
            out_dir = os.path.join(os.path.expanduser("~"), "Downloads")
            os.makedirs(out_dir, exist_ok=True)
            out_path = os.path.join(out_dir, f"resumen_global_{stats_cache['desde']}_{stats_cache['hasta']}.pdf")
            _build_pdf_to(out_path)
            messagebox.showinfo("PDF", f"PDF generado en Descargas:\n{os.path.basename(out_path)}")
            try: os.startfile(out_path)
            except Exception: pass
        except Exception as e:
            messagebox.showerror("PDF", f"No se pudo generar el PDF:\n{e}")

    def _enviar_pdf():
        if not HAS_PDF:
            messagebox.showwarning("PDF", "No está disponible ReportLab. Instálalo con: pip install reportlab")
            return
        if not stats_cache:
            messagebox.showinfo("Enviar", "Primero calcula el resumen."); return

        win_mail = ctk.CTkToplevel(win); win_mail.title("Enviar Resumen Global")
        try:
            win_mail.resizable(False, False); win_mail.transient(win); win_mail.grab_set()
        except Exception:
            pass
        cont = ctk.CTkFrame(win_mail); cont.pack(padx=14, pady=14)
        ctk.CTkLabel(cont, text="Para:").grid(row=0, column=0, sticky="e", padx=(0,6))
        ent_to = ctk.CTkEntry(cont, width=380); ent_to.grid(row=0, column=1, pady=4); ent_to.insert(0, "destinatario@empresa.cl")
        ctk.CTkLabel(cont, text="Asunto:").grid(row=1, column=0, sticky="e", padx=(0,6))
        ent_sub = ctk.CTkEntry(cont, width=380); ent_sub.grid(row=1, column=1, pady=4)
        ent_sub.insert(0, f"Resumen Global – {stats_cache['desde']} a {stats_cache['hasta']}")
        ctk.CTkLabel(cont, text="Mensaje:").grid(row=2, column=0, sticky="ne", padx=(0,6))
        TextBox = getattr(ctk, "CTkTextbox", None)
        if TextBox:
            tb = TextBox(cont, width=380, height=130); tb.grid(row=2, column=1, pady=4)
            tb.insert("1.0", "Estimado(a):\n\nAdjunto el Resumen Global de asistencia.\n\nSaludos.")
        else:
            import tkinter as tk
            tb = tk.Text(cont, width=48, height=7); tb.grid(row=2, column=1, pady=4)
            tb.insert("1.0", "Estimado(a):\n\nAdjunto el Resumen Global de asistencia.\n\nSaludos.")

        def _do_send():
            to = [p.strip() for p in ent_to.get().replace(",", ";").split(";") if p.strip()]
            if not to:
                messagebox.showwarning("Enviar", "Ingresa al menos un destinatario."); return
            subject = ent_sub.get().strip() or "Resumen Global"
            body = tb.get("1.0", "end").strip()

            tmp = tempfile.NamedTemporaryFile(prefix="resumen_global_", suffix=".pdf", delete=False)
            tmp.close()
            try:
                _build_pdf_to(tmp.name)
                _smtp_send(to, [], subject, body, html_body=None, attachment_path=tmp.name)
                messagebox.showinfo("Enviar", "Correo enviado correctamente.")
                win_mail.destroy()
            except Exception as e:
                messagebox.showerror("Enviar", f"No fue posible enviar el correo:\n{e}")
            finally:
                try: os.remove(tmp.name)
                except Exception: pass

        btns = ctk.CTkFrame(cont, fg_color="transparent"); btns.grid(row=3, column=0, columnspan=2, pady=(10,0))
        ctk.CTkButton(btns, text="Cancelar", fg_color="#6b7280", width=120, command=win_mail.destroy).pack(side="left", padx=6)
        ctk.CTkButton(btns, text="Enviar", fg_color="#22c55e", width=140, command=_do_send).pack(side="left", padx=6)

        _center_on_parent(win_mail, win)
        _lift_and_focus(win_mail, win)

    b_calc.configure(command=_calc)
    b_pdf.configure(command=_export_pdf)
    b_send.configure(command=_enviar_pdf)

    # Autocalcular al abrir (últimos 30 días)
    win.after(150, _calc)

    _center_on_parent(win, parent)
    _lift_and_focus(win, parent)

# --- Import robusto (packaged + dev) ---
def _load_runtime_module(mod_name: str):
    """
    1) Intenta import normal (sirve en el .exe si se agregó --hidden-import).
    2) Si falla, carga el archivo .py desde la carpeta junto al .py/.exe (modo dev).
    """
    try:
        return importlib.import_module(mod_name)
    except Exception as e_primary:
        base = _app_path()
        path = os.path.join(base, f"{mod_name}.py")
        try:
            if not os.path.exists(path):
                raise FileNotFoundError(path)
            spec = importlib.util.spec_from_file_location(mod_name, path)
            if not spec or not spec.loader:
                raise ImportError(f"spec inválido para {path}")
            mod = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(mod)  # type: ignore[attr-defined]
            return mod
        except Exception as e_fallback:
            raise ImportError(f"No se pudo importar '{mod_name}'. "
                              f"Primario: {e_primary} | Fallback: {e_fallback}")

# --- Abridores (usan import robusto) ---
def _abrir_asistencia_general(frame_padre):
    app_root = frame_padre.winfo_toplevel()
    try:
        mod = _load_runtime_module("asistencia_funcionarios")
        if not hasattr(mod, "abrir_asistencia"):
            raise AttributeError("El módulo 'asistencia_funcionarios' no define 'abrir_asistencia(app_root, db_path)'.")
        mod.abrir_asistencia(app_root, DB_PATH)
    except Exception as e:
        messagebox.showerror("Asistencia", f"No se pudo abrir la asistencia general:\n{e}")

def _abrir_asistencia_diaria(frame_padre):
    app_root = frame_padre.winfo_toplevel()
    try:
        mod = _load_runtime_module("asistencia_diaria")
        if not hasattr(mod, "abrir_asistencia_diaria"):
            raise AttributeError("El módulo 'asistencia_diaria' no define 'abrir_asistencia_diaria(app_root, db_path)'.")
        mod.abrir_asistencia_diaria(app_root, DB_PATH)
    except Exception as e:
        messagebox.showerror("Asistencia Diaria", f"No se pudo abrir la asistencia diaria:\n{e}")

# --- UI Panel ---
def construir_panel_avanzado(frame_padre):
    for w in frame_padre.winfo_children(): w.destroy()
    _ensure_panel_schema()
    _ensure_feriados_schema()

    root = frame_padre.winfo_toplevel()

    ctk.CTkLabel(
        frame_padre, text="Panel Avanzado (Herramientas Globales)",
        font=("Arial", 18, "bold"), text_color="#c2e7ff"
    ).pack(pady=(10, 20))

    ctk.CTkButton(
        frame_padre, width=340, fg_color="#FFA500",
        text="Permitir Salida Anticipada a Todos",
        command=lambda: mostrar_confirmacion_panel(
            root,
            "Permitir Salida Anticipada",
            "Permite a TODOS marcar salida en cualquier momento (usa hora oficial del turno).",
            habilitar_salida_anticipada_todos,
            "Salida anticipada por instrucción administrativa"
        )
    ).pack(pady=8)

    ctk.CTkButton(
        frame_padre, width=340, fg_color="#DC143C",
        text="Cerrar Jornada para Todos (Emergencia/Festivo)",
        command=lambda: mostrar_confirmacion_panel(
            root,
            "Cerrar Jornada",
            "Cierra ahora la jornada de todos los que no han marcado salida (usa hora oficial).",
            cerrar_dia_para_todos,
            "Cierre de jornada por instrucción administrativa (emergencia/festivo)"
        )
    ).pack(pady=8)

    ctk.CTkButton(
        frame_padre, width=340, fg_color="#00695C",
        text="Sincronizar Feriados 2025 y 2026",
        command=lambda: _sincronizar_feriados_anios([2025, 2026])
    ).pack(pady=8)

    # --- NUEVOS: Feriados manuales ---
    ctk.CTkButton(
        frame_padre, width=340, fg_color="#8E24AA",
        text="Asignar Feriado Manual (Normal/Irrenunciable)",
        command=lambda: _abrir_dialogo_feriado_manual(root)
    ).pack(pady=8)

    ctk.CTkButton(
        frame_padre, width=340, fg_color="#5E35B1",
        text="Administrar Feriados (Listar / Editar / Eliminar)",
        command=lambda: _abrir_gestor_feriados(root)
    ).pack(pady=8)

    # Accesos de asistencia
    ctk.CTkButton(
        frame_padre, width=340, fg_color="#1f6aa5",
        text="Asistencia (General)",
        command=lambda: _abrir_asistencia_general(frame_padre)
    ).pack(pady=8)

    ctk.CTkButton(
        frame_padre, width=340, fg_color="#1f6aa5",
        text="Asistencia Diaria (Matriz)",
        command=lambda: _abrir_asistencia_diaria(frame_padre)
    ).pack(pady=8)

    # Resumen global
    ctk.CTkButton(
        frame_padre, width=340, fg_color="#0b72b9",
        text="Resumen Global",
        command=lambda: _abrir_resumen_global(frame_padre)
    ).pack(pady=8)

def mostrar_confirmacion_panel(parent, titulo, mensaje, funcion_accion, observacion_default=""):
    win = ctk.CTkToplevel(parent)
    win.title(titulo)
    win.resizable(False, False)

    ctk.CTkLabel(win, text=mensaje, wraplength=380, justify="left").pack(padx=18, pady=(20, 12))
    ctk.CTkLabel(win, text="Observación para registro:").pack()
    entry_obs = ctk.CTkEntry(win, width=360); entry_obs.pack(pady=10); entry_obs.insert(0, observacion_default)
    def aceptar():
        obs = entry_obs.get().strip()
        if not obs:
            messagebox.showerror("Observación requerida", "Debe ingresar una observación.")
            return
        funcion_accion(obs); win.destroy()
    btns = ctk.CTkFrame(win, fg_color="transparent"); btns.pack(pady=16)
    ctk.CTkButton(btns, text="Aceptar", fg_color="green", width=120, command=aceptar).pack(side="left", padx=10)
    ctk.CTkButton(btns, text="Cancelar", fg_color="gray",  width=120, command=win.destroy).pack(side="left", padx=10)

    _center_on_parent(win, parent)
    _lift_and_focus(win, parent)
