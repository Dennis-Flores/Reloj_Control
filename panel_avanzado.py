import os
import sys
import customtkinter as ctk
import sqlite3
from tkinter import messagebox, ttk
import importlib.util
import importlib

# --- Rutas/BD ---
def _app_path():
    return os.path.dirname(sys.executable) if getattr(sys, "frozen", False) \
           else os.path.dirname(os.path.abspath(__file__))

DB_PATH = os.path.join(_app_path(), "reloj_control.db")

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
    # topmost temporal para garantizar primer plano
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

# ---------- Estilo oscuro para ttk.Treeview (coherente con el resto) ----------
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
    # Más ancho y un poco más alto
    win.geometry("900x560")
    win.minsize(860, 520)

    top = ctk.CTkFrame(win)
    top.pack(fill="x", padx=12, pady=(12, 6))

    ctk.CTkLabel(top, text="Buscar:").pack(side="left")
    entry_buscar = ctk.CTkEntry(top, width=260, placeholder_text="fecha o nombre")
    entry_buscar.pack(side="left", padx=(6,10))

    # Botones (Cerrar a la derecha)
    btn_close = ctk.CTkButton(top, text="Cerrar", width=110, fg_color="#4b5563", command=win.destroy)
    btn_del   = ctk.CTkButton(top, text="Eliminar", width=110, fg_color="#e53935")
    btn_edit  = ctk.CTkButton(top, text="Editar", width=110, fg_color="#1e88e5")
    btn_new   = ctk.CTkButton(top, text="Nuevo", width=110, fg_color="#2e7d32",
                              command=lambda: (_abrir_dialogo_feriado_manual(win), _refresh_after()))
    # Empaquetar en orden para que "Cerrar" quede totalmente a la derecha
    btn_close.pack(side="right", padx=4)
    btn_new.pack(side="right", padx=4)
    btn_edit.pack(side="right", padx=4)
    btn_del.pack(side="right", padx=4)

    mid = ctk.CTkFrame(win)
    mid.pack(fill="both", expand=True, padx=12, pady=(0, 12))

    # Treeview ttk con estilo oscuro
    tree = ttk.Treeview(mid, columns=("fecha","nombre","irren"), show="headings", height=16, style=style_name)
    tree.pack(fill="both", expand=True)
    tree.heading("fecha", text="Fecha")
    tree.heading("nombre", text="Nombre")
    tree.heading("irren", text="Irrenunciable")
    tree.column("fecha", width=120, anchor="center")
    tree.column("nombre", width=540, anchor="w")
    tree.column("irren", width=120, anchor="center")

    # Scrollbar vertical
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

# --- Import robusto (packaged + dev) ---
def _load_runtime_module(mod_name: str):
    """
    1) Intenta import normal (sirve en el .exe si se agregó --hidden-import).
    2) Si falla, carga el archivo .py desde la carpeta junto al .py/.exe (modo dev).
    """
    # 1) paquete / .exe
    try:
        return importlib.import_module(mod_name)
    except Exception as e_primary:
        # 2) fallback a .py en disco (desarrollo)
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
