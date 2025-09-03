import os
import sys
import customtkinter as ctk
import sqlite3
from tkinter import messagebox
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

# --- Acciones panel ---
def _hoy_iso():
    import datetime as _dt
    return _dt.date.today().strftime("%Y-%m-%d")

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
            dias = ["Lunes","Martes","Miércoles","Jueves","Viernes","Sábado","Domingo"]
            return dias[d.weekday()]

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

# --- IMPORTAR SIEMPRE DESDE DISCO (sin caché) ---
def _import_fresh(mod_name: str):
    base = _app_path()
    path = os.path.join(base, f"{mod_name}.py")
    if not os.path.exists(path):
        raise FileNotFoundError(f"No se encontró {mod_name}.py en: {path}")
    importlib.invalidate_caches()
    if mod_name in sys.modules:
        del sys.modules[mod_name]
    spec = importlib.util.spec_from_file_location(mod_name, path)
    if not spec or not spec.loader:
        raise ImportError(f"No se pudo crear spec para {path}")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    sys.modules[mod_name] = mod
    return mod

# --- Abridores (usan import "fresh") ---
def _abrir_asistencia_general(frame_padre):
    app_root = frame_padre.winfo_toplevel()
    try:
        mod = _import_fresh("asistencia_funcionarios")
        if not hasattr(mod, "abrir_asistencia"):
            raise AttributeError("El módulo 'asistencia_funcionarios' no define 'abrir_asistencia(app_root, db_path)'.")
        mod.abrir_asistencia(app_root, DB_PATH)
    except Exception as e:
        messagebox.showerror("Asistencia", f"No se pudo abrir la asistencia general:\n{e}")

def _abrir_asistencia_diaria(frame_padre):
    app_root = frame_padre.winfo_toplevel()
    try:
        mod = _import_fresh("asistencia_diaria")
        if not hasattr(mod, "abrir_asistencia_diaria"):
            raise AttributeError("El módulo 'asistencia_diaria' no define 'abrir_asistencia_diaria(app_root, db_path)'.")
        mod.abrir_asistencia_diaria(app_root, DB_PATH)
    except Exception as e:
        messagebox.showerror("Asistencia Diaria", f"No se pudo abrir la asistencia diaria:\n{e}")

# --- UI Panel ---
def construir_panel_avanzado(frame_padre):
    for w in frame_padre.winfo_children(): w.destroy()
    _ensure_panel_schema()

    ctk.CTkLabel(
        frame_padre, text="Panel Avanzado (Herramientas Globales)",
        font=("Arial", 18, "bold"), text_color="#c2e7ff"
    ).pack(pady=(10, 20))

    ctk.CTkButton(
        frame_padre, width=340, fg_color="#FFA500",
        text="Permitir Salida Anticipada a Todos",
        command=lambda: mostrar_confirmacion_panel(
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

def mostrar_confirmacion_panel(titulo, mensaje, funcion_accion, observacion_default=""):
    win = ctk.CTkToplevel(); win.title(titulo); win.geometry("420x280"); win.grab_set()
    ctk.CTkLabel(win, text=mensaje, wraplength=380, justify="left").pack(pady=(30, 12))
    ctk.CTkLabel(win, text="Observación para registro:").pack()
    entry_obs = ctk.CTkEntry(win, width=360); entry_obs.pack(pady=10); entry_obs.insert(0, observacion_default)
    def aceptar():
        obs = entry_obs.get().strip()
        if not obs:
            messagebox.showerror("Observación requerida", "Debe ingresar una observación.")
            return
        funcion_accion(obs); win.destroy()
    btns = ctk.CTkFrame(win, fg_color="transparent"); btns.pack(pady=18)
    ctk.CTkButton(btns, text="Aceptar", fg_color="green", width=120, command=aceptar).pack(side="left", padx=10)
    ctk.CTkButton(btns, text="Cancelar", fg_color="gray",  width=120, command=win.destroy).pack(side="left", padx=10)
