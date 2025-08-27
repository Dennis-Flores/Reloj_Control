# panel_avanzado.py
import customtkinter as ctk
import sqlite3
from tkinter import messagebox
import datetime

# ⇩ NUEVO: sincronización de feriados
try:
    from feriados import sincronizar_feriados_chile
except Exception as _e:
    sincronizar_feriados_chile = None

# ---------- Helpers de esquema/flags ----------
def _ensure_panel_schema():
    con = sqlite3.connect("reloj_control.db")
    cur = con.cursor()
    # Bandera por día (una fila por fecha)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS panel_flags (
            fecha TEXT PRIMARY KEY,                 -- YYYY-MM-DD
            salida_anticipada INTEGER NOT NULL DEFAULT 0,
            salida_anticipada_obs TEXT,
            cierre_forzado INTEGER NOT NULL DEFAULT 0,
            cierre_forzado_obs TEXT
        );
    """)
    # Índices útiles
    cur.execute("CREATE INDEX IF NOT EXISTS idx_registros_fecha_tipo ON registros(fecha, tipo);")
    con.commit()
    con.close()

def _hoy_iso():
    return datetime.date.today().strftime("%Y-%m-%d")

def _set_flag_salida_anticipada_activa(obs: str):
    _ensure_panel_schema()
    con = sqlite3.connect("reloj_control.db")
    cur = con.cursor()
    hoy = _hoy_iso()
    cur.execute("""
        INSERT INTO panel_flags (fecha, salida_anticipada, salida_anticipada_obs)
        VALUES (?, 1, ?)
        ON CONFLICT(fecha) DO UPDATE SET
          salida_anticipada=1,
          salida_anticipada_obs=excluded.salida_anticipada_obs
    """, (hoy, obs))
    con.commit()
    con.close()

def _get_flag_salida_anticipada(fecha_iso: str):
    _ensure_panel_schema()
    con = sqlite3.connect("reloj_control.db")
    cur = con.cursor()
    cur.execute("SELECT salida_anticipada, salida_anticipada_obs FROM panel_flags WHERE fecha=?", (fecha_iso,))
    row = cur.fetchone()
    con.close()
    if not row:
        return (0, "")
    return (row[0] or 0, row[1] or "")

# ---------- Cálculo hora salida programada ----------
def _dia_semana_es(fecha_iso: str) -> str:
    d = datetime.datetime.strptime(fecha_iso, "%Y-%m-%d").date()
    dias = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
    return dias[d.weekday()]

def _calcular_hora_salida_programada(cursor, rut: str, fecha_iso: str, hora_ingreso_hhmm: str) -> str:
    """
    Devuelve la hora de salida (HH:MM) según 'horarios' para el día/turno en que cayó el ingreso.
    Soporta turnos nocturnos (salida < entrada => día siguiente).
    Si no encuentra turno que abarque el ingreso: toma la salida más tarde del día. Fallback 17:30.
    """
    dia = _dia_semana_es(fecha_iso)

    cursor.execute("""
        SELECT hora_entrada, hora_salida, turno
        FROM horarios
        WHERE rut = ? AND dia = ?
    """, (rut, dia))
    turnos = cursor.fetchall()

    if not turnos:
        return "17:30"

    h_ing = datetime.datetime.strptime(hora_ingreso_hhmm[:5], "%H:%M")

    # Intentar encontrar el turno cuyo rango contenga la hora de ingreso
    for h_e, h_s, _t in turnos:
        if not h_e or not h_s:
            continue
        h_ini = datetime.datetime.strptime(h_e, "%H:%M")
        h_fin = datetime.datetime.strptime(h_s, "%H:%M")

        # Turno nocturno: salida al día siguiente
        if h_fin < h_ini:
            h_fin += datetime.timedelta(days=1)
            # si ingresó después de medianoche y antes de h_ini original, sumamos día
            if h_ing < h_ini:
                h_ing += datetime.timedelta(days=1)

        if h_ini <= h_ing <= h_fin:
            return h_s[:5]

    # Si no coincidió con un rango, usar la salida más tarde disponible
    salidas = [t[1] for t in turnos if t[1]]
    return (max(salidas)[:5] if salidas else "17:30")

# ---------- Botón 1: Permitir salida anticipada a TODOS ----------
def habilitar_salida_anticipada_todos(observacion):
    """
    Activa bandera para el día: a partir de ahora, cuando un funcionario marque 'salida',
    se registrará con la HORA OFICIAL del turno + la observación de autorización.
    No inserta salidas aquí (eso es tarea del botón de Cierre Forzado).
    """
    try:
        _set_flag_salida_anticipada_activa(observacion)
        messagebox.showinfo(
            "Salida Anticipada Activada",
            "Desde ahora, al marcar SALIDA, se registrará con la hora oficial de turno\n"
            "y se agregará la observación de autorización para todos los pendientes de hoy."
        )
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo activar la salida anticipada:\n{e}")

# ---------- Botón 2: Cerrar jornada para TODOS (forzado) ----------
def cerrar_dia_para_todos(observacion):
    """
    Cierra hoy la jornada de todos los RUT que tienen hora_ingreso registrada
    y NO tienen hora_salida. La salida se fija en la HORA OFICIAL del turno.
    Deja trazabilidad en panel_flags.
    """
    try:
        conexion = sqlite3.connect("reloj_control.db")
        cursor = conexion.cursor()

        # 1) Buscar pendientes con el esquema NUEVO (hora_ingreso / hora_salida)
        cursor.execute("""
            SELECT id, rut, fecha, hora_ingreso, COALESCE(observacion, '')
            FROM registros
            WHERE DATE(fecha) = DATE('now')
              AND hora_ingreso IS NOT NULL AND TRIM(hora_ingreso) <> ''
              AND (hora_salida IS NULL OR TRIM(hora_salida) = '')
        """)
        pendientes = cursor.fetchall()

        if not pendientes:
            messagebox.showinfo("Nada que cerrar", "No hay ingresos pendientes para cerrar hoy.")
            conexion.close()
            return

        # 2) Para cada pendiente, calcular hora oficial y cerrar
        for reg_id, rut, fecha, hora_ingreso, obs_prev in pendientes:
            hora_salida_prog = _calcular_hora_salida_programada(
                cursor, rut, fecha, hora_ingreso or "08:00"
            )

            # Actualizar la misma fila (no insertar otra)
            cursor.execute("""
                UPDATE registros
                   SET hora_salida = ?,
                       observacion = CASE
                           WHEN observacion IS NULL OR TRIM(observacion) = ''
                               THEN ?
                           ELSE observacion || ' | ' || ?
                       END
                 WHERE id = ?
            """, (hora_salida_prog, observacion, observacion, reg_id))

        # 3) Bitácora del cierre forzado del día
        _ensure_panel_schema()
        cursor.execute("""
            INSERT INTO panel_flags (fecha, cierre_forzado, cierre_forzado_obs)
            VALUES (DATE('now'), 1, ?)
            ON CONFLICT(fecha) DO UPDATE SET
              cierre_forzado=1,
              cierre_forzado_obs=excluded.cierre_forzado_obs
        """, (observacion,))

        conexion.commit()
        conexion.close()

        messagebox.showinfo(
            "Cierre de Jornada",
            f"Se cerró la jornada de {len(pendientes)} funcionario(s) con la hora de salida programada."
        )
    except Exception as e:
        messagebox.showerror("Error", f"Error al cerrar jornada:\n{e}")

# ---------- NUEVO: Sincronizar feriados (años específicos) ----------
def _sincronizar_feriados_anios(anios):
    if sincronizar_feriados_chile is None:
        messagebox.showerror(
            "Feriados",
            "No se encontró el módulo de feriados. Asegúrate de tener 'feriados.py' y haber instalado 'holidays'."
        )
        return
    try:
        for a in anios:
            sincronizar_feriados_chile(a)
        msg = "Se sincronizaron los feriados para: " + ", ".join(str(a) for a in anios)
        messagebox.showinfo("Feriados", msg)
    except Exception as e:
        messagebox.showerror("Error al sincronizar feriados", str(e))

# ---------- UI del Panel ----------
def construir_panel_avanzado(frame_padre):
    for widget in frame_padre.winfo_children():
        widget.destroy()

    _ensure_panel_schema()

    ctk.CTkLabel(
        frame_padre, text="Panel Avanzado (Herramientas Globales)",
        font=("Arial", 18, "bold"), text_color="#004080"
    ).pack(pady=(10, 20))

    ctk.CTkButton(
        frame_padre, width=320, fg_color="#FFA500",
        text="Permitir Salida Anticipada a Todos",
        command=lambda: mostrar_confirmacion_panel(
            "Permitir Salida Anticipada",
            "Esta acción permitirá que TODOS puedan marcar su salida en cualquier momento.\n"
            "Al hacerlo, se registrará la HORA OFICIAL del turno y se agregará la observación.",
            habilitar_salida_anticipada_todos,
            "Salida anticipada por instrucción administrativa"
        )
    ).pack(pady=10)

    ctk.CTkButton(
        frame_padre, width=320, fg_color="#DC143C",
        text="Cerrar Jornada para Todos (Emergencia/Festivo)",
        command=lambda: mostrar_confirmacion_panel(
            "Cerrar Jornada",
            "Cerrará AHORA la jornada de todos los que aún no han marcado salida.\n"
            "Se usará la hora oficial de turno y se guardará la observación.",
            cerrar_dia_para_todos,
            "Cierre de jornada por instrucción administrativa (emergencia/festivo)"
        )
    ).pack(pady=10)

    # ⇩ NUEVO botón: sincronizar feriados 2025 y 2026
    ctk.CTkButton(
        frame_padre, width=320, fg_color="#00695C",
        text="Sincronizar Feriados 2025 y 2026",
        command=lambda: _sincronizar_feriados_anios([2025, 2026])
    ).pack(pady=10)

def mostrar_confirmacion_panel(titulo, mensaje, funcion_accion, observacion_default=""):
    win = ctk.CTkToplevel()
    win.title(titulo)

    win_width, win_height = 420, 280
    sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
    x = int((sw/2) - (win_width/2))
    y = int((sh/2) - (win_height/2))
    win.geometry(f"{win_width}x{win_height}+{x}+{y}")
    win.grab_set()

    ctk.CTkLabel(win, text=mensaje, wraplength=380, justify="left").pack(pady=(30, 12))
    ctk.CTkLabel(win, text="Observación para registro:").pack()
    entry_obs = ctk.CTkEntry(win, width=360)
    entry_obs.pack(pady=10)
    entry_obs.insert(0, observacion_default)

    def aceptar():
        obs = entry_obs.get().strip()
        if not obs:
            messagebox.showerror("Observación requerida", "Debe ingresar una observación.")
            return
        funcion_accion(obs)
        win.destroy()

    btns = ctk.CTkFrame(win, fg_color="transparent")
    btns.pack(pady=18)
    ctk.CTkButton(btns, text="Aceptar", fg_color="green", width=120, command=aceptar).pack(side="left", padx=10)
    ctk.CTkButton(btns, text="Cancelar", fg_color="gray",  width=120, command=win.destroy).pack(side="left", padx=10)
