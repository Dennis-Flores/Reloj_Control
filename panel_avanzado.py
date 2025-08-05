import customtkinter as ctk
import sqlite3
from tkinter import messagebox
import datetime


def cerrar_dia_para_todos(observacion):
    import datetime
    try:
        conexion = sqlite3.connect("reloj_control.db")
        cursor = conexion.cursor()
        cursor.execute("""
            SELECT id, rut, fecha, hora FROM registros 
            WHERE DATE(fecha) = DATE('now') AND tipo = 'ingreso'
        """)
        ingresos = cursor.fetchall()
        for reg_id, rut, fecha, hora_ingreso in ingresos:
            fecha_obj = datetime.datetime.strptime(fecha, "%Y-%m-%d")
            dias = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
            dia_semana = dias[fecha_obj.weekday()]
            hora_ingreso_dt = datetime.datetime.strptime(hora_ingreso[:5], "%H:%M")

            # Buscar todos los turnos de ese trabajador y día
            cursor.execute("""
                SELECT hora_entrada, hora_salida, turno
                FROM horarios 
                WHERE rut = ? AND dia = ?
            """, (rut, dia_semana))
            turnos = cursor.fetchall()

            hora_salida_programada = None

            for hora_entrada, hora_salida, turno in turnos:
                if hora_entrada and hora_salida:
                    h_ini = datetime.datetime.strptime(hora_entrada, "%H:%M")
                    h_fin = datetime.datetime.strptime(hora_salida, "%H:%M")
                    # Si es turno nocturno, la salida puede ser menor que la entrada
                    if h_fin < h_ini:
                        h_fin += datetime.timedelta(days=1)
                        if hora_ingreso_dt < h_ini:
                            hora_ingreso_dt += datetime.timedelta(days=1)
                    # Si la hora de ingreso cae en el rango de este turno
                    if h_ini <= hora_ingreso_dt <= h_fin:
                        hora_salida_programada = hora_salida
                        break

            # Si no encontró turno, toma el fin más tarde del día
            if not hora_salida_programada and turnos:
                hora_salida_programada = max([t[1] for t in turnos if t[1]], default="17:30")
            if not hora_salida_programada:
                hora_salida_programada = "17:30"

            hora_salida_programada = hora_salida_programada[:5]

            cursor.execute("""
                UPDATE registros
                SET hora = ?, tipo = 'salida', nombre = 'Dirección', observacion = ?
                WHERE id = ?
            """, (hora_salida_programada, observacion, reg_id))

        conexion.commit()
        conexion.close()
        messagebox.showinfo(
            "Cierre de Ciclo",
            "Se cerró la jornada de todos los funcionarios pendientes y se registró la salida según el turno real y observación."
        )
    except Exception as e:
        messagebox.showerror("Error", f"Error al cerrar jornada:\n{e}")



def construir_panel_avanzado(frame_padre):
    for widget in frame_padre.winfo_children():
        widget.destroy()

    ctk.CTkLabel(
        frame_padre,
        text="Panel Avanzado (Herramientas Globales)",
        font=("Arial", 18, "bold"),
        text_color="#004080"
    ).pack(pady=(10, 20))

    btn_salida_anticipada = ctk.CTkButton(
        frame_padre,
        text="Permitir Salida Anticipada a Todos",
        fg_color="#FFA500",
        width=300,
        command=lambda: mostrar_confirmacion_panel(
            "Permitir Salida Anticipada",
            "Esta acción permitirá a TODOS los funcionarios marcar su salida a cualquier hora del día.\n\n"
            "La observación ingresada será registrada para todos los usuarios que aún no han marcado salida hoy.",
            habilitar_salida_anticipada_todos,
            "Salida anticipada por instrucción administrativa"
        )
    )
    btn_salida_anticipada.pack(pady=10)

    btn_cerrar_ciclo = ctk.CTkButton(
        frame_padre,
        text="Cerrar Jornada para Todos (Emergencia/Festivo)",
        fg_color="#DC143C",
        width=300,
        command=lambda: mostrar_confirmacion_panel(
            "Cerrar Jornada",
            "Esta acción cerrará la jornada de TODOS los funcionarios pendientes.\n\n"
            "La observación ingresada será registrada como motivo en la salida.",
            cerrar_dia_para_todos,
            "Cierre de jornada por instrucción administrativa (emergencia/festivo)"
        )
    )
    btn_cerrar_ciclo.pack(pady=10)

def mostrar_confirmacion_panel(titulo, mensaje, funcion_accion, observacion_default=""):
    win = ctk.CTkToplevel()
    win.title(titulo)

    win_width = 400
    win_height = 260
    screen_width = win.winfo_screenwidth()
    screen_height = win.winfo_screenheight()
    x = int((screen_width / 2) - (win_width / 2))
    y = int((screen_height / 2) - (win_height / 2))
    win.geometry(f"{win_width}x{win_height}+{x}+{y}")

    win.grab_set()
    win.grab_set()  # Bloquea la ventana principal hasta que se cierre esta

    label = ctk.CTkLabel(win, text=mensaje, wraplength=360, justify="left")
    label.pack(pady=(35, 10))

    ctk.CTkLabel(win, text="Observación para registro:").pack()
    entry_obs = ctk.CTkEntry(win, width=350)
    entry_obs.pack(pady=10)
    entry_obs.insert(0, observacion_default)

    def aceptar():
        observacion = entry_obs.get().strip()
        if not observacion:
            messagebox.showerror("Observación requerida", "Debe ingresar una observación.")
            return
        funcion_accion(observacion)
        win.destroy()

    btn_frame = ctk.CTkFrame(win, fg_color="transparent")
    btn_frame.pack(pady=15)

    ctk.CTkButton(btn_frame, text="Aceptar", fg_color="green", width=100, command=aceptar).pack(side="left", padx=10)
    ctk.CTkButton(btn_frame, text="Cancelar", fg_color="gray", width=100, command=win.destroy).pack(side="left", padx=10)

def habilitar_salida_anticipada_todos(observacion):
    try:
        conexion = sqlite3.connect("reloj_control.db")
        cursor = conexion.cursor()
        cursor.execute("""
            SELECT rut FROM registros 
            WHERE DATE(fecha) = DATE('now') AND tipo = 'ingreso'
            AND rut NOT IN (
                SELECT rut FROM registros WHERE DATE(fecha) = DATE('now') AND tipo = 'salida'
            )
        """)
        trabajadores = cursor.fetchall()
        for (rut,) in trabajadores:
            cursor.execute(
                "INSERT INTO registros (rut, fecha, tipo, observacion) VALUES (?, datetime('now'), 'salida', ?)",
                (rut, observacion)
            )
        conexion.commit()
        conexion.close()
        messagebox.showinfo("Salida Anticipada", "Se habilitó la salida anticipada a todos los funcionarios pendientes.")
    except Exception as e:
        messagebox.showerror("Error", f"Error al habilitar salida anticipada:\n{e}")
