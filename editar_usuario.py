import customtkinter as ctk
import sqlite3
from tkcalendar import DateEntry
from datetime import datetime
import tkinter as tk
import face_recognition
import cv2
import os
import pickle
import threading
from tkinter import ttk

def cargar_nombres_ruts():
    conexion = sqlite3.connect("reloj_control.db")
    cursor = conexion.cursor()
    cursor.execute("SELECT rut, nombre, apellido FROM trabajadores")
    nombres = []
    dict_nombre_rut = {}
    for rut, nombre, apellido in cursor.fetchall():
        nombre_completo = f"{nombre} {apellido}".strip()
        nombres.append(nombre_completo)
        dict_nombre_rut[nombre_completo] = rut
    conexion.close()
    return nombres, dict_nombre_rut


def construir_edicion(frame_padre, on_actualizacion=None):
    for widget in frame_padre.winfo_children():
        widget.destroy()
    
    ctk.CTkLabel(frame_padre, text="Editar / Eliminar Usuario", font=("Arial", 16)).pack(pady=(10, 10))

    contenedor = ctk.CTkFrame(frame_padre, fg_color="transparent")
    contenedor.pack(pady=10)

    cont_rut_buscar = ctk.CTkFrame(contenedor, fg_color="transparent")
    cont_rut_buscar.grid(row=0, column=0, columnspan=2, pady=10)

    fila_busqueda = ctk.CTkFrame(cont_rut_buscar, fg_color="transparent")
    fila_busqueda.pack(side="top", pady=(0, 5))

        # Cargar nombres y diccionario
    lista_nombres, dict_nombre_rut = cargar_nombres_ruts()

    combo_nombre = ctk.CTkComboBox(fila_busqueda, values=lista_nombres, width=250)
    combo_nombre.set("Buscar por Nombre")
    combo_nombre.pack(side="left", padx=(0, 5))

    def limpiar_placeholder(event):
        if combo_nombre.get() == "Buscar por Nombre":
            combo_nombre.set("")

    def restaurar_placeholder(event):
        if combo_nombre.get() == "":
            combo_nombre.set("Buscar por Nombre")

    combo_nombre.bind("<FocusIn>", limpiar_placeholder)
    combo_nombre.bind("<FocusOut>", restaurar_placeholder)



    # Autocompletado dinámico para el ComboBox de nombres
    def autocompletar_nombres(event):
        texto = combo_nombre.get().lower()
        if not texto or texto == "buscar por nombre":
            combo_nombre.configure(values=lista_nombres)
        else:
            filtrados = [n for n in lista_nombres if texto in n.lower()]
            combo_nombre.configure(values=filtrados if filtrados else ["No encontrado"])

    combo_nombre.bind("<KeyRelease>", autocompletar_nombres)

    def cargar_usuario():
        def tarea():
            rut = entry_rut_buscar.get().strip()
            if not rut:
                label_estado.configure(text="⚠️ Ingresa un RUT válido", text_color="orange")
                return

            conexion = sqlite3.connect("reloj_control.db")
            cursor = conexion.cursor()
            cursor.execute("SELECT nombre, apellido, rut, profesion, correo, cumpleanos, verificacion_facial FROM trabajadores WHERE rut = ?", (rut,))
            trabajador = cursor.fetchone()

            if not trabajador:
                label_estado.configure(text="❌ Usuario no encontrado", text_color="red")
                conexion.close()
                return

            entry_nombre.delete(0, 'end'); entry_nombre.insert(0, trabajador[0])
            entry_apellido.delete(0, 'end'); entry_apellido.insert(0, trabajador[1])
            entry_rut.configure(state="normal"); entry_rut.delete(0, 'end'); entry_rut.insert(0, trabajador[2]); entry_rut.configure(state="disabled")
            entry_profesion.delete(0, 'end'); entry_profesion.insert(0, trabajador[3])
            entry_correo.delete(0, 'end'); entry_correo.insert(0, trabajador[4])
            if trabajador[5]:
                try:
                    entry_cumple.set_date(datetime.strptime(trabajador[5], "%d/%m/%Y"))
                except:
                    pass

            cursor.execute("SELECT dia, turno, hora_entrada, hora_salida FROM horarios WHERE rut = ?", (rut,))
            horarios = cursor.fetchall()
            for dia, turno, entrada, salida in horarios:
                for i, (d, t, e, s) in enumerate(campos_horarios):
                    if d == dia and t == turno:
                        e.delete(0, 'end'); e.insert(0, entrada)
                        s.delete(0, 'end'); s.insert(0, salida)
                        break

            ruta_facial = os.path.join("rostros", f"{rut}.pkl")
            if trabajador[6] and os.path.exists(ruta_facial):
                label_verificacion.configure(text="✅ Rostro registrado", text_color="green")
            else:
                label_verificacion.configure(text="⚠️ Rostro no registrado", text_color="orange")

            conexion.close()
            label_estado.configure(text="✅ Usuario cargado", text_color="green")

        threading.Thread(target=tarea, daemon=True).start()

    
    btn_buscar = ctk.CTkButton(fila_busqueda, text="Buscar Usuario", command=cargar_usuario)
    btn_buscar.pack(side="left", padx=5)

    def cargar_por_nombre():
        nombre = combo_nombre.get()
        rut = dict_nombre_rut.get(nombre, "")
        if rut:
            entry_rut_buscar.delete(0, 'end')
            entry_rut_buscar.insert(0, rut)
            cargar_usuario()  # Ya tienes la función, solo la llamas
        else:
            label_estado.configure(text="⚠️ Selecciona un nombre válido", text_color="orange")

    btn_buscar_nombre = ctk.CTkButton(fila_busqueda, text="Buscar Nombre", command=cargar_por_nombre)
    btn_buscar_nombre.pack(side="left", padx=5)

    # También puedes permitir buscar con ENTER en el ComboBox
    combo_nombre.bind("<Return>", lambda event: cargar_por_nombre())

    entry_rut_buscar = ctk.CTkEntry(cont_rut_buscar, placeholder_text="Ingresa RUT del funcionario a buscar", width=250)
    entry_rut_buscar.pack(side="left", padx=(0, 5))
    entry_rut_buscar.bind("<Return>", lambda event: cargar_usuario())

    def limpiar_todo():
        entry_rut_buscar.delete(0, 'end')
        entry_nombre.delete(0, 'end')
        entry_apellido.delete(0, 'end')
        entry_rut.configure(state="normal")
        entry_rut.delete(0, 'end')
        entry_rut.configure(state="disabled")
        entry_profesion.delete(0, 'end')
        entry_correo.delete(0, 'end')
        entry_cumple.set_date(datetime.today())
        for _, _, ent, sal in campos_horarios:
            ent.delete(0, 'end')
            sal.delete(0, 'end')
        label_estado.configure(text="")
        label_verificacion.configure(text="Verificación facial: ---", text_color="white")

    btn_limpiar = ctk.CTkButton(fila_busqueda, text="🧹 Limpiar Formulario", fg_color="gray", width=40, command=limpiar_todo)
    btn_limpiar.pack(side="left", padx=5)


 

    entry_rut_buscar.bind("<Return>", lambda event: cargar_usuario())   

    

    panel_datos = ctk.CTkFrame(contenedor, corner_radius=10)
    panel_datos.grid(row=1, column=0, padx=30, sticky="n")

    ctk.CTkLabel(panel_datos, text="📋 Datos del Trabajador", font=("Arial", 14, "bold")).pack(pady=10)

    entry_nombre = ctk.CTkEntry(panel_datos, placeholder_text="Nombre", width=250)
    entry_apellido = ctk.CTkEntry(panel_datos, placeholder_text="Apellido", width=250)
    entry_rut = ctk.CTkEntry(panel_datos, placeholder_text="RUT", width=250, state="disabled")
    entry_profesion = ctk.CTkEntry(panel_datos, placeholder_text="Profesión", width=250)
    entry_correo = ctk.CTkEntry(panel_datos, placeholder_text="Correo", width=250)

    for e in [entry_nombre, entry_apellido, entry_rut, entry_profesion, entry_correo]:
        e.pack(pady=4)

    ctk.CTkLabel(panel_datos, text="Cumpleaños").pack(pady=(10, 2))
    entry_cumple = DateEntry(panel_datos, date_pattern="dd/mm/yyyy", width=18, background='darkblue', foreground='white', locale='es_CL')
    entry_cumple.set_date(datetime.today())
    entry_cumple.pack(pady=4)

    label_verificacion = ctk.CTkLabel(panel_datos, text="Verificación facial: ---")
    label_verificacion.pack(pady=5)

    def registrar_rostro():
        rut_actual = entry_rut.get().strip()
        if not rut_actual:
            label_verificacion.configure(text="⚠️ Ingresa primero el RUT", text_color="orange")
            return

        cap = cv2.VideoCapture(0)
        if not cap.isOpened():
            label_verificacion.configure(text="❌ No se pudo abrir la cámara", text_color="red")
            return

        label_verificacion.configure(
            text=(
                "🎥 Enfoca bien tu rostro. Asegúrate de estar en un lugar bien iluminado.\n"
                "Presiona ESPACIO para capturar o ESC para cancelar."
            ),
            text_color="white"
        )

        rostro_capturado = False

        while True:
            ret, frame = cap.read()
            if not ret:
                continue

            cv2.putText(frame, "Presiona ESPACIO para capturar / ESC para salir", (30, 40),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 255, 255), 2)

            cv2.imshow("Captura de rostro", cv2.resize(frame, (800, 600)))
            key = cv2.waitKey(1)

            if key == 27:
                label_verificacion.configure(text="❌ Captura cancelada por el usuario", text_color="orange")
                break
            elif key == 32:
                rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                encodings = face_recognition.face_encodings(rgb)

                if encodings:
                    ruta_guardado = os.path.join("rostros", f"{rut_actual}.pkl")
                    if not os.path.exists("rostros"):
                        os.makedirs("rostros")
                    with open(ruta_guardado, "wb") as f:
                        pickle.dump(encodings[0], f)

                    conexion = sqlite3.connect("reloj_control.db")
                    cursor = conexion.cursor()
                    cursor.execute(
                        "UPDATE trabajadores SET verificacion_facial = ? WHERE rut = ?",
                        (f"{rut_actual}.pkl", rut_actual)
                    )
                    conexion.commit()
                    conexion.close()

                    with open("log_capturas.txt", "a", encoding="utf-8") as log:
                        log.write(f"{datetime.now().strftime('%d/%m/%Y %H:%M:%S')} - Captura de rostro para RUT {rut_actual}\n")

                    rostro_capturado = True
                    label_verificacion.configure(text="✅ Rostro actualizado correctamente", text_color="green")
                    break
                else:
                    label_verificacion.configure(text="❌ No se detectó rostro, intenta nuevamente", text_color="red")

        cap.release()
        cv2.destroyAllWindows()

    ctk.CTkButton(panel_datos, text="📸 Volver a registrar rostro", font=("Arial", 14, "bold"), command=registrar_rostro).pack(pady=10)


    panel_horarios = ctk.CTkFrame(contenedor, fg_color="transparent")
    panel_horarios.grid(row=1, column=1, padx=30, sticky="n")

    ctk.CTkLabel(panel_horarios, text="Horario Semanal por Turno", font=("Arial", 16)).pack(pady=(5, 10))

    turnos = [("☀️ Mañana", "Mañana"), ("🕐 Tarde", "Tarde"), ("🌙 Nocturno", "Nocturno")]
    dias = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"]
    campos_horarios = []

    for titulo_turno, clave_turno in turnos:
        ctk.CTkLabel(panel_horarios, text=titulo_turno + ":", font=("Arial", 13, "bold")).pack(pady=(8, 4))
        for dia in dias:
            fila = ctk.CTkFrame(panel_horarios, fg_color="transparent")
            fila.pack(pady=1)
            ctk.CTkLabel(fila, text=dia, width=80, anchor="w").pack(side="left", padx=5)
            entrada = ctk.CTkEntry(fila, width=80, placeholder_text="Entrada")
            entrada.pack(side="left", padx=2)
            salida = ctk.CTkEntry(fila, width=80, placeholder_text="Salida")
            salida.pack(side="left", padx=2)
            campos_horarios.append((dia, clave_turno, entrada, salida))

    label_estado = ctk.CTkLabel(frame_padre, text="")
    label_estado.pack(pady=5)

    

    def guardar_cambios():
        rut = entry_rut.get().strip()
        if not rut:
            return
        conexion = sqlite3.connect("reloj_control.db")
        cursor = conexion.cursor()
        cursor.execute('''
            UPDATE trabajadores SET nombre = ?, apellido = ?, profesion = ?, correo = ?, cumpleanos = ?
            WHERE rut = ?
        ''', (
            entry_nombre.get().strip(),
            entry_apellido.get().strip(),
            entry_profesion.get().strip(),
            entry_correo.get().strip(),
            entry_cumple.get(),
            rut
        ))

        cursor.execute("DELETE FROM horarios WHERE rut = ?", (rut,))
        for dia, turno, entrada, salida in campos_horarios:
            cursor.execute("INSERT INTO horarios (rut, dia, turno, hora_entrada, hora_salida) VALUES (?, ?, ?, ?, ?)",
                           (rut, dia, turno, entrada.get().strip(), salida.get().strip()))
        conexion.commit()
        conexion.close()
        label_estado.configure(text="✅ Cambios guardados", text_color="green")
        if on_actualizacion:
            on_actualizacion()

    def eliminar_usuario():
        rut = entry_rut.get().strip()
        if not rut:
            return
        respuesta = tk.messagebox.askyesno("Confirmar", f"¿Deseas eliminar al usuario con RUT {rut}?")
        if respuesta:
            conexion = sqlite3.connect("reloj_control.db")
            cursor = conexion.cursor()
            cursor.execute("DELETE FROM trabajadores WHERE rut = ?", (rut,))
            cursor.execute("DELETE FROM horarios WHERE rut = ?", (rut,))
            cursor.execute("DELETE FROM registros WHERE rut = ?", (rut,))
            conexion.commit()
            conexion.close()
            limpiar_todo()
            label_estado.configure(text="🗑️ Usuario eliminado", text_color="red")
            if on_actualizacion:
                on_actualizacion()

    contenedor_botones = ctk.CTkFrame(frame_padre, fg_color="transparent")
    contenedor_botones.pack(pady=10)

    ctk.CTkButton(contenedor_botones, text="Guardar Cambios", command=guardar_cambios).pack(side="left", padx=10)
    ctk.CTkButton(contenedor_botones, text="Eliminar Usuario", command=eliminar_usuario, fg_color="red").pack(side="left", padx=10)
