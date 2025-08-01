import customtkinter as ctk
import sqlite3
from db import crear_bd
from tkcalendar import DateEntry
from datetime import datetime
import tkinter as tk
import face_recognition
import cv2
import os
import pickle
from tkinter import messagebox

crear_bd()

def construir_registro(frame_padre, on_guardado=None):
    def registrar_rostro():
        rut = entry_rut.get().strip()
        if not rut:
            messagebox.showwarning("Advertencia", "Debe ingresar un RUT antes de registrar rostro.")
            return

        cap = cv2.VideoCapture(0)
        if not cap.isOpened():
            messagebox.showerror("Error", "No se pudo acceder a la cámara.")
            return

        messagebox.showinfo("Info", "Presiona 's' para capturar rostro, 'q' para cancelar.")

        rostro_capturado = False
        while True:
            ret, frame = cap.read()
            if not ret:
                break

            cv2.imshow("Captura de Rostro - Presiona 's' para guardar", frame)
            key = cv2.waitKey(1)

            if key == ord("s"):
                rostro_capturado = True
                ruta_guardado = os.path.join("rostros", f"{rut}.pkl")

                if not os.path.exists("rostros"):
                    os.makedirs("rostros")

                rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                encodings = face_recognition.face_encodings(rgb)

                if encodings:
                    with open(ruta_guardado, "wb") as f:
                        pickle.dump(encodings[0], f)

                    conexion = sqlite3.connect("reloj_control.db")
                    cursor = conexion.cursor()
                    cursor.execute("UPDATE trabajadores SET verificacion_facial = ? WHERE rut = ?", (f"{rut}.pkl", rut))
                    conexion.commit()
                    conexion.close()

                    messagebox.showinfo("Éxito", "Rostro registrado correctamente.")
                else:
                    messagebox.showwarning("Advertencia", "No se detectó rostro. Intenta nuevamente.")

                break

            elif key == ord("q"):
                break

        cap.release()
        cv2.destroyAllWindows()
        if not rostro_capturado:
            messagebox.showinfo("Cancelado", "Registro de rostro cancelado.")

    def guardar_trabajador():
        nombre = entry_nombre.get().strip()
        apellido = entry_apellido.get().strip()
        rut = entry_rut.get().strip()
        profesion = entry_profesion.get().strip()
        correo = entry_correo.get().strip()
        cumpleanos = entry_cumple.get()

        if not (nombre and apellido and rut):
            label_estado.configure(text="❌ Faltan campos obligatorios", text_color="red")
            return

        conexion = sqlite3.connect("reloj_control.db")
        cursor = conexion.cursor()
        try:
            cursor.execute('''
                INSERT INTO trabajadores 
                (nombre, apellido, rut, profesion, correo, cumpleanos)
                VALUES (?, ?, ?, ?, ?, ?)
            ''', (nombre, apellido, rut, profesion, correo, cumpleanos))


            for dia, turno, entrada_entry, salida_entry in campos_horarios:
                hora_entrada = entrada_entry.get().strip()
                hora_salida = salida_entry.get().strip()

                if not hora_entrada or not hora_salida:
                    label_estado.configure(text=f"⚠️ Horario incompleto en {dia} - {turno}", text_color="orange")
                    conexion.close()
                    return
                
                cursor.execute('''
                    INSERT INTO horarios (rut, dia, turno, hora_entrada, hora_salida)
                    VALUES (?, ?, ?, ?, ?)
                ''', (rut, dia, turno, hora_entrada, hora_salida))
                


            conexion.commit()
            label_estado.configure(text="✅ Nuevo Usuario registrado correctamente", text_color="green")

            for entry in entradas:
                entry.delete(0, 'end')
            entry_cumple.set_date(datetime.today())

            for _, _, ent, sal in campos_horarios:
                ent.delete(0, 'end')
                sal.delete(0, 'end')

            entry_nombre.focus()

            if on_guardado:
                on_guardado()

        except sqlite3.IntegrityError:
            label_estado.configure(text="⚠️ RUT ya registrado", text_color="orange")
        conexion.close()

    # Limpiar contenido
    for widget in frame_padre.winfo_children():
        widget.destroy()

    contenedor_central = ctk.CTkFrame(frame_padre, fg_color="transparent")
    contenedor_central.pack(anchor="center", pady=20)

    # Panel izquierdo: Datos
    panel_datos = ctk.CTkFrame(contenedor_central, corner_radius=10)
    panel_datos.grid(row=0, column=0, padx=30, sticky="n")

    ctk.CTkLabel(panel_datos, text="⚙️ Datos del Trabajador", font=("Arial", 14, "bold")).pack(pady=10)

    entry_nombre = ctk.CTkEntry(panel_datos, placeholder_text="Nombre")
    entry_apellido = ctk.CTkEntry(panel_datos, placeholder_text="Apellido")
    entry_rut = ctk.CTkEntry(panel_datos, placeholder_text="RUT (Ej: 12345678-9)")
    entry_profesion = ctk.CTkEntry(panel_datos, placeholder_text="Cargo / Profesión")
    entry_correo = ctk.CTkEntry(panel_datos, placeholder_text="Correo Electrónico")

    for e in [entry_nombre, entry_apellido, entry_rut, entry_profesion, entry_correo]:
        e.pack(pady=4)

    ctk.CTkLabel(panel_datos, text="Cumpleaños").pack(pady=(10, 2))
    entry_cumple = DateEntry(panel_datos, date_pattern="dd/mm/yyyy", width=18, background='darkblue', foreground='white', locale='es_CL')
    entry_cumple.set_date(datetime.today())
    entry_cumple.pack(pady=4)

    # 📸 Botón de registrar rostro
    ctk.CTkButton(panel_datos, text="📸 Tomar Fotografía", font=("Arial", 16), command=registrar_rostro, height=40).pack(pady=8)

    entradas = [entry_nombre, entry_apellido, entry_rut, entry_profesion, entry_correo]

    # Panel derecho: Horarios
    panel_horarios = ctk.CTkFrame(contenedor_central, fg_color="transparent")
    panel_horarios.grid(row=0, column=1, padx=30, sticky="n")

    ctk.CTkLabel(panel_horarios, text="Horario Semanal por Turno", font=("Arial", 18)).pack(pady=(5, 10))

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

    def limpiar_campos():
        for entry in entradas:
            entry.delete(0, 'end')
        entry_cumple.set_date(datetime.today())
        for _, _, ent, sal in campos_horarios:
            ent.delete(0, 'end')
            sal.delete(0, 'end')
        label_estado.configure(text="")
        entry_nombre.focus()


        # ---- Frame para centrar los botones ----
    botones_frame = ctk.CTkFrame(frame_padre, fg_color="transparent")
    botones_frame.pack(pady=15, anchor="center")

    # Botón Limpiar
    ctk.CTkButton(
        botones_frame,
        text="🧹 Limpiar Formulario",
        fg_color="gray",
        font=("Arial", 16),
        command=limpiar_campos
    ).pack(side="left", padx=10)

    # Botón Guardar
    ctk.CTkButton(
        botones_frame,
        text="💾 Guardar Nuevo Usuario",
        fg_color="green", font=("Arial", 16),
        command=guardar_trabajador
    ).pack(side="left", padx=10)

    # Label estado debajo
    label_estado = ctk.CTkLabel(frame_padre, text="")
    label_estado.pack(pady=5)


    label_estado = ctk.CTkLabel(frame_padre, text="")
    label_estado.pack(pady=5)
    


    entry_nombre.focus()
