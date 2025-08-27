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
from PIL import Image, ImageTk  # ← para mostrar la foto

crear_bd()

def construir_registro(frame_padre, on_guardado=None):
    # ---------- helpers ----------
    def _buscar_foto_por_rut_archivo(rut: str):
        """Busca una foto en /rostros por RUT con extensiones comunes (normaliza puntos/espacios/guión)."""
        if not rut:
            return None
        rut_norm = rut.replace(".", "").replace(" ", "").strip()
        candidatos = [
            f"{rut_norm}.jpg", f"{rut_norm}.jpeg", f"{rut_norm}.png",
            f"{rut_norm.replace('-', '')}.jpg", f"{rut_norm.replace('-', '')}.jpeg", f"{rut_norm.replace('-', '')}.png",
        ]
        for nombre in candidatos:
            p = os.path.join("rostros", nombre)
            if os.path.isfile(p):
                return p
        return None

    # ---------- acciones ----------
    def registrar_rostro():
        rut = entry_rut.get().strip()
        if not rut:
            messagebox.showwarning("Advertencia", "Debe ingresar un RUT antes de registrar rostro.")
            return

        cap = cv2.VideoCapture(0)
        if not cap.isOpened():
            messagebox.showerror("Error", "No se pudo acceder a la cámara.")
            return

        messagebox.showinfo(
            "Captura de Rostro",
            "🎥 Enfoca tu rostro. Presiona ESPACIO para capturar o ESC para cancelar."
        )

        while True:
            ret, frame = cap.read()
            if not ret:
                continue

            cv2.putText(frame, "ESPACIO=Capturar | ESC=Salir", (30, 40),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0, 255, 255), 2)
            cv2.imshow("Captura de rostro", cv2.resize(frame, (800, 600)))
            key = cv2.waitKey(1)

            if key == 27:  # ESC
                break
            elif key == 32:  # SPACE
                rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
                encodings = face_recognition.face_encodings(rgb)

                if encodings:
                    if not os.path.exists("rostros"):
                        os.makedirs("rostros")

                    # normaliza para archivos
                    rut_norm = rut.replace(".", "").replace(" ", "")
                    ruta_pkl = os.path.join("rostros", f"{rut_norm}.pkl")
                    ruta_jpg = os.path.join("rostros", f"{rut_norm}.jpg")

                    # guarda encoding y foto
                    with open(ruta_pkl, "wb") as f:
                        pickle.dump(encodings[0], f)
                    try:
                        cv2.imwrite(ruta_jpg, frame)
                    except Exception:
                        pass

                    # guarda referencia en la BD (mantén el rut “tal cual” para compatibilidad)
                    con = sqlite3.connect("reloj_control.db")
                    cur = con.cursor()
                    cur.execute("UPDATE trabajadores SET verificacion_facial = ? WHERE rut = ?", (f"{rut}.pkl", rut))
                    con.commit()
                    con.close()

                    messagebox.showinfo("Éxito", "✅ Rostro registrado correctamente.")
                else:
                    messagebox.showwarning("Advertencia", "❌ No se detectó rostro. Intenta nuevamente.")
                break

        cap.release()
        cv2.destroyAllWindows()

    def ver_foto_registrada():
        rut = entry_rut.get().strip()
        nombre_completo = f"{entry_nombre.get().strip()} {entry_apellido.get().strip()}".strip()
        if not rut:
            messagebox.showinfo("Selecciona usuario", "Ingresa el RUT primero para buscar su foto.")
            return

        path = _buscar_foto_por_rut_archivo(rut)
        if not path:
            messagebox.showwarning("Sin foto", "No se encontró una fotografía registrada para este RUT.")
            return

        # Ventana modal con la foto
        Top = getattr(ctk, "CTkToplevel", None) or tk.Toplevel
        win = Top(frame_padre)
        win.title("Foto registrada")
        try:
            win.resizable(False, False)
            win.transient(frame_padre.winfo_toplevel())
            win.grab_set()
        except Exception:
            pass

        cont = ctk.CTkFrame(win, corner_radius=10)
        cont.pack(fill="both", expand=True, padx=16, pady=16)

        titulo = f"{nombre_completo}  |  {rut}" if nombre_completo else rut
        ctk.CTkLabel(cont, text=titulo, font=("Arial", 15, "bold")).pack(pady=(0, 8))

        img = Image.open(path)
        img.thumbnail((520, 520), Image.LANCZOS)
        img_tk = ImageTk.PhotoImage(img)

        lbl_img = ctk.CTkLabel(cont, text="")
        lbl_img.pack(pady=6)
        lbl_img.configure(image=img_tk)
        lbl_img.image = img_tk  # evita GC

        ctk.CTkButton(cont, text="Cerrar", command=win.destroy).pack(pady=(10, 0))

        # centrar
        frame_padre.update_idletasks()
        w, h = img_tk.width() + 64, img_tk.height() + 140
        x = frame_padre.winfo_rootx() + (frame_padre.winfo_width() // 2) - (w // 2)
        y = frame_padre.winfo_rooty() + (frame_padre.winfo_height() // 2) - (h // 2)
        win.geometry(f"{max(w, 360)}x{max(h, 260)}+{max(x, 0)}+{max(y, 0)}")

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

        con = sqlite3.connect("reloj_control.db")
        cur = con.cursor()
        try:
            cur.execute('''
                INSERT INTO trabajadores 
                (nombre, apellido, rut, profesion, correo, cumpleanos)
                VALUES (?, ?, ?, ?, ?, ?)
            ''', (nombre, apellido, rut, profesion, correo, cumpleanos))

            # Horarios opcionales: inserta solo si entrada y salida tienen valor
            for dia, turno, entrada_entry, salida_entry in campos_horarios:
                h_in = entrada_entry.get().strip()
                h_out = salida_entry.get().strip()
                if h_in and h_out:
                    cur.execute('''
                        INSERT INTO horarios (rut, dia, turno, hora_entrada, hora_salida)
                        VALUES (?, ?, ?, ?, ?)
                    ''', (rut, dia, turno, h_in, h_out))

            con.commit()
            label_estado.configure(text="✅ Nuevo Usuario registrado correctamente", text_color="green")

            # limpiar
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
        finally:
            con.close()

    def limpiar_campos():
        for entry in entradas:
            entry.delete(0, 'end')
        entry_cumple.set_date(datetime.today())
        for _, _, ent, sal in campos_horarios:
            ent.delete(0, 'end')
            sal.delete(0, 'end')
        label_estado.configure(text="")
        entry_nombre.focus()

    # ---------- UI ----------
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
    entry_cumple = DateEntry(panel_datos, date_pattern="dd/mm/yyyy", width=18,
                             background='darkblue', foreground='white', locale='es_CL')
    entry_cumple.set_date(datetime.today())
    entry_cumple.pack(pady=4)

    # Botones biometría
    ctk.CTkButton(panel_datos, text="📸 Tomar Fotografía", font=("Arial", 16),
                  command=registrar_rostro, height=40).pack(pady=(8, 4))
    ctk.CTkButton(panel_datos, text="👁️ Ver foto registrada",
                  command=ver_foto_registrada).pack(pady=(0, 10))

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

    # Botones inferiores
    botones_frame = ctk.CTkFrame(frame_padre, fg_color="transparent")
    botones_frame.pack(pady=15, anchor="center")

    ctk.CTkButton(
        botones_frame, text="🧹 Limpiar Formulario", fg_color="gray",
        font=("Arial", 16), command=limpiar_campos
    ).pack(side="left", padx=10)

    ctk.CTkButton(
        botones_frame, text="💾 Guardar Nuevo Usuario", fg_color="green",
        font=("Arial", 16), command=guardar_trabajador
    ).pack(side="left", padx=10)

    # Estado (único)
    label_estado = ctk.CTkLabel(frame_padre, text="")
    label_estado.pack(pady=5)

    # Focus inicial
    entry_nombre.focus()
