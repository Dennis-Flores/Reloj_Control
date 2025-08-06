import os
import sys
import shutil
import face_recognition
import dlib
import customtkinter as ctk
import sqlite3
import tkinter as tk
import pickle
import cv2
from datetime import datetime
import threading
from datetime import datetime, timedelta


# === SETUP MODELOS ===
if getattr(sys, 'frozen', False):
    ruta_base = sys._MEIPASS
else:
    ruta_base = os.path.dirname(__file__)

origen_dat = os.path.join(ruta_base, "face_recognition_models", "models", "shape_predictor_68_face_landmarks.dat")
destino_dir = os.path.join(os.environ.get("TEMP"), "face_recognition_models", "models")
os.makedirs(destino_dir, exist_ok=True)
destino_dat = os.path.join(destino_dir, "shape_predictor_68_face_landmarks.dat")

origen_5pt = os.path.join(ruta_base, "face_recognition_models", "models", "shape_predictor_5_face_landmarks.dat")
destino_5pt = os.path.join(destino_dir, "shape_predictor_5_face_landmarks.dat")

if not os.path.exists(destino_5pt):
    try:
        shutil.copy(origen_5pt, destino_5pt)
    except:
        pass

print(f"✅ Copiando modelo desde:\n  {origen_dat}\na:\n  {destino_dat}")
if not os.path.exists(destino_dat):
    shutil.copy(origen_dat, destino_dat)

# === FUNCIONES ===

def verificar_rostro(rut):
    archivo_rostro = os.path.join("rostros", f"{rut}.pkl")
    if not os.path.exists(archivo_rostro):
        return False

    with open(archivo_rostro, "rb") as f:
        rostro_guardado = pickle.load(f)

    cap = cv2.VideoCapture(0)
    if not cap.isOpened():
        return False

    verificado = False
    instrucciones_mostradas = False

    for _ in range(60):
        ret, frame = cap.read()
        if not ret:
            continue

        if not instrucciones_mostradas:
            cv2.putText(frame, "Mire al frente sin moverse", (40, 40), cv2.FONT_HERSHEY_SIMPLEX, 0.9, (0, 255, 255), 2)
            instrucciones_mostradas = True

        rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        rostros = face_recognition.face_encodings(rgb)

        for rostro in rostros:
            if face_recognition.compare_faces([rostro_guardado], rostro)[0]:
                verificado = True
                break

        cv2.imshow("Verificación Facial", cv2.resize(frame, (800, 600)))
        if cv2.waitKey(1) & 0xFF == ord('q') or verificado:
            break

    cap.release()
    cv2.destroyAllWindows()
    return verificado

def reconocer_rostro_sin_rut():
    cap = cv2.VideoCapture(0)
    if not cap.isOpened():
        return None

    rostro_detectado = None
    for _ in range(60):
        ret, frame = cap.read()
        if not ret:
            continue

        rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        rostros_en_vivo = face_recognition.face_encodings(rgb)

        if rostros_en_vivo:
            rostro_actual = rostros_en_vivo[0]
            for archivo in os.listdir("rostros"):
                if archivo.endswith(".pkl"):
                    rut_archivo = archivo.replace(".pkl", "")
                    with open(os.path.join("rostros", archivo), "rb") as f:
                        rostro_guardado = pickle.load(f)

                    if face_recognition.compare_faces([rostro_guardado], rostro_actual)[0]:
                        rostro_detectado = rut_archivo
                        break

        cv2.imshow("Reconocimiento Automático", cv2.resize(frame, (800, 600)))
        if cv2.waitKey(1) & 0xFF == ord('q') or rostro_detectado:
            break

    cap.release()
    cv2.destroyAllWindows()
    return rostro_detectado

def verificar_rostro_async(rut, callback_exito, callback_error):
    def proceso():
        try:
            exito = verificar_rostro(rut)
            if exito:
                callback_exito()
            else:
                callback_error()
        except Exception as e:
            print("Error en verificación:", e)
            callback_error()
    threading.Thread(target=proceso).start()

def reconocer_rostro_async(callback_exito, callback_error):
    def proceso():
        try:
            rut_detectado = reconocer_rostro_sin_rut()
            if rut_detectado:
                callback_exito(rut_detectado)
            else:
                callback_error()
        except Exception as e:
            print("Error en reconocimiento:", e)
            callback_error()
    threading.Thread(target=proceso).start()


def construir_ingreso_salida(frame_padre):
    def limpiar_campos():
        entry_rut.delete(0, tk.END)
        label_nombre.configure(text="Nombre: ---")
        label_profesion.configure(text="Profesión: ---")
        label_fecha.configure(text="Fecha: ---")
        label_hora.configure(text="Hora: ---")
        label_estado.configure(text="", text_color="white")
        boton_ingreso.pack_forget()
        boton_salida.pack_forget()

    def actualizar_estado_botones(rut):
        conexion = sqlite3.connect("reloj_control.db")
        cursor = conexion.cursor()
        # Busca registro solo para el día de hoy
        cursor.execute("""
            SELECT hora_ingreso, hora_salida FROM registros
            WHERE rut = ? AND DATE(fecha) = DATE('now')
        """, (rut,))
        resultado = cursor.fetchone()
        conexion.close()

        # Oculta ambos botones por defecto
        boton_ingreso.pack_forget()
        boton_salida.pack_forget()

        if not resultado:
            # No hay registro hoy: solo puede ingresar
            label_estado.configure(text="🔷 Puedes registrar el ingreso.", text_color="blue")
            boton_ingreso.pack(pady=10)
        else:
            hora_ingreso, hora_salida = resultado
            if not hora_ingreso:
                # Registro creado pero aún no ha ingresado (caso muy raro)
                label_estado.configure(text="🔷 Puedes registrar el ingreso.", text_color="blue")
                boton_ingreso.pack(pady=10)
            elif not hora_salida:
                # Ya registró ingreso, pero aún no salida
                label_estado.configure(
                    text="✅ Ingreso realizado correctamente. Ahora puedes registrar la salida.",
                    text_color="yellow"
                )
                boton_salida.pack(pady=10)
            else:
                # Ya marcó ingreso y salida
                label_estado.configure(
                    text="✔️ Ya se registró ingreso y salida hoy. Que tengas un excelente descanso.",
                    text_color="green"
                )



    def parse_hora(hora_str):
        for fmt in ("%H:%M:%S", "%H:%M"):
            try:
                return datetime.strptime(hora_str.strip(), fmt)
            except ValueError:
                continue
        raise ValueError(f"Formato de hora inválido: {hora_str}")


    def registrar(tipo):
        rut = entry_rut.get().strip()
        nombre = label_nombre.cget("text").replace("Nombre: ", "")
        fecha = datetime.now().strftime('%Y-%m-%d')
        hora_actual = datetime.now().strftime('%H:%M:%S')
        hora_actual_dt = parse_hora(hora_actual)

        dia_actual = datetime.now().strftime('%A')
        dias_traducidos = {
            'Monday': 'Lunes',
            'Tuesday': 'Martes',
            'Wednesday': 'Miércoles',
            'Thursday': 'Jueves',
            'Friday': 'Viernes',
            'Saturday': 'Sábado',
            'Sunday': 'Domingo'
        }
        dia_semana = dias_traducidos.get(dia_actual, '')

        conexion = sqlite3.connect("reloj_control.db")
        cursor = conexion.cursor()

        # Obtener todos los bloques del día (mañana, tarde, etc.)
        cursor.execute("""
            SELECT hora_entrada, hora_salida FROM horarios
            WHERE rut = ? AND dia = ?
        """, (rut, dia_semana))
        bloques = cursor.fetchall()

        def registrar_final(observacion=""):
            if tipo == "ingreso":
                # ... (ya lo tienes correcto)
                cursor.execute("""
                    SELECT hora_ingreso FROM registros WHERE rut = ? AND DATE(fecha) = DATE('now')
                """, (rut,))
                resultado = cursor.fetchone()
                if resultado:
                    if resultado[0]:
                        label_estado.configure(text="⚠️ Ya registraste un ingreso hoy.", text_color="orange")
                        conexion.close()
                        return
                    else:
                        cursor.execute("""
                            UPDATE registros SET hora_ingreso = ?, observacion = ? WHERE rut = ? AND DATE(fecha) = DATE('now')
                        """, (hora_actual, observacion, rut))
                        conexion.commit()
                else:
                    cursor.execute("""
                        INSERT INTO registros (rut, nombre, fecha, hora_ingreso, hora_salida, observacion)
                        VALUES (?, ?, ?, ?, ?, ?)
                    """, (rut, nombre, fecha, hora_actual, None, observacion))
                    conexion.commit()
                conexion.close()
                label_estado.configure(
                    text=f"Ingreso registrado correctamente ✅\nLimpieza automática en 60 seg...",
                    text_color="green"
                )
                boton_ingreso.pack_forget()
                boton_salida.pack_forget()
                actualizar_estado_botones(rut)
                frame.after(60000, limpiar_campos)

            elif tipo == "salida":
                # --- NUEVO FLUJO DE SALIDA ---
                cursor.execute("""
                    SELECT hora_salida FROM registros WHERE rut = ? AND DATE(fecha) = DATE('now')
                """, (rut,))
                resultado = cursor.fetchone()
                if resultado:
                    if resultado[0]:
                        label_estado.configure(text="⚠️ Ya registraste una salida hoy.", text_color="orange")
                        conexion.close()
                        return
                    else:
                        cursor.execute("""
                            UPDATE registros SET hora_salida = ?, observacion = ? WHERE rut = ? AND DATE(fecha) = DATE('now')
                        """, (hora_actual, observacion, rut))
                        conexion.commit()
                else:
                    # Si no hay fila de ingreso, permite crear la fila SOLO con salida (flujo de emergencia)
                    cursor.execute("""
                        INSERT INTO registros (rut, nombre, fecha, hora_ingreso, hora_salida, observacion)
                        VALUES (?, ?, ?, ?, ?, ?)
                    """, (rut, nombre, fecha, None, hora_actual, observacion))
                    conexion.commit()
                conexion.close()
                label_estado.configure(
                    text=f"Salida registrada correctamente ✅\nLimpieza automática en 60 seg...",
                    text_color="green"
                )
                boton_ingreso.pack_forget()
                boton_salida.pack_forget()
                actualizar_estado_botones(rut)
                frame.after(60000, limpiar_campos)



        def pedir_observacion(motivo):
            obs_win = tk.Toplevel()
            obs_win.title("📝 Observación requerida")
            obs_win.geometry("540x220")
            obs_win.resizable(False, False)

            # Centrado en pantalla
            obs_win.update_idletasks()
            x = (obs_win.winfo_screenwidth() // 2) - (540 // 2)
            y = (obs_win.winfo_screenheight() // 2) - (220 // 2)
            obs_win.geometry(f"+{x}+{y}")

            # Frame visual contenedor
            obs_frame = tk.Frame(obs_win, bg="#f9f9f9", bd=2, relief="groove")
            obs_frame.pack(fill="both", expand=True, padx=10, pady=10)

            # Mensaje de alerta
            label_mensaje = tk.Label(
                obs_frame,
                text=motivo,
                font=("Segoe UI", 10),
                justify="left",
                wraplength=500,
                bg="#f9f9f9",
                fg="#333"
            )
            label_mensaje.pack(pady=(10, 5), padx=15)

            # Campo de texto
            entry_obs = tk.Entry(obs_frame, width=55, font=("Segoe UI", 10))
            entry_obs.pack(padx=15, pady=(5, 10))

            # Botones
            contenedor_botones = tk.Frame(obs_frame, bg="#f9f9f9")
            contenedor_botones.pack(pady=(0, 10))

            def confirmar_obs():
                obs_text = entry_obs.get().strip()
                if obs_text:
                    obs_win.destroy()
                    registrar_final(obs_text)
                else:
                    tk.messagebox.showwarning("Campo vacío", "Debes ingresar una observación.")

            def cancelar_obs():
                obs_win.destroy()

            tk.Button(
                contenedor_botones, text="Guardar", width=14, bg="#d0f0c0",
                relief="ridge", font=("Segoe UI", 10), command=confirmar_obs
            ).pack(side="left", padx=15)

            tk.Button(
                contenedor_botones, text="Cancelar", width=14, bg="#f0d0d0",
                relief="ridge", font=("Segoe UI", 10), command=cancelar_obs
            ).pack(side="left", padx=15)

            # Ventana modal
            obs_win.transient(frame)
            obs_win.grab_set()
            obs_win.focus_set()



        requiere_observacion = False
        mensaje_motivo = ""

        if tipo == "ingreso":
            for hora_entrada, _ in bloques:
                if not hora_entrada:
                    continue

                hora_entrada_dt = parse_hora(hora_entrada)
                delta = (hora_actual_dt - hora_entrada_dt).total_seconds()
                if delta <= 5 * 60:  # margen de 5 minutos
                    registrar_final()
                    return
                else:
                    horas = abs(int(delta) // 3600)
                    minutos = (abs(int(delta)) % 3600) // 60
                    mensaje_motivo = f"Llegó con {int(delta // 60)} min de atraso ({horas:02}:{minutos:02}, hora esperada: {hora_entrada_dt.strftime('%H:%M')}). ¿Motivo?"
                    requiere_observacion = True
                    break

        elif tipo == "salida":
            if not bloques:
                mensaje_motivo = "No hay bloques definidos para hoy. ¿Motivo de salida?"
                requiere_observacion = True
            else:
                try:
                    # Elegir la hora_salida más tarde de todos los bloques del día
                    bloques_validos = [s for _, s in bloques if s and s.strip()]
                    if not bloques_validos:
                        raise ValueError("No se encontró una hora de salida válida.")

                    ultima_salida_str = max(bloques_validos, key=lambda h: parse_hora(h))
                    ultima_salida_dt = parse_hora(ultima_salida_str)


                    ultima_salida_dt = parse_hora(ultima_salida_str)
                    ultima_salida_dt = parse_hora(ultima_salida_str)

                    if hora_actual_dt >= ultima_salida_dt:
                        registrar_final()

                        return
                    else:
                        delta = (ultima_salida_dt - hora_actual_dt).total_seconds()
                        horas = int(delta // 3600)
                        minutos = int((delta % 3600) // 60)
                        mensaje_motivo = (
                            f"⚠️ ¡Atención! Estás registrando una salida antes del horario establecido.\n"
                            f"🕔 Hora de salida asignada: {ultima_salida_dt.strftime('%H:%M')}.\n"
                            f"🕒 Hora actual: {hora_actual_dt.strftime('%H:%M')}.\n"
                            f"⏳ Diferencia: {int(delta // 60)} minutos ({horas:02}:{minutos:02}), antes del horario esperado.\n"
                            f"✍️ Por favor, indica el motivo:"
                        )
                        requiere_observacion = True

                except Exception as e:
                    mensaje_motivo = (
                        f"⚠️ No se pudo determinar la hora de salida esperada.\n"
                        f"🕒 Hora actual: {hora_actual_dt.strftime('%H:%M')}\n"
                        f"✍️ Por favor, indica el motivo:"
                    )
                    requiere_observacion = True



        if requiere_observacion:
            pedir_observacion(mensaje_motivo)
        else:
            registrar_final()



    def cargar_info_usuario(rut):
        conexion = sqlite3.connect("reloj_control.db")
        cursor = conexion.cursor()
        cursor.execute("SELECT nombre, apellido, profesion, cumpleanos FROM trabajadores WHERE rut = ?", (rut,))
        resultado = cursor.fetchone()

        if resultado:
            nombre_completo = f"{resultado[0]} {resultado[1]}"
            profesion = resultado[2]
            label_nombre.configure(text=f"Nombre: {nombre_completo}")
            label_profesion.configure(text=f"Profesión: {profesion}")
            label_fecha.configure(text=f"Fecha: {datetime.now().strftime('%d/%m/%Y')}")
            label_hora.configure(text=f"Hora: {datetime.now().strftime('%H:%M:%S')}")

            cumpleanos = resultado[3] if resultado[3] else ""
            hoy = datetime.now().strftime('%d/%m')
            if cumpleanos and hoy == cumpleanos[:5]:
                label_estado.configure(text=f"🎂 ¡Hoy es el cumpleaños de {resultado[0]}! 🎉", text_color="yellow")
            else:
                actualizar_estado_botones(rut)

            boton_ingreso.configure(command=None)
            boton_ingreso.configure(command=lambda: registrar("ingreso"))

            boton_salida.configure(command=None)
            boton_salida.configure(command=lambda: registrar("salida"))

        else:
            label_nombre.configure(text="Nombre: ---")
            label_profesion.configure(text="Profesión: ---")
            label_fecha.configure(text="Fecha: ---")
            label_hora.configure(text="Hora: ---")
            label_estado.configure(text="⚠️ RUT no encontrado", text_color="red")
            boton_ingreso.pack_forget()
            boton_salida.pack_forget()
        conexion.close()

    def buscar_rut():
        rut = entry_rut.get().strip()
        if not rut:
            label_estado.configure(text="Ingresa un RUT válido", text_color="red")
            return

        label_estado.configure(text="🔄 Verificando rostro...", text_color="gray")
        frame.update()

        verificar_rostro_async(
            rut,
            callback_exito=lambda: cargar_info_usuario(rut),
            callback_error=lambda: label_estado.configure(text="❌ Rostro no verificado", text_color="red")
        )

    def buscar_automatico():
        rut = entry_rut.get().strip()
        if not rut:
            label_estado.configure(text="🔍 Buscando rostro...", text_color="gray")
            reconocer_rostro_async(
                callback_exito=lambda rut_detectado: [
                    entry_rut.delete(0, tk.END),
                    entry_rut.insert(0, rut_detectado),
                    cargar_info_usuario(rut_detectado)
                ],
                callback_error=lambda: label_estado.configure(text="❌ No se pudo identificar el rostro", text_color="red")
            )
            return

        buscar_rut()

    

    frame = ctk.CTkFrame(frame_padre)
    frame.pack(fill="both", expand=True)

    ctk.CTkLabel(frame, text="Ingreso / Salida de Funcionarios", font=("Arial", 16)).pack(pady=10)
    ctk.CTkLabel(frame, text="Ingresa el RUT del funcionario:").pack(pady=(10, 2))
    entry_rut = ctk.CTkEntry(frame, placeholder_text="Ej: 12345678-9")
    entry_rut.pack(pady=5)
    entry_rut.bind("<Return>", lambda event: buscar_automatico())

    ctk.CTkButton(frame, text="Buscar", command=buscar_automatico).pack(pady=5)
    ctk.CTkButton(frame, text="Limpiar", command=limpiar_campos).pack(pady=5)

    label_nombre = ctk.CTkLabel(frame, text="Nombre: ---", font=("Arial", 14))
    label_nombre.pack(pady=(15, 5))
    label_profesion = ctk.CTkLabel(frame, text="Profesión: ---", font=("Arial", 14))
    label_profesion.pack(pady=5)
    label_fecha = ctk.CTkLabel(frame, text="Fecha: ---", font=("Arial", 14))
    label_fecha.pack(pady=5)
    label_hora = ctk.CTkLabel(frame, text="Hora: ---", font=("Arial", 14))
    label_hora.pack(pady=5)

    boton_ingreso = ctk.CTkButton(frame, text="Registrar Ingreso", font=("Arial", 15), height=45, width=220)
    boton_salida = ctk.CTkButton(frame, text="Registrar Salida", font=("Arial", 15), height=45, width=220)

    label_estado = ctk.CTkLabel(frame, text="")
    label_estado.pack(pady=10)