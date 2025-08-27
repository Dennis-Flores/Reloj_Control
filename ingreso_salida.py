# ingreso_salida.py
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
from datetime import datetime, timedelta

from feriados import es_feriado  # <- NUEVO

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

# ====================== HELPERS ======================

def parse_hora(hora_str):
    for fmt in ("%H:%M:%S", "%H:%M"):
        try:
            return datetime.strptime(hora_str.strip(), fmt)
        except ValueError:
            continue
    raise ValueError(f"Formato de hora inválido: {hora_str}")

def _hoy_iso():
    return datetime.now().strftime("%Y-%m-%d")

def _dia_semana_es(fecha_iso: str) -> str:
    d = datetime.strptime(fecha_iso, "%Y-%m-%d")
    dias = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"]
    return dias[d.weekday()]

def _get_flag_salida_anticipada_local():
    """Lee bandera de salida anticipada del día (0/1, obs). No rompe si la tabla no existe."""
    try:
        con = sqlite3.connect("reloj_control.db")
        cur = con.cursor()
        cur.execute("""
            SELECT salida_anticipada, salida_anticipada_obs
            FROM panel_flags WHERE fecha=?
        """, (_hoy_iso(),))
        row = cur.fetchone()
        con.close()
        if not row:
            return (0, "")
        return (row[0] or 0, row[1] or "")
    except Exception:
        return (0, "")

def _hora_salida_oficial_por_horario(rut: str, fecha_iso: str, hora_ingreso_hhmm: str | None) -> str:
    """
    Calcula la hora de salida oficial (HH:MM) según 'horarios' para el día de 'fecha_iso'.
    Usa el bloque que contiene la hora de ingreso. Soporta turnos nocturnos.
    Si no encuentra bloque, usa la salida más tardía de ese día; si no hay, 17:30.
    """
    con = sqlite3.connect("reloj_control.db")
    cur = con.cursor()
    dia = _dia_semana_es(fecha_iso)
    cur.execute("SELECT hora_entrada, hora_salida FROM horarios WHERE rut=? AND dia=?", (rut, dia))
    turnos = cur.fetchall()
    con.close()

    if not turnos:
        return "17:30"

    if hora_ingreso_hhmm:
        h_ing = parse_hora(hora_ingreso_hhmm[:5])
        for h_e, h_s in turnos:
            if not h_e or not h_s:
                continue
            h_ini = parse_hora(h_e)
            h_fin = parse_hora(h_s)
            if h_fin < h_ini:  # nocturno
                h_fin = h_fin + timedelta(days=1)
                if h_ing < h_ini:
                    h_ing = h_ing + timedelta(days=1)
            if h_ini <= h_ing <= h_fin:
                return h_s[:5]

    salidas_validas = [s for (_e, s) in turnos if s]
    return (max(salidas_validas)[:5] if salidas_validas else "17:30")

# ======== EMERGENCIA: VALIDACIÓN POR RUT + CLAVE =========

CLAVE_MAESTRA = "2202225"

def _normalizar_rut(rut: str) -> str:
    """Elimina puntos y espacios. Mantiene guion si existe."""
    return rut.replace(".", "").replace(" ", "").strip()

def _clave_por_rut(rut: str) -> str:
    """
    Retorna los últimos 4 dígitos ANTES del guion del rut (sin puntos).
    Ej: 12.345.678-9 -> '5678'
    Si no hay guion, toma los últimos 4 del rut completo.
    """
    limpio = _normalizar_rut(rut)
    if "-" in limpio:
        parte_num = limpio.split("-")[0]
    else:
        parte_num = limpio
    solo_digitos = "".join(ch for ch in parte_num if ch.isdigit())
    return solo_digitos[-4:] if len(solo_digitos) >= 4 else solo_digitos

def validar_pass_rut(rut: str, clave: str) -> bool:
    """Acepta clave derivada del RUT o la CLAVE_MAESTRA."""
    if not rut or not clave:
        return False
    if clave == CLAVE_MAESTRA:
        return True
    return clave == _clave_por_rut(rut)

# ================== VERIFICACIÓN FACIAL ==================

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
    import threading
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
    import threading
    threading.Thread(target=proceso).start()

# ================== UI PRINCIPAL ==================

def construir_ingreso_salida(frame_padre):
    # ---------- Diálogo de EMERGENCIA ----------
    def pedir_emergencia(rut_sugerido="", mensaje="❌ Rostro no verificado. Usa clave de emergencia:"):
        TopLevelCls = getattr(ctk, "CTkToplevel", None) or tk.Toplevel
        master = frame.winfo_toplevel()
        win = TopLevelCls(master)
        win.title("🔐 Acceso de Emergencia")
        try:
            win.resizable(False, False)
            win.transient(master)
            win.grab_set()
        except Exception:
            pass

        cont = ctk.CTkFrame(win, corner_radius=12)
        cont.pack(fill="both", expand=True, padx=16, pady=16)

        # Cabecera de alerta/ayuda
        alert_box = ctk.CTkFrame(cont, corner_radius=10, fg_color="#122b39")
        alert_box.pack(fill="x", pady=(0, 10))
        ctk.CTkLabel(
            alert_box, text=mensaje,
            font=("Arial", 14, "bold"),
            text_color="#6FE6FF", justify="left", wraplength=520
        ).pack(padx=10, pady=8, anchor="w")

        # Campos
        form = ctk.CTkFrame(cont, fg_color="transparent")
        form.pack(fill="x", pady=(6, 8))

        ctk.CTkLabel(form, text="RUT:", width=110, anchor="w").grid(row=0, column=0, sticky="w", padx=(2, 6), pady=4)
        entry_rut_em = ctk.CTkEntry(form, placeholder_text="Ej: 12345678-9", width=300)
        entry_rut_em.grid(row=0, column=1, sticky="we", pady=4)

        ctk.CTkLabel(form, text="Clave:", width=110, anchor="w").grid(row=1, column=0, sticky="w", padx=(2, 6), pady=4)
        entry_pass_em = ctk.CTkEntry(form, placeholder_text="Últimos 4 del RUT o clave maestra", show="•", width=300)
        entry_pass_em.grid(row=1, column=1, sticky="we", pady=4)

        form.grid_columnconfigure(1, weight=1)

        if rut_sugerido:
            try:
                entry_rut_em.insert(0, rut_sugerido)
                entry_pass_em.focus_set()
            except Exception:
                pass

        # Pie de botones
        botones = ctk.CTkFrame(cont, fg_color="transparent")
        botones.pack(fill="x", pady=8)

        def _confirmar():
            rut_i = entry_rut_em.get().strip()
            pass_i = entry_pass_em.get().strip()
            if not rut_i or not pass_i:
                tk.messagebox.showwarning("Campo vacío", "Debes ingresar RUT y clave.")
                return

            if not validar_pass_rut(rut_i, pass_i):
                tk.messagebox.showerror("Acceso denegado", "RUT o clave inválidos.")
                return

            # OK -> continúa flujo como si estuviera verificado
            try:
                win.grab_release()
            except Exception:
                pass
            win.destroy()

            entry_rut.delete(0, tk.END)
            entry_rut.insert(0, rut_i)
            label_estado.configure(text="✅ Acceso por clave de emergencia.", text_color="yellow")
            label_hora_registro.configure(text="")
            cargar_info_usuario(rut_i, por_verificacion=True)

        ctk.CTkButton(botones, text="Cancelar", fg_color="gray", command=lambda: win.destroy()).pack(side="left", padx=6)
        ctk.CTkButton(botones, text="Validar", command=_confirmar).pack(side="right", padx=6)

        # Posición y binds
        master.update_idletasks()
        w, h = 640, 250
        x = master.winfo_x() + (master.winfo_width() // 2) - (w // 2)
        y = master.winfo_y() + (master.winfo_height() // 2) - (h // 2)
        win.geometry(f"{w}x{h}+{max(x,0)}+{max(y,0)}")
        win.bind("<Return>", lambda e: _confirmar())
        win.bind("<Escape>", lambda e: win.destroy())

    # ---------- Utilidades de esta vista ----------
    def limpiar_campos():
        entry_rut.delete(0, tk.END)
        label_nombre.configure(text="Nombre: ---")
        label_profesion.configure(text="Profesión: ---")
        label_fecha.configure(text="Fecha: ---")
        label_hora.configure(text="Hora: ---")
        label_estado.configure(text="", text_color="white")
        label_hora_registro.configure(text="")
        boton_ingreso.pack_forget()
        boton_salida.pack_forget()

    def actualizar_estado_botones(rut, por_reconocimiento=False):
        conexion = sqlite3.connect("reloj_control.db")
        cursor = conexion.cursor()
        cursor.execute("""
            SELECT hora_ingreso, hora_salida FROM registros
            WHERE rut = ? AND DATE(fecha) = DATE('now')
        """, (rut,))
        resultado = cursor.fetchone()
        conexion.close()

        boton_ingreso.pack_forget()
        boton_salida.pack_forget()

        if not resultado:
            label_estado.configure(text="🔷 Puedes registrar el ingreso.", text_color="blue")
            boton_ingreso.pack(pady=10)
            return

        hora_ingreso, hora_salida = resultado

        if not hora_ingreso:
            label_estado.configure(text="🔷 Puedes registrar el ingreso.", text_color="blue")
            boton_ingreso.pack(pady=10)
            return

        if not hora_salida:
            if por_reconocimiento:
                label_estado.configure(
                    text="✅ Verificación OK. Puedes registrar la salida.",
                    text_color="yellow"
                )
                boton_salida.pack(pady=10)
            else:
                label_estado.configure(
                    text=("✅ Ingreso realizado correctamente.\n"
                          "Para registrar la salida, verifica tu rostro nuevamente y presiona Buscar."),
                    text_color="yellow"
                )
            return

        label_estado.configure(
            text="✔️ Ya se registraron ingreso y salida hoy. Que tengas un excelente descanso.",
            text_color="green"
        )

    def registrar(tipo):
        rut = entry_rut.get().strip()
        nombre = label_nombre.cget("text").replace("Nombre: ", "")
        fecha_iso = _hoy_iso()
        hora_actual = datetime.now().strftime('%H:%M:%S')
        hora_actual_dt = parse_hora(hora_actual)

        # ------- FERIADO -------
        es_f, nombre_f, _ = es_feriado(datetime.now().date())
        obs_feriado = f"Feriado: {nombre_f}" if es_f else ""

        # Determinar día para bloques
        dia_actual = datetime.now().strftime('%A')
        dias_traducidos = {
            'Monday': 'Lunes', 'Tuesday': 'Martes', 'Wednesday': 'Miércoles',
            'Thursday': 'Jueves', 'Friday': 'Viernes', 'Saturday': 'Sábado', 'Sunday': 'Domingo'
        }
        dia_semana = dias_traducidos.get(dia_actual, '')

        conexion = sqlite3.connect("reloj_control.db")
        cursor = conexion.cursor()

        cursor.execute("""
            SELECT hora_entrada, hora_salida FROM horarios
            WHERE rut = ? AND dia = ?
        """, (rut, dia_semana))
        bloques = cursor.fetchall()

        def registrar_final(observacion=""):
            if es_f:
                observacion = (observacion + " | " if observacion else "") + obs_feriado

            if tipo == "ingreso":
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
                            UPDATE registros SET hora_ingreso = ?, observacion = ? 
                            WHERE rut = ? AND DATE(fecha) = DATE('now')
                        """, (hora_actual, observacion, rut))
                        conexion.commit()
                else:
                    cursor.execute("""
                        INSERT INTO registros (rut, nombre, fecha, hora_ingreso, hora_salida, observacion)
                        VALUES (?, ?, ?, ?, ?, ?)
                    """, (rut, nombre, fecha_iso, hora_actual, None, observacion))
                    conexion.commit()
                conexion.close()
                label_estado.configure(
                    text=("Ingreso registrado correctamente ✅\n"
                          "Limpieza automática en 60 seg..." if not es_f else
                          f"Ingreso en feriado registrado ✅ ({nombre_f}).\nLimpieza automática en 60 seg..."),
                    text_color="green"
                )

                boton_ingreso.pack_forget()
                boton_salida.pack_forget()
                actualizar_estado_botones(rut, por_reconocimiento=False)
                # Mostrar hora grande DESPUÉS de actualizar estado
                label_hora_registro.configure(text=f"⏰ Hora de registro: {hora_actual}", text_color="yellow")
                frame.after(60000, limpiar_campos)

            elif tipo == "salida":
                # Caso: salida anticipada autorizada por panel
                flag, obs_aut = _get_flag_salida_anticipada_local()
                if flag and not es_f:
                    cursor.execute("""
                        SELECT hora_ingreso, observacion FROM registros 
                        WHERE rut=? AND DATE(fecha)=DATE('now')
                    """, (rut,))
                    row = cursor.fetchone()
                    hora_ingreso_hhmm = row[0] if row else None
                    hora_oficial = _hora_salida_oficial_por_horario(rut, fecha_iso, hora_ingreso_hhmm)
                    obs_concat = (row[1] + " | " if row and row[1] else "") + (obs_aut or "Salida anticipada autorizada")

                    if row:
                        cursor.execute("SELECT hora_salida FROM registros WHERE rut=? AND DATE(fecha)=DATE('now')", (rut,))
                        hs = cursor.fetchone()
                        if hs and hs[0]:
                            label_estado.configure(text="⚠️ Ya registraste una salida hoy.", text_color="orange")
                            conexion.close()
                            return
                        cursor.execute("""
                            UPDATE registros SET hora_salida=?, observacion=? 
                            WHERE rut=? AND DATE(fecha)=DATE('now')
                        """, (hora_oficial, obs_concat, rut))
                    else:
                        cursor.execute("""
                            INSERT INTO registros (rut, nombre, fecha, hora_ingreso, hora_salida, observacion)
                            VALUES (?, ?, ?, ?, ?, ?)
                        """, (rut, nombre, fecha_iso, None, hora_oficial, obs_aut or "Salida anticipada autorizada"))
                    conexion.commit()
                    conexion.close()

                    label_estado.configure(
                        text=f"Salida anticipada registrada ✅ (hora oficial {hora_oficial}).\nLimpieza automática en 60 seg...",
                        text_color="green"
                    )

                    boton_ingreso.pack_forget()
                    boton_salida.pack_forget()
                    actualizar_estado_botones(rut, por_reconocimiento=False)
                    # Hora grande del registro realizado ahora
                    label_hora_registro.configure(text=f"⏰ Hora de registro: {hora_actual}", text_color="yellow")
                    frame.after(60000, limpiar_campos)
                    return

                # Salida normal (sin flag)
                cursor.execute("""
                    SELECT hora_salida FROM registros WHERE rut = ? AND DATE(fecha) = DATE('now')
                """, (rut,))
                resultado = cursor.fetchone()
                if resultado and resultado[0]:
                    label_estado.configure(text="⚠️ Ya registraste una salida hoy.", text_color="orange")
                    conexion.close()
                    return

                if resultado:
                    cursor.execute("""
                        UPDATE registros SET hora_salida = ?, observacion = ? 
                        WHERE rut = ? AND DATE(fecha) = DATE('now')
                    """, (hora_actual, observacion, rut))
                    conexion.commit()
                else:
                    cursor.execute("""
                        INSERT INTO registros (rut, nombre, fecha, hora_ingreso, hora_salida, observacion)
                        VALUES (?, ?, ?, ?, ?, ?)
                    """, (rut, nombre, fecha_iso, None, hora_actual, observacion))
                    conexion.commit()

                conexion.close()
                label_estado.configure(
                    text=("Salida registrada correctamente ✅\n"
                          "Limpieza automática en 60 seg..." if not es_f else
                          f"Salida en feriado registrada ✅ ({nombre_f}).\nLimpieza automática en 60 seg..."),
                    text_color="green"
                )

                boton_ingreso.pack_forget()
                boton_salida.pack_forget()
                actualizar_estado_botones(rut, por_reconocimiento=False)
                # Hora grande también en salida normal
                label_hora_registro.configure(text=f"⏰ Hora de registro: {hora_actual}", text_color="yellow")
                frame.after(60000, limpiar_campos)

        # ---------- Reglas para pedir observación (NO en feriado) ----------
        if es_f:
            registrar_final("")
            return

        requiere_observacion = False
        mensaje_motivo = ""

        if tipo == "ingreso":
            for hora_entrada, _ in bloques:
                if not hora_entrada:
                    continue
                hora_entrada_dt = parse_hora(hora_entrada)
                delta = (hora_actual_dt - hora_entrada_dt).total_seconds()
                if delta <= 5 * 60:
                    registrar_final()
                    return
                else:
                    mensaje_motivo = (
                        f"⚠️ ¡Atención! Llegas con {int(delta // 60)} min de atraso.\n"
                        f"⏱ Hora esperada: {hora_entrada_dt.strftime('%H:%M')}."
                    )
                    requiere_observacion = True
                    break

        elif tipo == "salida":
            flag, _ = _get_flag_salida_anticipada_local()
            if not flag:
                if not bloques:
                    mensaje_motivo = "⚠️ ¡Atención! No hay bloques definidos para hoy.\n✍️ Indica el motivo de salida:"
                    requiere_observacion = True
                else:
                    try:
                        bloques_validos = [s for _, s in bloques if s and s.strip()]
                        if not bloques_validos:
                            raise ValueError("No se encontró una hora de salida válida.")
                        ultima_salida_str = max(bloques_validos, key=lambda h: parse_hora(h))
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
                                f"ℹ️ Hora de salida asignada: {ultima_salida_dt.strftime('%H:%M')}.\n"
                                f"🕒 Hora actual: {hora_actual_dt.strftime('%H:%M')}.\n"
                                f"⏳ Diferencia: {int(delta // 60)} minutos ({horas:02}:{minutos:02}).\n"
                                f"✍️ Por favor, indica el motivo:"
                            )
                            requiere_observacion = True
                    except Exception:
                        mensaje_motivo = (
                            f"⚠️ ¡Atención! No se pudo determinar la hora de salida esperada.\n"
                            f"🕒 Hora actual: {hora_actual_dt.strftime('%H:%M')}\n"
                            f"✍️ Por favor, indica el motivo:"
                        )
                        requiere_observacion = True

        # --- Diálogo de observación (solo si se requiere y NO feriado) ---
        def pedir_observacion(motivo):
            TopLevelCls = getattr(ctk, "CTkToplevel", None) or tk.Toplevel
            master = frame.winfo_toplevel()
            win = TopLevelCls(master)
            win.title("📝 Observación requerida")
            try:
                win.resizable(False, False)
                win.transient(master)
                win.grab_set()
            except Exception:
                pass

            cont = ctk.CTkFrame(win, corner_radius=12)
            cont.pack(fill="both", expand=True, padx=16, pady=16)

            # --------- Cabecera (alerta) + detalle alineado en 2 columnas ----------
            lines = [l for l in motivo.split("\n") if l.strip()]
            alert_line = ""
            body_lines = []

            if lines and (lines[0].startswith("⚠️") or "¡Atención" in lines[0]):
                alert_line = lines[0]
                body_lines = lines[1:]
            else:
                body_lines = lines

            if alert_line:
                alert_box = ctk.CTkFrame(cont, corner_radius=10, fg_color="#3a2e00")
                alert_box.pack(fill="x", pady=(0, 8))
                ctk.CTkLabel(
                    alert_box, text=alert_line, font=("Arial", 14, "bold"),
                    text_color="#FFC857", justify="left", wraplength=560
                ).pack(padx=10, pady=8, anchor="w")

            ICONS = {"ℹ️", "🕒", "⏳", "✍️", "⏱"}

            def split_icon(line: str):
                parts = line.split(" ", 1)
                if len(parts) == 2 and parts[0] in ICONS:
                    return parts[0], parts[1]
                return "•", line

            rows = ctk.CTkFrame(cont, fg_color="transparent")
            rows.pack(fill="x", pady=(0, 10))

            for l in body_lines:
                icon, text = split_icon(l)
                row = ctk.CTkFrame(rows, fg_color="transparent")
                row.pack(fill="x", pady=1)

                ctk.CTkLabel(row, text=icon, width=28, font=("Arial", 14), anchor="center")\
                    .grid(row=0, column=0, sticky="nw", padx=(4, 6))

                ctk.CTkLabel(row, text=text, font=("Arial", 13), justify="left", wraplength=560)\
                    .grid(row=0, column=1, sticky="w")
                row.grid_columnconfigure(1, weight=1)

            Textbox = getattr(ctk, "CTkTextbox", None)
            if Textbox:
                entry_obs = Textbox(cont, width=560, height=130)
                entry_obs.pack(fill="both", expand=False)
                def get_text(): return entry_obs.get("1.0", "end").strip()
                def focus_text():
                    try: entry_obs.focus_set()
                    except Exception: pass
            else:
                wrap = tk.Frame(cont); wrap.pack(fill="both", expand=False)
                entry_obs = tk.Text(wrap, width=68, height=6, font=("Segoe UI", 11))
                entry_obs.pack(fill="both", expand=False)
                def get_text(): return entry_obs.get("1.0", "end").strip()
                def focus_text():
                    try: entry_obs.focus_set()
                    except Exception: pass

            counter = ctk.CTkLabel(cont, text="0/200 caracteres", text_color="gray")
            counter.pack(anchor="e", pady=(4, 0))

            def actualizar_contador(_=None):
                txt = get_text()
                if len(txt) > 200:
                    entry_obs.delete("1.0", "end"); entry_obs.insert("1.0", txt[:200])
                    txt = txt[:200]
                counter.configure(text=f"{len(txt)}/200 caracteres")

            entry_obs.bind("<KeyRelease>", actualizar_contador)

            botones = ctk.CTkFrame(cont, fg_color="transparent")
            botones.pack(fill="x", pady=12)

            def confirmar():
                txt = get_text()
                if not txt:
                    tk.messagebox.showwarning("Campo vacío", "Debes ingresar una observación.")
                    return
                try: win.grab_release()
                except Exception: pass
                win.destroy()
                registrar_final(txt)

            ctk.CTkButton(botones, text="Cancelar", fg_color="gray", command=lambda: win.destroy())\
                .pack(side="left", padx=6)
            ctk.CTkButton(botones, text="Guardar", command=confirmar)\
                .pack(side="right", padx=6)

            master.update_idletasks()
            w, h = 640, 360
            x = master.winfo_x() + (master.winfo_width() // 2) - (w // 2)
            y = master.winfo_y() + (master.winfo_height() // 2) - (h // 2)
            win.geometry(f"{w}x{h}+{max(x,0)}+{max(y,0)}")
            win.bind("<Return>", lambda e: confirmar())
            win.bind("<Escape>", lambda e: win.destroy())
            win.after(80, focus_text)
            actualizar_contador()

        if requiere_observacion:
            pedir_observacion(mensaje_motivo)
        else:
            registrar_final()

    def cargar_info_usuario(rut, por_verificacion=False):
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
                # Al cargar info para una búsqueda nueva, limpia la hora grande previa
                label_hora_registro.configure(text="")
                actualizar_estado_botones(rut, por_reconocimiento=por_verificacion)

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
            label_hora_registro.configure(text="")
            boton_ingreso.pack_forget()
            boton_salida.pack_forget()
        conexion.close()

    def buscar_rut():
        rut = entry_rut.get().strip()
        if not rut:
            label_estado.configure(text="Ingresa un RUT válido", text_color="red")
            label_hora_registro.configure(text="")
            return
        label_estado.configure(text="🔄 Verificando rostro...", text_color="gray")
        label_hora_registro.configure(text="")  # limpia posible hora previa
        frame.update()
        verificar_rostro_async(
            rut,
            callback_exito=lambda: cargar_info_usuario(rut, por_verificacion=True),
            callback_error=lambda: pedir_emergencia(rut_sugerido=rut)
        )

    def buscar_automatico():
        rut = entry_rut.get().strip()
        if not rut:
            label_estado.configure(text="🔍 Buscando rostro...", text_color="gray")
            label_hora_registro.configure(text="")
            reconocer_rostro_async(
                callback_exito=lambda rut_detectado: [
                    entry_rut.delete(0, tk.END),
                    entry_rut.insert(0, rut_detectado),
                    cargar_info_usuario(rut_detectado, por_verificacion=True)
                ],
                callback_error=lambda: pedir_emergencia(rut_sugerido="", mensaje="❌ No se pudo identificar el rostro. Usa clave de emergencia:")
            )
            return
        # Si el usuario ya puso RUT, intentamos con rostro; si falla, emergencia:
        label_estado.configure(text="🔄 Verificando rostro...", text_color="gray")
        verificar_rostro_async(
            rut,
            callback_exito=lambda: cargar_info_usuario(rut, por_verificacion=True),
            callback_error=lambda: pedir_emergencia(rut_sugerido=rut)
        )

    # ---------- UI ----------
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
    boton_salida  = ctk.CTkButton(frame, text="Registrar Salida",  font=("Arial", 15), height=45, width=220)

    # Mensaje más grande (+2)
    label_estado = ctk.CTkLabel(frame, text="", font=("Arial", 16))
    label_estado.pack(pady=10)

    # Hora grande separada
    label_hora_registro = ctk.CTkLabel(
        frame,
        text="",
        font=("Arial", 36, "bold"),
        text_color="yellow"
    )
    label_hora_registro.pack(pady=(25, 35))
