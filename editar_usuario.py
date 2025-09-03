# editar_usuario.py
import os
import sqlite3
import threading
import pickle
from datetime import datetime

import customtkinter as ctk
from tkcalendar import DateEntry
import tkinter as tk
from tkinter import messagebox

from PIL import Image, ImageTk

import cv2
import face_recognition
import numpy as np


# ========================== PARÁMETROS BIOMÉTRICOS ==========================
N_TARGET_SAMPLES = 8

# Calidad base (ligeramente más permisivos que antes)
MIN_FACE_SIZE    = 110            # px del lado menor de la caja del rostro (antes 120)
MIN_LAPLACIAN    = 100.0          # nitidez mínima (varianza del Laplaciano) (antes 120)
BRIGHT_MIN, BRIGHT_MAX = 55, 195  # brillo medio aceptable (0–255)

# Reglas de cuadro
MAX_FACES_FRAME  = 1              # exigir una sola cara

# Anti-duplicados (si distancia < umbral => muestra “muy similar”)
DEDUP_DISTANCE_BASE = 0.34        # antes 0.36

# -------- Fluidez (auto-captura/relajación) --------
AUTO_CAPTURE_ENABLED = True        # auto-disparo sin tecla
STABLE_FRAMES_NEEDED = 8           # frames consecutivos OK para auto-disparo
RELAX_AFTER_SEC      = 12          # tras 12s sin capturar, relajar calidad
RELAX_FACTORS = {
    "laplacian": 0.80,             # 20% menos exigente en nitidez
    "face_size": 0.90,             # 10% menos exigente en tamaño
    "bright_pad": 15,              # ampliar banda de brillo ±15
    "dedup_delta": -0.04,          # baja umbral anti-duplicados (acepta algo más parecido)
}

# Guía paso a paso (se muestra 1 por cada muestra a capturar).
INSTRUCCIONES_MUESTRA = [
    "Mira al centro, rostro completo en cuadro",
    "Gira levemente a tu IZQUIERDA",
    "Gira levemente a tu DERECHA",
    "Levanta un poco la barbilla",
    "Baja un poco la barbilla",
    "Expresión neutra (sin sonrisa)",
    "Sonríe levemente / cambia expresión",
    "Muévete a luz distinta si puedes",
]


# =============================== HELPERS ====================================
def _norm_rut_filename(rut: str) -> str:
    """Normaliza RUT para nombre de archivo (sin puntos/espacios). Mantiene guion si viene."""
    return rut.replace(".", "").replace(" ", "").strip()

def _face_box_size(face_location):
    top, right, bottom, left = face_location
    return min(right - left, bottom - top)

def _crop_gray(frame_bgr, face_location):
    """Recorta rostro y aplica CLAHE para estabilizar contraste."""
    top, right, bottom, left = face_location
    top = max(0, top); left = max(0, left)
    gray = cv2.cvtColor(frame_bgr[top:bottom, left:right], cv2.COLOR_BGR2GRAY)
    if gray.size == 0:
        return gray
    # CLAHE para mejorar contraste en distintas luces
    clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
    gray = clahe.apply(gray)
    return gray

def _quality_ok(frame_bgr, loc, *, min_face_size, min_lap, bright_min, bright_max):
    """
    Devuelve (ok, motivo) tras chequear:
      - tamaño mínimo
      - nitidez (Laplaciano)
      - brillo medio dentro del recorte
    con parámetros ajustables (para relajación dinámica).
    """
    if _face_box_size(loc) < min_face_size:
        return False, "Acércate a la cámara"
    roi = _crop_gray(frame_bgr, loc)
    if roi.size == 0:
        return False, "Reencuadra"
    focus = cv2.Laplacian(roi, cv2.CV_64F).var()
    if focus < min_lap:
        return False, "Imagen borrosa"
    mean_b = float(roi.mean())
    if not (bright_min <= mean_b <= bright_max):
        return False, "Iluminación deficiente"
    return True, ""

def _guardar_biometria(rut_original: str, frame_bgr, encodings_list):
    """
    Guarda lista de encodings en PKL + foto JPG de referencia.
    Actualiza la BD con el nombre de archivo NORMALIZADO (consistente con disco).
    """
    os.makedirs("rostros", exist_ok=True)
    rut_norm = _norm_rut_filename(rut_original)
    ruta_pkl = os.path.join("rostros", f"{rut_norm}.pkl")
    ruta_jpg = os.path.join("rostros", f"{rut_norm}.jpg")

    try:
        cv2.imwrite(ruta_jpg, frame_bgr)
    except Exception:
        pass

    with open(ruta_pkl, "wb") as f:
        pickle.dump(list(encodings_list), f, protocol=pickle.HIGHEST_PROTOCOL)

    con = sqlite3.connect("reloj_control.db")
    cur = con.cursor()
    cur.execute("UPDATE trabajadores SET verificacion_facial = ? WHERE rut = ?",
                (f"{rut_norm}.pkl", rut_original))
    con.commit()
    con.close()

def _buscar_foto_por_rut_archivo(rut: str):
    """Busca una foto en /rostros por RUT (normalizado) con extensiones comunes."""
    if not rut:
        return None
    base = _norm_rut_filename(rut)
    for nombre in (f"{base}.jpg", f"{base}.jpeg", f"{base}.png"):
        p = os.path.join("rostros", nombre)
        if os.path.isfile(p):
            return p
    return None

def cargar_nombres_ruts():
    con = sqlite3.connect("reloj_control.db")
    cur = con.cursor()
    cur.execute("SELECT rut, nombre, apellido FROM trabajadores")
    nombres = []
    dict_nombre_rut = {}
    for rut, nombre, apellido in cur.fetchall():
        nom = f"{nombre} {apellido}".strip()
        nombres.append(nom)
        dict_nombre_rut[nom] = rut
    con.close()
    return nombres, dict_nombre_rut


# =========================== UI: CONSTRUIR EDICIÓN ==========================
def construir_edicion(frame_padre, on_actualizacion=None):
    # --------- limpiar contenedor ---------
    for w in frame_padre.winfo_children():
        w.destroy()

    ctk.CTkLabel(frame_padre, text="Editar / Eliminar Usuario", font=("Arial", 16, "bold")).pack(pady=(10, 10))

    contenedor = ctk.CTkFrame(frame_padre, fg_color="transparent")
    contenedor.pack(pady=10)

    cont_rut_buscar = ctk.CTkFrame(contenedor, fg_color="transparent")
    cont_rut_buscar.grid(row=0, column=0, columnspan=2, pady=10)

    # --- Cargar nombres y diccionario para el combo ---
    lista_nombres, dict_nombre_rut = cargar_nombres_ruts()

    # =================== Búsqueda por NOMBRE ===================
    fila_nombres = ctk.CTkFrame(cont_rut_buscar, fg_color="transparent")
    fila_nombres.pack(side="top", pady=(0, 5))

    combo_nombre = ctk.CTkComboBox(fila_nombres, values=lista_nombres, width=260)
    combo_nombre.set("Buscar por Nombre")
    combo_nombre.pack(side="left", padx=(0, 6))

    def _combo_placeholder_in(_e):
        if combo_nombre.get() == "Buscar por Nombre":
            combo_nombre.set("")
    def _combo_placeholder_out(_e):
        if combo_nombre.get() == "":
            combo_nombre.set("Buscar por Nombre")
    combo_nombre.bind("<FocusIn>", _combo_placeholder_in)
    combo_nombre.bind("<FocusOut>", _combo_placeholder_out)

    def autocompletar_nombres(_e):
        texto = combo_nombre.get().lower()
        if not texto or texto == "buscar por nombre":
            combo_nombre.configure(values=lista_nombres)
        else:
            filtrados = [n for n in lista_nombres if texto in n.lower()]
            combo_nombre.configure(values=filtrados if filtrados else ["No encontrado"])
    combo_nombre.bind("<KeyRelease>", autocompletar_nombres)

    def cargar_por_nombre():
        nombre = combo_nombre.get()
        rut_v = dict_nombre_rut.get(nombre, "")
        if rut_v:
            entry_rut_buscar.delete(0, 'end')
            entry_rut_buscar.insert(0, rut_v)
            cargar_usuario()
        else:
            label_estado.configure(text="⚠️ Selecciona un nombre válido", text_color="orange")

    ctk.CTkButton(fila_nombres, text="Buscar por Nombre", command=cargar_por_nombre).pack(side="left", padx=5)

    # =================== Búsqueda por RUT ===================
    fila_rut = ctk.CTkFrame(cont_rut_buscar, fg_color="transparent")
    fila_rut.pack(side="left", pady=(0, 5))

    entry_rut_buscar = ctk.CTkEntry(fila_rut, placeholder_text="Ingresa RUT a buscar", width=260)
    entry_rut_buscar.pack(side="left", padx=(10, 6))
    entry_rut_buscar.bind("<Return>", lambda e: cargar_usuario())

    ctk.CTkButton(fila_rut, text="Buscar por RUT", command=lambda: cargar_usuario()).pack(side="left", padx=5)

    # =================== Panel Datos ===================
    panel_datos = ctk.CTkFrame(contenedor, corner_radius=10)
    panel_datos.grid(row=1, column=0, padx=30, sticky="n")

    ctk.CTkLabel(panel_datos, text="📋 Datos del Trabajador", font=("Arial", 14, "bold")).pack(pady=10)

    entry_nombre   = ctk.CTkEntry(panel_datos, placeholder_text="Nombre",   width=260)
    entry_apellido = ctk.CTkEntry(panel_datos, placeholder_text="Apellido", width=260)
    entry_rut      = ctk.CTkEntry(panel_datos, placeholder_text="RUT",      width=260, state="disabled")
    entry_profesion= ctk.CTkEntry(panel_datos, placeholder_text="Profesión",width=260)
    entry_correo   = ctk.CTkEntry(panel_datos, placeholder_text="Correo",   width=260)

    for e in (entry_nombre, entry_apellido, entry_rut, entry_profesion, entry_correo):
        e.pack(pady=4)

    ctk.CTkLabel(panel_datos, text="Cumpleaños").pack(pady=(10, 2))
    entry_cumple = DateEntry(panel_datos, date_pattern="dd/mm/yyyy",
                             width=18, background='darkblue', foreground='white', locale='es_CL')
    entry_cumple.set_date(datetime.today())
    entry_cumple.pack(pady=4)

    label_verificacion = ctk.CTkLabel(panel_datos, text="Verificación facial: ---")
    label_verificacion.pack(pady=5)

    # =================== Panel Horarios ===================
    panel_horarios = ctk.CTkFrame(contenedor, fg_color="transparent")
    panel_horarios.grid(row=1, column=1, padx=30, sticky="n")

    ctk.CTkLabel(panel_horarios, text="Horario Semanal por Turno", font=("Arial", 16)).pack(pady=(5, 10))

    turnos = [("☀️ Mañana", "Mañana"), ("🕐 Tarde", "Tarde"), ("🌙 Nocturno", "Nocturno")]
    dias   = ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes"]
    campos_horarios = []

    for titulo_turno, clave_turno in turnos:
        ctk.CTkLabel(panel_horarios, text=titulo_turno + ":", font=("Arial", 13, "bold")).pack(pady=(8, 4))
        for dia in dias:
            fila = ctk.CTkFrame(panel_horarios, fg_color="transparent")
            fila.pack(pady=1)
            ctk.CTkLabel(fila, text=dia, width=80, anchor="w").pack(side="left", padx=5)
            entrada = ctk.CTkEntry(fila, width=80, placeholder_text="Entrada")
            salida  = ctk.CTkEntry(fila, width=80, placeholder_text="Salida")
            entrada.pack(side="left", padx=2)
            salida.pack(side="left", padx=2)
            campos_horarios.append((dia, clave_turno, entrada, salida))

    # =================== Estado y botones inferiores ===================
    label_estado = ctk.CTkLabel(frame_padre, text="")
    label_estado.pack(pady=5)

    contenedor_botones = ctk.CTkFrame(frame_padre, fg_color="transparent")
    contenedor_botones.pack(pady=10)

    # ---------- Acciones ----------
    def cargar_usuario():
        def tarea():
            rut = entry_rut_buscar.get().strip()
            if not rut:
                label_estado.configure(text="⚠️ Ingresa un RUT válido", text_color="orange")
                return

            con = sqlite3.connect("reloj_control.db")
            cur = con.cursor()
            cur.execute("""
                SELECT nombre, apellido, rut, profesion, correo, cumpleanos, verificacion_facial
                FROM trabajadores WHERE rut = ?
            """, (rut,))
            trabajador = cur.fetchone()

            if not trabajador:
                label_estado.configure(text="❌ Usuario no encontrado", text_color="red")
                con.close()
                return

            entry_nombre.delete(0, 'end');   entry_nombre.insert(0, trabajador[0] or "")
            entry_apellido.delete(0, 'end'); entry_apellido.insert(0, trabajador[1] or "")
            entry_rut.configure(state="normal"); entry_rut.delete(0, 'end'); entry_rut.insert(0, trabajador[2] or ""); entry_rut.configure(state="disabled")
            entry_profesion.delete(0, 'end'); entry_profesion.insert(0, trabajador[3] or "")
            entry_correo.delete(0, 'end');    entry_correo.insert(0, trabajador[4] or "")
            if trabajador[5]:
                try:
                    entry_cumple.set_date(datetime.strptime(trabajador[5], "%d/%m/%Y"))
                except Exception:
                    pass

            # Horarios
            cur.execute("SELECT dia, turno, hora_entrada, hora_salida FROM horarios WHERE rut = ?", (rut,))
            horarios = cur.fetchall()
            for _, _, ent, sal in campos_horarios:
                ent.delete(0, 'end'); sal.delete(0, 'end')
            for dia, turno, h_in, h_out in horarios:
                for d, t, e, s in campos_horarios:
                    if d == dia and t == turno:
                        if h_in:  e.insert(0, h_in)
                        if h_out: s.insert(0, h_out)
                        break

            # Verificación facial
            ruta_pkl = os.path.join("rostros", f"{_norm_rut_filename(rut)}.pkl")
            if os.path.exists(ruta_pkl):
                label_verificacion.configure(text="✅ Rostro registrado", text_color="green")
            else:
                label_verificacion.configure(text="⚠️ Rostro no registrado", text_color="orange")

            con.close()
            label_estado.configure(text="✅ Usuario cargado", text_color="green")

        threading.Thread(target=tarea, daemon=True).start()

    def guardar_cambios():
        rut = entry_rut.get().strip()
        if not rut:
            return
        con = sqlite3.connect("reloj_control.db")
        cur = con.cursor()
        cur.execute("""
            UPDATE trabajadores SET nombre=?, apellido=?, profesion=?, correo=?, cumpleanos=?
            WHERE rut=?
        """, (
            entry_nombre.get().strip(),
            entry_apellido.get().strip(),
            entry_profesion.get().strip(),
            entry_correo.get().strip(),
            entry_cumple.get(),
            rut
        ))
        cur.execute("DELETE FROM horarios WHERE rut=?", (rut,))
        for dia, turno, entrada, salida in campos_horarios:
            h_in, h_out = entrada.get().strip(), salida.get().strip()
            if h_in and h_out:
                cur.execute("""
                    INSERT INTO horarios (rut, dia, turno, hora_entrada, hora_salida)
                    VALUES (?, ?, ?, ?, ?)
                """, (rut, dia, turno, h_in, h_out))
        con.commit(); con.close()
        label_estado.configure(text="✅ Cambios guardados", text_color="green")
        if on_actualizacion:
            on_actualizacion()

    def eliminar_usuario():
        rut = entry_rut.get().strip()
        if not rut:
            return
        if not tk.messagebox.askyesno("Confirmar", f"¿Eliminar al usuario con RUT {rut}?"):
            return
        con = sqlite3.connect("reloj_control.db")
        cur = con.cursor()
        cur.execute("DELETE FROM trabajadores WHERE rut=?", (rut,))
        cur.execute("DELETE FROM horarios WHERE rut=?", (rut,))
        cur.execute("DELETE FROM registros WHERE rut=?", (rut,))
        con.commit(); con.close()
        # limpiar UI
        entry_rut_buscar.delete(0, 'end')
        for e in (entry_nombre, entry_apellido, entry_profesion, entry_correo):
            e.delete(0, 'end')
        entry_rut.configure(state="normal"); entry_rut.delete(0, 'end'); entry_rut.configure(state="disabled")
        entry_cumple.set_date(datetime.today())
        for _, _, ent, sal in campos_horarios:
            ent.delete(0, 'end'); sal.delete(0, 'end')
        label_verificacion.configure(text="Verificación facial: ---", text_color="white")
        label_estado.configure(text="🗑️ Usuario eliminado", text_color="red")
        if on_actualizacion:
            on_actualizacion()

    ctk.CTkButton(contenedor_botones, text="Guardar Cambios", command=guardar_cambios).pack(side="left", padx=10)
    ctk.CTkButton(contenedor_botones, text="Eliminar Usuario", command=eliminar_usuario, fg_color="red").pack(side="left", padx=10)

    # =================== Botones biometría ===================
    def ver_foto_registrada():
        rut = entry_rut.get().strip()
        nombre_completo = f"{entry_nombre.get().strip()} {entry_apellido.get().strip()}".strip()
        if not rut:
            tk.messagebox.showinfo("Selecciona usuario", "Primero carga un usuario (RUT) para ver su foto.")
            return
        path = _buscar_foto_por_rut_archivo(rut)
        if not path:
            tk.messagebox.showwarning(
                "Sin foto",
                "No se encontró una fotografía registrada para este RUT.\n\n"
                "Sugerencia: usa “📸 Volver a registrar rostro” para crearla."
            )
            return

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
        lbl_img.image = img_tk

        ctk.CTkButton(cont, text="Cerrar", command=win.destroy).pack(pady=(10, 0))

        frame_padre.update_idletasks()
        w, h = img_tk.width() + 64, img_tk.height() + 140
        x = frame_padre.winfo_rootx() + (frame_padre.winfo_width() // 2) - (w // 2)
        y = frame_padre.winfo_rooty() + (frame_padre.winfo_height() // 2) - (h // 2)
        win.geometry(f"{max(w, 360)}x{max(h, 260)}+{max(x, 0)}+{max(y, 0)}")

    def registrar_rostro():
        """Captura fluida: auto-captura, relajación progresiva y HUD claro."""
        rut_actual = entry_rut.get().strip()
        if not rut_actual:
            label_verificacion.configure(text="⚠️ Ingresa/carga primero el RUT", text_color="orange")
            return

        cap = cv2.VideoCapture(0)
        if not cap.isOpened():
            label_verificacion.configure(text="❌ No se pudo abrir la cámara", text_color="red")
            return

        label_verificacion.configure(
            text=("🎥 Captura: puedes usar ESPACIO/ENTER, o dejar que se auto-capture cuando se vea bien.\n"
                  "Consejo: varía ángulo / expresión / luz de forma natural."),
            text_color="white"
        )

        samples = []
        last_frame_ok = None
        dedup_thr = DEDUP_DISTANCE_BASE

        # Control de fluidez
        import time
        t_start = time.time()
        t_last_capture = t_start
        stable_ok = 0  # frames consecutivos con calidad OK

        cv2.namedWindow("Captura de rostro", cv2.WINDOW_NORMAL)
        cv2.resizeWindow("Captura de rostro", 900, 680)

        while True:
            ok_ret, frame = cap.read()
            if not ok_ret:
                continue

            rgb  = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            locs = face_recognition.face_locations(rgb, model="hog")

            # ¿Necesitamos relajar?
            elapsed_since_last = time.time() - t_last_capture
            relax = elapsed_since_last >= RELAX_AFTER_SEC

            # Parámetros actuales (con relajación si aplica)
            min_face = int(MIN_FACE_SIZE * (RELAX_FACTORS["face_size"] if relax else 1.0))
            min_lap  = float(MIN_LAPLACIAN * (RELAX_FACTORS["laplacian"] if relax else 1.0))
            bright_lo = BRIGHT_MIN - (RELAX_FACTORS["bright_pad"] if relax else 0)
            bright_hi = BRIGHT_MAX + (RELAX_FACTORS["bright_pad"] if relax else 0)
            dedup_thr = DEDUP_DISTANCE_BASE + (RELAX_FACTORS["dedup_delta"] if relax else 0.0)

            # ------------ HUD: progreso ------------
            done, total = len(samples), N_TARGET_SAMPLES
            bar_w = 340
            progress = int((done / total) * bar_w)
            cv2.rectangle(frame, (20, 20), (20 + bar_w, 40), (40, 40, 40), -1)
            cv2.rectangle(frame, (20, 20), (20 + progress, 40), (0, 200, 0), -1)
            cv2.rectangle(frame, (20, 20), (20 + bar_w, 40), (255, 255, 255), 1)
            cv2.putText(frame, f"{done}/{total}", (24 + bar_w + 8, 38), cv2.FONT_HERSHEY_SIMPLEX, 0.65, (255, 255, 255), 2)

            # ------------ Instrucción sugerida ------------
            paso_idx = min(done, len(INSTRUCCIONES_MUESTRA)-1)
            instr = INSTRUCCIONES_MUESTRA[paso_idx]
            cv2.putText(frame, instr, (20, 72), cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0,255,255), 2)

            guide_msg = "ESPACIO/ENTER = Capturar   |   ESC = Cancelar"
            guide_color = (200, 200, 200)

            can_capture_now = False
            loc_ok = None

            if len(locs) == 0:
                stable_ok = 0
                cv2.putText(frame, "Ubica el rostro dentro del cuadro", (20, 104), cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0,255,255), 2)
                key = cv2.waitKey(1) & 0xFF
            elif len(locs) > MAX_FACES_FRAME:
                stable_ok = 0
                cv2.putText(frame, "Solo 1 rostro en cuadro", (20, 104), cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0,0,255), 2)
                key = cv2.waitKey(1) & 0xFF
            else:
                loc = locs[0]
                ok_q, motivo = _quality_ok(frame, loc,
                                           min_face_size=min_face,
                                           min_lap=min_lap,
                                           bright_min=bright_lo,
                                           bright_max=bright_hi)
                top, right, bottom, left = loc
                cv2.rectangle(frame, (left, top), (right, bottom), (0, 255, 0) if ok_q else (0, 255, 255), 2)

                if ok_q:
                    stable_ok += 1
                    can_capture_now = True
                    loc_ok = loc
                    cv2.putText(frame, f"Estable: {stable_ok}/{STABLE_FRAMES_NEEDED}", (20, 104),
                                cv2.FONT_HERSHEY_SIMPLEX, 0.7, (124,255,124), 2)
                    if AUTO_CAPTURE_ENABLED and stable_ok >= STABLE_FRAMES_NEEDED:
                        key = 32  # simulamos SPACE
                    else:
                        key = cv2.waitKey(1) & 0xFF
                else:
                    stable_ok = 0
                    cv2.putText(frame, motivo, (20, 104), cv2.FONT_HERSHEY_SIMPLEX, 0.7, (0,255,255), 2)
                    key = cv2.waitKey(1) & 0xFF

            # Leyenda
            cv2.putText(frame, guide_msg, (20, 136), cv2.FONT_HERSHEY_SIMPLEX, 0.6, guide_color, 2)

            # Mostrar
            cv2.imshow("Captura de rostro", frame)

            # ---------- Teclas ----------
            if key == 27:  # ESC
                label_verificacion.configure(text="❌ Captura cancelada", text_color="orange")
                break

            elif key in (32, ord(' '), 13):  # SPACE/ENTER (o auto-disparo)
                if not can_capture_now or loc_ok is None:
                    continue

                encs = face_recognition.face_encodings(rgb, known_face_locations=[loc_ok])
                if not encs:
                    stable_ok = 0
                    continue

                enc_new = encs[0]
                # Anti-duplicados dinámico
                if samples:
                    d = face_recognition.face_distance(np.vstack(samples), enc_new).min()
                    if d < dedup_thr:
                        cv2.putText(frame, "Muy similar. Cambia angulo/luz/expresion.", (20, 168),
                                    cv2.FONT_HERSHEY_SIMPLEX, 0.6, (0,255,255), 2)
                        cv2.imshow("Captura de rostro", frame); cv2.waitKey(400)
                        stable_ok = max(0, STABLE_FRAMES_NEEDED - 3)  # no pierdas todo el progreso
                        continue

                samples.append(enc_new)
                last_frame_ok = frame.copy()
                t_last_capture = time.time()
                stable_ok = 0  # reinicia contador de estabilidad

                # feedback breve
                cv2.putText(frame, "¡Muestra guardada!", (20, 200), cv2.FONT_HERSHEY_SIMPLEX, 0.8, (0,255,0), 2)
                cv2.imshow("Captura de rostro", frame); cv2.waitKey(250)

                if len(samples) >= N_TARGET_SAMPLES:
                    break

        cap.release()
        cv2.destroyAllWindows()

        if samples:
            ref_frame = last_frame_ok if last_frame_ok is not None else np.zeros((200, 200, 3), dtype=np.uint8)
            _guardar_biometria(rut_actual, ref_frame, samples)
            label_verificacion.configure(
                text=f"✅ Rostro actualizado ({len(samples)}/{N_TARGET_SAMPLES})",
                text_color="green"
            )
            try:
                with open("log_capturas.txt", "a", encoding="utf-8") as log:
                    log.write(f"{datetime.now().strftime('%d/%m/%Y %H:%M:%S')} - Captura (edición) {rut_actual} ({len(samples)} muestras)\n")
            except Exception:
                pass
        else:
            label_verificacion.configure(text="❌ No se guardaron muestras", text_color="red")

    # Botones biometría
    ctk.CTkButton(panel_datos, text="📸 Volver a registrar rostro", font=("Arial", 14, "bold"),
                  command=registrar_rostro).pack(pady=10)
    ctk.CTkButton(panel_datos, text="👁️ Ver foto registrada",
                  command=ver_foto_registrada).pack(pady=(0, 10))

    # Binds rápidos
    combo_nombre.bind("<Return>", lambda e: cargar_por_nombre())
    entry_rut_buscar.bind("<Return>", lambda e: cargar_usuario())
