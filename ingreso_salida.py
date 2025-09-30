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
import time
from datetime import datetime, timedelta

from feriados import es_feriado

# ======== OpenCV: forzar backend estable y silenciar logs ========
os.environ.setdefault("OPENCV_VIDEOIO_PRIORITY_MSMF", "0")
os.environ.setdefault("OPENCV_LOG_LEVEL", "SILENT")

# ============== RUTAS ROBUSTAS PARA EJECUTABLE / INTERPRETADO ==============
def _base_dir():
    # En onedir: carpeta del .exe; en interpretado: carpeta del .py
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

BASE = _base_dir()

DB_PATH        = os.path.join(BASE, "reloj_control.db")
ROSTROS_DIR    = os.path.join(BASE, "rostros")
MODELOS_DIR    = os.path.join(BASE, "face_recognition_models", "models")
SALIDAS_DIR    = os.path.join(BASE, "salidas_solicitudes")
ASSETS_DIR     = os.path.join(BASE, "assets")

os.makedirs(ROSTROS_DIR, exist_ok=True)
os.makedirs(SALIDAS_DIR, exist_ok=True)

# ================== LOG SENCILLO A ARCHIVO (junto al .exe) ==================
LOG_PATH = os.path.join(BASE, "log_capturas.txt")
def log(msg: str):
    try:
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}\n")
    except Exception:
        pass

# Info de arranque (útil al empaquetar)
print(f"[init] BASE={BASE}")
print(f"[init] DB_PATH={DB_PATH}")
print(f"[init] ROSTROS_DIR={ROSTROS_DIR}")
log(f"Inicio app | BASE={BASE}")

# ============== MODELOS DLIB (copiamos a %TEMP% si hace falta) =============
origen_dat = os.path.join(MODELOS_DIR, "shape_predictor_68_face_landmarks.dat")
origen_5pt = os.path.join(MODELOS_DIR, "shape_predictor_5_face_landmarks.dat")
destino_dir = os.path.join(os.environ.get("TEMP", ""), "face_recognition_models", "models")
os.makedirs(destino_dir, exist_ok=True)
destino_dat = os.path.join(destino_dir, "shape_predictor_68_face_landmarks.dat")
destino_5pt = os.path.join(destino_dir, "shape_predictor_5_face_landmarks.dat")

print(f"✅ Copiando modelo desde:\n  {origen_dat}\na:\n  {destino_dat}")
try:
    if os.path.exists(origen_dat) and not os.path.exists(destino_dat):
        shutil.copy(origen_dat, destino_dat)
    if os.path.exists(origen_5pt) and not os.path.exists(destino_5pt):
        shutil.copy(origen_5pt, destino_5pt)
except Exception as e:
    print("Aviso: no se pudo copiar a TEMP:", e)
    log(f"Aviso copia modelos: {e}")

# ====================== HELPERS DE TIEMPO/BD ======================

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
    try:
        con = sqlite3.connect(DB_PATH)
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
    except Exception as e:
        log(f"_get_flag_salida_anticipada_local error: {e}")
        return (0, "")

def _hora_salida_oficial_por_horario(rut: str, fecha_iso: str, hora_ingreso_hhmm: str | None) -> str:
    """
    DEVUELVE LA SALIDA OFICIAL (FIN DE JORNADA): la ÚLTIMA hora_salida del día para el RUT y día de semana.
    Ignora colación. Si no hay turnos, retorna '17:30'.
    """
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    dia = _dia_semana_es(fecha_iso)
    cur.execute("SELECT hora_entrada, hora_salida FROM horarios WHERE rut=? AND dia=?", (rut, dia))
    turnos = cur.fetchall()
    con.close()

    if not turnos:
        return "17:30"

    salidas_validas = [s for (_e, s) in turnos if s and s.strip()]
    if not salidas_validas:
        return "17:30"

    # Elegimos la mayor por hora del día (no contempla nocturnos cruzando medianoche)
    try:
        ultima = max(salidas_validas, key=lambda h: parse_hora(h))
    except Exception:
        ultima = max(salidas_validas)  # fallback por string
    return ultima[:5]

# ================== EMERGENCIA: VALIDACIÓN POR RUT + CLAVE ==================
CLAVE_MAESTRA = "2202225"

def _normalizar_rut(rut: str) -> str:
    return rut.replace(".", "").replace(" ", "").strip()

def _clave_por_rut(rut: str) -> str:
    limpio = _normalizar_rut(rut)
    parte_num = limpio.split("-")[0] if "-" in limpio else limpio
    solo_digitos = "".join(ch for ch in parte_num if ch.isdigit())
    return solo_digitos[-4:] if len(solo_digitos) >= 4 else solo_digitos

def validar_pass_rut(rut: str, clave: str) -> bool:
    if not rut or not clave:
        return False
    if clave == CLAVE_MAESTRA:
        return True
    return clave == _clave_por_rut(rut)

# -------- utilidades de RUT (formateo y validación) --------
def _rut_limpio(r: str) -> str:
    # Solo letras/números, sin puntos ni guiones, DV en mayúscula
    s = "".join(ch for ch in r if ch.isalnum()).upper()
    return s

def _rut_formatear(r: str) -> str:
    """
    Convierte '123456789' -> '12345678-9', '12.345.678-9' -> '12345678-9'.
    Si no hay longitud suficiente, retorna tal cual.
    """
    s = _rut_limpio(r)
    if len(s) < 2:
        return r.strip()
    cuerpo, dv = s[:-1], s[-1]
    if not cuerpo.isdigit():
        return r.strip()
    return f"{cuerpo}-{dv}"

def _rut_valido(r: str) -> bool:
    """
    Valida DV por algoritmo Módulo 11. Acepta DV 'K'.
    """
    s = _rut_limpio(r)
    if len(s) < 2 or not s[:-1].isdigit():
        return False
    cuerpo, dv = s[:-1], s[-1]
    suma = 0
    factor = 2
    for c in reversed(cuerpo):
        suma += int(c) * factor
        factor += 1
        if factor > 7:
            factor = 2
    resto = suma % 11
    res = 11 - resto
    dv_calc = '0' if res == 11 else 'K' if res == 10 else str(res)
    return dv_calc == dv

# ================== VERIFICACIÓN FACIAL ==================
FACIAL_TOLERANCE = 0.50      # más bajo = más estricto
DISTANCE_MARGIN  = 0.04
FRIDAY_FLEX_MINUTES = 30     # margen de “colación” (viernes)
LATE_AFTER_EXIT_MINUTES = 40 # observación si supera la salida final por 40+ min

# ---------- configuración de extras ----------
EXTRA_MINUTES_THRESHOLD = 30
EXTRA_COUNT_ONLY_ABOVE_THRESHOLD = True

def _ensure_list_encodings(obj):
    if isinstance(obj, (list, tuple)):
        return list(obj)
    return [obj]

# ===== Loader tolerante a variantes de nombre de archivo =====
def _load_encodings_for_rut(rut: str):
    candidatos = set()

    r_fmt = _rut_formatear(rut)          # 12345678-K
    r_raw = _rut_limpio(rut)             # 12345678K

    # Variantes comunes con y sin guion, K/k
    candidatos.update([
        r_fmt,                            # 12345678-K
        r_fmt.lower(),                    # 12345678-k
        r_raw,                            # 12345678K
        (r_raw[:-1] + "-" + r_raw[-1]) if len(r_raw) >= 2 else r_fmt,          # 12345678-K
        ((r_raw[:-1] + "-" + r_raw[-1]).lower()) if len(r_raw) >= 2 else r_fmt  # 12345678-k
    ])

    for base in candidatos:
        p = os.path.join(ROSTROS_DIR, f"{base}.pkl")
        if os.path.exists(p):
            with open(p, "rb") as f:
                data = pickle.load(f)
            return _ensure_list_encodings(data)
    return []

def _load_all_known_encodings():
    encs, ruts = [], []
    if not os.path.isdir(ROSTROS_DIR):
        return encs, ruts
    for archivo in os.listdir(ROSTROS_DIR):
        if archivo.endswith(".pkl"):
            rut = archivo[:-4]
            for enc in _load_encodings_for_rut(rut):
                encs.append(enc)
                ruts.append(rut)
    return encs, ruts

def _debug_listar_pkl():
    try:
        archivos = [a for a in os.listdir(ROSTROS_DIR) if a.endswith(".pkl")]
        print(f"[enc] dir={ROSTROS_DIR} | archivos_pkl={len(archivos)}")
    except Exception as e:
        print(f"[enc] no se pudo listar pkl: {e}")

# ------------------ Apertura robusta de cámara ------------------
def _open_camera(max_index: int = 2):
    try:
        backends = [cv2.CAP_DSHOW, cv2.CAP_MSMF, cv2.CAP_VFW, cv2.CAP_ANY]
    except Exception:
        backends = [cv2.CAP_ANY]

    errores = []
    for i in range(0, max_index + 1):
        for be in backends:
            try:
                cap = cv2.VideoCapture(i, be)
            except TypeError:
                cap = cv2.VideoCapture(i)
            if cap is not None and cap.isOpened():
                print(f"[cam] OK index={i} backend={be}", flush=True)
                try:
                    cap.set(cv2.CAP_PROP_FRAME_WIDTH, 640)
                    cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 480)
                except Exception:
                    pass
                return cap, f"index={i}, backend={be}"
            if cap is not None:
                cap.release()
            errores.append(f"falló index={i}, backend={be}")
    detalle = "; ".join(errores) if errores else "no backends probados"
    print("[cam] No se pudo abrir cámara:", detalle, flush=True)
    log("No se pudo abrir cámara: " + detalle)
    return None, detalle

def _warmup_camera(cap, n=12):
    for _ in range(n):
        cap.read()
        time.sleep(0.02)

def verificar_rostro(rut):
    _debug_listar_pkl()
    expected_encs = _load_encodings_for_rut(rut)
    print(f"[enc] para {rut}: {len(expected_encs)} encodings")
    if not expected_encs:
        log(f"verificar_rostro: no hay encodings para {rut}")
        return False

    cap, info = _open_camera()
    if not cap:
        log("verificar_rostro: cámara no abierta (" + info + ")")
        return False

    print(f"[cam] iniciado verificación; encodings={len(expected_encs)}", flush=True)
    _warmup_camera(cap, 12)

    cv2.namedWindow("Verificación Facial", cv2.WINDOW_NORMAL)
    cv2.resizeWindow("Verificación Facial", 800, 600)

    t0 = time.time()
    duracion = 10.0  # segundos mínimos de captura
    verificado = False
    instrucciones_mostradas = False

    while time.time() - t0 < duracion:
        ret, frame = cap.read()
        if not ret:
            time.sleep(0.02)
            continue

        restante = max(0, int(duracion - (time.time() - t0)))
        cv2.putText(frame, f"Tiempo restante: {restante}s", (40, 30),
                    cv2.FONT_HERSHEY_SIMPLEX, 0.8, (255, 255, 255), 2)

        if not instrucciones_mostradas:
            cv2.putText(frame, "Mire al frente sin moverse", (40, 70),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.9, (0, 255, 255), 2)
            instrucciones_mostradas = True

        rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        try:
            rostros = face_recognition.face_encodings(rgb)
        except Exception as e:
            print("Error en face_encodings:", e, flush=True)
            log(f"face_encodings error: {e}")
            break

        if rostros:
            if len(rostros) > 1:
                cv2.putText(frame, "Por favor, solo 1 persona frente a la camara", (40, 110),
                            cv2.FONT_HERSHEY_SIMPLEX, 0.8, (0, 0, 255), 2)
            cand = rostros[0]
            dists = face_recognition.face_distance(expected_encs, cand)
            best = float(min(dists)) if len(dists) else 1.0
            if best <= FACIAL_TOLERANCE:
                verificado = True

        cv2.imshow("Verificación Facial", frame)
        if (cv2.waitKey(1) & 0xFF == ord('q')) or verificado:
            break

    cap.release()
    cv2.destroyAllWindows()
    print(f"[cam] verificación fin -> {'OK' if verificado else 'FAIL'}", flush=True)
    return verificado

def reconocer_rostro_sin_rut():
    _debug_listar_pkl()
    known_encs, rut_map = _load_all_known_encodings()
    print(f"[enc] total encodings en carpeta: {len(known_encs)}")
    if not known_encs:
        log("reconocer_rostro_sin_rut: no hay encodings en carpeta 'rostros'")
        return None

    import numpy as np
    known_matrix = np.vstack(known_encs)

    cap, info = _open_camera()
    if not cap:
        log("reconocer_rostro_sin_rut: cámara no abierta (" + info + ")")
        return None

    print("[cam] iniciado reconocimiento abierto", flush=True)
    _warmup_camera(cap, 12)

    cv2.namedWindow("Reconocimiento Automático", cv2.WINDOW_NORMAL)
    cv2.resizeWindow("Reconocimiento Automático", 800, 600)

    t0 = time.time()
    duracion = 10.0
    rostro_detectado = None

    while time.time() - t0 < duracion:
        ret, frame = cap.read()
        if not ret:
            time.sleep(0.02)
            continue

        restante = max(0, int(duracion - (time.time() - t0)))
        rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        try:
            rostros_en_vivo = face_recognition.face_encodings(rgb)
        except Exception as e:
            print("Error en face_encodings:", e, flush=True)
            log(f"face_encodings error: {e}")
            break

        if rostros_en_vivo:
            cand = rostros_en_vivo[0]
            dists = face_recognition.face_distance(known_matrix, cand)

            idxs = np.argsort(dists)
            best_idx = int(idxs[0])
            best_dist = float(dists[best_idx])
            second_best = float(dists[idxs[1]]) if len(dists) > 1 else None

            seguro = (best_dist <= FACIAL_TOLERANCE) and \
                     (second_best is None or (second_best - best_dist) >= DISTANCE_MARGIN)

            if seguro:
                rostro_detectado = rut_map[best_idx]

        cv2.imshow("Reconocimiento Automático", frame)
        if (cv2.waitKey(1) & 0xFF == ord('q')) or rostro_detectado:
            break

    cap.release()
    cv2.destroyAllWindows()
    print(f"[cam] reconocimiento fin -> {rostro_detectado}", flush=True)
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
            log(f"Error en verificación: {e}")
            print("Error en verificación:", e)
            callback_error()
    import threading
    threading.Thread(target=proceso, daemon=True).start()

def reconocer_rostro_async(callback_exito, callback_error):
    def proceso():
        try:
            rut_detectado = reconocer_rostro_sin_rut()
            if rut_detectado:
                callback_exito(rut_detectado)
            else:
                callback_error()
        except Exception as e:
            log(f"Error en reconocimiento: {e}")
            print("Error en reconocimiento:", e)
            callback_error()
    import threading
    threading.Thread(target=proceso, daemon=True).start()

# ============== EXTRAS MENSUALES (TABLA Y UTILIDADES) ==============
def _extras_ensure_schema():
    try:
        con = sqlite3.connect(DB_PATH)
        cur = con.cursor()
        cur.execute("""
            CREATE TABLE IF NOT EXISTS extras_mensuales (
                rut TEXT NOT NULL,
                anio_mes TEXT NOT NULL,           -- 'YYYY-MM'
                minutos_extra INTEGER NOT NULL DEFAULT 0,
                updated_at TEXT,
                PRIMARY KEY (rut, anio_mes)
            )
        """)
        con.commit()
        con.close()
    except Exception as e:
        log(f"extras schema error: {e}")

def _extras_sumar_minutos(rut: str, fecha_iso: str, minutos: int):
    """Suma 'minutos' al registro mensual (YYYY-MM). Si no existe, lo crea."""
    try:
        if minutos <= 0:
            return
        anio_mes = fecha_iso[:7]  # 'YYYY-MM'
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        con = sqlite3.connect(DB_PATH)
        cur = con.cursor()
        cur.execute("""
            INSERT INTO extras_mensuales (rut, anio_mes, minutos_extra, updated_at)
            VALUES (?, ?, ?, ?)
            ON CONFLICT(rut, anio_mes) DO UPDATE SET
                minutos_extra = minutos_extra + excluded.minutos_extra,
                updated_at = excluded.updated_at
        """, (rut, anio_mes, int(minutos), now))
        con.commit()
        con.close()
    except Exception as e:
        log(f"extras sumar error: {e}")

# Crear esquema de extras al cargar el módulo
_extras_ensure_schema()

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

        alert_box = ctk.CTkFrame(cont, corner_radius=10, fg_color="#122b39")
        alert_box.pack(fill="x", pady=(0, 10))
        ctk.CTkLabel(
            alert_box, text=mensaje,
            font=("Arial", 14, "bold"),
            text_color="#6FE6FF", justify="left", wraplength=520
        ).pack(padx=10, pady=8, anchor="w")

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
                entry_rut_em.insert(0, _rut_formatear(rut_sugerido))
                entry_pass_em.focus_set()
            except Exception:
                pass

        botones = ctk.CTkFrame(cont, fg_color="transparent")
        botones.pack(fill="x", pady=8)

        def _confirmar():
            rut_i_raw = entry_rut_em.get().strip()
            rut_i = _rut_formatear(rut_i_raw)  # ← formateo automático
            if not rut_i or not entry_pass_em.get().strip():
                tk.messagebox.showwarning("Campo vacío", "Debes ingresar RUT y clave.")
                return
            if not validar_pass_rut(rut_i, entry_pass_em.get().strip()):
                tk.messagebox.showerror("Acceso denegado", "RUT o clave inválidos.")
                return
            # Validación suave del DV
            if not _rut_valido(rut_i):
                tk.messagebox.showwarning("RUT", "El RUT ingresado no supera la validación de DV. Revisa si es correcto.")

            try:
                win.grab_release()
            except Exception:
                pass
            win.destroy()

            entry_rut.delete(0, tk.END)
            entry_rut.insert(0, rut_i)  # ← ya formateado
            label_estado.configure(text="✅ Acceso por clave de emergencia.", text_color="yellow")
            label_hora_registro.configure(text="")
            cargar_info_usuario(rut_i, por_verificacion=True)

        ctk.CTkButton(botones, text="Cancelar", fg_color="gray",
                      command=lambda: win.destroy()).pack(side="left", padx=6)
        ctk.CTkButton(botones, text="Validar", command=_confirmar).pack(side="right", padx=6)

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
        conexion = sqlite3.connect(DB_PATH)
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
                label_estado.configure(text="✅ Verificación OK. Puedes registrar la salida.", text_color="yellow")
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

    # --------- función para registrar (ingreso/salida) ---------
    def registrar(tipo):
        # Normalizar RUT leído del entry para dejar BD consistente
        rut = _rut_formatear(entry_rut.get().strip())

        nombre = label_nombre.cget("text").replace("Nombre: ", "")
        fecha_iso = _hoy_iso()
        hora_actual = datetime.now().strftime('%H:%M:%S')
        hora_actual_dt = parse_hora(hora_actual)

        es_f, nombre_f, _ = es_feriado(datetime.now().date())
        obs_feriado = f"Feriado: {nombre_f}" if es_f else ""

        dia_actual = datetime.now().strftime('%A')
        dias_traducidos = {
            'Monday': 'Lunes', 'Tuesday': 'Martes', 'Wednesday': 'Miércoles',
            'Thursday': 'Jueves', 'Friday': 'Viernes', 'Saturday': 'Sábado', 'Sunday': 'Domingo'
        }
        dia_semana = dias_traducidos.get(dia_actual, '')

        conexion = sqlite3.connect(DB_PATH)
        cursor = conexion.cursor()
        cursor.execute("""
            SELECT hora_entrada, hora_salida FROM horarios
            WHERE rut = ? AND dia = ?
        """, (rut, dia_semana))
        bloques = cursor.fetchall()

        # ---- REGISTRA EN BD (aplica panel/feriado/viernes-flex) ----
        def registrar_final(observacion="", usar_hora_oficial_salida=False):
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
                          "Limpieza automática en 5 seg..." if not es_f else
                          f"Ingreso en feriado registrado ✅ ({nombre_f}).\nLimpieza automática en 60 seg..."),
                    text_color="green"
                )
                boton_ingreso.pack_forget()
                boton_salida.pack_forget()
                actualizar_estado_botones(rut, por_reconocimiento=False)
                label_hora_registro.configure(text=f"⏰ Hora de registro: {hora_actual}", text_color="yellow")
                frame.after(5000 if not es_f else 60000, limpiar_campos)

            elif tipo == "salida":
                # Panel de salida anticipada (guarda hora oficial) → NO suma extras
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

                    # Doble-check salida ya registrada
                    cursor.execute("SELECT hora_salida FROM registros WHERE rut=? AND DATE(fecha)=DATE('now')", (rut,))
                    hs = cursor.fetchone()
                    if hs and hs[0]:
                        label_estado.configure(text="⚠️ Ya registraste una salida hoy.", text_color="orange")
                        conexion.close()
                        return

                    if row:
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

                    # No sumar extras (se guardó hora oficial)
                    label_estado.configure(
                        text=f"Salida anticipada registrada ✅ (hora oficial {hora_oficial}).\nLimpieza automática en 5 seg...",
                        text_color="green"
                    )
                    boton_ingreso.pack_forget()
                    boton_salida.pack_forget()
                    actualizar_estado_botones(rut, por_reconocimiento=False)
                    label_hora_registro.configure(text=f"⏰ Hora de registro: {hora_actual}", text_color="yellow")
                    frame.after(5000, limpiar_campos)
                    return

                # Flujo normal (sin panel)
                cursor.execute("""
                    SELECT hora_salida FROM registros WHERE rut = ? AND DATE(fecha) = DATE('now')
                """, (rut,))
                resultado = cursor.fetchone()
                if resultado and resultado[0]:
                    label_estado.configure(text="⚠️ Ya registraste una salida hoy.", text_color="orange")
                    conexion.close()
                    return

                # ¿Usar hora oficial (regla viernes) o hora actual?
                hora_db = hora_actual
                hora_oficial_usada = None
                if usar_hora_oficial_salida:
                    cursor.execute("""
                        SELECT hora_ingreso, observacion FROM registros 
                        WHERE rut=? AND DATE(fecha)=DATE('now')
                    """, (rut,))
                    row = cursor.fetchone()
                    hora_ingreso_hhmm = row[0] if row else None
                    hora_oficial = _hora_salida_oficial_por_horario(rut, fecha_iso, hora_ingreso_hhmm)
                    hora_db = f"{hora_oficial}:00" if len(hora_oficial) == 5 else hora_oficial
                    hora_oficial_usada = hora_oficial

                # Guardar salida
                if resultado:
                    cursor.execute("""
                        UPDATE registros SET hora_salida = ?, observacion = ? 
                        WHERE rut = ? AND DATE(fecha) = DATE('now')
                    """, (hora_db, observacion, rut))
                    conexion.commit()
                else:
                    cursor.execute("""
                        INSERT INTO registros (rut, nombre, fecha, hora_ingreso, hora_salida, observacion)
                        VALUES (?, ?, ?, ?, ?, ?)
                    """, (rut, nombre, fecha_iso, None, hora_db, observacion))
                    conexion.commit()
                conexion.close()

                # ---- SUMA DE EXTRAS (si corresponde) ----
                if not usar_hora_oficial_salida:
                    try:
                        con2 = sqlite3.connect(DB_PATH)
                        cur2 = con2.cursor()
                        cur2.execute("SELECT hora_ingreso FROM registros WHERE rut=? AND DATE(fecha)=DATE('now')", (rut,))
                        row = cur2.fetchone()
                        con2.close()
                        hora_ingreso_hhmm = row[0] if row and row[0] else None
                        hora_oficial_str = _hora_salida_oficial_por_horario(rut, fecha_iso, hora_ingreso_hhmm)
                        if hora_oficial_str:
                            base = datetime.now().date()
                            t_act = datetime.combine(base, hora_actual_dt.time())
                            t_ofi_time = parse_hora(hora_oficial_str)
                            t_ofi = datetime.combine(base, t_ofi_time.time())
                            exceso_min = int(max(0, (t_act - t_ofi).total_seconds() // 60))
                            if EXTRA_COUNT_ONLY_ABOVE_THRESHOLD:
                                exceso_contable = exceso_min if exceso_min >= EXTRA_MINUTES_THRESHOLD else 0
                            else:
                                exceso_contable = exceso_min
                            if exceso_contable > 0:
                                _extras_sumar_minutos(rut, fecha_iso, exceso_contable)
                    except Exception as e:
                        log(f"calc extras error: {e}")

                # Mensajes finales
                if usar_hora_oficial_salida and hora_oficial_usada:
                    label_estado.configure(
                        text=(f"Salida de viernes (colación) registrada ✅\n"
                              f"Se guarda la hora oficial {hora_oficial_usada}.\n"
                              f"Limpieza automática en 5 seg..."),
                        text_color="green"
                    )
                else:
                    label_estado.configure(
                        text=("Salida registrada correctamente ✅\n"
                              "Limpieza automática en 5 seg..." if not es_f else
                              f"Salida en feriado registrada ✅ ({nombre_f}).\nLimpieza automática en 60 seg..."),
                        text_color="green"
                    )

                boton_ingreso.pack_forget()
                boton_salida.pack_forget()
                actualizar_estado_botones(rut, por_reconocimiento=False)
                label_hora_registro.configure(text=f"⏰ Hora de registro: {hora_actual}", text_color="yellow")
                frame.after(5000 if not es_f else 60000, limpiar_campos)

        # ---------- Reglas para pedir observación ----------
        if es_f:
            registrar_final("")
            return

        requiere_observacion = False
        mensaje_motivo = ""

        if tipo == "ingreso":
            # Observación sólo si llegas con atraso (> 5 min)
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
                        # Hora final pactada del día (fin de jornada)
                        ultima_salida_str = max(bloques_validos, key=lambda h: parse_hora(h))
                        ultima_salida_dt = parse_hora(ultima_salida_str)

                        # --- Regla de viernes (colación) ---
                        if dia_semana == 'Viernes':
                            margen = timedelta(minutes=FRIDAY_FLEX_MINUTES)
                            if ultima_salida_dt - margen <= hora_actual_dt < ultima_salida_dt:
                                registrar_final("", usar_hora_oficial_salida=True)
                                return
                        # ------------------------------------

                        if hora_actual_dt < ultima_salida_dt:
                            # Salida ANTES de la pactada → pedir observación
                            mensaje_motivo = (
                                f"⚠️ Vas a salir antes de la hora pactada ({ultima_salida_dt.strftime('%H:%M')}).\n"
                                f"🕒 Hora actual: {hora_actual_dt.strftime('%H:%M')}.\n"
                                f"✍️ Por favor, indica el motivo:"
                            )
                            requiere_observacion = True
                        else:
                            # Salida DESPUÉS de la pactada: observación solo si supera umbral (40 min)
                            delta_seg = (hora_actual_dt - ultima_salida_dt).total_seconds()
                            delta_min = int(delta_seg // 60)
                            if delta_min >= LATE_AFTER_EXIT_MINUTES:
                                horas = int(delta_min // 60)
                                minutos = int(delta_min % 60)
                                mensaje_motivo = (
                                    f"⚠️ Estás registrando salida {delta_min} minuto(s) "
                                    f"después de la hora pactada ({ultima_salida_dt.strftime('%H:%M')}).\n"
                                    f"🕒 Hora actual: {hora_actual_dt.strftime('%H:%M')}.\n"
                                    f"⏱ Exceso: {horas:02}:{minutos:02}.\n"
                                    f"✍️ Por favor, indica el motivo:"
                                )
                                requiere_observacion = True
                            else:
                                # Entre hora pactada y +40 min → registro directo
                                registrar_final()
                                return

                    except Exception:
                        mensaje_motivo = (
                            f"⚠️ ¡Atención! No se pudo determinar la hora de salida esperada.\n"
                            f"🕒 Hora actual: {hora_actual_dt.strftime('%H:%M')}\n"
                            f"✍️ Por favor, indica el motivo:"
                        )
                        requiere_observacion = True

        if requiere_observacion:
            # Diálogo para anotar motivo
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

            # Encabezado
            alert_box = ctk.CTkFrame(cont, corner_radius=10, fg_color="#3a2e00")
            alert_box.pack(fill="x", pady=(0, 8))
            ctk.CTkLabel(
                alert_box, text=mensaje_motivo, font=("Arial", 13),
                text_color="#FFC857", justify="left", wraplength=560
            ).pack(padx=10, pady=8, anchor="w")

            Textbox = getattr(ctk, "CTkTextbox", None)
            if Textbox:
                entry_obs = Textbox(cont, width=560, height=130)
                entry_obs.pack(fill="both", expand=False)
                def get_text(): return entry_obs.get("1.0", "end").strip()
            else:
                wrap = tk.Frame(cont); wrap.pack(fill="both", expand=False)
                entry_obs = tk.Text(wrap, width=68, height=6, font=("Segoe UI", 11))
                entry_obs.pack(fill="both", expand=False)
                def get_text(): return entry_obs.get("1.0", "end").strip()

            counter = ctk.CTkLabel(cont, text="0/200 caracteres", text_color="gray")
            counter.pack(anchor="e", pady=(4, 0))

            def actualizar_contador(_=None):
                txt = get_text()
                if len(txt) > 200:
                    entry_obs.delete("1.0", "end"); entry_obs.insert("1.0", txt[:200])
                    txt = txt[:200]
                counter.configure(text=f"{len(txt)}/200 caracteres")
            try:
                entry_obs.bind("<KeyRelease>", actualizar_contador)
            except Exception:
                pass

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
            try:
                win.lift()
                win.attributes("-topmost", True)
                win.after(50, lambda: win.attributes("-topmost", False))
            except Exception:
                pass

            win.bind("<Return>", lambda e: confirmar())
            win.bind("<Escape>", lambda e: win.destroy())
            actualizar_contador()

        else:
            registrar_final()

    # ---- Wrapper para forzar uso del RUT canónico en botones ----
    def registrar_con_rut(rut_param, tipo):
        entry_rut.delete(0, tk.END)
        entry_rut.insert(0, _rut_formatear(rut_param))
        registrar(tipo)

    def cargar_info_usuario(rut, por_verificacion=False):
        # ---- Reset de estado para evitar arrastres de mensajes previos ----
        label_estado.configure(text="", text_color="white")

        conexion = sqlite3.connect(DB_PATH)
        cursor = conexion.cursor()

        # Comparar por RUT "limpio" (sin puntos/guion y DV mayúscula)
        rut_fmt = _rut_formatear(rut)
        rut_raw = _rut_limpio(rut_fmt)  # p.ej. 10063936K
        cursor.execute("""
            SELECT nombre, apellido, profesion, cumpleanos, rut
            FROM trabajadores
            WHERE REPLACE(REPLACE(UPPER(rut),'.',''),'-','') = ?
        """, (rut_raw,))
        resultado = cursor.fetchone()

        if resultado:
            nombre_completo = f"{resultado[0]} {resultado[1]}"
            profesion = resultado[2]
            rut_db = _rut_formatear(resultado[4])  # canónico desde BD

            label_nombre.configure(text=f"Nombre: {nombre_completo}")
            label_profesion.configure(text=f"Profesión: {profesion}")
            label_fecha.configure(text=f"Fecha: {datetime.now().strftime('%d/%m/%Y')}")
            label_hora.configure(text=f"Hora: {datetime.now().strftime('%H:%M:%S')}")

            cumpleanos = resultado[3] if resultado[3] else ""
            hoy = datetime.now().strftime('%d/%m')
            if cumpleanos and hoy == cumpleanos[:5]:
                label_estado.configure(
                    text=f"🎂 ¡Hoy es el cumpleaños de {resultado[0]}! 🎉\nPuedes registrar tu ingreso o salida.",
                    text_color="yellow"
                )
            else:
                label_hora_registro.configure(text="")

            # Mostrar/ocultar botones SIEMPRE, usando rut_db
            actualizar_estado_botones(rut_db, por_reconocimiento=por_verificacion)

            # Botones usan el RUT canónico
            boton_ingreso.configure(command=lambda: registrar_con_rut(rut_db, "ingreso"))
            boton_salida.configure(command=lambda: registrar_con_rut(rut_db, "salida"))
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

    # -------- formateo automático del RUT en UI --------
    def _formatear_entry_rut():
        raw = entry_rut.get().strip()
        if not raw:
            return
        fmt = _rut_formatear(raw)
        # Validación suave: no bloquea, solo informa si DV no cuadra
        if len(_rut_limpio(raw)) >= 2 and not _rut_valido(fmt):
            label_estado.configure(text="ℹ️ Revisa el DV del RUT (validación no coincide).", text_color="orange")
        entry_rut.delete(0, tk.END)
        entry_rut.insert(0, fmt)

    def buscar_automatico():
        # Formatear antes de usar
        _formatear_entry_rut()
        rut = entry_rut.get().strip()
        if not rut:
            label_estado.configure(text="🔍 Buscando rostro...", text_color="gray")
            label_hora_registro.configure(text="")
            reconocer_rostro_async(
                callback_exito=lambda rut_detectado: [
                    entry_rut.delete(0, tk.END),
                    entry_rut.insert(0, _rut_formatear(rut_detectado)),
                    cargar_info_usuario(_rut_formatear(rut_detectado), por_verificacion=True)
                ],
                callback_error=lambda: pedir_emergencia(rut_sugerido="", mensaje="❌ No se pudo identificar el rostro. Usa clave de emergencia:")
            )
            return

        label_estado.configure(text="🔄 Verificando rostro...", text_color="gray")
        label_hora_registro.configure(text="")
        frame.update()
        verificar_rostro_async(
            rut,
            callback_exito=lambda: cargar_info_usuario(rut, por_verificacion=True),
            callback_error=lambda: pedir_emergencia(rut_sugerido=rut)
        )

    def limpiar_campos_btn():
        limpiar_campos()

    # ---------- UI ----------
    frame = ctk.CTkFrame(frame_padre)
    frame.pack(fill="both", expand=True)

    ctk.CTkLabel(frame, text="Ingreso / Salida de Funcionarios", font=("Arial", 16)).pack(pady=10)
    ctk.CTkLabel(frame, text="Ingresa el RUT del funcionario:").pack(pady=(10, 2))
    entry_rut = ctk.CTkEntry(frame, placeholder_text="Ej: 12345678-9")
    entry_rut.pack(pady=5)
    entry_rut.bind("<Return>", lambda event: buscar_automatico())
    # Formateo automático al salir del campo
    try:
        entry_rut.bind("<FocusOut>", lambda e: _formatear_entry_rut())
    except Exception:
        pass

    ctk.CTkButton(frame, text="Buscar", command=buscar_automatico).pack(pady=5)
    ctk.CTkButton(frame, text="Limpiar", command=limpiar_campos_btn).pack(pady=5)

    label_nombre = ctk.CTkLabel(frame, text="Nombre: ---", font=("Arial", 14)); label_nombre.pack(pady=(15, 5))
    label_profesion = ctk.CTkLabel(frame, text="Profesión: ---", font=("Arial", 14)); label_profesion.pack(pady=5)
    label_fecha = ctk.CTkLabel(frame, text="Fecha: ---", font=("Arial", 14)); label_fecha.pack(pady=5)
    label_hora = ctk.CTkLabel(frame, text="Hora: ---", font=("Arial", 14)); label_hora.pack(pady=5)

    boton_ingreso = ctk.CTkButton(frame, text="Registrar Ingreso", font=("Arial", 15), height=45, width=220)
    boton_salida  = ctk.CTkButton(frame, text="Registrar Salida",  font=("Arial", 15), height=45, width=220)

    label_estado = ctk.CTkLabel(frame, text="", font=("Arial", 16)); label_estado.pack(pady=10)
    label_hora_registro = ctk.CTkLabel(frame, text="", font=("Arial", 36, "bold"), text_color="yellow")
    label_hora_registro.pack(pady=(25, 35))
