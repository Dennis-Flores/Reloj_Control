import os, pickle, cv2, numpy as np, face_recognition

OUTPUT_DIR = "rostros"
N_TARGET = 8                 # cantidad objetivo de muestras por persona (6–10)
MIN_LAPLACIAN = 120.0
BRIGHT_MIN, BRIGHT_MAX = 60, 190
MIN_FACE_SIZE = 120
DEDUP_DISTANCE = 0.36        # si una nueva muestra es muy similar a una previa, se descarta

def _quality_ok(frame, loc):
    top, right, bottom, left = loc
    w = right - left; h = bottom - top
    if min(w,h) < MIN_FACE_SIZE:
        return False, "Acércate a la cámara"
    roi = cv2.cvtColor(frame[top:bottom, left:right], cv2.COLOR_BGR2GRAY)
    if roi.size == 0:
        return False, "Reencuadra"
    focus = cv2.Laplacian(roi, cv2.CV_64F).var()
    if focus < MIN_LAPLACIAN:
        return False, "Imagen borrosa"
    mean_b = float(roi.mean())
    if not (BRIGHT_MIN <= mean_b <= BRIGHT_MAX):
        return False, "Iluminación deficiente"
    return True, ""

def _load_existing(rut):
    path = os.path.join(OUTPUT_DIR, f"{rut}.pkl")
    if not os.path.exists(path):
        return []
    with open(path, "rb") as f:
        data = pickle.load(f)
    return list(data) if isinstance(data, (list, tuple)) else [data]

def enrolar(rut: str):
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    samples = _load_existing(rut)

    cap = cv2.VideoCapture(0)
    if not cap.isOpened():
        print("❌ No se pudo abrir la cámara")
        return

    print(f"Enrolando RUT {rut}. Objetivo: {N_TARGET} muestras.")
    while len(samples) < N_TARGET:
        ret, frame = cap.read()
        if not ret:
            continue

        rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        locs = face_recognition.face_locations(rgb, model="hog")

        msg = "Mire al frente. Cambie levemente ángulo/luz entre capturas."
        if len(locs) == 0:
            msg = "Ubique su rostro"
        elif len(locs) > 1:
            msg = "Solo 1 rostro por favor"
        else:
            ok, motivo = _quality_ok(frame, locs[0])
            if not ok:
                msg = motivo
            else:
                enc = face_recognition.face_encodings(rgb, known_face_locations=[locs[0]])
                if enc:
                    enc_new = enc[0]
                    # deduplicación (evita casi iguales)
                    if samples:
                        d = face_recognition.face_distance(np.vstack(samples), enc_new).min()
                        if d < DEDUP_DISTANCE:
                            msg = "Muestra muy similar; gire levemente o cambie iluminación"
                        else:
                            samples.append(enc_new); msg = f"Muestra OK ({len(samples)}/{N_TARGET})"
                    else:
                        samples.append(enc_new); msg = f"Muestra OK (1/{N_TARGET})"

        cv2.putText(frame, msg, (20, 40), cv2.FONT_HERSHEY_SIMPLEX, 0.8, (0,255,255), 2)
        cv2.imshow("Enrolamiento", cv2.resize(frame, (800, 600)))
        key = cv2.waitKey(1) & 0xFF
        if key == ord('q'): break

    cap.release(); cv2.destroyAllWindows()

    if samples:
        out = os.path.join(OUTPUT_DIR, f"{rut}.pkl")
        with open(out, "wb") as f:
            pickle.dump(samples, f, protocol=pickle.HIGHEST_PROTOCOL)
        print(f"✅ Guardado {len(samples)} muestras en {out}")
    else:
        print("⚠️ No se guardaron muestras")

if __name__ == "__main__":
    rut = input("RUT a enrolar (ej 12345678-9): ").strip()
    enrolar(rut)
