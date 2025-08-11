import customtkinter as ctk
import sqlite3

# mantiene qué RUT está expandido aunque se filtre
expandido_por_rut = {}

def construir_nomina(frame_padre):
    # ---------- ROOT ----------
    frame = ctk.CTkFrame(frame_padre)
    frame.pack(fill="both", expand=True)

    ctk.CTkLabel(frame, text="Nómina de Funcionarios", font=("Arial", 16)).pack(pady=10)

    # ---------- BUSCADOR ----------
    contenedor_busqueda = ctk.CTkFrame(frame, fg_color="transparent")
    contenedor_busqueda.pack(pady=5, padx=20, anchor="w")

    ctk.CTkLabel(contenedor_busqueda, text="🔍", font=("Arial", 18)).pack(side="left", padx=(0, 5))
    entry_buscar = ctk.CTkEntry(contenedor_busqueda, placeholder_text="Buscar por nombre o RUT", width=300)
    entry_buscar.pack(side="left")

    def limpiar_busqueda():
        entry_buscar.delete(0, 'end')
        actualizar_lista()

    ctk.CTkButton(contenedor_busqueda, text="Limpiar", width=60, command=limpiar_busqueda).pack(side="left", padx=(6, 0))

    # ---------- ENCABEZADOS ----------
    hdr = ctk.CTkFrame(frame)
    hdr.pack(fill="x", padx=20, pady=(10, 2))
    cols = ["", "Nombre", "RUT", "Profesión", "Correo", "Cumpleaños"]
    widths = [32, 220, 130, 160, 220, 110]
    for i, (t, w) in enumerate(zip(cols, widths)):
        ctk.CTkLabel(hdr, text=t, font=("Arial", 13, "bold")).pack(side="left", padx=(6 if i else 0, 6))
        hdr.pack_propagate(False)

    # ---------- LISTA SCROLLABLE ----------
    lista = ctk.CTkScrollableFrame(frame)
    lista.pack(fill="both", expand=True, padx=20, pady=(0, 16))

    # ---------- HELPERS DB ----------
    def _fetch_nomina(filtro_txt=""):
        conn = sqlite3.connect("reloj_control.db")
        cur = conn.cursor()
        if filtro_txt:
            f = f"%{filtro_txt.lower()}%"
            cur.execute("""
                SELECT nombre, apellido, rut, profesion, correo, cumpleanos 
                FROM trabajadores 
                WHERE lower(nombre || ' ' || apellido || rut) LIKE ?
                ORDER BY apellido, nombre
            """, (f,))
        else:
            cur.execute("""
                SELECT nombre, apellido, rut, profesion, correo, cumpleanos 
                FROM trabajadores 
                ORDER BY apellido, nombre
            """)
        rows = cur.fetchall()
        conn.close()
        return rows

    def _fetch_horarios(rut):
        conn = sqlite3.connect("reloj_control.db")
        cur = conn.cursor()
        cur.execute("""
            SELECT dia, hora_entrada, hora_salida
            FROM horarios
            WHERE rut = ?
            ORDER BY dia, hora_entrada
        """, (rut,))
        horarios = cur.fetchall()
        conn.close()
        return horarios

    # ---------- BUILDER DE FILA ----------
    def build_row(datos):
        nombre_completo = f"{datos[0]} {datos[1]}"
        rut, profesion, correo, cumple = datos[2], datos[3] or "-", datos[4] or "-", datos[5] or "-"

        # contenedor de un “bloque” de funcionario (header + details)
        row = ctk.CTkFrame(lista, fg_color="transparent")
        row.pack(fill="x", pady=2)

        header = ctk.CTkFrame(row)
        header.pack(fill="x")

        # columna 0: toggle
        is_expanded = expandido_por_rut.get(rut, False)
        btn_toggle = ctk.CTkButton(header, text=("▼" if is_expanded else "▶"), width=28, height=28)
        btn_toggle.pack(side="left", padx=4, pady=3)

        # resto de columnas como labels (usamos pack para layout horizontal)
        def _lbl(parent, text, width):
            l = ctk.CTkLabel(parent, text=text, anchor="w")
            f = ctk.CTkFrame(parent, width=width, height=28, fg_color="transparent")
            f.pack_propagate(False)
            l.pack(in_=f, fill="both", expand=True, padx=6)
            f.pack(side="left")
            return l

        _lbl(header, nombre_completo or "-", 220)
        _lbl(header, rut or "-", 130)
        _lbl(header, profesion, 160)
        _lbl(header, correo, 220)
        _lbl(header, cumple, 110)

        # detalles (accordion)
        details = ctk.CTkFrame(row, fg_color="#1f2630", corner_radius=8)
        details.pack(fill="x", padx=34, pady=(4, 8))
        if not is_expanded:
            details.pack_forget()

        # función que arma/recarga la tabla de horarios
        def populate_details():
            for w in details.winfo_children():
                w.destroy()

            dias_orden = ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes']
            bloques = {'Mañana': {d: [] for d in dias_orden},
                       'Tarde': {d: [] for d in dias_orden},
                       'Nocturno': {d: [] for d in dias_orden}}

            for dia, entrada, salida in _fetch_horarios(rut):
                if dia in dias_orden:
                    e = entrada or "-"
                    s = salida or "-"
                    try:
                        h_ent = int(e.split(":")[0]) if e != "-" else 0
                    except Exception:
                        h_ent = 0
                    if h_ent < 12:
                        bloques['Mañana'][dia].append((e, s))
                    elif h_ent < 18:
                        bloques['Tarde'][dia].append((e, s))
                    else:
                        bloques['Nocturno'][dia].append((e, s))

            # encabezado de tarjetas
            head = ctk.CTkFrame(details, fg_color="transparent")
            head.pack(fill="x", pady=(8, 4))
            ctk.CTkLabel(head, text=f"Horario de {nombre_completo}", font=("Arial", 13, "bold")).pack(side="left", padx=10)

            grid = ctk.CTkFrame(details)
            grid.pack(fill="x", padx=10, pady=(4, 10))

            # títulos de columnas
            for i, titulo in enumerate(["☀️ Mañana", "🌇 Tarde", "🌙 Nocturno"]):
                ctk.CTkLabel(grid, text=titulo, font=("Arial", 12, "bold")).grid(row=0, column=i+1, padx=8, pady=4, sticky="w")

            # filas por día
            for r, dia in enumerate(dias_orden, start=1):
                ctk.CTkLabel(grid, text=dia + ":", text_color="gray").grid(row=r, column=0, padx=8, pady=3, sticky="w")
                for c, bloque in enumerate(["Mañana", "Tarde", "Nocturno"], start=1):
                    tramos = bloques[bloque][dia]
                    texto = "\n".join([f"{e} a {s}" for (e, s) in tramos]) if tramos else "-"
                    ctk.CTkLabel(grid, text=texto, justify="left").grid(row=r, column=c, padx=8, pady=3, sticky="w")

        # toggle handler
        def toggle():
            expanded = expandido_por_rut.get(rut, False)
            if expanded:
                details.pack_forget()
                expandido_por_rut[rut] = False
                btn_toggle.configure(text="▶")
            else:
                populate_details()
                details.pack(fill="x", padx=34, pady=(4, 8))
                expandido_por_rut[rut] = True
                btn_toggle.configure(text="▼")

        btn_toggle.configure(command=toggle)

        # si viene expandido desde un filtro anterior, poblar
        if is_expanded:
            populate_details()

    # ---------- REFRESH LIST ----------
    def actualizar_lista():
        for w in lista.winfo_children():
            w.destroy()
        filtro = (entry_buscar.get() or "").strip()
        for row in _fetch_nomina(filtro):
            build_row(row)

    entry_buscar.bind("<KeyRelease>", lambda e: actualizar_lista())
    actualizar_lista()
