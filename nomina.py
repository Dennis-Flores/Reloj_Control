import customtkinter as ctk
import sqlite3

expandido_por_rut = {}

def construir_nomina(frame_padre):
    frame = ctk.CTkFrame(frame_padre)
    frame.pack(fill="both", expand=True)

    ctk.CTkLabel(frame, text="Nómina de Funcionarios", font=("Arial", 16)).pack(pady=10)

    contenedor_busqueda = ctk.CTkFrame(frame, fg_color="transparent")
    contenedor_busqueda.pack(pady=5, padx=20, anchor="w")

    icono = ctk.CTkLabel(contenedor_busqueda, text="🔍", font=("Arial", 18))
    icono.pack(side="left", padx=(0, 5))

    entry_buscar = ctk.CTkEntry(contenedor_busqueda, placeholder_text="Buscar por nombre o RUT", width=300)
    entry_buscar.pack(side="left")

    def limpiar_busqueda():
        entry_buscar.delete(0, 'end')
        actualizar_tabla()

    boton_limpiar = ctk.CTkButton(contenedor_busqueda, text="Limpiar", width=40, command=limpiar_busqueda)
    boton_limpiar.pack(side="left", padx=(5, 0))

    tabla_scroll = ctk.CTkScrollableFrame(frame)
    tabla_scroll.pack(fill="both", expand=True, padx=20, pady=10)

    encabezados = ["Nombre", "RUT", "Profesión", "Correo", "Cumpleaños"]
    for col, titulo in enumerate(encabezados):
        ctk.CTkLabel(tabla_scroll, text=titulo, font=("Arial", 13, "bold")).grid(row=0, column=col, padx=10, pady=5)

    def mostrar_u_ocultar_horario(rut, fila):
        global expandido_por_rut
        clave = (rut, fila)

        fila_widgets = [w for w in tabla_scroll.winfo_children() if w.grid_info().get("row") == fila and w.grid_info().get("column") == 0]
        if fila_widgets:
            label = fila_widgets[0]
            texto_actual = label.cget("text")
            nueva_flecha = "🔽" if clave not in expandido_por_rut else "▶"
            label.configure(text=f"{nueva_flecha} {texto_actual[2:]}")

        if clave in expandido_por_rut:
            for widget in expandido_por_rut[clave]:
                widget.destroy()
            del expandido_por_rut[clave]
        else:
            conexion = sqlite3.connect("reloj_control.db")
            cursor = conexion.cursor()
            cursor.execute("SELECT dia, hora_entrada, hora_salida FROM horarios WHERE rut = ? ORDER BY dia, hora_entrada", (rut,))
            horarios = cursor.fetchall()
            conexion.close()

            widgets = []
            dias_orden = ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes']
            bloques = {
                'Mañana': {dia: [] for dia in dias_orden},
                'Tarde': {dia: [] for dia in dias_orden},
                'Nocturno': {dia: [] for dia in dias_orden},
            }
            for dia, entrada, salida in horarios:
                if dia in dias_orden:
                    entrada = entrada or "-"
                    salida = salida or "-"
                    hora_ent = int(entrada.split(':')[0]) if entrada != "-" else 0

                    if hora_ent < 12:
                        bloques['Mañana'][dia].append((entrada, salida))
                    elif hora_ent < 18:
                        bloques['Tarde'][dia].append((entrada, salida))
                    else:
                        bloques['Nocturno'][dia].append((entrada, salida))

            encabezados_bloques = ["☀️ Mañana", "🌇 Tarde", "🌙 Nocturno"]
            fila_base = fila + 1

            for col_bloque, bloque_nombre in enumerate(bloques.keys(), start=1):
                titulo = ctk.CTkLabel(tabla_scroll, text=encabezados_bloques[col_bloque - 1], font=("Arial", 12, "bold"), text_color="gray")
                titulo.grid(row=fila_base, column=col_bloque + 1, padx=5, sticky='w')
                widgets.append(titulo)

            for i, dia in enumerate(dias_orden):
                lbl_dia = ctk.CTkLabel(tabla_scroll, text=dia + ":", text_color="gray")
                lbl_dia.grid(row=fila_base + 1 + i, column=1, sticky='w', padx=10)
                widgets.append(lbl_dia)

                for col_bloque, bloque_nombre in enumerate(bloques.keys(), start=1):
                    horarios_dia = bloques[bloque_nombre][dia]
                    if horarios_dia:
                        texto = "\n".join([f"{e} a {s}" for e, s in horarios_dia])
                    else:
                        texto = "-"
                    lbl = ctk.CTkLabel(tabla_scroll, text=texto, text_color="gray", justify="left")
                    lbl.grid(row=fila_base + 1 + i, column=col_bloque + 1, sticky='w', padx=5)
                    widgets.append(lbl)

            expandido_por_rut[clave] = widgets

            total_filas_insertadas = len(dias_orden) + 2
            for w in tabla_scroll.winfo_children():
                info = w.grid_info()
                fila_actual = info.get("row")
                if fila_actual and fila_actual > fila:
                    w.grid(row=fila_actual + total_filas_insertadas)

    def actualizar_tabla(filtro=""):
        global expandido_por_rut
        expandido_por_rut.clear()
        for widget in tabla_scroll.winfo_children():
            info = widget.grid_info()
            if int(info['row']) > 0:
                widget.destroy()

        conexion = sqlite3.connect("reloj_control.db")
        cursor = conexion.cursor()

        if filtro:
            filtro = f"%{filtro.lower()}%"
            cursor.execute("""
                SELECT nombre, apellido, rut, profesion, correo, cumpleanos 
                FROM trabajadores 
                WHERE lower(nombre || ' ' || apellido || rut) LIKE ?
                ORDER BY apellido
            """, (filtro,))
        else:
            cursor.execute("SELECT nombre, apellido, rut, profesion, correo, cumpleanos FROM trabajadores ORDER BY apellido")

        datos = cursor.fetchall()
        conexion.close()

        if not datos:
            ctk.CTkLabel(tabla_scroll, text="No se encontraron resultados", text_color="orange").grid(row=1, column=0, columnspan=5, pady=10)
        else:
            fila_actual = 1
            for trabajador in datos:
                nombre = f"{trabajador[0]} {trabajador[1]}"
                celdas = [nombre, trabajador[2], trabajador[3], trabajador[4], trabajador[5]]
                for col, dato in enumerate(celdas):
                    icono_estado = "🔽" if (trabajador[2], fila_actual) in expandido_por_rut else "▶"
                    if col == 0:
                        label = ctk.CTkLabel(tabla_scroll, text=f"{icono_estado} {dato or '-'}", anchor='w', justify='left')
                    elif col == 4:
                        label = ctk.CTkLabel(tabla_scroll, text=dato or "-", anchor='w', justify='left')
                    else:
                        label = ctk.CTkLabel(tabla_scroll, text=dato or "-", anchor='center')
                    label.grid(row=fila_actual, column=col, padx=10, pady=2)
                    if col == 0:
                        label.bind("<Double-Button-1>", lambda e, rut=trabajador[2], fila=fila_actual: mostrar_u_ocultar_horario(rut, fila))
                fila_actual += 1

    entry_buscar.bind("<KeyRelease>", lambda event: actualizar_tabla(entry_buscar.get().strip()))
    actualizar_tabla()
