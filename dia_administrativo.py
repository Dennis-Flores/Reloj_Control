import customtkinter as ctk
import sqlite3
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import messagebox
from tkcalendar import Calendar

def construir_dia_administrativo(frame_padre):
    # --- FUNCIÓN PARA SELECCIÓN DE FECHA ---
    def seleccionar_fecha(entry_target):
        top = tk.Toplevel()
        cal = Calendar(top, date_pattern='dd/mm/yyyy', locale='es_CL')
        cal.pack(padx=10, pady=10)
        def poner_fecha():
            entry_target.delete(0, tk.END)
            entry_target.insert(0, cal.get_date())
            top.destroy()
        tk.Button(top, text="Seleccionar", command=poner_fecha).pack(pady=5)

    # --- CONFIGURACIÓN DE PERMISOS Y DÍAS ---
    opciones_permiso = [
        "Elija tipo de Solicitud o Permiso",
        "Día Administrativo",
        "Permiso por Matrimonio/Acuerdo Unión Civil (5 días hábiles)",
        "Permiso por Defunción de Hijo (10 días corridos)",
        "Permiso por Defunción de Cónyuge/Conviviente Civil (7 días corridos)",
        "Permiso por Defunción de Hijo en Gestación (7 días hábiles)",
        "Permiso por Defunción de Padre/Madre/Hermano(a) (4 días hábiles)",
        "Permiso de Nacimiento Paternal (5 días corridos)",
        "Permiso de Alimentación (1 hora diaria)",
        "Permiso sin Goce de Sueldo (máx 6 meses)",
        "Cometido de Servicio",
        "Otro (especificar)"
    ]
    dias_por_permiso = {
        "Permiso por Matrimonio/Acuerdo Unión Civil (5 días hábiles)": 5,
        "Permiso por Defunción de Hijo (10 días corridos)": 10,
        "Permiso por Defunción de Cónyuge/Conviviente Civil (7 días corridos)": 7,
        "Permiso por Defunción de Hijo en Gestación (7 días hábiles)": 7,
        "Permiso por Defunción de Padre/Madre/Hermano(a) (4 días hábiles)": 4,
        "Permiso de Nacimiento Paternal (5 días corridos)": 5,
        "Permiso sin Goce de Sueldo (máx 6 meses)": 180
    }

    # --- FRAME PRINCIPAL ---
    frame = ctk.CTkFrame(frame_padre)
    frame.pack(fill="both", expand=True, padx=20, pady=20)

    ctk.CTkLabel(frame, text="Asignar, Ver o Editar Días Administrativos, Permisos y Licencias", font=("Arial", 18)).pack(pady=10)

    # --- CARGA NOMBRES/RUTS Y AUTOCOMPLETADO ---
    def cargar_nombres_ruts():
        conn = sqlite3.connect("reloj_control.db")
        cursor = conn.cursor()
        cursor.execute("SELECT rut, nombre, apellido FROM trabajadores")
        nombres, dict_nombre_rut = [], {}
        for rut, nombre, apellido in cursor.fetchall():
            nombre_completo = f"{nombre} {apellido}".strip()
            nombres.append(nombre_completo)
            dict_nombre_rut[nombre_completo] = rut
        conn.close()
        return nombres, dict_nombre_rut

    lista_nombres, dict_nombre_rut = cargar_nombres_ruts()
    primeros_10 = lista_nombres[:10]

    # --- COMBOBOX Y BOTÓN DE BÚSQUEDA ---
    fila_nombre = ctk.CTkFrame(frame, fg_color="transparent")
    fila_nombre.pack(pady=10)
    combo_nombre = ctk.CTkComboBox(fila_nombre, values=primeros_10, width=250)
    combo_nombre.set("Buscar por Nombre")
    combo_nombre.pack(side="left", padx=(0, 5))
    btn_buscar = ctk.CTkButton(fila_nombre, text="🔍 Buscar", command=lambda: buscar_solicitudes())
    btn_buscar.pack(side="left", padx=(5, 0))

    # --- ENTRY DE RUT ---
    entry_rut = ctk.CTkEntry(frame, placeholder_text="RUT: (Ej: 12345678-9)", width=150)
    entry_rut.pack(pady=5)
    entry_rut.bind("<Return>", lambda event: mostrar_vista_previa())

    # --- LABELS PARA RESUMEN ---
    label_resumen_admin = ctk.CTkLabel(frame, text="", text_color="green", font=("Arial", 13))
    label_resumen_admin.pack(pady=(0, 5))
    label_resumen_extras = ctk.CTkLabel(frame, text="", text_color="red", font=("Arial", 13))
    label_resumen_extras.pack(pady=(0, 10))

    # --- SUBTITULO SECCIÓN DE NUEVA SOLICITUD ---
    ctk.CTkLabel(
        frame, 
        text="IMPORTANTE",
        font=("Arial", 13, "bold"),
        text_color="#d32f2f"
    ).pack(pady=(12, 2))
    ctk.CTkLabel(
        frame, 
        text="Nueva Solicitud de Permiso o Licencia   ▼",
        font=("Arial", 15, "bold"),
        text_color="#607D8B",
        anchor="center"
    ).pack(pady=(0, 8))

    # --- PERMISO (COMBOBOX) ---
    combo_tipo_permiso = ctk.CTkComboBox(frame, values=opciones_permiso, width=400)
    combo_tipo_permiso.set("Elija tipo de Solicitud o Permiso")
    combo_tipo_permiso.pack(pady=5)

    # --- FECHAS DESDE/HASTA ---
    cont_desde = ctk.CTkFrame(frame, fg_color="transparent")
    cont_desde.pack(pady=2)
    entry_fecha_desde = ctk.CTkEntry(cont_desde, placeholder_text="Fecha Desde (dd/mm/aaaa)", width=200)
    entry_fecha_desde.pack(side="left")
    ctk.CTkButton(cont_desde, text="📅", width=40, command=lambda: seleccionar_fecha(entry_fecha_desde)).pack(side="left", padx=5)

    cont_hasta = ctk.CTkFrame(frame, fg_color="transparent")
    cont_hasta.pack(pady=2)
    entry_fecha_hasta = ctk.CTkEntry(cont_hasta, placeholder_text="Fecha Hasta (dd/mm/aaaa)", width=200)
    entry_fecha_hasta.pack(side="left")
    ctk.CTkButton(cont_hasta, text="📅", width=40, command=lambda: seleccionar_fecha(entry_fecha_hasta)).pack(side="left", padx=5)

    # --- MOTIVO EXTRA Y LABEL INFO ---
    entry_motivo_otro = ctk.CTkEntry(frame, placeholder_text="Especificar motivo...", width=400)
    entry_motivo_otro.pack(pady=(0, 10))
    entry_motivo_otro.pack_forget()
    label_info_cometido = ctk.CTkLabel(frame, text="", text_color="orange", font=("Arial", 12))
    label_info_cometido.pack()

    # --- TABLA PREVIEW ---
    tabla_preview = ctk.CTkScrollableFrame(frame)
    tabla_preview.pack(fill="both", expand=True, pady=10)

    # ========== FUNCIONES DE AUTOCOMPLETADO Y BÚSQUEDA ==========
    def limpiar_placeholder(event):
        if combo_nombre.get() == "Buscar por Nombre":
            combo_nombre.set("")
    def restaurar_placeholder(event):
        if combo_nombre.get() == "":
            combo_nombre.set("Buscar por Nombre")
    combo_nombre.bind("<FocusIn>", limpiar_placeholder)
    combo_nombre.bind("<FocusOut>", restaurar_placeholder)

    def mostrar_sugerencias(event):
        texto = combo_nombre.get().lower()
        if not texto or texto == "buscar por nombre":
            combo_nombre.configure(values=primeros_10)
    combo_nombre.bind("<FocusIn>", mostrar_sugerencias)

    def autocompletar_nombres(event):
        texto = combo_nombre.get().lower()
        if not texto or texto == "buscar por nombre":
            combo_nombre.configure(values=primeros_10)
        else:
            filtrados = [n for n in lista_nombres if texto in n.lower()]
            combo_nombre.configure(values=filtrados[:10] if filtrados else ["No encontrado"])
    combo_nombre.bind("<KeyRelease>", autocompletar_nombres)

    def buscar_por_nombre(event=None):
        nombre = combo_nombre.get()
        rut = dict_nombre_rut.get(nombre, "")
        if rut:
            entry_rut.delete(0, tk.END)
            entry_rut.insert(0, rut)
            entry_fecha_desde.delete(0, tk.END)
            entry_fecha_hasta.delete(0, tk.END)
            mostrar_vista_previa()
        else:
            messagebox.showwarning("Nombre no válido", "Selecciona un nombre válido.")
    combo_nombre.bind("<Return>", buscar_por_nombre)
    combo_nombre.bind("<<ComboboxSelected>>", buscar_por_nombre)

    def buscar_solicitudes():
        nombre = combo_nombre.get()
        rut_combo = dict_nombre_rut.get(nombre, "")
        rut_entry = entry_rut.get().strip()
        if not rut_entry and rut_combo:
            entry_rut.delete(0, tk.END)
            entry_rut.insert(0, rut_combo)
        mostrar_vista_previa()

    def sumar_dias_habiles(fecha_inicio, dias_habiles):
        dias_agregados = 0
        fecha = fecha_inicio
        while dias_agregados < dias_habiles:
            if fecha.weekday() < 5:  # 0 = lunes, 4 = viernes
                dias_agregados += 1
                if dias_agregados == dias_habiles:
                    break
            fecha += timedelta(days=1)
        return fecha


    # ========== FUNCIÓN PARA AUTO-CÁLCULO DE FECHAS ==========
    def actualizar_fecha_hasta_auto(event=None):
        tipo = combo_tipo_permiso.get()
        dias = dias_por_permiso.get(tipo)
        fecha_inicio_str = entry_fecha_desde.get().strip()
        if dias and fecha_inicio_str:
            try:
                fecha_inicio = datetime.strptime(fecha_inicio_str, "%d/%m/%Y")
                # Permisos con días hábiles (nombre contiene 'días hábiles')
                if "días hábiles" in tipo.lower():
                    fecha_fin = sumar_dias_habiles(fecha_inicio, dias)
                else:
                    fecha_fin = fecha_inicio + timedelta(days=dias - 1)
                entry_fecha_hasta.configure(state="normal")
                entry_fecha_hasta.delete(0, tk.END)
                entry_fecha_hasta.insert(0, fecha_fin.strftime("%d/%m/%Y"))
                entry_fecha_hasta.configure(state="disabled")
            except ValueError:
                entry_fecha_hasta.configure(state="normal")
                entry_fecha_hasta.delete(0, tk.END)
        else:
            entry_fecha_hasta.configure(state="normal")
            entry_fecha_hasta.delete(0, tk.END)

    entry_fecha_desde.bind("<FocusOut>", actualizar_fecha_hasta_auto)
    entry_fecha_desde.bind("<Return>", actualizar_fecha_hasta_auto)

    def guardar_dias_admin():
        rut = entry_rut.get().strip()
        fecha_inicio_str = entry_fecha_desde.get().strip()
        fecha_fin_str = entry_fecha_hasta.get().strip()
        motivo_seleccionado = combo_tipo_permiso.get()
        if motivo_seleccionado == "Elija tipo de Solicitud o Permiso":
            messagebox.showerror("Error", "Debes elegir un tipo de solicitud o permiso.")
            return
        if motivo_seleccionado == "Otro (especificar)":
            motivo = entry_motivo_otro.get().strip() or motivo_seleccionado
        elif motivo_seleccionado == "Cometido de Servicio":
            detalle = entry_motivo_otro.get().strip()
            motivo = f"Cometido de Servicio - {detalle}" if detalle else motivo_seleccionado
        else:
            motivo = motivo_seleccionado
        if motivo_seleccionado == "Día Administrativo" and fecha_inicio_str and not fecha_fin_str:
            fecha_fin_str = fecha_inicio_str
        try:
            fecha_inicio = datetime.strptime(fecha_inicio_str, "%d/%m/%Y")
            fecha_fin = datetime.strptime(fecha_fin_str, "%d/%m/%Y")
        except ValueError:
            messagebox.showerror("Error", "Fechas inválidas. Usa formato dd/mm/aaaa.")
            return
        if not rut or not fecha_inicio_str or not fecha_fin_str:
            messagebox.showerror("Error", "Debes ingresar todos los campos.")
            return
        if fecha_fin < fecha_inicio:
            messagebox.showerror("Error", "La fecha final no puede ser anterior a la inicial.")
            return
        conn = sqlite3.connect("reloj_control.db")
        cur = conn.cursor()
        dias_registrados = 0
        fecha_actual = fecha_inicio
        while fecha_actual <= fecha_fin:
            fecha_str = fecha_actual.strftime("%Y-%m-%d")
            cur.execute("SELECT COUNT(*) FROM dias_libres WHERE rut = ? AND fecha = ?", (rut, fecha_str))
            if cur.fetchone()[0] == 0:
                cur.execute("INSERT INTO dias_libres (rut, fecha, motivo) VALUES (?, ?, ?)", (rut, fecha_str, motivo))
                dias_registrados += 1
            fecha_actual += timedelta(days=1)
        conn.commit()
        conn.close()
        messagebox.showinfo("Éxito", f"{dias_registrados} día(s) guardado(s) correctamente.")
        mostrar_vista_previa()

    def limpiar_campos():
        entry_rut.delete(0, tk.END)
        entry_fecha_desde.delete(0, tk.END)
        entry_fecha_hasta.delete(0, tk.END)
        combo_tipo_permiso.set("Elija tipo de Solicitud o Permiso")
        entry_motivo_otro.delete(0, tk.END)
        entry_motivo_otro.pack_forget()
        combo_nombre.set("Buscar por Nombre")
        limpiar_placeholder(None)
        label_resumen_admin.configure(text="")
        for widget in tabla_preview.winfo_children():
            widget.destroy()

    # --- BOTONES GUARDAR Y LIMPIAR ---
    botones_frame = ctk.CTkFrame(frame, fg_color="transparent")
    botones_frame.pack(pady=(10, 10))  # Espaciado entre fechas y botones

    ctk.CTkButton(botones_frame, text="Guardar Nueva Solicitud", command=guardar_dias_admin, width=180).pack(side="left", padx=10)
    ctk.CTkButton(botones_frame, text="🧹 Limpiar Formulario", fg_color="gray", command=limpiar_campos, width=160).pack(side="left", padx=10)


    # ========== FUNCIÓN PARA CAMBIOS DE PERMISO ==========
    def actualizar_visibilidad_entry_otro(choice):
        entry_fecha_desde.configure(state="normal")
        entry_fecha_hasta.configure(state="normal")

        if choice == "Cometido de Servicio":
            label_info_cometido.configure(
                text="⚠ Este permiso simula una jornada completa. No se requiere hora, solo el motivo."
            )
            entry_motivo_otro.pack(pady=(0, 10))
            entry_fecha_desde.delete(0, tk.END)
            entry_fecha_hasta.delete(0, tk.END)
        elif choice == "Otro (especificar)":
            label_info_cometido.configure(text="")
            entry_motivo_otro.pack(pady=(0, 10))
            entry_fecha_desde.delete(0, tk.END)
            entry_fecha_hasta.delete(0, tk.END)
        elif choice == "Día Administrativo":
            label_info_cometido.configure(text="")
            entry_motivo_otro.pack_forget()
            entry_fecha_desde.delete(0, tk.END)
            entry_fecha_hasta.delete(0, tk.END)
        elif choice in dias_por_permiso:
            label_info_cometido.configure(text="")
            entry_motivo_otro.pack_forget()
            # No borres fechas aquí
        else:
            label_info_cometido.configure(text="")
            entry_motivo_otro.pack_forget()
            entry_fecha_desde.delete(0, tk.END)
            entry_fecha_hasta.delete(0, tk.END)

        # SIEMPRE
        actualizar_fecha_hasta_auto()
    

    # ========== FUNCIONES PRINCIPALES ==========
    def mostrar_vista_previa():
        for widget in tabla_preview.winfo_children():
            widget.destroy()
        rut = entry_rut.get().strip()
        if not rut:
            return
        anio_actual = datetime.now().year
        conn = sqlite3.connect("reloj_control.db")
        cur = conn.cursor()
        cur.execute("""
            SELECT COUNT(*) FROM dias_libres 
            WHERE rut = ? 
            AND strftime('%Y', fecha) = ? 
            AND motivo = 'Día Administrativo'
        """, (rut, str(anio_actual)))
        total_admin = cur.fetchone()[0]
        MAX_ADMIN_ANUAL = 6
        color = "green" if total_admin < MAX_ADMIN_ANUAL else ("orange" if total_admin == MAX_ADMIN_ANUAL else "red")
        texto = f"📅 Días administrativos solicitados este año: {total_admin} / {MAX_ADMIN_ANUAL} disponibles"
        label_resumen_admin.configure(text=texto, text_color=color)
        cur.execute("""
            SELECT COUNT(*) FROM dias_libres 
            WHERE rut = ? 
            AND strftime('%Y', fecha) = ? 
            AND motivo != 'Día Administrativo'
        """, (rut, str(anio_actual)))
        total_extras = cur.fetchone()[0]
        texto_extras = f"🧾 Permisos extras este año: {total_extras}"
        label_resumen_extras.configure(text=texto_extras)
        cur.execute("SELECT id, fecha, motivo FROM dias_libres WHERE rut = ? ORDER BY fecha", (rut,))
        registros = cur.fetchall()
        conn.close()
        headers = ["Fecha", "Motivo", "Guardar", "Eliminar"]
        for i, header in enumerate(headers):
            ctk.CTkLabel(tabla_preview, text=header, font=("Arial", 13, "bold")).grid(row=0, column=i, padx=10, pady=5)
        for idx, (id_, fecha, motivo) in enumerate(registros, start=1):
            fecha_legible = datetime.strptime(fecha, "%Y-%m-%d").strftime("%d/%m/%Y")
            ctk.CTkLabel(tabla_preview, text=fecha_legible).grid(row=idx, column=0, padx=10, pady=2)
            entry_motivo_edit = ctk.CTkEntry(tabla_preview, width=400, font=("Arial", 14))
            entry_motivo_edit.insert(0, motivo or "")
            entry_motivo_edit.grid(row=idx, column=1, padx=10, pady=2)
            ctk.CTkButton(tabla_preview, text="💾", width=30, fg_color="green",
                          command=lambda i=id_, e=entry_motivo_edit: actualizar_motivo(i, e)).grid(row=idx, column=2, padx=5)
            ctk.CTkButton(tabla_preview, text="❌", width=30, fg_color="red",
                          command=lambda i=id_: eliminar_dia_admin(i)).grid(row=idx, column=3, padx=5)

    

    def actualizar_motivo(id_, entry_widget):
        nuevo_motivo = entry_widget.get().strip()
        conn = sqlite3.connect("reloj_control.db")
        cur = conn.cursor()
        cur.execute("UPDATE dias_libres SET motivo = ? WHERE id = ?", (nuevo_motivo, id_))
        conn.commit()
        conn.close()
        messagebox.showinfo("Actualizado", "Motivo actualizado correctamente.")

    def eliminar_dia_admin(id_):
        confirm = messagebox.askyesno("Confirmar", "¿Deseas eliminar este día administrativo?")
        if not confirm:
            return
        conn = sqlite3.connect("reloj_control.db")
        cur = conn.cursor()
        cur.execute("DELETE FROM dias_libres WHERE id = ?", (id_,))
        conn.commit()
        conn.close()
        mostrar_vista_previa()
        messagebox.showinfo("Eliminado", "Día administrativo eliminado correctamente.")

    