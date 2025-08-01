import customtkinter as ctk
import sqlite3
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import messagebox
from tkcalendar import Calendar

def construir_dia_administrativo(frame_padre):
    def seleccionar_fecha(entry_target):
        top = tk.Toplevel()
        cal = Calendar(top, date_pattern='dd/mm/yyyy', locale='es_CL')
        cal.pack(padx=10, pady=10)
        def poner_fecha():
            entry_target.delete(0, tk.END)
            entry_target.insert(0, cal.get_date())
            top.destroy()
        tk.Button(top, text="Seleccionar", command=poner_fecha).pack(pady=5)

    def guardar_dias_admin():
        rut = entry_rut.get().strip()
        fecha_inicio_str = entry_fecha_desde.get().strip()
        fecha_fin_str = entry_fecha_hasta.get().strip()
        motivo_seleccionado = combo_tipo_permiso.get()
        if motivo_seleccionado == "Otro (especificar)" or motivo_seleccionado == "Cometido de Servicio":
            motivo = entry_motivo_otro.get().strip()
            if motivo_seleccionado == "Cometido de Servicio":
                motivo = f"Cometido de Servicio - {motivo}"
            else:
                motivo = motivo_seleccionado



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

    def mostrar_vista_previa():
        for widget in tabla_preview.winfo_children():
            widget.destroy()

        rut = entry_rut.get().strip()
        if not rut:
            return

        # Mostrar resumen de días administrativos del año actual
        anio_actual = datetime.now().year
        conn = sqlite3.connect("reloj_control.db")
        cur = conn.cursor()
        # Total días administrativos reales (filtro exacto)
        cur.execute("""
            SELECT COUNT(*) FROM dias_libres 
            WHERE rut = ? 
            AND strftime('%Y', fecha) = ? 
            AND motivo = 'Día Administrativo'
        """, (rut, str(anio_actual)))
        total_admin = cur.fetchone()[0]

        MAX_ADMIN_ANUAL = 6
        disponibles = MAX_ADMIN_ANUAL - total_admin

        if total_admin < MAX_ADMIN_ANUAL:
            color = "green"
        elif total_admin == MAX_ADMIN_ANUAL:
            color = "orange"
        else:
            color = "red"

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
            
            

    def limpiar_campos():
        entry_rut.delete(0, tk.END)
        entry_fecha_desde.delete(0, tk.END)
        entry_fecha_hasta.delete(0, tk.END)
        combo_tipo_permiso.set("Día Administrativo")
        entry_motivo_otro.delete(0, tk.END)
        entry_motivo_otro.pack_forget()
        label_resumen_admin.configure(text="")  # limpia contador
        for widget in tabla_preview.winfo_children():
            widget.destroy()

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

    # === INTERFAZ ===
    frame = ctk.CTkFrame(frame_padre)
    frame.pack(fill="both", expand=True, padx=20, pady=20)

    ctk.CTkLabel(frame, text="Asignar, Ver o Editar Días Administrativos, Permisos y Licencias", font=("Arial", 18)).pack(pady=10)

    entry_rut = ctk.CTkEntry(frame, placeholder_text="RUT del Trabajador (Ej: 12345678-9)", width=250)
    entry_rut.pack(pady=5)
    entry_rut.bind("<Return>", lambda event: mostrar_vista_previa())

    label_resumen_admin = ctk.CTkLabel(frame, text="", text_color="green", font=("Arial", 13))
    label_resumen_admin.pack(pady=(0, 5))

    label_resumen_extras = ctk.CTkLabel(frame, text="", text_color="lightblue", font=("Arial", 13))
    label_resumen_extras.pack(pady=(0, 10))


    ctk.CTkButton(frame, text="🔍 Buscar solicitudes del trabajador", command=mostrar_vista_previa).pack(pady=5)

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

    # Lista de tipos de permisos disponibles
    opciones_permiso = [
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
    "Permiso sin Goce de Sueldo (máx 6 meses)": 180  # puedes cambiar si lo harás manual
    }


    combo_tipo_permiso = ctk.CTkComboBox(frame, values=opciones_permiso, width=400)
    combo_tipo_permiso.set("Día Administrativo")
    combo_tipo_permiso.pack(pady=5)

    entry_motivo_otro = ctk.CTkEntry(frame, placeholder_text="Especificar motivo...", width=400)
    entry_motivo_otro.pack(pady=(0, 10))
    entry_motivo_otro.pack_forget()  # oculto por defecto
    label_info_cometido = ctk.CTkLabel(frame, text="", text_color="orange", font=("Arial", 12))
    label_info_cometido.pack()

    def actualizar_visibilidad_entry_otro(choice):
        hoy = datetime.today()

        if choice == "Cometido de Servicio":
            label_info_cometido.configure(
                text="⚠ Este permiso simula una jornada completa. No se requiere hora, solo el motivo."
            )
            entry_motivo_otro.pack(pady=(0, 10))
            entry_fecha_desde.configure(state="normal")
            entry_fecha_hasta.configure(state="normal")
            entry_fecha_desde.delete(0, tk.END)
            entry_fecha_hasta.delete(0, tk.END)

        elif choice == "Otro (especificar)":
            label_info_cometido.configure(text="")
            entry_motivo_otro.pack(pady=(0, 10))
            entry_fecha_desde.configure(state="normal")
            entry_fecha_hasta.configure(state="normal")
            entry_fecha_desde.delete(0, tk.END)
            entry_fecha_hasta.delete(0, tk.END)

        elif choice == "Día Administrativo":
            label_info_cometido.configure(text="")
            entry_motivo_otro.pack_forget()
            entry_fecha_desde.configure(state="normal")
            entry_fecha_hasta.configure(state="normal")
            entry_fecha_desde.delete(0, tk.END)
            entry_fecha_hasta.delete(0, tk.END)

        elif choice in dias_por_permiso:
            label_info_cometido.configure(text="")
            dias = dias_por_permiso[choice]
            fecha_manual = entry_fecha_desde.get().strip()
            try:
                fecha_inicio = datetime.strptime(fecha_manual, "%d/%m/%Y") if fecha_manual else datetime.today()
            except ValueError:
                fecha_inicio = datetime.today()
            fecha_fin = fecha_inicio + timedelta(days=dias - 1)

            entry_fecha_desde.configure(state="normal")
            entry_fecha_hasta.configure(state="normal")
            entry_fecha_desde.delete(0, tk.END)
            entry_fecha_hasta.delete(0, tk.END)
            entry_fecha_desde.insert(0, fecha_inicio.strftime("%d/%m/%Y"))
            entry_fecha_hasta.insert(0, fecha_fin.strftime("%d/%m/%Y"))
            entry_fecha_desde.configure(state="disabled")
            entry_fecha_hasta.configure(state="disabled")
            entry_motivo_otro.pack_forget()

        else:
            label_info_cometido.configure(text="")
            entry_fecha_desde.configure(state="normal")
            entry_fecha_hasta.configure(state="normal")
            entry_fecha_desde.delete(0, tk.END)
            entry_fecha_hasta.delete(0, tk.END)
            entry_motivo_otro.pack_forget()



    combo_tipo_permiso.configure(command=actualizar_visibilidad_entry_otro)


    ctk.CTkButton(frame, text="Guardar Días Administrativos", command=guardar_dias_admin).pack(pady=15)
    ctk.CTkButton(frame, text="🧹 Limpiar Formulario", fg_color="gray", command=limpiar_campos).pack(pady=(0, 10))
    tabla_preview = ctk.CTkScrollableFrame(frame)
    tabla_preview.pack(fill="both", expand=True, pady=10)