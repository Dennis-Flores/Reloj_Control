import customtkinter as ctk
import sqlite3
from datetime import datetime
from tkcalendar import Calendar
import tkinter as tk
import pandas as pd
import os
from tkinter import messagebox
import time
import math


def parse_hora_flexible(hora_str):
    """Devuelve un objeto datetime.time para HH:MM o HH:MM:SS."""
    for fmt in ("%H:%M:%S", "%H:%M"):
        try:
            return datetime.strptime(hora_str.strip(), fmt)
        except ValueError:
            continue
    return None

def construir_reportes(frame_padre):
    registros_por_dia = {}
    modo_exportable = False

    label_estado = None
    label_datos = None
    label_total = None
    label_leyenda_admin = None

    def seleccionar_fecha(entry_target):
        top = tk.Toplevel()
        top.grab_set()
        cal = Calendar(top, date_pattern='dd/mm/yyyy', locale='es_CL')
        cal.pack(padx=10, pady=10)

        def poner_fecha():
            fecha = cal.get_date()
            entry_target.delete(0, 'end')
            entry_target.insert(0, fecha)
            top.destroy()

        ctk.CTkButton(top, text="Seleccionar", command=poner_fecha).pack(pady=5)

    def obtener_horas_administrativo(rut, fecha_str, como_minutos=False):
        try:
            conexion = sqlite3.connect("reloj_control.db")
            cursor = conexion.cursor()

            if isinstance(fecha_str, str):
                fecha_dt = datetime.strptime(fecha_str, "%Y-%m-%d")
            else:
                fecha_dt = fecha_str

            nombre_dia = fecha_dt.strftime("%A")
            dias_map = {
                'Monday': 'lunes', 'Tuesday': 'martes', 'Wednesday': 'miércoles',
                'Thursday': 'jueves', 'Friday': 'viernes', 'Saturday': 'sábado', 'Sunday': 'domingo'
            }
            dia = dias_map.get(nombre_dia, '').lower()

            cursor.execute("SELECT hora_entrada, hora_salida FROM horarios WHERE rut = ? AND LOWER(dia) = ?", (rut, dia))
            filas = cursor.fetchall()
            conexion.close()

            total_minutos = 0

            total_minutos = 0

            for entrada_str, salida_str in filas:
                entrada_str = entrada_str.strip()
                salida_str = salida_str.strip()

                if not entrada_str or not salida_str:
                    continue  # Ignorar registros incompletos

                try:
                    entrada = datetime.strptime(entrada_str, "%H:%M") if len(entrada_str) <= 5 else datetime.strptime(entrada_str, "%H:%M:%S")
                    salida = datetime.strptime(salida_str, "%H:%M") if len(salida_str) <= 5 else datetime.strptime(salida_str, "%H:%M:%S")
                except ValueError as ve:
                    print(f"⚠️ Error en formato hora: {entrada_str} o {salida_str}")
                    continue

                duracion = salida - entrada
                total_minutos += int(duracion.total_seconds() // 60)


            if como_minutos:
                return total_minutos

            h, m = divmod(total_minutos, 60)
            return f"{int(h)}h {int(m)}min"

        except Exception as e:
            print("Error al calcular horas administrativas:", e)
            return "00:00" if not como_minutos else 0



    def agregar_dias_administrativos(registros_por_dia, rut, desde_dt, hasta_dt):
        conexion = sqlite3.connect("reloj_control.db")
        cursor = conexion.cursor()

        desde_str = desde_dt.strftime('%Y-%m-%d')
        hasta_str = hasta_dt.strftime('%Y-%m-%d')

        cursor.execute("""
            SELECT fecha, motivo FROM dias_libres 
            WHERE rut = ? AND fecha BETWEEN ? AND ?
        """, (rut, desde_str, hasta_str))

        dias_admin = cursor.fetchall()
        conexion.close()

        for fecha, motivo in dias_admin:
            try:
                fecha_dt = datetime.strptime(fecha, '%Y-%m-%d').date()
                fecha_str = fecha

                trabajado = obtener_horas_administrativo(rut, fecha_dt)

                if fecha_str not in registros_por_dia:
                    registros_por_dia[fecha_str] = {
                        "ingreso": "--",
                        "salida": "--",
                        "obs_ingreso": motivo,
                        "obs_salida": "",
                        "trabajado": trabajado,
                        "es_admin": True
                    }
                else:
                    registros_por_dia[fecha_str]["obs_ingreso"] = motivo
                    registros_por_dia[fecha_str]["trabajado"] = trabajado
                    registros_por_dia[fecha_str]["es_admin"] = True

            except Exception as e:
                print(f"Error al procesar día administrativo {fecha}: {str(e)}")
                continue

                
        conexion.close()

    
    def exportar_excel():
        nonlocal label_estado, label_leyenda_admin

        def parse_hora(hora_str):
            for fmt in ("%H:%M:%S", "%H:%M"):
                try:
                    return datetime.strptime(hora_str.strip(), fmt)
                except ValueError:
                    continue
            raise ValueError(f"Formato de hora inválido: {hora_str}")

        rut = entry_rut.get().strip()
        if not modo_exportable:
            label_estado.configure(text="⚠️ Ingresa un rango de fechas para exportar", text_color="orange")
            return

        datos = []
        total_minutos = 0
        dias_ok = 0
        dias_incompletos = 0
        dias_administrativos = 0
        permisos_extra = 0
        minutos_atraso_real = 0

        for fecha in sorted(registros_por_dia):
            minutos_atraso_dia = 0
            minutos_atraso_dia_para_excel = 0
            ingreso = registros_por_dia[fecha].get("ingreso", "--")
            salida = registros_por_dia[fecha].get("salida", "--")
            obs_ingreso = registros_por_dia[fecha].get("obs_ingreso", "")
            obs_salida = registros_por_dia[fecha].get("obs_salida", "")
            es_admin = registros_por_dia[fecha].get("es_admin", False)
            obs_texto = obs_ingreso or obs_salida
            motivo = obs_ingreso.strip().lower()

            trabajado_dia = "--:--"
            estado_dia = "⚠️ Incompleto"

            # CALCULAR ATRASO AUNQUE NO HAYA SALIDA
            if ingreso not in (None, "--"):
                try:
                    fecha_dt = datetime.strptime(fecha, "%Y-%m-%d")
                    nombre_dia = fecha_dt.strftime("%A")
                    dias_map = {
                        "Monday": "Lunes", "Tuesday": "Martes", "Wednesday": "Miércoles",
                        "Thursday": "Jueves", "Friday": "Viernes", "Saturday": "Sábado", "Sunday": "Domingo"
                    }
                    nombre_dia_es = dias_map.get(nombre_dia, "")
                    conexion = sqlite3.connect("reloj_control.db")
                    cursor = conexion.cursor()
                    cursor.execute("""
                        SELECT hora_entrada FROM horarios
                        WHERE rut = ? AND dia = ?
                    """, (rut, nombre_dia_es))
                    resultado = cursor.fetchone()
                    conexion.close()

                    if resultado:
                        hora_esperada_entrada = resultado[0]
                        fecha_base = fecha_dt.date()

                        hora_ingreso_parsed = parse_hora(ingreso)
                        hora_esperada_parsed = parse_hora(hora_esperada_entrada)

                        t_ingreso = datetime.combine(fecha_base, hora_ingreso_parsed.time())
                        t_esperado_ingreso = datetime.combine(fecha_base, hora_esperada_parsed.time())

                        delta_ingreso = (t_ingreso - t_esperado_ingreso).total_seconds()
                        ingreso_valido = delta_ingreso <= 300

                        if not ingreso_valido:
                            segundos_excedentes = delta_ingreso - 300
                            minutos_atraso_dia = max(0, math.ceil(segundos_excedentes / 60))
                            minutos_atraso_dia_para_excel = minutos_atraso_dia
                            minutos_atraso_real += minutos_atraso_dia
                            print(f"🚨 Atraso registrado en {fecha}: {minutos_atraso_dia} minuto(s)")
                except Exception as e:
                    print(f"❌ Error al calcular atraso para {fecha}: {e}")

            # CALCULAR HORAS TRABAJADAS SOLO SI HAY INGRESO Y SALIDA
            if ingreso not in (None, "--") and salida not in (None, "--"):
                t1 = parse_hora_flexible(ingreso)
                t2 = parse_hora_flexible(salida)
                if t1 and t2:
                    duracion = (t2 - t1)
                    minutos = int(duracion.total_seconds() // 60)
                    h, m = divmod(minutos, 60)
                    trabajado_dia = f"{int(h):02}:{int(m):02}"
                    total_minutos += minutos
                else:
                    trabajado_dia = "--:--"


            # DÍA OK o INCOMPLETO
            if ingreso not in (None, "--") and salida not in (None, "--"):
                if minutos_atraso_dia_para_excel == 0:
                    estado_dia = "✅ Día Completado Satisfactoriamente"
                    dias_ok += 1
                else:
                    estado_dia = "⚠️ Incompleto"
                    dias_incompletos += 1
            elif ingreso or salida:
                dias_incompletos += 1

            # ADMINISTRATIVO O PERMISO
            if es_admin:
                trabajado_dia = registros_por_dia[fecha].get("trabajado", trabajado_dia)
                minutos = obtener_horas_administrativo(rut, fecha, como_minutos=True)
                total_minutos += minutos

                if "administrativo" in motivo:
                    dias_administrativos += 1
                else:
                    permisos_extra += 1

                if estado_dia == "⚠️ Incompleto" and (ingreso in (None, "--") or salida in (None, "--")):
                    estado_dia = "📌 Permiso Aceptado"

            tipo_dia = "Administrativo" if es_admin else ("Normal" if ingreso and salida else "Incompleto")

            print(f"📦 {fecha} → atraso: {minutos_atraso_dia_para_excel}")
            datos.append([
                fecha, ingreso, salida, trabajado_dia,
                minutos_atraso_dia_para_excel,
                obs_texto, estado_dia, tipo_dia
            ])
            print(f"✅ Excel recibirá {minutos_atraso_dia_para_excel} minutos para {fecha}")

        # TOTALES Y LEYENDA
        h_tot, m_tot = divmod(total_minutos, 60)
        total_trabajado = f"{int(h_tot):02}:{int(m_tot):02}"
        datos.append(["", "", "", "", "", ""])
        datos.append(["", "", "", f"Total trabajado: {total_trabajado}", "", f"✅ Días OK: {dias_ok} / ⚠️ Incompletos: {dias_incompletos}"])
        datos.append(["", "", "", "", f"⏱️ Total Minutos Atraso: {minutos_atraso_real}", "", "", ""])
        datos.append(["", "", "", "", "", f"📌 Días Administrativos usados (Anual): {dias_administrativos} / 6"])
        datos.append(["", "", "", "", "", f"📝 Permisos adicionales usados: {permisos_extra}"])
        datos.append(["", "", "", "", "", ""])
        datos.append(["", "", "", "📌 Notas:", "", ""])
        datos.append(["", "", "", "✅ Día Completado:", "", "Ingreso antes o hasta 5 min tarde + salida igual o posterior al horario."])
        datos.append(["", "", "", "📌 Permiso Aceptado:", "", "Día con permiso aprobado sin registro válido de entrada/salida."])
        datos.append(["", "", "", "⚠️ Incompleto:", "", "Día con errores en el ingreso o salida respecto al horario."])

        if label_leyenda_admin:
            label_leyenda_admin.configure(
                text=f"📌 Días Administrativos usados: {dias_administrativos} / 6\n📝 Permisos adicionales: {permisos_extra}",
                text_color="white"
            )

        columnas = ["Fecha", "Ingreso", "Salida", "Horas Trabajadas Por Día", "Minutos Atrasados del Mes", "Observación", "Estado del Día", "Tipo de Día"]
        df = pd.DataFrame(datos, columns=columnas)

        carpeta_descargas = os.path.join(os.path.expanduser("~"), "Downloads")
        nombre_archivo = f"reporte_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        ruta_completa = os.path.join(carpeta_descargas, nombre_archivo)

        try:
            df.to_excel(ruta_completa, index=False)
            time.sleep(1)
            label_estado.configure(
                text=f"✅ El archivo fue exportado exitosamente.\nSe guardó en Descargas como:\n{nombre_archivo}",
                text_color="green"
            )
            messagebox.showinfo("Exportación exitosa", f"El archivo fue guardado en Descargas:\n{nombre_archivo}")
            os.startfile(ruta_completa)
        except Exception as e:
            label_estado.configure(text=f"❌ Error al exportar: {str(e)}", text_color="red")






    def buscar_reportes():
        nonlocal registros_por_dia, modo_exportable, label_estado, label_datos, label_total, label_leyenda_admin
        rut = entry_rut.get().strip()
        fecha_desde = entry_desde.get().strip()
        fecha_hasta = entry_hasta.get().strip()

        for widget in frame_tabla.winfo_children():
            widget.destroy()

        if not rut:
            label_estado.configure(text="⚠️ Ingresa un RUT", text_color="red")
            return

        conexion = sqlite3.connect("reloj_control.db")
        cursor = conexion.cursor()
        cursor.execute("SELECT nombre, apellido, profesion FROM trabajadores WHERE rut = ?", (rut,))
        info_trabajador = cursor.fetchone()

        if info_trabajador:
            label_datos.configure(
                text=f"RUT: {rut}\nNombre: {info_trabajador[0]} {info_trabajador[1]}\nProfesión: {info_trabajador[2]}"
            )
        else:
            label_datos.configure(text="RUT no encontrado")
            label_estado.configure(text="⚠️ RUT no registrado", text_color="orange")
            return

        # En la función buscar_reportes():
        from datetime import timedelta

        if not fecha_desde or not fecha_hasta:
            hoy = datetime.today()
            desde_dt = hoy.replace(day=1)
            hasta_dt = hoy + timedelta(days=45)
            modo_exportable = True

        else:
            try:
                desde_dt = datetime.strptime(fecha_desde, "%d/%m/%Y")
                hasta_dt = datetime.strptime(fecha_hasta, "%d/%m/%Y")
                modo_exportable = True
            except ValueError:
                label_estado.configure(text="⚠️ Formato de fecha incorrecto", text_color="red")
                return

        cursor.execute("""
            SELECT fecha, hora_ingreso, hora_salida, observacion
            FROM registros
            WHERE rut = ? AND fecha BETWEEN ? AND ?
            ORDER BY fecha
        """, (rut, desde_dt.strftime('%Y-%m-%d'), hasta_dt.strftime('%Y-%m-%d')))

        registros = cursor.fetchall()

        for fecha, hora_ingreso, hora_salida, observacion in registros:
            registros_por_dia[fecha] = {
                "ingreso": hora_ingreso if hora_ingreso else "--",
                "salida": hora_salida if hora_salida else "--",
                "obs_ingreso": observacion if hora_ingreso else "",
                "obs_salida": observacion if hora_salida else "",
                "motivo": ""
            }
        else:
                # Si por alguna razón hay más de un registro por día (no debería), solo actualiza si el campo está vacío
                if not registros_por_dia[fecha]["ingreso"] and hora_ingreso:
                    registros_por_dia[fecha]["ingreso"] = hora_ingreso
                    registros_por_dia[fecha]["obs_ingreso"] = observacion
                if not registros_por_dia[fecha]["salida"] and hora_salida:
                    registros_por_dia[fecha]["salida"] = hora_salida
                    registros_por_dia[fecha]["obs_salida"] = observacion    


        agregar_dias_administrativos(registros_por_dia, rut, desde_dt, hasta_dt)
        conexion.close()

        # Revisar si hay días administrativos en el futuro
        hoy = datetime.today().date()
        dias_admin_futuros = [
            f for f in registros_por_dia
            if registros_por_dia[f].get("es_admin", False) and datetime.strptime(f, "%Y-%m-%d").date() > hoy
        ]

        if dias_admin_futuros:
            label_estado.configure(
                text=f"✅ Reporte generado\n📌 Incluye {len(dias_admin_futuros)} día(s) administrativo(s) futuro(s)",
                text_color="green"
            )
        else:
            label_estado.configure(text="✅ Reporte generado", text_color="green")


        if registros_por_dia:
            encabezados = ["Fecha", "Ingreso", "Salida", "Trabajado", "Obs. Ingreso", "Obs. Salida"]
            for col, texto in enumerate(encabezados):
                ctk.CTkLabel(frame_tabla, text=texto, font=("Arial", 13, "bold")).grid(row=0, column=col, padx=10, pady=4)

            fila = 1
            total_minutos = 0
            for fecha in sorted(registros_por_dia):
                ingreso = registros_por_dia[fecha].get("ingreso", "--")
                salida = registros_por_dia[fecha].get("salida", "--")
                obs_ingreso = registros_por_dia[fecha].get("obs_ingreso", "")
                obs_salida = registros_por_dia[fecha].get("obs_salida", "")
                es_admin = registros_por_dia[fecha].get("es_admin", False)

                # Valor por defecto:
                trabajado = registros_por_dia[fecha].get("trabajado", "Incompleto")

                # Si ambos existen y no es admin, calcula las horas:
                if not es_admin and ingreso not in (None, "--") and salida not in (None, "--"):
                    t1 = parse_hora_flexible(ingreso)
                    t2 = parse_hora_flexible(salida)
                    if t1 and t2:
                        duracion = (t2 - t1)
                        minutos = duracion.total_seconds() // 60
                        h, m = divmod(minutos, 60)
                        trabajado = f"{int(h)}h {int(m)}min"
                        registros_por_dia[fecha]["trabajado"] = trabajado
                        total_minutos += minutos
                    else:
                        trabajado = "Incompleto"
                        registros_por_dia[fecha]["trabajado"] = trabajado

                elif ingreso in (None, "--") or salida in (None, "--"):
                    trabajado = "Incompleto"
                    registros_por_dia[fecha]["trabajado"] = trabajado

                fecha_legible = datetime.strptime(fecha, "%Y-%m-%d").strftime("%d/%m/%Y")
                fila_datos = [
                    fecha_legible,
                    ingreso or "--",
                    salida or "--",
                    trabajado,
                    obs_ingreso or "",
                    obs_salida or ""
                ]

                for col, val in enumerate(fila_datos):
                    color = None
                    if es_admin:
                        color = "#00bfff"  # celeste
                    elif col == 3 and trabajado != "Incompleto":
                        color = "green"
                    elif col == 3 and trabajado == "Incompleto":
                        color = "orange"

                    if color:
                        ctk.CTkLabel(frame_tabla, text=val, text_color=color).grid(row=fila, column=col, padx=10, pady=2)
                    else:
                        ctk.CTkLabel(frame_tabla, text=val).grid(row=fila, column=col, padx=10, pady=2)
                fila += 1

            h_tot, m_tot = divmod(total_minutos, 60)
            label_total.configure(text=f"Total trabajado: {int(h_tot)}h {int(m_tot)}min", text_color="white")
            from datetime import timedelta

            # Calcular inicio y fin de la semana actual
            hoy = datetime.today().date()
            inicio_semana = hoy - timedelta(days=hoy.weekday())  # lunes
            fin_semana = inicio_semana + timedelta(days=6)       # domingo

            # Total por semana
            minutos_semana = 0
            for fecha in registros_por_dia:
                fecha_dt = datetime.strptime(fecha, "%Y-%m-%d").date()
                if inicio_semana <= fecha_dt <= fin_semana:
                    trabajado = registros_por_dia[fecha].get("trabajado", "0h 0min")
                    if "h" in trabajado:
                        partes = trabajado.replace("min", "").split("h")
                        try:
                            horas = int(partes[0].strip())
                            mins = int(partes[1].strip())
                            minutos_semana += horas * 60 + mins
                        except:
                            pass

            h_sem, m_sem = divmod(minutos_semana, 60)
            ctk.CTkLabel(frame, text=f"Total trabajado esta semana: {h_sem}h {m_sem}min", font=("Arial", 12)).pack(pady=2)

            label_estado.configure(text="✅ Reporte generado", text_color="green")
        else:
            label_total.configure(text="")
            ctk.CTkLabel(frame_tabla, text="No hay registros en este rango", text_color="orange").grid(row=0, column=0, columnspan=6, pady=10)
            label_estado.configure(text="⚠️ Sin registros", text_color="orange")

    def exportar_mes_completo():
        try:
            conexion = sqlite3.connect("reloj_control.db")
            cursor = conexion.cursor()

            # Obtener todos los trabajadores
            cursor.execute("SELECT rut, nombre, apellido FROM trabajadores")
            trabajadores = cursor.fetchall()

            hoy = datetime.today()
            desde_dt = hoy.replace(day=1)
            hasta_dt = hoy

            datos_globales = []

            for rut, nombre, apellido in trabajadores:
                registros_por_dia = {}

                # Obtener registros normales
                cursor.execute("""
                SELECT fecha, hora, tipo, observacion FROM registros
                WHERE rut = ? AND fecha BETWEEN ? AND ?
                ORDER BY fecha, hora
            """, (rut, desde_dt.strftime('%Y-%m-%d'), hasta_dt.strftime('%Y-%m-%d')))
            registros = cursor.fetchall()

            for fecha, hora, tipo, obs in registros:
                if fecha not in registros_por_dia:
                    registros_por_dia[fecha] = {"ingreso": None, "salida": None, "motivo": ""}
                registros_por_dia[fecha][tipo] = hora
                if tipo == "ingreso":
                    registros_por_dia[fecha]["obs_ingreso"] = obs
                elif tipo == "salida":
                    registros_por_dia[fecha]["obs_salida"] = obs


                # Agregar días administrativos con motivo
                cursor.execute("""
                    SELECT fecha, motivo FROM dias_libres
                    WHERE rut = ? AND fecha BETWEEN ? AND ?
                """, (rut, desde_dt.strftime('%Y-%m-%d'), hasta_dt.strftime('%Y-%m-%d')))
                dias_admin = cursor.fetchall()

                for fecha, motivo in dias_admin:
                    if fecha not in registros_por_dia:
                        registros_por_dia[fecha] = {
                            "ingreso": "--", "salida": "--", "motivo": motivo, "obs_ingreso": motivo, "obs_salida": ""
                        }
                    else:
                        registros_por_dia[fecha]["motivo"] = motivo
                        registros_por_dia[fecha]["obs_ingreso"] = motivo

                for fecha in sorted(registros_por_dia):
                    ingreso = registros_por_dia[fecha].get("ingreso", "--")
                    salida = registros_por_dia[fecha].get("salida", "--")
                    motivo = registros_por_dia[fecha].get("motivo", "")
                    obs_ingreso = registros_por_dia[fecha].get("obs_ingreso", "")
                    obs_salida = registros_por_dia[fecha].get("obs_salida", "")

                    # Calcular horas trabajadas
                    if ingreso not in (None, "--") and salida not in (None, "--"):
                        try:
                            t1 = datetime.strptime(ingreso, "%H:%M:%S") if len(ingreso) > 5 else datetime.strptime(ingreso, "%H:%M")
                            t2 = datetime.strptime(salida, "%H:%M:%S") if len(salida) > 5 else datetime.strptime(salida, "%H:%M")
                            duracion = t2 - t1
                            minutos = int(duracion.total_seconds() // 60)
                            h, m = divmod(minutos, 60)
                            trabajado = f"{h}h {m}min"
                        except:
                            trabajado = "Error"
                    elif motivo:
                        # Si es administrativo, calcular según horario del contrato
                        trabajado = obtener_horas_administrativo(rut, fecha)
                    else:
                        trabajado = "Incompleto"

                    datos_globales.append([
                        rut, nombre, apellido, fecha,
                        ingreso or "--", salida or "--",
                        trabajado, obs_ingreso or "", obs_salida or "", motivo or ""
                    ])

            conexion.close()

            # Crear DataFrame
            columnas = [
                "RUT", "Nombre", "Apellido", "Fecha",
                "Ingreso", "Salida", "Horas Trabajadas",
                "Obs. Ingreso", "Obs. Salida", "Motivo"
            ]
            df = pd.DataFrame(datos_globales, columns=columnas)

            # Guardar archivo
            carpeta_descargas = os.path.join(os.path.expanduser("~"), "Downloads")
            nombre_archivo = f"reporte_mensual_completo_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            ruta_completa = os.path.join(carpeta_descargas, nombre_archivo)
            df.to_excel(ruta_completa, index=False)

            messagebox.showinfo("Exportación completa", f"Archivo guardado en:\n{ruta_completa}")
            os.startfile(ruta_completa)

        except Exception as e:
            messagebox.showerror("Error al exportar", str(e))



    # === INTERFAZ ===
    frame = ctk.CTkFrame(frame_padre)
    frame.pack(fill="both", expand=True)

    ctk.CTkLabel(frame, text="Reportes por Funcionario", font=("Arial", 16)).pack(pady=10)
        # --- Campo de búsqueda por nombre ---
    lista_funcionarios = []
    dict_nombre_rut = {}

    try:
        conexion = sqlite3.connect("reloj_control.db")
        cursor = conexion.cursor()
        cursor.execute("SELECT rut, nombre, apellido FROM trabajadores")
        for rut, nombre, apellido in cursor.fetchall():
            nombre_completo = f"{nombre} {apellido}"
            lista_funcionarios.append(nombre_completo)
            dict_nombre_rut[nombre_completo] = rut
        conexion.close()
    except Exception as e:
        print(f"Error al cargar funcionarios: {e}")

    # === FUNCIONARIO BUSCADOR COMBO ===

    combo_var = tk.StringVar()
    lista_nombres = []
    dict_nombre_rut = {}

    # Cargar nombres desde la base de datos
    try:
        conexion = sqlite3.connect("reloj_control.db")
        cursor = conexion.cursor()
        cursor.execute("SELECT rut, nombre, apellido FROM trabajadores")
        for rut, nombre, apellido in cursor.fetchall():
            nombre_completo = f"{nombre} {apellido}"
            lista_nombres.append(nombre_completo)
            dict_nombre_rut[nombre_completo] = rut
        conexion.close()
    except Exception as e:
        print(f"Error al cargar funcionarios: {e}")

    # Función para actualizar sugerencias
    def actualizar_opciones(valor_actual):
        filtro = valor_actual.lower()
        coincidencias = [n for n in lista_nombres if filtro in n.lower()]
        if not coincidencias:
            coincidencias = ["No hay coincidencias"]
        combo_funcionarios.configure(values=coincidencias)
        combo_funcionarios.set(valor_actual)

        # Autolimpia el campo RUT si el texto no coincide exactamente
        if valor_actual not in dict_nombre_rut:
            entry_rut.delete(0, "end")

    # Al seleccionar opción
    def seleccionar_funcionario(valor):
        if valor == "No hay coincidencias":
            return
        rut = dict_nombre_rut.get(valor, "")
        entry_rut.delete(0, "end")
        entry_rut.insert(0, rut)

    # Crear un subframe horizontal para el ComboBox y el botón Limpiar
    combo_frame = ctk.CTkFrame(frame, fg_color="transparent")
    combo_frame.pack(pady=5)

    def al_hacer_click(event):
        if combo_funcionarios.get() == "Buscar Usuario por nombre":
            combo_funcionarios.set("")

    # Luego crear el ComboBox
    combo_funcionarios = ctk.CTkComboBox(
        combo_frame,
        values=["Buscar Usuario por nombre"],
        width=400,
        height=35,
        font=("Arial", 14),
        corner_radius=8
    )
    combo_funcionarios.set("Buscar Usuario por nombre")
    combo_funcionarios.pack(side="left", padx=(0, 10))

    # Finalmente conectar el evento
    combo_funcionarios.bind("<FocusIn>", al_hacer_click)
    

    # Función del botón Limpiar
    def limpiar_combobox():
        combo_funcionarios.set("Buscar Usuario por nombre")
        combo_funcionarios.configure(values=lista_nombres)
        entry_rut.delete(0, 'end')

    # Botón Limpiar
    boton_limpiar = ctk.CTkButton(
        combo_frame,
        text="Limpiar",
        width=80,
        height=35,
        command=limpiar_combobox
    )
    boton_limpiar.pack(side="left")

    def buscar_por_enter(event):
        texto = combo_funcionarios.get().lower()
        if texto and texto != "buscar usuario por nombre":
            resultados = [nombre for nombre in lista_nombres if texto in nombre.lower()]
            if resultados:
                combo_funcionarios.configure(values=resultados)
                combo_funcionarios.set(resultados[0])  # Opcional: seleccionar primero
                seleccionar_funcionario(resultados[0])  # Simular acción
            else:
                combo_funcionarios.configure(values=["Sin resultados"])
                combo_funcionarios.set("Sin resultados")

    combo_funcionarios.bind("<Return>", buscar_por_enter)


    # 🔍 Detectar escritura en tiempo real
    combo_funcionarios.bind("<KeyRelease>", lambda event: actualizar_opciones(combo_funcionarios.get()))

    # ⌨️ Acción extra al presionar Enter
    combo_funcionarios.bind("<Return>", lambda event: actualizar_opciones(combo_funcionarios.get()))




    entry_rut = ctk.CTkEntry(frame, placeholder_text="RUT (Ej: 12345678-9)")
    entry_rut.pack(pady=5)
    entry_rut.bind("<Return>", lambda event: buscar_reportes())

    cont_fecha_desde = ctk.CTkFrame(frame, fg_color="transparent")
    cont_fecha_desde.pack(pady=2)
    entry_desde = ctk.CTkEntry(cont_fecha_desde, placeholder_text="Desde (dd/mm/aaaa)", width=200)
    entry_desde.pack(side="left")
    ctk.CTkButton(cont_fecha_desde, text="📅", width=40, command=lambda: seleccionar_fecha(entry_desde)).pack(side="left", padx=5)

    cont_fecha_hasta = ctk.CTkFrame(frame, fg_color="transparent")
    cont_fecha_hasta.pack(pady=2)
    entry_hasta = ctk.CTkEntry(cont_fecha_hasta, placeholder_text="Hasta (dd/mm/aaaa)", width=200)
    entry_hasta.pack(side="left")
    ctk.CTkButton(cont_fecha_hasta, text="📅", width=40, command=lambda: seleccionar_fecha(entry_hasta)).pack(side="left", padx=5)

    ctk.CTkButton(frame, text="Buscar Reporte", command=buscar_reportes).pack(pady=5)
    ctk.CTkButton(frame, text="Exportar Excel", command=exportar_excel).pack(pady=5)
    ctk.CTkButton(frame, text="Exportar Mes Completo De Todos Los Usuarios", command=exportar_mes_completo).pack(pady=5)

    label_datos = ctk.CTkLabel(frame, text="RUT: ---\nNombre: ---\nProfesión: ---", font=("Arial", 13), justify="left")
    label_datos.pack(pady=10)

    frame_tabla = ctk.CTkScrollableFrame(frame)
    frame_tabla.pack(pady=10, fill="both", expand=True)

    label_total = ctk.CTkLabel(frame, text="", font=("Arial", 12))
    label_total.pack(pady=5)

    label_leyenda_admin = ctk.CTkLabel(frame, text="", font=("Arial", 12))
    label_leyenda_admin.pack(pady=2)

    label_estado = ctk.CTkLabel(frame, text="")
    label_estado.pack(pady=10)
