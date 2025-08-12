import customtkinter as ctk
import sqlite3
from datetime import datetime, timedelta
from tkcalendar import Calendar
import tkinter as tk
import pandas as pd
import os
from tkinter import messagebox
import time
import math


def parse_hora_flexible(hora_str):
    """Devuelve un objeto datetime.datetime (solo tiempo relevante) para HH:MM o HH:MM:SS."""
    if not hora_str:
        return None
    for fmt in ("%H:%M:%S", "%H:%M"):
        try:
            return datetime.strptime(hora_str.strip(), fmt)
        except ValueError:
            continue
    return None


def obtener_horario_del_dia(rut, fecha_dt):
    """Devuelve (hora_entrada, hora_salida) del horario pactado para ese día (string HH:MM[:SS])."""
    if isinstance(fecha_dt, str):
        fecha_dt = datetime.strptime(fecha_dt, "%Y-%m-%d")

    dias_en = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
    dias_es = ['lunes', 'martes', 'miercoles', 'jueves', 'viernes', 'sabado', 'domingo']
    dia_norm = dias_es[dias_en.index(fecha_dt.strftime("%A"))]  # 'miercoles' sin tilde

    con = sqlite3.connect("reloj_control.db")
    cur = con.cursor()
    cur.execute("""
        SELECT hora_entrada, hora_salida
        FROM horarios
        WHERE rut = ?
          AND lower(replace(replace(replace(replace(replace(dia,'á','a'),'é','e'),'í','i'),'ó','o'),'ú','u')) = ?
        ORDER BY time(hora_entrada)
    """, (rut, dia_norm))
    filas = cur.fetchall()
    con.close()

    if not filas:
        return (None, None)

    he = filas[0][0].strip() if filas[0][0] else None
    hs = filas[-1][1].strip() if filas[-1][1] else None
    return (he, hs)


def calcular_minutos_pactados_semana(rut, colacion_min_por_dia=None):
    """
    Suma los minutos pactados en la semana para un RUT según la tabla 'horarios'.
    - Admite múltiples bloques por día.
    - Normaliza tildes de 'dia'.
    - Si colacion_min_por_dia es un entero (minutos), se suma por cada día con jornada (>0).
    """
    con = sqlite3.connect("reloj_control.db")
    cur = con.cursor()
    cur.execute("""
        SELECT lower(replace(replace(replace(replace(replace(dia,'á','a'),'é','e'),'í','i'),'ó','o'),'ú','u')) AS d,
               hora_entrada, hora_salida
        FROM horarios
        WHERE rut = ?
    """, (rut,))
    filas = cur.fetchall()
    con.close()

    minutos_por_dia = {'lunes':0,'martes':0,'miercoles':0,'jueves':0,'viernes':0,'sabado':0,'domingo':0}

    for d, he, hs in filas:
        t1 = parse_hora_flexible(he)
        t2 = parse_hora_flexible(hs)
        if t1 and t2:
            mins = int((t2 - t1).total_seconds() // 60)
            if mins > 0:
                minutos_por_dia[d] = minutos_por_dia.get(d, 0) + mins

    minutos_total = sum(minutos_por_dia.values())

    if isinstance(colacion_min_por_dia, int) and colacion_min_por_dia > 0:
        dias_con_jornada = sum(1 for v in minutos_por_dia.values() if v > 0)
        minutos_total += colacion_min_por_dia * dias_con_jornada

    return minutos_total


def construir_reportes(frame_padre):
    registros_por_dia = {}
    modo_exportable = False

    label_estado = None
    label_datos = None
    label_total = None
    label_leyenda_admin = None
    label_total_semana = None
    label_pactadas = None
    label_completadas = None

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

    def obtener_horas_administrativo(rut, fecha_str_o_dt, como_minutos=False):
        """Suma de (salida-entrada) en horarios pactados del día de la semana; normaliza tildes."""
        try:
            if isinstance(fecha_str_o_dt, str):
                fecha_dt = datetime.strptime(fecha_str_o_dt, "%Y-%m-%d")
            else:
                fecha_dt = fecha_str_o_dt

            nombre_dia = fecha_dt.strftime("%A")
            dias_map = {
                'Monday': 'lunes', 'Tuesday': 'martes', 'Wednesday': 'miercoles',
                'Thursday': 'jueves', 'Friday': 'viernes', 'Saturday': 'sabado', 'Sunday': 'domingo'
            }
            dia = dias_map.get(nombre_dia, '').lower()
            dia_norm = (dia
                        .replace('á', 'a').replace('é', 'e')
                        .replace('í', 'i').replace('ó', 'o').replace('ú', 'u'))

            conexion = sqlite3.connect("reloj_control.db")
            cursor = conexion.cursor()
            cursor.execute("""
                SELECT hora_entrada, hora_salida FROM horarios
                WHERE rut = ?
                  AND lower(replace(replace(replace(replace(replace(dia,'á','a'),'é','e'),'í','i'),'ó','o'),'ú','u')) = ?
            """, (rut, dia_norm))
            filas = cursor.fetchall()
            conexion.close()

            total_minutos = 0
            for entrada_str, salida_str in filas:
                if not entrada_str or not salida_str:
                    continue
                try:
                    entrada = parse_hora_flexible(entrada_str)
                    salida = parse_hora_flexible(salida_str)
                    if not entrada or not salida:
                        continue
                    duracion = salida - entrada
                    minutos = int(duracion.total_seconds() // 60)
                    total_minutos += int(minutos)
                except Exception:
                    continue

            if como_minutos:
                return int(total_minutos)

            h, m = divmod(int(total_minutos), 60)
            return f"{int(h)}h {int(m)}min"

        except Exception as e:
            print("Error al calcular horas administrativas:", e)
            return 0 if como_minutos else "00:00"

    def agregar_dias_administrativos(regs_por_dia, rut, desde_dt, hasta_dt):
        """Mezcla en regs_por_dia los registros de dias_libres con motivo; calcula trabajado según horario."""
        conexion = sqlite3.connect("reloj_control.db")
        cursor = conexion.cursor()

        desde_str = desde_dt.strftime('%Y-%m-%d')
        hasta_str = hasta_dt.strftime('%Y-%m-%d')

        cursor.execute("""
            SELECT fecha, motivo FROM dias_libres 
            WHERE rut = ? AND fecha BETWEEN ? AND ?
        """, (rut, desde_str, hasta_str))

        dias_admin = cursor.fetchall()
        conexion.close()  # cierre único

        for fecha, motivo in dias_admin:
            try:
                fecha_dt = datetime.strptime(fecha, '%Y-%m-%d').date()
                trabajado = obtener_horas_administrativo(rut, fecha_dt)
                if fecha not in regs_por_dia:
                    regs_por_dia[fecha] = {
                        "ingreso": "--",
                        "salida": "--",
                        "obs_ingreso": motivo,
                        "obs_salida": "",
                        "trabajado": trabajado,
                        "es_admin": True
                    }
                else:
                    regs_por_dia[fecha]["obs_ingreso"] = motivo
                    regs_por_dia[fecha]["trabajado"] = trabajado
                    regs_por_dia[fecha]["es_admin"] = True
            except Exception as e:
                print(f"Error al procesar día administrativo {fecha}: {str(e)}")
                continue

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
            motivo = (obs_ingreso or "").strip().lower()

            trabajado_dia = "--:--"
            estado_dia = "⚠️ Incompleto"

            # ATRASO (aunque no haya salida)
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
                        ingreso_valido = delta_ingreso <= 300  # 5 min

                        if not ingreso_valido:
                            segundos_excedentes = delta_ingreso - 300
                            minutos_atraso_dia = max(0, math.ceil(segundos_excedentes / 60))
                            minutos_atraso_dia_para_excel = minutos_atraso_dia
                            minutos_atraso_real += minutos_atraso_dia
                except Exception as e:
                    print(f"❌ Error al calcular atraso para {fecha}: {e}")

            # HORAS TRABAJADAS (solo si hay ingreso y salida)
            if ingreso not in (None, "--") and salida not in (None, "--"):
                t1 = parse_hora_flexible(ingreso)
                t2 = parse_hora_flexible(salida)
                if t1 and t2:
                    duracion = (t2 - t1)
                    minutos = int(duracion.total_seconds() // 60)
                    h, m = divmod(minutos, 60)
                    trabajado_dia = f"{int(h):02}:{int(m):02}"
                    total_minutos += int(minutos)
                else:
                    trabajado_dia = "--:--"

            # ESTADO DEL DÍA
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
                minutos_admin = obtener_horas_administrativo(rut, fecha, como_minutos=True)
                total_minutos += int(minutos_admin)

                if "administrativo" in motivo:
                    dias_administrativos += 1
                else:
                    permisos_extra += 1

                if estado_dia == "⚠️ Incompleto" and (ingreso in (None, "--") or salida in (None, "--")):
                    estado_dia = "📌 Permiso Aceptado"

            tipo_dia = "Administrativo" if es_admin else ("Normal" if ingreso and salida else "Incompleto")

            datos.append([
                fecha, ingreso, salida, trabajado_dia,
                minutos_atraso_dia_para_excel,
                obs_texto, estado_dia, tipo_dia
            ])

        # TOTALES Y LEYENDA
        h_tot, m_tot = divmod(int(total_minutos), 60)
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

        # === RESUMEN: PACTADAS vs COMPLETADAS (SEMANA ACTUAL) ===
        # 1) Minutos efectivos completados esta semana, desde registros_por_dia
        hoy = datetime.today().date()
        inicio_semana = hoy - timedelta(days=hoy.weekday())  # Lunes
        fin_semana = inicio_semana + timedelta(days=6)       # Domingo
        minutos_semana = 0
        for f, info in registros_por_dia.items():
            fecha_dt = datetime.strptime(f, "%Y-%m-%d").date()
            if inicio_semana <= fecha_dt <= fin_semana:
                trabajado_str = info.get("trabajado", "")
                if isinstance(trabajado_str, str) and "h" in trabajado_str:
                    try:
                        hh, mm = trabajado_str.replace("min", "").split("h")
                        minutos_semana += int(hh.strip()) * 60 + int(mm.strip())
                    except Exception:
                        pass
                elif isinstance(trabajado_str, str) and ":" in trabajado_str:
                    try:
                        hh, mm = trabajado_str.split(":")
                        minutos_semana += int(hh) * 60 + int(mm)
                    except Exception:
                        pass

        # 2) Colación por día (si existe en trabajadores, úsala; si no, 30 min)
        colacion_min_por_dia = 30
        try:
            con = sqlite3.connect("reloj_control.db")
            cur = con.cursor()
            cur.execute("SELECT colacion_min FROM trabajadores WHERE rut = ? LIMIT 1", (rut,))
            r = cur.fetchone()
            con.close()
            if r and r[0] is not None:
                colacion_min_por_dia = int(r[0])
        except Exception:
            pass

        # 3) Pactadas dinámicas desde horarios + colación
        minutos_pactados_semana = calcular_minutos_pactados_semana(rut, colacion_min_por_dia=colacion_min_por_dia)

        h_sem, m_sem = divmod(int(minutos_semana), 60)
        h_pac, m_pac = divmod(int(minutos_pactados_semana), 60)

        datos.append(["", "", "", "", "", ""])
        datos.append(["", "", "", "Resumen semanal (semana actual):", "", ""])
        datos.append(["", "", "", f"Horas semanales pactadas según horario: {h_pac}h {m_pac:02d}min", "", ""])
        datos.append(["", "", "", f"Horas completadas esta semana (efectivas): {h_sem}h {m_sem:02d}min", "", ""])

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
            try:
                os.startfile(ruta_completa)  # Windows
            except Exception:
                pass
        except Exception as e:
            label_estado.configure(text=f"❌ Error al exportar: {str(e)}", text_color="red")

    def buscar_reportes():
        nonlocal registros_por_dia, modo_exportable, label_estado, label_datos, label_total, label_leyenda_admin, label_total_semana, label_pactadas, label_completadas
        rut = entry_rut.get().strip()
        fecha_desde = entry_desde.get().strip()
        fecha_hasta = entry_hasta.get().strip()

        # Limpiar la tabla visual
        for widget in frame_tabla.winfo_children():
            widget.destroy()
        # Reiniciar contenedor de datos para no acumular
        registros_por_dia = {}

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
            conexion.close()
            return

        # Rango de fechas (por defecto: desde el 1 del mes actual hasta hoy +45d)
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
                conexion.close()
                return

        cursor.execute("""
            SELECT fecha, hora_ingreso, hora_salida, observacion
            FROM registros
            WHERE rut = ? AND fecha BETWEEN ? AND ?
            ORDER BY fecha
        """, (rut, desde_dt.strftime('%Y-%m-%d'), hasta_dt.strftime('%Y-%m-%d')))
        registros = cursor.fetchall()
        conexion.close()

        for fecha, hora_ingreso, hora_salida, observacion in registros:
            if fecha not in registros_por_dia:
                registros_por_dia[fecha] = {
                    "ingreso": "--", "salida": "--",
                    "obs_ingreso": "", "obs_salida": "",
                    "motivo": ""
                }
            if hora_ingreso:
                registros_por_dia[fecha]["ingreso"] = hora_ingreso
                registros_por_dia[fecha]["obs_ingreso"] = observacion or registros_por_dia[fecha]["obs_ingreso"]
            if hora_salida:
                registros_por_dia[fecha]["salida"] = hora_salida
                registros_por_dia[fecha]["obs_salida"] = observacion or registros_por_dia[fecha]["obs_salida"]

        # Mezclar días administrativos
        agregar_dias_administrativos(registros_por_dia, rut, desde_dt, hasta_dt)

        # Regla: "Cometido" → autocompletar salida con horario pactado
        for fecha, info in registros_por_dia.items():
            obs_txt = (info.get("obs_ingreso") or "") + " " + (info.get("obs_salida") or "") + " " + (info.get("motivo") or "")
            obs_txt = obs_txt.lower()

            if "cometid" in obs_txt and (not info.get("salida") or info.get("salida") == "--"):
                try:
                    fecha_dt = datetime.strptime(fecha, "%Y-%m-%d")
                except Exception:
                    continue

                _, salida_horario = obtener_horario_del_dia(rut, fecha_dt)
                if not salida_horario:
                    continue

                info["salida"] = salida_horario
                t1 = parse_hora_flexible(info.get("ingreso")) if info.get("ingreso") and info.get("ingreso") != "--" else None
                t2 = parse_hora_flexible(salida_horario)
                if t1 and t2:
                    mins = int((t2 - t1).total_seconds() // 60)
                    h, m = divmod(mins, 60)
                    info["trabajado"] = f"{h}h {m}min"
                else:
                    info["trabajado"] = obtener_horas_administrativo(rut, fecha_dt)

                info["es_admin"] = True
                if not info.get("obs_salida"):
                    info["obs_salida"] = "Salida autocompletada por Cometido"

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

        # Render de tabla y totales
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

                # Si no es admin y hay ingreso+salida válidos, calcula:
                if not es_admin and ingreso not in (None, "--") and salida not in (None, "--"):
                    t1 = parse_hora_flexible(ingreso)
                    t2 = parse_hora_flexible(salida)
                    if t1 and t2:
                        duracion = (t2 - t1)
                        minutos = int(duracion.total_seconds() // 60)
                        h, m = divmod(minutos, 60)
                        trabajado = f"{int(h)}h {int(m)}min"
                        registros_por_dia[fecha]["trabajado"] = trabajado
                        total_minutos += int(minutos)
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

            # Total acumulado del rango buscado (para el label del mes/rango)
            h_tot, m_tot = divmod(int(total_minutos), 60)
            label_total.configure(text=f"Total trabajado en el mes: {int(h_tot)}h {int(m_tot)}min", text_color="white")

            # === TOTAL SEMANAL EFECTIVO ===
            hoy = datetime.today().date()
            inicio_semana = hoy - timedelta(days=hoy.weekday())  # Lunes
            fin_semana = inicio_semana + timedelta(days=6)       # Domingo

            minutos_semana = 0
            for f, info in registros_por_dia.items():
                fecha_dt = datetime.strptime(f, "%Y-%m-%d").date()
                if inicio_semana <= fecha_dt <= fin_semana:
                    trabajado_str = info.get("trabajado", "")
                    if isinstance(trabajado_str, str) and "h" in trabajado_str:
                        try:
                            hh, mm = trabajado_str.replace("min", "").split("h")
                            minutos_semana += int(hh.strip()) * 60 + int(mm.strip())
                        except Exception:
                            pass
                    elif isinstance(trabajado_str, str) and ":" in trabajado_str:
                        try:
                            hh, mm = trabajado_str.split(":")
                            minutos_semana += int(hh) * 60 + int(mm)
                        except Exception:
                            pass

            h_sem, m_sem = divmod(int(minutos_semana), 60)
            label_total_semana.configure(text=f"Total trabajado en la semana (efectivo): {h_sem}h {m_sem}min")

            # === PACTADAS vs COMPLETADAS (dinámico por funcionario) ===
            # Colación por día (si existe en trabajadores, úsala; si no, 30 min)
            colacion_min_por_dia = 30
            try:
                con = sqlite3.connect("reloj_control.db")
                cur = con.cursor()
                cur.execute("SELECT colacion_min FROM trabajadores WHERE rut = ? LIMIT 1", (rut,))
                r = cur.fetchone()
                con.close()
                if r and r[0] is not None:
                    colacion_min_por_dia = int(r[0])
            except Exception:
                pass

            minutos_pactados_semana = calcular_minutos_pactados_semana(rut, colacion_min_por_dia=colacion_min_por_dia)
            h_pac, m_pac = divmod(int(minutos_pactados_semana), 60)

            # Mostrar pactadas y completadas (colores según cumplimiento)
            cumple = minutos_semana >= minutos_pactados_semana
            label_pactadas.configure(text=f"Horas semanales pactadas según horario: {h_pac}h {m_pac:02d}min")
            label_completadas.configure(
                text=f"Horas completadas esta semana: {h_sem}h {m_sem:02d}min",
                text_color="green" if cumple else "orange"
            )

            label_estado.configure(text="✅ Reporte generado", text_color="green")
        else:
            label_total.configure(text="")
            ctk.CTkLabel(frame_tabla, text="No hay registros en este rango", text_color="orange").grid(row=0, column=0, columnspan=6, pady=10)
            label_estado.configure(text="⚠️ Sin registros", text_color="orange")

    def exportar_mes_completo():
        """Exporta un Excel con el mes actual (1..hoy) para TODOS los usuarios."""
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

            for rut_t, nombre, apellido in trabajadores:
                regs = {}

                # Registros normales
                cursor.execute("""
                    SELECT fecha, hora_ingreso, hora_salida, observacion
                    FROM registros
                    WHERE rut = ? AND fecha BETWEEN ? AND ?
                    ORDER BY fecha
                """, (rut_t, desde_dt.strftime('%Y-%m-%d'), hasta_dt.strftime('%Y-%m-%d')))
                registros = cursor.fetchall()

                for fecha, hora_ingreso, hora_salida, obs in registros:
                    if fecha not in regs:
                        regs[fecha] = {"ingreso": "--", "salida": "--", "obs_ingreso": "", "obs_salida": "", "motivo": ""}
                    if hora_ingreso:
                        regs[fecha]["ingreso"] = hora_ingreso
                        regs[fecha]["obs_ingreso"] = obs or regs[fecha]["obs_ingreso"]
                    if hora_salida:
                        regs[fecha]["salida"] = hora_salida
                        regs[fecha]["obs_salida"] = obs or regs[fecha]["obs_salida"]

                # Días administrativos
                cursor.execute("""
                    SELECT fecha, motivo FROM dias_libres
                    WHERE rut = ? AND fecha BETWEEN ? AND ?
                """, (rut_t, desde_dt.strftime('%Y-%m-%d'), hasta_dt.strftime('%Y-%m-%d')))
                dias_admin = cursor.fetchall()

                for fecha, motivo in dias_admin:
                    if fecha not in regs:
                        regs[fecha] = {
                            "ingreso": "--", "salida": "--",
                            "motivo": motivo, "obs_ingreso": motivo, "obs_salida": "", "es_admin": True
                        }
                    else:
                        regs[fecha]["motivo"] = motivo
                        regs[fecha]["obs_ingreso"] = motivo
                        regs[fecha]["es_admin"] = True

                # Calcular trabajado por día
                for fecha in sorted(regs):
                    ingreso = regs[fecha].get("ingreso", "--")
                    salida = regs[fecha].get("salida", "--")
                    motivo = regs[fecha].get("motivo", "")
                    obs_ingreso = regs[fecha].get("obs_ingreso", "")
                    obs_salida = regs[fecha].get("obs_salida", "")
                    es_admin = regs[fecha].get("es_admin", False)

                    if ingreso not in (None, "--") and salida not in (None, "--") and not es_admin:
                        try:
                            t1 = parse_hora_flexible(ingreso)
                            t2 = parse_hora_flexible(salida)
                            if t1 and t2:
                                minutos = int((t2 - t1).total_seconds() // 60)
                                h, m = divmod(int(minutos), 60)
                                trabajado = f"{h}h {m}min"
                            else:
                                trabajado = "Incompleto"
                        except Exception:
                            trabajado = "Error"
                    elif motivo:
                        # Si es administrativo, calcular según horario del contrato
                        trabajado = obtener_horas_administrativo(rut_t, fecha)
                    else:
                        trabajado = "Incompleto"

                    datos_globales.append([
                        rut_t, nombre, apellido, fecha,
                        ingreso or "--", salida or "--",
                        trabajado, obs_ingreso or "", obs_salida or "", motivo or ""
                    ])

            conexion.close()

            # Crear DataFrame y guardar
            columnas = [
                "RUT", "Nombre", "Apellido", "Fecha",
                "Ingreso", "Salida", "Horas Trabajadas",
                "Obs. Ingreso", "Obs. Salida", "Motivo"
            ]
            df = pd.DataFrame(datos_globales, columns=columnas)

            carpeta_descargas = os.path.join(os.path.expanduser("~"), "Downloads")
            nombre_archivo = f"reporte_mensual_completo_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            ruta_completa = os.path.join(carpeta_descargas, nombre_archivo)
            df.to_excel(ruta_completa, index=False)

            messagebox.showinfo("Exportación completa", f"Archivo guardado en:\n{ruta_completa}")
            try:
                os.startfile(ruta_completa)  # Windows
            except Exception:
                pass

        except Exception as e:
            messagebox.showerror("Error al exportar", str(e))

    # === INTERFAZ ===
    frame = ctk.CTkFrame(frame_padre)
    frame.pack(fill="both", expand=True)

    ctk.CTkLabel(frame, text="Reportes por Funcionario", font=("Arial", 16)).pack(pady=10)

    # === BUSCADOR POR NOMBRE ===
    lista_nombres = []
    dict_nombre_rut = {}
    try:
        conexion = sqlite3.connect("reloj_control.db")
        cursor = conexion.cursor()
        cursor.execute("SELECT rut, nombre, apellido FROM trabajadores")
        for rut_db, nombre_db, apellido_db in cursor.fetchall():
            nombre_completo = f"{nombre_db} {apellido_db}"
            lista_nombres.append(nombre_completo)
            dict_nombre_rut[nombre_completo] = rut_db
        conexion.close()
    except Exception as e:
        print(f"Error al cargar funcionarios: {e}")

    def actualizar_opciones(valor_actual):
        filtro = (valor_actual or "").lower()
        coincidencias = [n for n in lista_nombres if filtro in n.lower()]
        if not coincidencias:
            coincidencias = ["No hay coincidencias"]
        combo_funcionarios.configure(values=coincidencias)
        combo_funcionarios.set(valor_actual)
        if valor_actual not in dict_nombre_rut:
            entry_rut.delete(0, "end")

    def seleccionar_funcionario(valor):
        if valor == "No hay coincidencias":
            return
        rut_sel = dict_nombre_rut.get(valor, "")
        entry_rut.delete(0, "end")
        entry_rut.insert(0, rut_sel)

    combo_frame = ctk.CTkFrame(frame, fg_color="transparent")
    combo_frame.pack(pady=5)

    def al_hacer_click(event):
        if combo_funcionarios.get() == "Buscar Usuario por nombre":
            combo_funcionarios.set("")

    combo_funcionarios = ctk.CTkComboBox(
        combo_frame,
        values=["Buscar Usuario por nombre"],
        width=400,
        height=35,
        font=("Arial", 14),
        corner_radius=8,
        command=seleccionar_funcionario
    )
    combo_funcionarios.set("Buscar Usuario por nombre")
    combo_funcionarios.pack(side="left", padx=(0, 10))
    combo_funcionarios.bind("<FocusIn>", al_hacer_click)
    combo_funcionarios.bind("<KeyRelease>", lambda event: actualizar_opciones(combo_funcionarios.get()))
    combo_funcionarios.bind("<Return>", lambda event: actualizar_opciones(combo_funcionarios.get()))

    # Botón Limpiar
    def limpiar_combobox():
        combo_funcionarios.set("Buscar Usuario por nombre")
        combo_funcionarios.configure(values=lista_nombres)
        entry_rut.delete(0, 'end')

    ctk.CTkButton(combo_frame, text="Limpiar", width=80, height=35, command=limpiar_combobox).pack(side="left")

    # RUT y fechas
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

    # Totales y comparativos
    label_total = ctk.CTkLabel(frame, text="", font=("Arial", 12))
    label_total.pack(pady=2)

    label_total_semana = ctk.CTkLabel(frame, text="", font=("Arial", 12))
    label_total_semana.pack(pady=2)

    label_pactadas = ctk.CTkLabel(frame, text="", font=("Arial", 12))
    label_pactadas.pack(pady=2)

    label_completadas = ctk.CTkLabel(frame, text="", font=("Arial", 12))
    label_completadas.pack(pady=2)

    label_leyenda_admin = ctk.CTkLabel(frame, text="", font=("Arial", 12))
    label_leyenda_admin.pack(pady=2)

    label_estado = ctk.CTkLabel(frame, text="")
    label_estado.pack(pady=10)
