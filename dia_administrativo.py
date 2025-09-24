import os
import sqlite3
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import messagebox
import customtkinter as ctk
from tkcalendar import Calendar

# ===== PDF =====
try:
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib import colors
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.lib.enums import TA_LEFT
    PDF_OK = True
except Exception:
    PDF_OK = False

# ===== Correo =====
import smtplib
from email.message import EmailMessage
import mimetypes
import traceback
import sys
import subprocess

# Config SMTP (igual que otros informes)
SMTP_HOST = "mail.bioaccess.cl"
SMTP_PORT = 465
SMTP_USER = "documentos_bd@bioaccess.cl"
SMTP_PASS = "documentos@2025"
USAR_SMTP = True


# ----------------------- Utilidades varias -----------------------
def _downloads_dir():
    home = os.path.expanduser("~")
    cand = os.path.join(home, "Downloads")
    if os.path.isdir(cand):
        return cand
    xdg = os.environ.get("XDG_DOWNLOAD_DIR")
    if xdg and os.path.isdir(xdg):
        return xdg
    return home

def _abrir_archivo(path):
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)  # type: ignore
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception as e:
        messagebox.showinfo("Archivo generado", f"Se guardó en:\n{path}\n\nNo se pudo abrir automáticamente: {e}")

def _ruta_unica(path: str) -> str:
    base, ext = os.path.splitext(path)
    i = 1
    out = path
    while os.path.exists(out):
        out = f"{base} ({i}){ext}"
        i += 1
    return out

def _categoria_motivo(motivo: str) -> str:
    s = (motivo or "").lower()
    if "día administrativo" in s:
        return "Día Administrativo"
    if "matrimonio" in s or "unión civil" in s:
        return "Matrimonio/AUC"
    if "defunción" in s:
        if "hijo en gestación" in s:
            return "Defunción: Hijo en gestación"
        if "hijo" in s:
            return "Defunción: Hijo"
        if "cónyuge" in s or "conviviente civil" in s:
            return "Defunción: Cónyuge/Conviviente"
        if "padre" in s or "madre" in s or "hermano" in s:
            return "Defunción: Padre/Madre/Hermano(a)"
        return "Defunción: Otro"
    if "nacimiento paternal" in s:
        return "Nacimiento Paternal"
    if "alimentación" in s:
        return "Permiso de Alimentación"
    if "sin goce" in s:
        return "Permiso sin Goce de Sueldo"
    if "cometido de servicio" in s:
        return "Cometido de Servicio"
    if "licencia" in s:
        return "Licencia Médica"
    return "Otros permisos"

def _emails_para_rut(rut: str):
    try:
        con = sqlite3.connect("reloj_control.db")
        cur = con.cursor()
        cur.execute("SELECT correo FROM trabajadores WHERE rut = ?", (rut,))
        row = cur.fetchone()
        con.close()
        c = (row[0] or "").strip() if row else ""
        return [c] if c and "@" in c else []
    except Exception:
        return []

def _send_pdf_email(to_list, cc_list, subject, text_body, html_body, attach_path) -> bool:
    if not USAR_SMTP:
        messagebox.showinfo("Correo (simulado)", f"Para: {', '.join(to_list)}\nCC: {', '.join(cc_list)}\nAsunto: {subject}")
        return True
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = SMTP_USER
    msg["To"] = ", ".join([t for t in to_list if t])
    if cc_list:
        msg["Cc"] = ", ".join([c for c in cc_list if c])
    msg.set_content(text_body or "Adjunto reporte en PDF.")
    if html_body:
        msg.add_alternative(html_body, subtype="html")
    if attach_path and os.path.exists(attach_path):
        ctype, _ = mimetypes.guess_type(attach_path)
        maintype, subtype = (ctype.split("/",1) if ctype else ("application","pdf"))
        with open(attach_path, "rb") as f:
            msg.add_attachment(
                f.read(), maintype=maintype, subtype=subtype,
                filename=os.path.basename(attach_path)
            )
    try:
        with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, timeout=20) as s:
            s.login(SMTP_USER, SMTP_PASS)
            s.send_message(msg)
        return True
    except Exception as e:
        traceback.print_exc()
        messagebox.showerror("Correo", f"No fue posible enviar el correo:\n{e}")
        return False

def _html_email_permisos(period_text: str, nombre: str, rut: str, total_reg: int) -> str:
    PAGE_BG = "#f3f4f6"; CARD_BG = "#ffffff"; TEXT = "#111827"
    MUTED = "#6b7280"; BORDER = "#e5e7eb"; ACCENT = "#0b5ea8"
    gen = datetime.now().strftime("%d-%m-%Y %H:%M")
    return f"""<!doctype html>
<html lang="es"><meta charset="utf-8">
<body style="margin:0;padding:24px;background:{PAGE_BG};color:{TEXT};font-family:Arial,Helvetica,sans-serif">
  <div style="max-width:900px;margin:0 auto;background:{CARD_BG};border:1px solid {BORDER};border-radius:12px;padding:20px;box-sizing:border-box">
    <h1 style="margin:0 0 8px;font-size:22px;color:{TEXT}">Permisos y Días Administrativos</h1>
    <p style="margin:8px 0 14px;line-height:1.55;color:{TEXT}">
      Estimado(a), adjuntamos el informe de <strong>Permisos y Días Administrativos</strong>.
      Documento generado automáticamente por <strong>BioAccess – Control de Horarios</strong>.
    </p>
    <div style="margin:0 0 12px;line-height:1.7">
      <span style="display:inline-block;background:{ACCENT};color:#fff;border-radius:8px;padding:4px 10px;font-size:12px;margin-right:6px">
        Periodo: {period_text}
      </span>
      <span style="display:inline-block;background:{ACCENT};color:#fff;border-radius:8px;padding:4px 10px;font-size:12px;margin-right:6px">
        Funcionario: {nombre}
      </span>
      <span style="display:inline-block;background:{ACCENT};color:#fff;border-radius:8px;padding:4px 10px;font-size:12px">
        RUT: {rut}
      </span>
    </div>
    <p style="margin:0;color:{MUTED};font-size:12px">Registros incluidos: {total_reg}. Generado el {gen}.</p>
  </div>
</body>
</html>"""


# ----------------------- PDF builder -----------------------
def _build_pdf_permisos(path_pdf, rut, nombre_funcionario, filas, period_text, resumen_info):
    """
    filas: lista [(id, fecha_iso, motivo), ...] ordenadas por fecha
    resumen_info: dict con llaves:
        total, admin_usados, admin_limite, admin_restantes, licencias, extras, categorias (dict)
    """
    if not PDF_OK:
        raise RuntimeError("ReportLab no está instalado.")

    styles = getSampleStyleSheet()
    normal = styles["Normal"]
    title  = styles["Title"]; title.alignment = TA_LEFT
    note   = ParagraphStyle("Note", parent=normal, fontSize=9, textColor=colors.HexColor("#6b7280"))
    obs    = ParagraphStyle("Obs",  parent=normal, fontSize=10, leading=12, textColor=colors.HexColor("#111827"))

    elems = []

    # Título
    elems.append(Paragraph("Permisos y Días Administrativos", title))
    band = Table([[""]], colWidths=[27*cm], rowHeights=[0.25*cm])
    band.setStyle(TableStyle([("BACKGROUND", (0,0), (-1,-1), colors.HexColor("#0b5ea8"))]))
    elems.append(band)
    elems.append(Spacer(0, 6))

    # Cabecera con datos
    gen = datetime.now().strftime("%d/%m/%Y %H:%M")
    left = Paragraph(
        f"<b>Funcionario:</b> {nombre_funcionario or '—'}<br/>"
        f"<b>RUT:</b> {rut or '—'}",
        normal
    )
    right = Paragraph(
        f"<b>Periodo:</b> {period_text}<br/>"
        f"<b>Emitido el:</b> {gen}",
        normal
    )
    meta_tbl = Table([[left, right]], colWidths=[13.2*cm, 13.8*cm])
    meta_tbl.setStyle(TableStyle([
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("LEFTPADDING",  (0,0), (-1,-1), 0),
        ("RIGHTPADDING", (0,0), (-1,-1), 0),
        ("TOPPADDING",   (0,0), (-1,-1), 2),
        ("BOTTOMPADDING",(0,0), (-1,-1), 2),
    ]))
    elems.append(meta_tbl)
    elems.append(Spacer(0, 6))

    # Nota
    elems.append(Paragraph(
        "Documento generado automáticamente para fines administrativos. "
        "Revise la categoría y el detalle de cada registro en la tabla siguiente.",
        note
    ))
    elems.append(Spacer(0, 8))

    # Resumen compacto (panel)
    res = resumen_info
    res_tbl = Table([
        ["Total registros", str(res["total"])],
        ["Días administrativos (usados)", str(res["admin_usados"])],
        ["Límite anual", str(res["admin_limite"])],
        ["Restantes", str(res["admin_restantes"])],
        ["Permisos extras", str(res["extras"])],
        ["Licencias médicas", str(res["licencias"])],
    ], colWidths=[8.5*cm, 3.0*cm])
    res_tbl.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#D9E1F2")),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("ALIGN", (1,0), (1,-1), "RIGHT"),
        ("GRID", (0,0), (-1,-1), 0.25, colors.lightgrey),
        ("LEFTPADDING",  (0,0), (-1,-1), 4),
        ("RIGHTPADDING", (0,0), (-1,-1), 4),
        ("TOPPADDING",   (0,0), (-1,-1), 2),
        ("BOTTOMPADDING",(0,0), (-1,-1), 2),
    ]))
    elems.append(res_tbl)
    elems.append(Spacer(0, 8))

    # Tabla principal
    headers = ["ID", "Fecha", "Motivo"]
    data = [headers]
    for id_, fecha_iso, motivo in filas:
        fecha_leg = datetime.strptime(fecha_iso, "%Y-%m-%d").strftime("%d/%m/%Y")
        data.append([str(id_), fecha_leg, Paragraph(motivo or "—", obs)])

    # 27 cm útiles aprox.
    col_w = [2.0*cm, 3.0*cm, 22.0*cm]
    tbl = Table(data, colWidths=col_w, repeatRows=1)
    tbl.setStyle(TableStyle([
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE", (0,0), (-1,0), 10),
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#0b5ea8")),
        ("TEXTCOLOR",  (0,0), (-1,0), colors.white),
        ("ALIGN", (0,1), (0,-1), "CENTER"),
        ("ALIGN", (1,1), (1,-1), "CENTER"),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("GRID", (0,0), (-1,-1), 0.25, colors.lightgrey),
        ("LEFTPADDING",  (0,0), (-1,-1), 4),
        ("RIGHTPADDING", (0,0), (-1,-1), 4),
        ("TOPPADDING",   (0,0), (-1,-1), 2),
        ("BOTTOMPADDING",(0,0), (-1,-1), 2),
    ]))
    for i in range(1, len(data)):
        bg = "#f6f7fb" if i % 2 == 0 else "#ffffff"
        tbl.setStyle(TableStyle([("BACKGROUND", (0,i), (-1,i), colors.HexColor(bg))]))

    elems.append(tbl)
    elems.append(Spacer(0, 8))

    # Categorías
    cats = res.get("categorias", {})
    if cats:
        elems.append(Paragraph("Resumen por categoría", styles["Heading4"]))
        cdata = [["Categoría", "Días"]] + [[k, str(v)] for k, v in sorted(cats.items(), key=lambda x: (-x[1], x[0]))]
        ctbl = Table(cdata, colWidths=[14.0*cm, 3.0*cm])
        ctbl.setStyle(TableStyle([
            ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#D9E1F2")),
            ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
            ("ALIGN", (1,1), (1,-1), "RIGHT"),
            ("GRID", (0,0), (-1,-1), 0.25, colors.lightgrey),
        ]))
        elems.append(ctbl)
        elems.append(Spacer(0, 6))

    doc = SimpleDocTemplate(
        path_pdf,
        pagesize=landscape(A4),
        leftMargin=36, rightMargin=36, topMargin=36, bottomMargin=36
    )
    doc.build(elems)


# ----------------------- Construcción de UI -----------------------
def construir_dia_administrativo(frame_padre):
    # ==== helpers de calendario ====
    def seleccionar_fecha(entry_target):
        top = tk.Toplevel()
        cal = Calendar(top, date_pattern='dd/mm/yyyy', locale='es_CL')
        cal.pack(padx=10, pady=10)
        def poner_fecha():
            entry_target.delete(0, tk.END)
            entry_target.insert(0, cal.get_date())
            top.destroy()
            if entry_target is entry_fecha_desde and auto_calc_var.get():
                actualizar_fecha_hasta_auto()
        tk.Button(top, text="Seleccionar", command=poner_fecha).pack(pady=5)

    # ==== exportar / enviar (PDF) ====
    def _consultar_filas_y_nombre(rut, solo_ids=None, fecha_ini=None, fecha_fin=None):
        """
        retorna (filas, nombre_completo)
        filas = [(id, fecha_iso, motivo), ...] ordenado por fecha
        """
        try:
            conn = sqlite3.connect("reloj_control.db"); cur = conn.cursor()
            if solo_ids:
                placeholders = ",".join("?" for _ in solo_ids)
                cur.execute(f"""
                    SELECT id, fecha, motivo
                    FROM dias_libres
                    WHERE id IN ({placeholders})
                    ORDER BY fecha
                """, tuple(solo_ids))
            else:
                sql = "SELECT id, fecha, motivo FROM dias_libres WHERE rut = ?"
                params = [rut]
                if fecha_ini and fecha_fin:
                    sql += " AND fecha BETWEEN ? AND ?"
                    params.extend([fecha_ini, fecha_fin])
                sql += " ORDER BY fecha"
                cur.execute(sql, params)

            filas = cur.fetchall()
            cur.execute("SELECT nombre, apellido FROM trabajadores WHERE rut = ?", (rut,))
            row = cur.fetchone()
            nombre = (f"{row[0] or ''} {row[1] or ''}".strip()) if row else ""
            conn.close()
            return filas, nombre
        except Exception as e:
            raise RuntimeError(f"No se pudo consultar datos: {e}")

    def _armar_resumen(filas, rut):
        from collections import Counter
        cats = Counter()
        motivos = Counter()
        for _id, _f, mot in filas:
            motivos[mot or ""] += 1
            cats[_categoria_motivo(mot)] += 1
        admin_usados = cats.get("Día Administrativo", 0)
        lic = cats.get("Licencia Médica", 0)
        total = sum(motivos.values())
        extras = total - admin_usados
        # límite anual = 6 (regla local)
        anio_actual = datetime.now().year
        try:
            con = sqlite3.connect("reloj_control.db"); cur = con.cursor()
            cur.execute("""
                SELECT COUNT(*) FROM dias_libres 
                WHERE rut = ? AND strftime('%Y', fecha) = ? AND motivo = 'Día Administrativo'
            """, (rut, str(anio_actual)))
            admin_year = cur.fetchone()[0]
            con.close()
        except Exception:
            admin_year = admin_usados
        admin_rest = max(0, 6 - admin_year)
        return {
            "total": total,
            "admin_usados": admin_year,
            "admin_limite": 6,
            "admin_restantes": admin_rest,
            "licencias": lic,
            "extras": extras,
            "categorias": dict(cats),
        }

    def exportar_informe():
        rut = entry_rut.get().strip()
        if not rut:
            messagebox.showwarning("Falta RUT", "Primero selecciona/busca un trabajador (RUT).")
            return

        # ¿solo seleccionados?
        exportar_solo_seleccion = False
        if selected_ids:
            exportar_solo_seleccion = messagebox.askyesno(
                "Exportar",
                "¿Exportar SOLO los registros seleccionados?\n\nPulsa 'No' para exportar todos los del trabajador."
            )

        # rango
        fi = entry_fecha_desde.get().strip()
        ff = entry_fecha_hasta.get().strip()
        fecha_ini = fecha_fin = None
        period_text = "Todos"
        if fi and ff:
            try:
                fecha_ini = datetime.strptime(fi, "%d/%m/%Y").strftime("%Y-%m-%d")
                fecha_fin = datetime.strptime(ff, "%d/%m/%Y").strftime("%Y-%m-%d")
                period_text = f"{fi} a {ff}"
            except ValueError:
                fecha_ini = fecha_fin = None
                period_text = "Todos"

        try:
            if exportar_solo_seleccion:
                filas, nombre = _consultar_filas_y_nombre(rut, solo_ids=sorted(selected_ids))
            else:
                filas, nombre = _consultar_filas_y_nombre(rut, fecha_ini=fecha_ini, fecha_fin=fecha_fin)
        except Exception as e:
            messagebox.showerror("Error", str(e)); return

        if not filas:
            messagebox.showinfo("Sin datos", "No hay registros para exportar con los filtros actuales.")
            return

        if not PDF_OK:
            return messagebox.showerror("PDF", "Instala reportlab (pip install reportlab) para exportar en PDF.")

        carpeta = _downloads_dir()
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        base = f"Permisos_{rut}_{ts}"
        if exportar_solo_seleccion:
            base += f"_sel{len(filas)}"
        ruta = _ruta_unica(os.path.join(carpeta, f"{base}.pdf"))

        try:
            resumen = _armar_resumen(filas, rut)
            _build_pdf_permisos(ruta, rut, nombre, filas, period_text, resumen)
            _abrir_archivo(ruta)
            messagebox.showinfo("Exportación exitosa", f"El informe se guardó en:\n{ruta}")
        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("PDF", f"No se pudo generar el PDF:\n{e}")

    def enviar_informe():
        rut = entry_rut.get().strip()
        if not rut:
            messagebox.showwarning("Falta RUT", "Primero selecciona/busca un trabajador (RUT).")
            return

        # rango
        fi = entry_fecha_desde.get().strip()
        ff = entry_fecha_hasta.get().strip()
        fecha_ini = fecha_fin = None
        period_text = "Todos"
        if fi and ff:
            try:
                fecha_ini = datetime.strptime(fi, "%d/%m/%Y").strftime("%Y-%m-%d")
                fecha_fin = datetime.strptime(ff, "%d/%m/%Y").strftime("%Y-%m-%d")
                period_text = f"{fi} a {ff}"
            except ValueError:
                fecha_ini = fecha_fin = None
                period_text = "Todos"

        # ¿solo selección si hay?
        exportar_solo_seleccion = False
        if selected_ids:
            exportar_solo_seleccion = messagebox.askyesno(
                "Enviar",
                "¿Enviar SOLO los registros seleccionados?\n\nPulsa 'No' para enviar todos los del trabajador."
            )

        try:
            if exportar_solo_seleccion:
                filas, nombre = _consultar_filas_y_nombre(rut, solo_ids=sorted(selected_ids))
            else:
                filas, nombre = _consultar_filas_y_nombre(rut, fecha_ini=fecha_ini, fecha_fin=fecha_fin)
        except Exception as e:
            messagebox.showerror("Error", str(e)); return

        if not filas:
            messagebox.showinfo("Sin datos", "No hay registros para enviar con los filtros actuales.")
            return

        if not PDF_OK:
            return messagebox.showerror("PDF", "Instala reportlab (pip install reportlab) para generar el adjunto PDF.")

        # Generar PDF temporal en Descargas
        carpeta = _downloads_dir()
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        base = f"Permisos_{rut}_{ts}"
        if exportar_solo_seleccion:
            base += f"_sel{len(filas)}"
        ruta_pdf = _ruta_unica(os.path.join(carpeta, f"{base}.pdf"))
        try:
            resumen = _armar_resumen(filas, rut)
            _build_pdf_permisos(ruta_pdf, rut, nombre, filas, period_text, resumen)
        except Exception as e:
            traceback.print_exc()
            return messagebox.showerror("PDF", f"No se pudo preparar el adjunto:\n{e}")

        # Sugerir destinatarios (correo del funcionario si existe)
        sugeridos = _emails_para_rut(rut)

        # Diálogo envío
        TopSend = getattr(ctk, "CTkToplevel", None) or tk.Toplevel
        dlg = TopSend(frame); dlg.title("Enviar por correo (PDF)")
        try:
            dlg.transient(frame); dlg.grab_set(); dlg.resizable(False, False)
        except Exception: pass

        box = ctk.CTkFrame(dlg, corner_radius=12); box.pack(fill="both", expand=True, padx=16, pady=16)
        ctk.CTkLabel(box, text="Enviar reporte — Permisos y Días Administrativos", font=("Arial", 16, "bold"))\
            .grid(row=0, column=0, columnspan=2, pady=(0,8), sticky="w")

        ctk.CTkLabel(box, text="Para:").grid(row=1, column=0, sticky="e", padx=(0,8), pady=6)
        entry_to = ctk.CTkEntry(box, width=520, placeholder_text="correo1@dominio.cl, correo2@dominio.cl")
        entry_to.grid(row=1, column=1, sticky="w", pady=6)
        if sugeridos:
            entry_to.insert(0, ", ".join(sugeridos))

        ctk.CTkLabel(box, text="CC:").grid(row=2, column=0, sticky="e", padx=(0,8), pady=(0,6))
        entry_cc = ctk.CTkEntry(box, width=520, placeholder_text="(opcional) separar por coma")
        entry_cc.grid(row=2, column=1, sticky="w", pady=(0,6))

        ctk.CTkLabel(
            box,
            text=f"Periodo: {period_text}   |   Funcionario: {nombre} ({rut})   |   Registros: {len(filas)}",
            text_color="#9ca3af"
        ).grid(row=3, column=0, columnspan=2, sticky="w", pady=(4,2))

        btns = ctk.CTkFrame(box, fg_color="transparent")
        btns.grid(row=4, column=0, columnspan=2, sticky="ew", pady=(10,0))
        btn_cancel = ctk.CTkButton(btns, text="Cancelar", fg_color="#475569", command=dlg.destroy, width=120)
        btn_send2 = ctk.CTkButton(btns, text="Enviar", fg_color="#22c55e", width=160)
        btn_cancel.pack(side="left", padx=(0,6))
        btn_send2.pack(side="right", padx=(6,0))
        box.grid_columnconfigure(1, weight=1)

        def do_send_now():
            to_raw = entry_to.get().strip()
            cc_raw = entry_cc.get().strip()
            to_list = [x.strip() for x in to_raw.split(",") if x.strip()]
            cc_list = [x.strip() for x in cc_raw.split(",") if x.strip()]
            if not to_list:
                return messagebox.showerror("Correo", "Indica al menos un destinatario en 'Para'.")
            gen = datetime.now().strftime("%d-%m-%Y %H:%M")
            text_body = (
                "Estimado(a), adjuntamos el informe de Permisos y Días Administrativos.\n"
                "Documento generado automáticamente por BioAccess – Control de Horarios.\n\n"
                f"Funcionario : {nombre} ({rut})\n"
                f"Periodo     : {period_text}\n"
                f"Registros   : {len(filas)}\n"
                f"Generado    : {gen}\n"
            )
            html_body = _html_email_permisos(period_text, nombre or "—", rut or "—", len(filas))
            subject   = f"Permisos y Días Administrativos — {nombre or rut} — {period_text}"

            ok = _send_pdf_email(to_list, cc_list, subject, text_body, html_body, ruta_pdf)
            if ok:
                messagebox.showinfo("Correo", "Reporte enviado correctamente.")
                try: dlg.destroy()
                except Exception: pass

        btn_send2.configure(command=do_send_now)

        frame.update_idletasks()
        try:
            w = max(640, dlg.winfo_reqwidth() + 32)
            h = max(260, dlg.winfo_reqheight() + 16)
        except Exception:
            w, h = 640, 260
        x = frame.winfo_rootx() + (frame.winfo_width() // 2) - (w // 2)
        y = frame.winfo_rooty()  + (frame.winfo_height() // 2) - (h // 2)
        dlg.geometry(f"{w}x{h}+{max(x,0)}+{max(y,0)}")

    # --------------- UI y lógica existente (ajustada) ---------------
    frame = ctk.CTkFrame(frame_padre)
    frame.pack(fill="both", expand=True, padx=20, pady=20)

    ctk.CTkLabel(frame, text="Asignar, Ver o Editar Días Administrativos, Permisos y Licencias",
                 font=("Arial", 18)).pack(pady=10)

    # Carga nombres/ruts
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

    # Buscador nombre
    fila_nombre = ctk.CTkFrame(frame, fg_color="transparent")
    fila_nombre.pack(pady=10)
    combo_nombre = ctk.CTkComboBox(fila_nombre, values=primeros_10, width=250)
    combo_nombre.set("Buscar por Nombre")
    combo_nombre.pack(side="left", padx=(0, 5))
    btn_buscar = ctk.CTkButton(fila_nombre, text="🔍 Buscar", command=lambda: buscar_solicitudes())
    btn_buscar.pack(side="left", padx=(5, 0))

    # RUT
    entry_rut = ctk.CTkEntry(frame, placeholder_text="RUT: (Ej: 12345678-9)", width=150)
    entry_rut.pack(pady=5)
    entry_rut.bind("<Return>", lambda event: mostrar_vista_previa())

    label_resumen_admin  = ctk.CTkLabel(frame, text="", text_color="green", font=("Arial", 13))
    label_resumen_admin.pack(pady=(0, 5))
    label_resumen_extras = ctk.CTkLabel(frame, text="", text_color="red", font=("Arial", 13))
    label_resumen_extras.pack(pady=(0, 10))

    ctk.CTkLabel(frame, text="IMPORTANTE", font=("Arial", 13, "bold"), text_color="#d32f2f").pack(pady=(12, 2))
    ctk.CTkLabel(frame, text="Nueva Solicitud de Permiso o Licencia   ▼",
                 font=("Arial", 15, "bold"), text_color="#607D8B", anchor="center").pack(pady=(0, 8))

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

    combo_tipo_permiso = ctk.CTkComboBox(frame, values=opciones_permiso, width=400)
    combo_tipo_permiso.set("Elija tipo de Solicitud o Permiso")
    combo_tipo_permiso.pack(pady=5)

    auto_calc_var = tk.BooleanVar(value=True)
    chk_auto = ctk.CTkCheckBox(frame, text="Calcular automáticamente la fecha hasta", variable=auto_calc_var)
    chk_auto.pack(pady=(2, 8))

    cont_desde = ctk.CTkFrame(frame, fg_color="transparent"); cont_desde.pack(pady=2)
    entry_fecha_desde = ctk.CTkEntry(cont_desde, placeholder_text="Fecha Desde (dd/mm/aaaa)", width=200)
    entry_fecha_desde.pack(side="left")
    ctk.CTkButton(cont_desde, text="📅", width=40, command=lambda: seleccionar_fecha(entry_fecha_desde)).pack(side="left", padx=5)

    cont_hasta = ctk.CTkFrame(frame, fg_color="transparent"); cont_hasta.pack(pady=2)
    entry_fecha_hasta = ctk.CTkEntry(cont_hasta, placeholder_text="Fecha Hasta (dd/mm/aaaa)", width=200)
    entry_fecha_hasta.pack(side="left")
    ctk.CTkButton(cont_hasta, text="📅", width=40, command=lambda: seleccionar_fecha(entry_fecha_hasta)).pack(side="left", padx=5)
    entry_fecha_hasta.configure(state="disabled")

    entry_motivo_otro = ctk.CTkEntry(frame, placeholder_text="Especificar motivo...", width=400)
    entry_motivo_otro.pack_forget()
    label_info_cometido = ctk.CTkLabel(frame, text="", text_color="orange", font=("Arial", 12))
    label_info_cometido.pack()

    # Autocomplete / búsqueda
    def limpiar_placeholder(_):
        if combo_nombre.get() == "Buscar por Nombre":
            combo_nombre.set("")
    def restaurar_placeholder(_):
        if combo_nombre.get() == "":
            combo_nombre.set("Buscar por Nombre")
    combo_nombre.bind("<FocusIn>", limpiar_placeholder)
    combo_nombre.bind("<FocusOut>", restaurar_placeholder)

    def mostrar_sugerencias(_):
        texto = combo_nombre.get().lower()
        if not texto or texto == "buscar por nombre":
            combo_nombre.configure(values=primeros_10)
    combo_nombre.bind("<FocusIn>", mostrar_sugerencias)

    def autocompletar_nombres(_):
        texto = combo_nombre.get().lower()
        if not texto or texto == "buscar por nombre":
            combo_nombre.configure(values=primeros_10)
        else:
            filtrados = [n for n in lista_nombres if texto in n.lower()]
            combo_nombre.configure(values=filtrados[:10] if filtrados else ["No encontrado"])
    combo_nombre.bind("<KeyRelease>", autocompletar_nombres)

    def buscar_por_nombre(event=None):
        nombre = combo_nombre.get()
        rutx = dict_nombre_rut.get(nombre, "")
        if rutx:
            entry_rut.delete(0, tk.END)
            entry_rut.insert(0, rutx)
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
    btn_buscar.configure(command=buscar_solicitudes)

    # helpers fecha
    def sumar_dias_habiles(fecha_inicio, dias_habiles):
        dias_agregados = 0
        fecha = fecha_inicio
        while dias_agregados < dias_habiles:
            if fecha.weekday() < 5:
                dias_agregados += 1
                if dias_agregados == dias_habiles:
                    break
            fecha += timedelta(days=1)
        return fecha

    def sumar_dias_corridos_inclusivo(fecha_inicio, dias_corridos):
        return fecha_inicio + timedelta(days=dias_corridos - 1)

    # multiselección
    selected_ids = set()
    checkbox_vars = {}
    def actualizar_boton_eliminar():
        count = len(selected_ids)
        try:
            btn_eliminar_multi.configure(
                text=f"🗑️ Eliminar seleccionados ({count})",
                state=("normal" if count > 0 else "disabled")
            )
        except NameError:
            pass
    def on_toggle_checkbox(id_, checked):
        if checked: selected_ids.add(id_)
        else: selected_ids.discard(id_)
        actualizar_boton_eliminar()
    def seleccionar_todo():
        marcar = not all(var.get() for var in checkbox_vars.values())
        for id_, var in checkbox_vars.items():
            var.set(marcar)
            if marcar: selected_ids.add(id_)
            else: selected_ids.discard(id_)
        actualizar_boton_eliminar()
    def eliminar_seleccionados():
        if not selected_ids: return
        if not messagebox.askyesno("Confirmar", f"¿Eliminar {len(selected_ids)} registro(s) seleccionado(s)?"):
            return
        conn = sqlite3.connect("reloj_control.db")
        cur = conn.cursor()
        placeholders = ",".join("?" for _ in selected_ids)
        cur.execute(f"DELETE FROM dias_libres WHERE id IN ({placeholders})", tuple(selected_ids))
        conn.commit(); conn.close()
        cant = len(selected_ids)
        selected_ids.clear(); checkbox_vars.clear()
        mostrar_vista_previa(); actualizar_boton_eliminar()
        messagebox.showinfo("Eliminado", f"Se eliminaron {cant} registro(s).")

    # tabla
    tabla_preview = ctk.CTkScrollableFrame(frame)
    tabla_preview.pack(fill="both", expand=True, pady=10)

    acciones_sel_frame = ctk.CTkFrame(frame, fg_color="transparent")
    acciones_sel_frame.pack(fill="x", pady=(4, 6))
    btn_select_all = ctk.CTkButton(acciones_sel_frame, text="Seleccionar todo",
                                   command=seleccionar_todo, width=140)
    btn_select_all.pack(side="left", padx=(0, 8))
    btn_eliminar_multi = ctk.CTkButton(
        acciones_sel_frame, text="🗑️ Eliminar seleccionados (0)",
        fg_color="red", state="disabled", command=eliminar_seleccionados, width=220
    )
    btn_eliminar_multi.pack(side="left")
    actualizar_boton_eliminar()

    # auto-cálculo de fecha hasta
    def actualizar_fecha_hasta_auto(event=None):
        tipo = combo_tipo_permiso.get()
        dias = dias_por_permiso.get(tipo)
        fecha_inicio_str = entry_fecha_desde.get().strip()
        if not auto_calc_var.get():
            return
        if dias and fecha_inicio_str:
            try:
                fecha_inicio = datetime.strptime(fecha_inicio_str, "%d/%m/%Y")
                if "días hábiles" in tipo.lower():
                    fecha_fin = sumar_dias_habiles(fecha_inicio, dias)
                else:
                    fecha_fin = sumar_dias_corridos_inclusivo(fecha_inicio, dias)
                entry_fecha_hasta.configure(state="normal")
                entry_fecha_hasta.delete(0, tk.END)
                entry_fecha_hasta.insert(0, fecha_fin.strftime("%d/%m/%Y"))
                entry_fecha_hasta.configure(state="disabled")
            except ValueError:
                entry_fecha_hasta.configure(state="normal"); entry_fecha_hasta.delete(0, tk.END)
        else:
            entry_fecha_hasta.configure(state="normal"); entry_fecha_hasta.delete(0, tk.END)

    entry_fecha_desde.bind("<FocusOut>", actualizar_fecha_hasta_auto)
    entry_fecha_desde.bind("<Return>", actualizar_fecha_hasta_auto)

    def guardar_dias_admin():
        rut = entry_rut.get().strip()
        fecha_inicio_str = entry_fecha_desde.get().strip()
        fecha_fin_str = entry_fecha_hasta.get().strip()
        motivo_sel = combo_tipo_permiso.get()

        if motivo_sel == "Elija tipo de Solicitud o Permiso":
            messagebox.showerror("Error", "Debes elegir un tipo de solicitud o permiso."); return

        if motivo_sel == "Otro (especificar)":
            motivo = entry_motivo_otro.get().strip() or motivo_sel
        elif motivo_sel == "Cometido de Servicio":
            detalle = entry_motivo_otro.get().strip()
            motivo = f"Cometido de Servicio - {detalle}" if detalle else motivo_sel
        else:
            motivo = motivo_sel

        if fecha_inicio_str and not fecha_fin_str:
            dias = dias_por_permiso.get(motivo_sel)
            if motivo_sel == "Día Administrativo":
                fecha_fin_str = fecha_inicio_str
            elif dias:
                try:
                    f0 = datetime.strptime(fecha_inicio_str, "%d/%m/%Y")
                    f1 = sumar_dias_habiles(f0, dias) if "días hábiles" in motivo_sel.lower() else sumar_dias_corridos_inclusivo(f0, dias)
                    fecha_fin_str = f1.strftime("%d/%m/%Y")
                except ValueError:
                    pass

        if not rut or not fecha_inicio_str or not fecha_fin_str:
            messagebox.showerror("Error", "Debes ingresar todos los campos."); return

        try:
            fecha_inicio = datetime.strptime(fecha_inicio_str, "%d/%m/%Y")
            fecha_fin    = datetime.strptime(fecha_fin_str, "%d/%m/%Y")
        except ValueError:
            messagebox.showerror("Error", "Fechas inválidas. Usa formato dd/mm/aaaa."); return

        if fecha_fin < fecha_inicio:
            messagebox.showerror("Error", "La fecha final no puede ser anterior a la inicial."); return

        conn = sqlite3.connect("reloj_control.db"); cur = conn.cursor()
        dias_registrados = 0
        fecha_actual = fecha_inicio
        while fecha_actual <= fecha_fin:
            if "días hábiles" in motivo_sel.lower() and fecha_actual.weekday() >= 5:
                fecha_actual += timedelta(days=1); continue
            fecha_iso = fecha_actual.strftime("%Y-%m-%d")
            cur.execute("SELECT COUNT(*) FROM dias_libres WHERE rut = ? AND fecha = ?", (rut, fecha_iso))
            if cur.fetchone()[0] == 0:
                cur.execute("INSERT INTO dias_libres (rut, fecha, motivo) VALUES (?, ?, ?)", (rut, fecha_iso, motivo))
                dias_registrados += 1
            fecha_actual += timedelta(days=1)
        conn.commit(); conn.close()
        messagebox.showinfo("Éxito", f"{dias_registrados} día(s) guardado(s) correctamente.")
        mostrar_vista_previa()

    def limpiar_campos():
        entry_rut.delete(0, tk.END)
        entry_fecha_desde.delete(0, tk.END)
        entry_fecha_hasta.delete(0, tk.END)
        combo_tipo_permiso.set("Elija tipo de Solicitud o Permiso")
        entry_motivo_otro.delete(0, tk.END); entry_motivo_otro.pack_forget()
        combo_nombre.set("Buscar por Nombre"); limpiar_placeholder(None)
        label_resumen_admin.configure(text=""); label_resumen_extras.configure(text="", text_color="red")
        for w in tabla_preview.winfo_children(): w.destroy()
        if auto_calc_var.get(): entry_fecha_hasta.configure(state="disabled")
        else: entry_fecha_hasta.configure(state="normal")
        selected_ids.clear(); checkbox_vars.clear(); actualizar_boton_eliminar()

    # botones acción
    botones_frame = ctk.CTkFrame(frame, fg_color="transparent"); botones_frame.pack(pady=(10, 10))
    ctk.CTkButton(botones_frame, text="Guardar Nueva Solicitud", command=guardar_dias_admin, width=180)\
        .pack(side="left", padx=10)
    ctk.CTkButton(botones_frame, text="⬇️ Descargar (PDF)", command=exportar_informe, width=170)\
        .pack(side="left", padx=10)
    ctk.CTkButton(botones_frame, text="✉️ Enviar (PDF)", command=enviar_informe, width=160, fg_color="#22c55e")\
        .pack(side="left", padx=10)
    ctk.CTkButton(botones_frame, text="🧹 Limpiar Formulario", fg_color="gray", command=limpiar_campos, width=160)\
        .pack(side="left", padx=10)

    # visibilidad "otro/cometido"
    def actualizar_visibilidad_entry_otro(choice):
        entry_fecha_desde.configure(state="normal")
        entry_fecha_hasta.configure(state="normal")
        if choice == "Cometido de Servicio":
            label_info_cometido.configure(text="⚠ Este permiso simula una jornada completa. No se requiere hora, solo el motivo.")
            entry_motivo_otro.pack(pady=(0, 10))
            entry_fecha_desde.delete(0, tk.END); entry_fecha_hasta.delete(0, tk.END)
        elif choice == "Otro (especificar)":
            label_info_cometido.configure(text="")
            entry_motivo_otro.pack(pady=(0, 10))
            entry_fecha_desde.delete(0, tk.END); entry_fecha_hasta.delete(0, tk.END)
        elif choice == "Día Administrativo":
            label_info_cometido.configure(text="")
            entry_motivo_otro.pack_forget()
            entry_fecha_desde.delete(0, tk.END); entry_fecha_hasta.delete(0, tk.END)
        elif choice in dias_por_permiso:
            label_info_cometido.configure(text="")
            entry_motivo_otro.pack_forget()
        else:
            label_info_cometido.configure(text="")
            entry_motivo_otro.pack_forget()
            entry_fecha_desde.delete(0, tk.END); entry_fecha_hasta.delete(0, tk.END)

        if auto_calc_var.get(): entry_fecha_hasta.configure(state="disabled")
        else: entry_fecha_hasta.configure(state="normal")
        actualizar_fecha_hasta_auto()

    combo_tipo_permiso.configure(command=actualizar_visibilidad_entry_otro)

    def _on_toggle_auto():
        if auto_calc_var.get():
            entry_fecha_hasta.configure(state="disabled"); actualizar_fecha_hasta_auto()
        else:
            entry_fecha_hasta.configure(state="normal")
    chk_auto.configure(command=_on_toggle_auto)

    # vista previa
    def mostrar_vista_previa():
        for w in tabla_preview.winfo_children(): w.destroy()
        selected_ids.clear(); checkbox_vars.clear(); actualizar_boton_eliminar()

        rut = entry_rut.get().strip()
        if not rut:
            label_resumen_admin.configure(text="", text_color="green")
            label_resumen_extras.configure(text="", text_color="red")
            return

        anio_actual = datetime.now().year
        conn = sqlite3.connect("reloj_control.db"); cur = conn.cursor()
        cur.execute("""
            SELECT COUNT(*) FROM dias_libres 
            WHERE rut = ? AND strftime('%Y', fecha) = ? AND motivo = 'Día Administrativo'
        """, (rut, str(anio_actual)))
        total_admin = cur.fetchone()[0]
        MAX_ADMIN_ANUAL = 6
        color = "green" if total_admin < MAX_ADMIN_ANUAL else ("orange" if total_admin == MAX_ADMIN_ANUAL else "red")
        label_resumen_admin.configure(
            text=f"📅 Días administrativos solicitados este año: {total_admin} / {MAX_ADMIN_ANUAL} disponibles",
            text_color=color
        )

        cur.execute("""
            SELECT COUNT(*) FROM dias_libres 
            WHERE rut = ? AND strftime('%Y', fecha) = ? AND motivo != 'Día Administrativo'
        """, (rut, str(anio_actual)))
        total_extras = cur.fetchone()[0]
        label_resumen_extras.configure(text=f"🧾 Permisos extras este año: {total_extras}")

        cur.execute("SELECT id, fecha, motivo FROM dias_libres WHERE rut = ? ORDER BY fecha", (rut,))
        registros = cur.fetchall()
        conn.close()

        headers = ["Sel.", "Fecha", "Motivo", "Guardar", "Eliminar"]
        for i, header in enumerate(headers):
            ctk.CTkLabel(tabla_preview, text=header, font=("Arial", 13, "bold")).grid(row=0, column=i, padx=10, pady=5)

        for idx, (id_, fecha, motivo) in enumerate(registros, start=1):
            var = tk.BooleanVar(value=False)
            chk = ctk.CTkCheckBox(tabla_preview, text="", variable=var,
                                  command=lambda i=id_, v=var: on_toggle_checkbox(i, v.get()))
            chk.grid(row=idx, column=0, padx=6, pady=2); checkbox_vars[id_] = var

            fecha_legible = datetime.strptime(fecha, "%Y-%m-%d").strftime("%d/%m/%Y")
            ctk.CTkLabel(tabla_preview, text=fecha_legible).grid(row=idx, column=1, padx=10, pady=2)

            entry_motivo_edit = ctk.CTkEntry(tabla_preview, width=400, font=("Arial", 14))
            entry_motivo_edit.insert(0, motivo or ""); entry_motivo_edit.grid(row=idx, column=2, padx=10, pady=2)

            ctk.CTkButton(tabla_preview, text="💾", width=30, fg_color="green",
                          command=lambda i=id_, e=entry_motivo_edit: actualizar_motivo(i, e)).grid(row=idx, column=3, padx=5)

            ctk.CTkButton(tabla_preview, text="❌", width=30, fg_color="red",
                          command=lambda i=id_: eliminar_dia_admin(i)).grid(row=idx, column=4, padx=5)

    def actualizar_motivo(id_, entry_widget):
        nuevo_motivo = entry_widget.get().strip()
        conn = sqlite3.connect("reloj_control.db"); cur = conn.cursor()
        cur.execute("UPDATE dias_libres SET motivo = ? WHERE id = ?", (nuevo_motivo, id_))
        conn.commit(); conn.close()
        messagebox.showinfo("Actualizado", "Motivo actualizado correctamente.")

    def eliminar_dia_admin(id_):
        if not messagebox.askyesno("Confirmar", "¿Deseas eliminar este día administrativo/permisos?"):
            return
        conn = sqlite3.connect("reloj_control.db"); cur = conn.cursor()
        cur.execute("DELETE FROM dias_libres WHERE id = ?", (id_,))
        conn.commit(); conn.close()
        mostrar_vista_previa()
        messagebox.showinfo("Eliminado", "Registro eliminado correctamente.")

    mostrar_vista_previa()
