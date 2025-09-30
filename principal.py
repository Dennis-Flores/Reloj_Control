# principal.py
import os
import sys
import sqlite3
import tkinter as tk
import customtkinter as ctk
from tkinter import messagebox
import traceback

from db import crear_bd
from panel_avanzado import construir_panel_avanzado
from solicitudes import construir_solicitudes
from cambio_clave_admin import abrir_cambio_clave
from dia_administrativo import construir_dia_administrativo
from resumen_dia import construir_resumen_dia

# ========== Utilidades de ruta ==========
def app_path() -> str:
    # En .exe: carpeta del ejecutable; en .py: carpeta del archivo
    return os.path.dirname(sys.executable) if getattr(sys, "frozen", False) \
           else os.path.dirname(os.path.abspath(__file__))

BASE = app_path()
# Fijar directorio de trabajo (crítico para rutas relativas como "reloj_control.db")
os.chdir(BASE)

DB_PATH = os.path.join(BASE, "reloj_control.db")
CARPETA_ROSTROS = os.path.join(BASE, "rostros")
MODELS_PATH = os.path.join(BASE, "face_recognition_models")

# ========== Configuración general de UI ==========
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

admin_info = None  # Almacena los datos del administrador logueado

# ========== Ventana ==========
app = ctk.CTk()
app.title("BioAccess/Control de Horarios | www.bioaccess.cl")

# Tamaño mínimo si se desmaximiza (opcional pero recomendado)
app.minsize(1200, 700)  # ← NUEVO

# Maximizar con fallback multiplataforma y reintento tras el primer render
def _maximizar():
    try:
        app.state("zoomed")          # Windows
    except Exception:
        try:
            app.attributes("-zoomed", True)  # Linux/otros
        except Exception:
            # Fallback: ocupar toda la pantalla
            app.update_idletasks()
            sw, sh = app.winfo_screenwidth(), app.winfo_screenheight()
            app.geometry(f"{sw}x{sh}+0+0")

# Llamar ahora...
_maximizar()
# ...y volver a llamarlo después del primer dibujo para evitar que se “achique”
app.after_idle(_maximizar)  # ← NUEVO


# ========== Inicializar base de datos ==========
try:
    crear_bd(DB_PATH)   # <<--- ahora recibe la ruta
except Exception as e:
    messagebox.showerror("Base de datos", f"No se pudo crear/abrir la BD:\n{e}")
    app.destroy()
    sys.exit(1)

# ========== Helpers ==========
def safe_focus(widget):
    """Intenta enfocar un widget sin reventar si ya no existe."""
    try:
        if widget and widget.winfo_exists():
            widget.focus_set()
            return True
    except Exception:
        pass
    return False

# ========== Etiqueta contador ==========
label_contador = ctk.CTkLabel(app, text="", font=("Arial", 13), text_color="gray")
label_contador.pack(pady=(5, 10))

def actualizar_contador():
    con = sqlite3.connect(DB_PATH)
    cur = con.cursor()
    cur.execute("SELECT COUNT(*) FROM trabajadores")
    total = cur.fetchone()[0]
    con.close()
    label_contador.configure(text=f"📋 Trabajadores registrados: {total}")

actualizar_contador()

# ========== Mostrar nombre del admin logueado ==========
label_admin_activo = ctk.CTkLabel(app, text="", font=("Arial", 16), text_color="lightgreen")
label_admin_activo.pack(pady=(0, 5))

def mostrar_info_admin():
    if admin_info:
        nombre = admin_info["nombre"]
        permiso = admin_info["permiso"]
        label_admin_activo.configure(text=f"👤 {nombre} | {permiso}")
    else:
        label_admin_activo.configure(text="")

# ========== Contenedor de menú ==========
menu = ctk.CTkFrame(app)
menu.pack(side="top", fill="x", padx=20, pady=(10, 5))

# ========== Contenedor principal ==========
frame_contenedor = ctk.CTkFrame(app)
frame_contenedor.pack(fill="both", expand=True, padx=20, pady=(5, 20))

# ========== Funciones de navegación ==========
def limpiar_frame():
    for widget in frame_contenedor.winfo_children():
        widget.destroy()

def mostrar_registro_con_refresco():
    limpiar_frame()
    from registrar import construir_registro
    construir_registro(frame_contenedor, actualizar_contador)

def mostrar_ingreso_salida():
    limpiar_frame()
    from ingreso_salida import construir_ingreso_salida
    construir_ingreso_salida(frame_contenedor)

def mostrar_reportes():
    limpiar_frame()
    from reportes import construir_reportes
    construir_reportes(frame_contenedor)

def mostrar_nomina():
    limpiar_frame()
    from nomina import construir_nomina
    construir_nomina(frame_contenedor)

def mostrar_editar_usuario():
    limpiar_frame()
    from editar_usuario import construir_edicion
    construir_edicion(
        frame_contenedor,
        on_actualizacion=actualizar_contador,
        on_volver_inicio=lambda: [
            mostrar_ingreso_salida(),
            resaltar_boton_activo("ingreso")
        ]
    )

# ========== Ventana de Bienvenida ==========
def mostrar_bienvenida(nombre: str, autoclose_ms: int = 2500):
    TopLevelCls = getattr(ctk, "CTkToplevel", None) or tk.Toplevel
    win = TopLevelCls(app)
    win.title("Acceso permitido")
    try:
        win.resizable(False, False)
        win.transient(app)
        win.grab_set()
    except Exception:
        pass

    cont = ctk.CTkFrame(win, corner_radius=12)
    cont.pack(fill="both", expand=True, padx=16, pady=16)

    lbl = ctk.CTkLabel(cont, text=f"Bienvenido: {nombre}",
                       font=("Arial", 18, "bold"), text_color="lightgreen")
    lbl.pack(pady=(20, 10))

    def cerrar_bienvenida():
        safe_focus(app)
        if win and win.winfo_exists():
            win.destroy()

    btn = ctk.CTkButton(cont, text="Aceptar", command=cerrar_bienvenida)
    btn.pack(pady=(0, 16))

    # Centrado
    app.update_idletasks()
    w, h = 340, 160
    try:
        w = max(w, win.winfo_reqwidth() + 32)
        h = max(h, win.winfo_reqheight() + 16)
    except Exception:
        pass
    x = app.winfo_x() + (app.winfo_width() // 2) - (w // 2)
    y = app.winfo_y() + (app.winfo_height() // 2) - (h // 2)
    win.geometry(f"{w}x{h}+{max(x,0)}+{max(y,0)}")

    # Fade-in
    try:
        win.attributes("-alpha", 0.0)
        def fade_in(a=0.0):
            if not (win and win.winfo_exists()):
                return
            a = round(a + 0.08, 2)
            try:
                win.attributes("-alpha", min(a, 1.0))
            except Exception:
                return
            if a < 1.0 and win.winfo_exists():
                win.after(15, fade_in, a)
        fade_in()
    except Exception:
        pass

    if autoclose_ms and autoclose_ms > 0:
        win.after(autoclose_ms, cerrar_bienvenida)
    win.bind("<Escape>", lambda e: cerrar_bienvenida())

# ========== Manejo de botones activos ==========
botones_menu = {}

def resaltar_boton_activo(nombre_activo):
    for nombre, boton in botones_menu.items():
        boton.configure(fg_color="#004080" if nombre == nombre_activo else "#1f6aa5")

# ========== Botones visibles siempre ==========
btn_ingreso = ctk.CTkButton(
    menu, text="Ingreso / Salida",
    command=lambda: [mostrar_ingreso_salida(), resaltar_boton_activo("ingreso")]
)
btn_ingreso.pack(side="left", padx=5)
botones_menu["ingreso"] = btn_ingreso

btn_solicitud = ctk.CTkButton(
    menu, text="Solicitar Permiso",
    command=lambda: [
        construir_solicitudes(
            frame_contenedor,
            on_volver=lambda: [mostrar_ingreso_salida(), resaltar_boton_activo("ingreso")]
        ),
        resaltar_boton_activo("solicitud")
    ]
)
btn_solicitud.pack(side="left", padx=5)
botones_menu["solicitud"] = btn_solicitud

# ========== Botones solo admin ==========
btn_registro = ctk.CTkButton(menu, text="Agregar Nuevo Usuario",
                             command=lambda: [mostrar_registro_con_refresco(), resaltar_boton_activo("registro")])
btn_reportes = ctk.CTkButton(menu, text="Reportes",
                             command=lambda: [mostrar_reportes(), resaltar_boton_activo("reportes")])
btn_nomina = ctk.CTkButton(menu, text="Nómina ICP",
                           command=lambda: [mostrar_nomina(), resaltar_boton_activo("nomina")])
btn_editar = ctk.CTkButton(menu, text="Editar Usuario",
                           command=lambda: [mostrar_editar_usuario(), resaltar_boton_activo("editar")])
btn_dia_admin = ctk.CTkButton(menu, text="Administrativos/Permisos",
                              command=lambda: [limpiar_frame(), construir_dia_administrativo(frame_contenedor), resaltar_boton_activo("diaadmin")])
botones_menu["diaadmin"] = btn_dia_admin

btn_clave = ctk.CTkButton(menu, text="Cambiar Clave Administrador", command=abrir_cambio_clave)
btn_panel_avanzado = ctk.CTkButton(menu, text="Panel Avanzado",
                                   command=lambda: [limpiar_frame(), construir_panel_avanzado(frame_contenedor), resaltar_boton_activo("panelavanzado")])
botones_menu["panelavanzado"] = btn_panel_avanzado

# Ocultar inicialmente los de admin
for btn in [btn_registro, btn_reportes, btn_nomina, btn_editar, btn_dia_admin, btn_panel_avanzado, btn_clave]:
    btn.pack_forget()

btn_logout_admin = ctk.CTkButton(menu, text="Cerrar Sesión Admin", fg_color="gray",
                                 command=lambda: cerrar_sesion_admin())
btn_logout_admin.pack_forget()

def mostrar_opciones_admin():
    btn_registro.pack(side="left", padx=5); botones_menu["registro"] = btn_registro
    btn_reportes.pack(side="left", padx=5); botones_menu["reportes"] = btn_reportes
    btn_nomina.pack(side="left", padx=5);   botones_menu["nomina"]   = btn_nomina
    btn_editar.pack(side="left", padx=5);   botones_menu["editar"]   = btn_editar
    btn_dia_admin.pack(side="left", padx=5)
    btn_panel_avanzado.pack(side="left", padx=5)
    btn_clave.pack(side="left", padx=5)
    btn_logout_admin.pack(side="right", padx=5)

def cerrar_sesion_admin():
    global admin_info
    admin_info = None
    label_admin_activo.configure(text="")
    for btn in [btn_registro, btn_reportes, btn_nomina, btn_editar, btn_dia_admin, btn_clave, btn_panel_avanzado]:
        btn.pack_forget()
    btn_logout_admin.pack_forget()
    btn_admin.pack(side="left", padx=5)
    messagebox.showinfo("Cierre de sesión", "Has cerrado sesión de administrador.")

def abrir_login_admin():
    """Diálogo moderno centrado para login de administrador con cierre seguro."""
    TopLevelCls = getattr(ctk, "CTkToplevel", None) or tk.Toplevel
    login = TopLevelCls(app)
    login.title("Acceso Administrativo")
    try:
        login.resizable(False, False)
        login.transient(app)
        login.grab_set()
    except Exception:
        pass

    usar_ctk = TopLevelCls is not tk.Toplevel
    focus_after_id = None

    def safe_close_login():
        nonlocal focus_after_id
        try:
            if focus_after_id:
                login.after_cancel(focus_after_id)
        except Exception:
            pass
        safe_focus(app)
        try:
            login.withdraw()
        except Exception:
            pass
        login.after(200, lambda: login.winfo_exists() and login.destroy())

    if usar_ctk:
        cont = ctk.CTkFrame(login, corner_radius=12)
        cont.pack(fill="both", expand=True, padx=16, pady=16)

        ctk.CTkLabel(cont, text="Ingreso de Administrador",
                     font=("Arial", 18, "bold")).grid(row=0, column=0, columnspan=2, pady=(4, 12))

        ctk.CTkLabel(cont, text="RUT").grid(row=1, column=0, sticky="w", padx=(4, 8))
        entry_rut = ctk.CTkEntry(cont, placeholder_text="12.345.678-9", width=220)
        entry_rut.grid(row=1, column=1, sticky="ew", pady=6)

        ctk.CTkLabel(cont, text="Clave").grid(row=2, column=0, sticky="w", padx=(4, 8))
        entry_clave = ctk.CTkEntry(cont, placeholder_text="********", show="*", width=220)
        entry_clave.grid(row=2, column=1, sticky="ew", pady=6)

        ver_var = tk.BooleanVar(value=False)
        def toggle_ver():
            entry_clave.configure(show="" if ver_var.get() else "*")
        ctk.CTkCheckBox(cont, text="Mostrar", variable=ver_var, command=toggle_ver)\
            .grid(row=3, column=1, sticky="w", pady=(0, 8))

        def validar_admin():
            rut = entry_rut.get().strip()
            clave = entry_clave.get().strip()
            con = sqlite3.connect(DB_PATH)
            cur = con.cursor()
            cur.execute("SELECT * FROM admins WHERE rut = ? AND clave = ?", (rut, clave))
            admin = cur.fetchone()
            con.close()
            if admin:
                global admin_info
                admin_info = {"nombre": admin[1], "permiso": "Administrador"}
                mostrar_bienvenida(admin[1])
                btn_admin.pack_forget()
                mostrar_opciones_admin()
                mostrar_info_admin()
                safe_close_login()
            else:
                messagebox.showerror("Error", "Credenciales incorrectas")

        ctk.CTkButton(cont, text="Cancelar", fg_color="gray", command=safe_close_login)\
            .grid(row=4, column=0, pady=(8, 0), padx=4, sticky="ew")
        ctk.CTkButton(cont, text="Entrar", command=validar_admin)\
            .grid(row=4, column=1, pady=(8, 0), padx=4, sticky="ew")

        cont.grid_columnconfigure(1, weight=1)
        focus_after_id = login.after(80, lambda: safe_focus(entry_rut))

        login.bind("<Return>", lambda _: validar_admin())
        login.bind("<Escape>", lambda _: safe_close_login())
        try:
            login.protocol("WM_DELETE_WINDOW", safe_close_login)
        except Exception:
            pass

    else:
        marco = tk.Frame(login, padx=14, pady=14)
        marco.pack(fill="both", expand=True)

        tk.Label(marco, text="Ingreso de Administrador", font=("Arial", 14, "bold"))\
            .grid(row=0, column=0, columnspan=2, pady=(0, 10))

        tk.Label(marco, text="RUT").grid(row=1, column=0, sticky="e", padx=(0, 8))
        entry_rut = tk.Entry(marco, width=28); entry_rut.grid(row=1, column=1, pady=4)

        tk.Label(marco, text="Clave").grid(row=2, column=0, sticky="e", padx=(0, 8))
        entry_clave = tk.Entry(marco, show="*", width=28); entry_clave.grid(row=2, column=1, pady=4)

        def validar_admin():
            rut = entry_rut.get().strip()
            clave = entry_clave.get().strip()
            con = sqlite3.connect(DB_PATH)
            cur = con.cursor()
            cur.execute("SELECT * FROM admins WHERE rut = ? AND clave = ?", (rut, clave))
            admin = cur.fetchone()
            con.close()
            if admin:
                global admin_info
                admin_info = {"nombre": admin[1], "permiso": "Administrador"}
                mostrar_bienvenida(admin[1])
                btn_admin.pack_forget()
                mostrar_opciones_admin()
                mostrar_info_admin()
                safe_close_login()
            else:
                messagebox.showerror("Error", "Credenciales incorrectas")

        tk.Button(marco, text="Cancelar", command=safe_close_login)\
            .grid(row=3, column=0, pady=(10, 0), sticky="ew")
        tk.Button(marco, text="Entrar", command=validar_admin)\
            .grid(row=3, column=1, pady=(10, 0), sticky="ew")

        marco.grid_columnconfigure(1, weight=1)
        focus_after_id = login.after(80, lambda: safe_focus(entry_rut))
        login.bind("<Return>", lambda _: validar_admin())
        login.bind("<Escape>", lambda _: safe_close_login())
        try:
            login.protocol("WM_DELETE_WINDOW", safe_close_login)
        except Exception:
            pass

    # Centrar el login
    app.update_idletasks()
    try:
        w = max(360, login.winfo_reqwidth() + 32)
        h = max(240, login.winfo_reqheight() + 16)
    except Exception:
        w, h = 360, 240
    x = app.winfo_x() + (app.winfo_width() // 2) - (w // 2)
    y = app.winfo_y() + (app.winfo_height() // 2) - (h // 2)
    login.geometry(f"{w}x{h}+{max(x,0)}+{max(y,0)}")

# Botón Admin (visible siempre al inicio)
btn_admin = ctk.CTkButton(menu, text="🔒 Admin", command=abrir_login_admin)
btn_admin.pack(side="left", padx=5)

# Botón Salir (lado derecho)
btn_salir = ctk.CTkButton(menu, text="Salir", command=app.destroy)
btn_salir.pack(side="right", padx=5)

# Botón Resumen Diario (visible para todos)
btn_resumen = ctk.CTkButton(
    menu, text="Resumen Diario",
    command=lambda: [limpiar_frame(), construir_resumen_dia(frame_contenedor), resaltar_boton_activo("resumen")]
)
btn_resumen.pack(side="left", padx=5)
botones_menu["resumen"] = btn_resumen

# ========== Inicio por defecto ==========
try:
    resaltar_boton_activo("ingreso")
    mostrar_ingreso_salida()
except Exception as e:
    traceback.print_exc()
    messagebox.showerror("Error al cargar Ingreso/Salida", str(e))

app.mainloop()
