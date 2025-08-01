import customtkinter as ctk
from db import crear_bd
import sqlite3
import tkinter as tk
from tkinter import messagebox

from cambio_clave_admin import abrir_cambio_clave
from dia_administrativo import construir_dia_administrativo

# ========== Inicializar base de datos ==========
crear_bd()

# ========== Configuración general ==========
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

admin_info = None  # Almacena los datos del administrador logueado

app = ctk.CTk()
app.geometry("1200x600")
app.title("BioAccess/Control de Horarios | www.bioaccess.cl")

# ========== Etiqueta contador ==========
label_contador = ctk.CTkLabel(app, text="", font=("Arial", 13), text_color="gray")
label_contador.pack(pady=(5, 10))

def actualizar_contador():
    conexion = sqlite3.connect("reloj_control.db")
    cursor = conexion.cursor()
    cursor.execute("SELECT COUNT(*) FROM trabajadores")
    total = cursor.fetchone()[0]
    conexion.close()
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

# ========== Funciones ==========
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
    construir_edicion(frame_contenedor, actualizar_contador)

# ========== Manejo de botones activos ==========
botones_menu = {}

def resaltar_boton_activo(nombre_activo):
    for nombre, boton in botones_menu.items():
        boton.configure(fg_color="#004080" if nombre == nombre_activo else "#1f6aa5")

# ========== Botones ==========
btn_ingreso = ctk.CTkButton(menu, text="Ingreso / Salida", command=lambda: [mostrar_ingreso_salida(), resaltar_boton_activo("ingreso")])
btn_ingreso.pack(side="left", padx=5)
botones_menu["ingreso"] = btn_ingreso

btn_registro = ctk.CTkButton(menu, text="Registrar Nuevo Usuario", command=lambda: [mostrar_registro_con_refresco(), resaltar_boton_activo("registro")])
btn_reportes = ctk.CTkButton(menu, text="Reportes", command=lambda: [mostrar_reportes(), resaltar_boton_activo("reportes")])
btn_nomina = ctk.CTkButton(menu, text="Nómina ICP", command=lambda: [mostrar_nomina(), resaltar_boton_activo("nomina")])
btn_editar = ctk.CTkButton(menu, text="Editar/Eliminar Usuario", command=lambda: [mostrar_editar_usuario(), resaltar_boton_activo("editar")])
btn_dia_admin = ctk.CTkButton(menu, text="Días Administrativos/Permisos", command=lambda: [limpiar_frame(), construir_dia_administrativo(frame_contenedor), resaltar_boton_activo("diaadmin")])
botones_menu["diaadmin"] = btn_dia_admin

btn_logout_admin = ctk.CTkButton(menu, text="Cerrar Sesión Admin", fg_color="gray", command=lambda: cerrar_sesion_admin())
btn_logout_admin.pack_forget()

for btn in [btn_registro, btn_reportes, btn_nomina, btn_editar, btn_dia_admin]:
    btn.pack_forget()

btn_clave = ctk.CTkButton(menu, text="Cambiar Clave Administrador", command=abrir_cambio_clave)
btn_clave.pack_forget()

def mostrar_opciones_admin():
    btn_registro.pack(side="left", padx=5)
    botones_menu["registro"] = btn_registro

    btn_reportes.pack(side="left", padx=5)
    botones_menu["reportes"] = btn_reportes

    btn_nomina.pack(side="left", padx=5)
    botones_menu["nomina"] = btn_nomina

    btn_editar.pack(side="left", padx=5)
    botones_menu["editar"] = btn_editar

    btn_dia_admin.pack(side="left", padx=5)
    btn_clave.pack(side="left", padx=5)
    btn_logout_admin.pack(side="right", padx=5)

def cerrar_sesion_admin():
    global admin_info
    admin_info = None
    label_admin_activo.configure(text="")
    for btn in [btn_registro, btn_reportes, btn_nomina, btn_editar, btn_dia_admin, btn_clave]:
        btn.pack_forget()
    btn_logout_admin.pack_forget()
    messagebox.showinfo("Cierre de sesión", "Has cerrado sesión de administrador.")

def abrir_login_admin():
    login = tk.Toplevel()
    login.title("Acceso Administrativo")
    login.geometry("300x180")
    login.resizable(False, False)

    tk.Label(login, text="RUT:", font=("Arial", 11)).pack(pady=5)
    entry_rut = tk.Entry(login)
    entry_rut.pack()

    tk.Label(login, text="Clave:", font=("Arial", 11)).pack(pady=5)
    entry_clave = tk.Entry(login, show="*")
    entry_clave.pack()

    def validar_admin():
        rut = entry_rut.get().strip()
        clave = entry_clave.get().strip()

        conexion = sqlite3.connect("reloj_control.db")
        cursor = conexion.cursor()
        cursor.execute("SELECT * FROM admins WHERE rut = ? AND clave = ?", (rut, clave))
        admin = cursor.fetchone()
        conexion.close()

        if admin:
            global admin_info
            admin_info = {"nombre": admin[1], "permiso": "Administrador"}  # Puedes ajustar esto si agregas un campo tipo_permiso
            login.destroy()
            messagebox.showinfo("Acceso permitido", f"Bienvenido: {admin[1]}")
            mostrar_opciones_admin()
            mostrar_info_admin()
        else:
            messagebox.showerror("Error", "Credenciales incorrectas")

    tk.Button(login, text="Entrar", command=validar_admin).pack(pady=10)

btn_admin = ctk.CTkButton(menu, text="🔒 Admin", command=abrir_login_admin)
btn_admin.pack(side="left", padx=5)

btn_salir = ctk.CTkButton(menu, text="Salir", command=app.destroy)
btn_salir.pack(side="right", padx=5)

resaltar_boton_activo("ingreso")
mostrar_ingreso_salida()

app.mainloop()
