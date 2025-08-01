import tkinter as tk
import sqlite3
from tkinter import messagebox

def abrir_cambio_clave():
    ventana = tk.Toplevel()
    ventana.title("Cambiar Clave de Administrador")
    ventana.geometry("420x280")
    ventana.resizable(False, False)

    tk.Label(ventana, text="RUT:", font=("Arial", 11)).pack(pady=5)
    entry_rut = tk.Entry(ventana)
    entry_rut.pack()

    tk.Label(ventana, text="Clave Actual:", font=("Arial", 11)).pack(pady=5)
    entry_clave_actual = tk.Entry(ventana, show="*")
    entry_clave_actual.pack()

    tk.Label(ventana, text="Nueva Clave:", font=("Arial", 11)).pack(pady=5)
    entry_clave_nueva = tk.Entry(ventana, show="*")
    entry_clave_nueva.pack()

    tk.Label(ventana, text="Confirmar Nueva Clave:", font=("Arial", 11)).pack(pady=5)
    entry_confirmar = tk.Entry(ventana, show="*")
    entry_confirmar.pack()

    def cambiar():
        rut = entry_rut.get().strip()
        clave_actual = entry_clave_actual.get().strip()
        clave_nueva = entry_clave_nueva.get().strip()
        confirmar = entry_confirmar.get().strip()

        if clave_nueva != confirmar:
            messagebox.showerror("Error", "Las nuevas claves no coinciden.")
            return

        conexion = sqlite3.connect("reloj_control.db")
        cursor = conexion.cursor()
        cursor.execute("SELECT * FROM admins WHERE rut = ? AND clave = ?", (rut, clave_actual))
        admin = cursor.fetchone()

        if admin:
            cursor.execute("UPDATE admins SET clave = ? WHERE rut = ?", (clave_nueva, rut))
            conexion.commit()
            messagebox.showinfo("Éxito", "Clave actualizada correctamente.")
            ventana.destroy()
        else:
            messagebox.showerror("Error", "RUT o clave actual incorrectos.")

        conexion.close()

    # Frame para los botones
    frame_botones = tk.Frame(ventana)
    frame_botones.pack(pady=15)

    # Botón Actualizar Clave
    btn_actualizar = tk.Button(frame_botones, text="Actualizar Clave", command=cambiar)
    btn_actualizar.pack(side="left", padx=10)

    # Botón Cancelar
    btn_cancelar = tk.Button(frame_botones, text="Cancelar", command=ventana.destroy)
    btn_cancelar.pack(side="left", padx=10)
