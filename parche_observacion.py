import sqlite3

conexion = sqlite3.connect("reloj_control.db")
cursor = conexion.cursor()

# Agregar columna 'observacion' si no existe
try:
    cursor.execute("ALTER TABLE registros ADD COLUMN observacion TEXT")
    print("✅ Columna 'observacion' agregada con éxito.")
except sqlite3.OperationalError:
    print("⚠️ La columna 'observacion' ya existe.")

conexion.commit()
conexion.close()
