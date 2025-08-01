import sqlite3

def crear_bd():
    conexion = sqlite3.connect("reloj_control.db")
    cursor = conexion.cursor()

    # Crear tabla trabajadores
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS trabajadores (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT NOT NULL,
            apellido TEXT NOT NULL,
            rut TEXT NOT NULL UNIQUE,
            profesion TEXT,
            correo TEXT,
            cumpleanos TEXT
        )
    ''')

    # Agregar columna 'verificacion_facial' si aún no está
    cursor.execute("PRAGMA table_info(trabajadores)")
    columnas_trabajadores = [col[1] for col in cursor.fetchall()]
    if "verificacion_facial" not in columnas_trabajadores:
        cursor.execute("ALTER TABLE trabajadores ADD COLUMN verificacion_facial TEXT")

    # Crear tabla registros
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS registros (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            rut TEXT NOT NULL,
            nombre TEXT,
            fecha TEXT,
            hora TEXT,
            tipo TEXT
        )
    ''')

    # Agregar campo 'observacion' si aún no está
    cursor.execute("PRAGMA table_info(registros)")
    columnas_registros = [col[1] for col in cursor.fetchall()]
    if "observacion" not in columnas_registros:
        cursor.execute("ALTER TABLE registros ADD COLUMN observacion TEXT")

    # Crear tabla horarios
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS horarios (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            rut TEXT NOT NULL,
            dia TEXT NOT NULL,
            hora_entrada TEXT,
            hora_salida TEXT
        )
    ''')

    # Agregar columna 'turno' si aún no está
    cursor.execute("PRAGMA table_info(horarios)")
    columnas_horarios = [col[1] for col in cursor.fetchall()]
    if "turno" not in columnas_horarios:
        cursor.execute("ALTER TABLE horarios ADD COLUMN turno TEXT DEFAULT 'general'")

    # Crear tabla admins
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS admins (
            rut TEXT PRIMARY KEY,
            nombre TEXT NOT NULL,
            clave TEXT NOT NULL
        )
    ''')

    # Agregar columna 'tipo_permiso' si no está
    cursor.execute("PRAGMA table_info(admins)")
    columnas_admins = [col[1] for col in cursor.fetchall()]
    if "tipo_permiso" not in columnas_admins:
        cursor.execute("ALTER TABLE admins ADD COLUMN tipo_permiso TEXT DEFAULT 'Administrador'")
        cursor.execute("UPDATE admins SET tipo_permiso = 'Administrador' WHERE tipo_permiso IS NULL")

    # Insertar administrador inicial si no existe
    cursor.execute("SELECT * FROM admins WHERE rut = ?", ("16632174-3",))
    if not cursor.fetchone():
        cursor.execute("INSERT INTO admins (rut, nombre, clave, tipo_permiso) VALUES (?, ?, ?, ?)", (
            "16632174-3", "Administrador General", "admin123", "Administrador"
        ))

    # Crear tabla dias_libres
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS dias_libres (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            rut TEXT NOT NULL,
            fecha TEXT NOT NULL,
            motivo TEXT DEFAULT 'Día administrativo autorizado',
            anio INTEGER
        )
    ''')

    conexion.commit()
    conexion.close()
