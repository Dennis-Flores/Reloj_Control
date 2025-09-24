# db.py
import os
import sqlite3

def crear_bd(db_path: str = "reloj_control.db") -> None:
    # Asegura la carpeta del archivo
    base_dir = os.path.dirname(os.path.abspath(db_path))
    os.makedirs(base_dir or ".", exist_ok=True)

    con = sqlite3.connect(db_path)
    cur = con.cursor()

    # ---------- TRABAJADORES ----------
    cur.execute("""
        CREATE TABLE IF NOT EXISTS trabajadores (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT NOT NULL,
            apellido TEXT NOT NULL,
            rut TEXT NOT NULL UNIQUE,
            profesion TEXT,
            correo TEXT,
            cumpleanos TEXT
        )
    """)
    # Columna extra usada en algunas pantallas
    cur.execute("PRAGMA table_info(trabajadores)")
    cols_trab = [c[1] for c in cur.fetchall()]
    if "verificacion_facial" not in cols_trab:
        cur.execute("ALTER TABLE trabajadores ADD COLUMN verificacion_facial TEXT")

    # ---------- REGISTROS (nuevo esquema) ----------
    # Si ya existe con las columnas nuevas, solo garantizamos índices.
    cur.execute("PRAGMA table_info(registros)")
    cols_reg = [c[1] for c in cur.fetchall()]

    def create_registros_new_table():
        cur.execute("""
            CREATE TABLE IF NOT EXISTS registros (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                rut TEXT NOT NULL,
                nombre TEXT,
                fecha TEXT,
                hora_ingreso TEXT,
                hora_salida TEXT,
                observacion TEXT
            )
        """)
        cur.execute("CREATE INDEX IF NOT EXISTS idx_registros_rut_fecha ON registros(rut, fecha)")

    if not cols_reg:
        # No existía la tabla: crear directamente con el esquema nuevo
        create_registros_new_table()
    elif "hora_ingreso" in cols_reg and "hora_salida" in cols_reg:
        # Ya está migrada: asegurar índice
        cur.execute("CREATE INDEX IF NOT EXISTS idx_registros_rut_fecha ON registros(rut, fecha)")
    else:
        # Esquema viejo: registros(rut, nombre, fecha, hora, tipo, observacion)
        # => Migramos a un esquema por día con hora_ingreso/hora_salida
        cur.execute("""
            CREATE TABLE IF NOT EXISTS registros_nuevo (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                rut TEXT NOT NULL,
                nombre TEXT,
                fecha TEXT,
                hora_ingreso TEXT,
                hora_salida TEXT,
                observacion TEXT
            )
        """)
        # Insertar agregando por (rut, fecha). Tomamos:
        # - nombre: el máximo (cualquiera consistente)
        # - hora_ingreso: MIN(hora) donde tipo='ingreso'
        # - hora_salida: MAX(hora) donde tipo='salida'
        # - observacion: la última no nula (usamos MAX como aproximación)
        cur.execute("""
            INSERT INTO registros_nuevo (rut, nombre, fecha, hora_ingreso, hora_salida, observacion)
            SELECT  r.rut,
                    MAX(r.nombre),
                    r.fecha,
                    MIN(CASE WHEN LOWER(r.tipo)='ingreso' THEN r.hora END),
                    MAX(CASE WHEN LOWER(r.tipo)='salida'  THEN r.hora END),
                    MAX(COALESCE(r.observacion, ''))
            FROM registros AS r
            GROUP BY r.rut, r.fecha
        """)
        # Reemplazar tabla
        cur.execute("DROP TABLE registros")
        cur.execute("ALTER TABLE registros_nuevo RENAME TO registros")
        cur.execute("CREATE INDEX IF NOT EXISTS idx_registros_rut_fecha ON registros(rut, fecha)")

    # ---------- HORARIOS ----------
    cur.execute("""
        CREATE TABLE IF NOT EXISTS horarios (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            rut TEXT NOT NULL,
            dia TEXT NOT NULL,
            hora_entrada TEXT,
            hora_salida TEXT
        )
    """)
    cur.execute("PRAGMA table_info(horarios)")
    cols_hor = [c[1] for c in cur.fetchall()]
    if "turno" not in cols_hor:
        cur.execute("ALTER TABLE horarios ADD COLUMN turno TEXT DEFAULT 'general'")

    # ---------- ADMINS ----------
    cur.execute("""
        CREATE TABLE IF NOT EXISTS admins (
            rut TEXT PRIMARY KEY,
            nombre TEXT NOT NULL,
            clave TEXT NOT NULL
        )
    """)
    cur.execute("PRAGMA table_info(admins)")
    cols_admins = [c[1] for c in cur.fetchall()]
    if "tipo_permiso" not in cols_admins:
        cur.execute("ALTER TABLE admins ADD COLUMN tipo_permiso TEXT DEFAULT 'Administrador'")
        cur.execute("UPDATE admins SET tipo_permiso='Administrador' WHERE tipo_permiso IS NULL")

    # Admin inicial (ajusta si quieres)
    cur.execute("SELECT 1 FROM admins WHERE rut=?", ("16632174-3",))
    if not cur.fetchone():
        cur.execute("""
            INSERT INTO admins (rut, nombre, clave, tipo_permiso)
            VALUES (?, ?, ?, ?)
        """, ("16632174-3", "Administrador General", "admin123", "Administrador"))

    # ---------- DIAS LIBRES ----------
    cur.execute("""
        CREATE TABLE IF NOT EXISTS dias_libres (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            rut TEXT NOT NULL,
            fecha TEXT NOT NULL,
            motivo TEXT DEFAULT 'Día administrativo autorizado',
            anio INTEGER
        )
    """)

    # ---------- PANEL FLAGS (lo usa ingreso_salida para salida anticipada) ----------
    cur.execute("""
        CREATE TABLE IF NOT EXISTS panel_flags (
            fecha TEXT PRIMARY KEY,
            salida_anticipada INTEGER DEFAULT 0,
            salida_anticipada_obs TEXT
        )
    """)

    # ---------- SOLICITUDES (para módulo Solicitar Permiso) ----------
    cur.execute("""
        CREATE TABLE IF NOT EXISTS solicitudes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            rut TEXT,
            tipo TEXT,                 -- ej: "permiso", "administrativo"
            fecha_solicitud TEXT,
            fecha_desde TEXT,
            fecha_hasta TEXT,
            estado TEXT,               -- ej: "pendiente", "aprobado", "rechazado"
            motivo TEXT,
            archivo_pdf TEXT
        )
    """)

    con.commit()
    con.close()
