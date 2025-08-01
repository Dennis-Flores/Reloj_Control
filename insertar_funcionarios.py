import sqlite3

conexion = sqlite3.connect("reloj_control.db")
cursor = conexion.cursor()

funcionarios = [
    ("YANETT SILVANA", "ALLENDE MILLAQUÉN", "12029647-7", "EDUCADORA PIE", "yanette.allende@slepllanquihue.cl", "06-12-1980"),
    ("CRISTINA YOCELYNN", "ANTILEF GONZÁLEZ", "18231824-8", "EDUCADORA PIE", "cristina.antilef@slepllanquihue.cl", "13-08-1992"),
    ("IVONNE MACARENA", "AÑAZCO GALINDO", "17241257-2", "DOCENTE PIE", "ivonne.anazco@slepllanquihue.cl", "21-02-1990"),
    ('CAROLINA PAZ ARAYA', 'SAAVEDRA', '13687406-3', 'DOCENTE ARTES', 'carolina.araya@slepllanquihue.cl', '1979-06-02'),
    ('VIVIANA BAUDELIA BASTIDAS', 'ÁLVAREZ', '14095437-3', 'DOCENTE INGLES', 'viviana.bastidas@slepllanquihue.cl', '1981-03-29'),
    ('PAOLA PILAR CARCAMO', 'QUIROZ', '14245774-1', 'DOCENTE UTP', 'paola.carcamo@slepllanquihue.cl', '1974-03-03'),
    ('FLOR KARINA CARRASCO', 'OJEDA', '15305911-k', 'DOCENTE MATEMÁTICA', 'flor.carrasco@slepllanquihue.cl', '1982-07-26'),
    ('ALEJANDRA PILAR CASANOVA', 'BARRÍA', '14517070-2', 'DOCENTE LENGUAJE', 'alejandra.casanova@slepllanquihue.cl', '1974-08-21'),
    ('LORENA PAOLA CHEUQUEMÁN', 'VELÁSQUEZ', '15687483-3', 'DOCENTE LENGUAJE', 'lorena.velasquez@slepllanquihue.cl', '1983-02-11'),
    ('VALENTINA MARGARITA COFRÉ', 'JARA', '17584651-4', 'DOCENTE MATEMÁTICA', 'valentina.cofre@slepllanquihue.cl', '1990-12-21'),
    ('CATHERINE DONOSO', 'QUEZADA', '17341810-8', 'DOCENTE DE CIENCIA', 'catherine.donoso@slepllanquihue.cl', '1990-07-01'),
    ('ALEJANDRA ROSSANA DUARTE', 'ALMONACID', '13448766-6', 'CONVIVENCIA ESCOLAR', 'alejandra.duarte@slepllanquihue.cl', '1978-01-14'),
    ('ROXANA VIVIANA ESPINOSA', 'BRAVO', '13094383-7', 'DOCENTE CIENCIAS', 'roxana.espinosa@slepllanquihue.cl', '1976-09-24'),
    ('LUIS ANDRÉS FURRIANCA', 'ALVARADO', '17985545-3', 'DOCENTE - AULA DE REINGRESO', 'luis.furrianca@slepllanquihue.cl', '1991-08-27'),
    ('CRISTIAN ALEJANDRO GÓMEZ', 'CÁRCAMO', '13166192-4', 'DOCENTE ADMINISTRACIÓN', 'cristian.gomez@slepllanquihue.cl', '1977-05-23'),
    ('MÓNICA GONZÁLEZ', 'LIZAMA', '12917030-1', 'EPJA - LENGUAJE', 'monica.gonzalez@slepllanquihue.cl', '1975-01-07'),
    ('ANGÉLICA SOLEDAD GONZÁLEZ', 'LEVICÁN', '9144301-5', 'EPJA - CIENCIAS', 'angelica.gonzalez@slepllanquihue.cl', 'NULL'),
    ('MARLEN MARISOL HERNANDEZ', 'ALMONACID', '17604769-0', 'EPJA - HISTORIA', 'marlen.hernandez@slepllanquihue.cl', '1990-09-30'),
    ('FABIOLA CAROLINA HERNÁNDEZ', 'CRUCES', '14440767-9', 'DOCENTE FILOSOFIA', 'fabiola.hernández@slepllanquihue.cl', '1974-10-17'),
    ('GERALD HERRERA', 'LIMARÍ', '13246940-7', 'DOCENTE MATEMÁTICA', 'gerald.herrera@slepllanquihue.cl', '1977-01-30'),
    ('DANIELA WANGULEN HUAIQUILAO', 'HUENCHUFIL', '18726930-k', 'EDUCADORA PIE', 'daniela.huaiquilao@slepllanquihue.cl', '1993-12-27'),
    ('KAMILA FERNANDA LARA', 'ZAMBRANO', '18501447-9', 'EDUCADORA PIE', 'kamila.lara@slepllanquihue.cl', '1993-06-07'),
    ('LORETO CRISTINA MONSALVE', 'NÚÑEZ', '16589196-1', 'COORDINADORA PIE', 'loreto.monsalve@slepllanquihue.cl', '1987-03-09'),
    ('PAULETTE ROSSANNA MONTES', 'VÁSQUEZ', '14615439-5', 'DOCENTE LENGUAJE', 'paulette.montes@slepllanquihue.cl', '1981-01-30'),
    ('JUAN CLAUDIO MORALES', 'MARTÍNEZ', '12791157-6', 'DOCENTE ADMINISTRACIÓN', 'juan.morales@slepllanquihue.cl', '1975-10-29'),
    ('ROCÍO FERNANDA MORALES', 'NANNIG', '17301668-9', 'DOCENTE HISTORIA', 'rocio.morales@slepllanquihue.cl', '1989-08-18'),
    ('RICARDO ALEXIS NAVARRO', 'SANHUEZA', '17632447-3', 'DOCENTE EDUCACIÓN FÍSICA', 'ricardo.navarro@slepllanquihue.cl', '1991-02-16'),
    ('NATALIA BEATRIZ NEGRÓN', 'JARAMILLO', '15844179-9', 'JEFA UTP', 'natalia.negron@slepllanquihue.cl', '1985-01-01'),
    ('MARCIA ELENA PIZARRO', 'HERNÁNDEZ', '9972395-5', 'DOCENTE RELIGIÓN CATÓLICA', 'marcia.pizarro@slepllanquihue.cl', '1964-10-23'),
    ('ANA KAREN QUINTUL', 'GONZÁLEZ', '17641850-8', 'EDUCADORA PIE', 'ana.quintul@slepllanquihue.cl', '1990-11-01'),
    ('MAURICIO FERNANDO REGENTE', 'AYALA', '10063936-k', 'DOCENTE HISTORIA', 'mauricio.regente@slepllanquihue.cl', '1967-06-03'),
    ('KAREN VANESSA ROJAS', 'AZOCAR', '16551947-7', 'EDUCADORA PIE', 'karen.rojas@slepllanquihue.cl', '1987-11-09'),
    ('DIEGO ALEXIS RUIZ', 'GALLARDO', '17299246-3', 'DOCENTE RELIGIÓN EVANGÉLICA', 'diego.ruiz@slepllanquihue.cl', '1988-12-26'),
    ('ANDRES SALOMON SANCHEZ VERA', 'JORGE', '20112746-7', 'DOCENTE MÚSICA', 'jorge.sanchez@slepllanquihue.cl', '1999-11-04'),
    ('JORDAN HORACIO SOLIS', 'VILLARROEL', '18708279-k', 'DOCENTE CRA', 'jordan.solis@slepllanquihue.cl', '1994-08-30'),
    ('CAROLINA SOLEDAD SOLÍS', 'AGUILEF', '17219688-8', 'DOCENTE GASTRONOMÍA', 'carolina.solis@slepllanquihue.cl', '1989-06-27'),
    ('PAMELA JUDITZA SOTO', 'VELÁSQUEZ', '16551805-5', 'DOCENTE INGLÉS', 'pamela.soto@slepllanquihue.cl', '1987-02-03'),
    ('KARLA YANINNE TORREALBA', 'YEFI', '18368959-2', 'EDUCADORA PIE', 'karla.torrealba@slepllanquihue.cl', '1994-11-10'),
    ('GABRIELA PAZ VARGAS', 'BARRIENTOS', '18819673-K', 'EDUCADORA PIE', 'gabriela.vargas@slepllanquihue.cl', '1995-10-08'),
    ('LUIS FERNANDO VARGAS', 'AGUILAR', '17604739-9', 'DOCENTE GASTRONOMÍA', 'luis.vargas@slepllanquihue.cl', '1990-07-31'),
    ('MARCELO PATRICIO VARGAS', 'GALLARDO', '16523838-9', 'INSPECTOR GENERAL', 'marcelo.vargas@slepllanquihue.cl', '1987-01-16'),
    ('FABIANA ELIZABETH VELÁSQUEZ', 'SALDIVIA', '20272546-5', 'DOCENTE MATEMÁTICA', 'fabiana.velasquez@slepllanquihue.cl', '1999-09-28'),
    ('DEL CARMEN VILLARRUEL CÓRDOVA', 'FIDELIA', '9442883-1', 'DOCENTE EDUCACIÓN FÍSICA', 'fidelia.villarruel@slepllanquihue.cl', '1964-04-03'),
    ('BERNABE JOSUE GONZALEZ', 'LEIVA', '20979939-1', 'Reemplazo DOCENTE GASTRONOMÍA', 'bernabe.gonzalez@slepllanquihue.cl', '2002-04-27'),
    ('DEL PILAR ZÚÑIGA SILVA', 'MARCELA', '16727401-3', 'Reemplazo DOCENTE DE ARTES', 'marcela.zuniga@slepllanquihue.cl', '1987-11-27'),
    ('CARMEN JANNETH SEPÚLVEDA', 'OSSES', '11214726-8', 'Reemplazo COORDINADORA EPJA + docencia', 'carmen.sepulveda@slepllanquihue.cl', '1977-06-09'),
    ('DANIELA NATALIA MUÑOZ', 'MOLINA', '19609236-6', 'Reemplazo EDUCADORA DIFERENCIAL', 'daniela.munoz.m@slepllanquihue.cl', '1998-06-18'),
    ('JAVIERA CASTILLO', 'SANCHEZ', '20491666-7', 'Reemplazo MARLEN HERNANDEZ', 'javiera.castillo@slepllanquihue.cl', 'NULL'),
    ('PATRICIA JEANETTE ALBARRACÍN', 'IGOR', '12756019-6', 'ASISTENTE SOCIAL', 'patricia.albarracin@slepllanquihue.cl', '1975-03-15'),
    ('PAOLA BELÉN ALVARADO', 'ANGEL', '15961712-2', 'PSICOLOGA', 'paola.alvaradoa@slepllanquihue.cl', '1984-10-23'),
    ('ORIANA PAMELA ARRIAGADA', 'VILLARROEL', '15282140-9', 'SECRETARIA DIRECCIÓN', 'oriana.arriagada@slepllanquihue.cl', '1982-02-09'),
    ('DEL CARMEN AZOCAR ARRIAGADA', 'ROSA', '10445433-k', 'ADMINISTRATIVA', 'rosa.azocar@slepllanquihue.cl', '1965-08-30'),
    ('ANGÉLICA BEATRIZ BARRÍA', 'OJEDA', '11925135-4', 'INSPECTORA', 'angelica.barria@slepllanquihue.cl', '1972-06-20'),
    ('GUISELLA MARGARITA BOLLMANN', 'VIDAL', '14038388-0', 'BIBLIOTECARIA', 'guisella.bollmann@slepllanquihue.cl', '1981-02-17'),
    ('JAIME HUBERTO CÁRDENAS', 'BARRIENTOS', '12141944-0', 'AUXILIAR', 'jaime.cardenas@slepllanquihue.cl', 'NULL'),
    ('JOSÉ LUIS ELGUETA', 'SOTO', '11712127-5', 'INSPECTOR', 'jose.elgueta@slepllanquihue.cl', '1971-04-15'),
    ('BRUNO OMAR HERNÁNDEZ', 'CÁRCAMO', '8827369-4', 'INSPECTOR EPJA', 'bruno.hernandez@slepllanquihue.cl', '1963-08-26'),
    ('CARMEN GLORIA NAVARRO', 'SOTO', '11001548-8', 'AUXILIAR', 'carmen.navarro@slepllanquihue.cl', '1965-07-16'),
    ('JORGE ALEJANDRO OVANDO', 'ALTAMIRANO', '13166620-9', 'PSICOPEDAGOGO', 'jorge.ovando@slepllanquihue.cl', '1977-09-16'),
    ('DANIEL ALFREDO PÉREZ', 'NÚÑEZ', '14038886-6', 'INSPECTOR', 'daniel.perez@slepllanquihue.cl', '1981-05-16'),
    ('JAIME JESÚS SANCHEZ', 'TIZNADO', '19789996-4', 'PSICOLOGO PIE - 44 horas', 'jaime.sanchez@slepllanquihue.cl', '1997-08-20'),
    ('OSCAR HERNÁN SANTANA', 'DELGADO', '19027600-7', 'PSICOLOGO PIE - 22 horas', 'oscar.santana@slepllanquihue.cl', '1994-12-28'),
    ('YARENLLA SCARLETT SCHOLER', 'HERNÁNDEZ', '13823511-4', 'INSPECTORA', 'yarenlla.scholer@slepllanquihue.cl', '1980-01-14'),
    ('MARCELO ANDRÉS SOTO', 'ALMONACID', '12756841-3', 'AUXILIAR', 'marcelo.soto@slepllanquihue.cl', '1975-09-18'),
    ('ALEXIS JAVIER TOLEDO', 'DELGADO', '20624188-8', 'PSICOLOGO - AULA DE REINGRESO', 'alexis.toledo@slepllanquihue.cl', '2000-11-08'),
    ('CLAUDIA ALEJANDRA VARGAS', 'GONZÁLEZ', '12998842-8', 'AUXILIAR', 'claudia.vargasg@slepllanquihue.cl', '1976-01-26'),
    ('ROSA EDITH VELÁSQUEZ', 'SOTO', '14087113-3', 'AUXILIAR', 'rosa.velasquez@slepllanquihue.cl', '1976-03-15'),
    ('MARIANELA JOVITA NAVARRO', 'PEREIRA', '10449669-5', 'Reemplazo - Auxiliar', 'marianela.navarro@slepllanquihue.cl', '1967-10-07'),
    ('JHONSON ROJAS MICHELLE', 'KARY', '', 'MONITORA DANZA RECREO INTERACTIVO', '', 'NULL'),
    ('ESPAÑA CORNEJO', 'VALERIA', '', 'MATRONA', '', 'NULL'),
    ('SANTANA DELGADO OSCAR', 'ANDRÉS', '', 'PSICOLOGO - 22 horas', '', 'NULL'),
    ('GALLARDO BARRIA PATRICIO', 'DOMINGO', '', 'DOCENTE APOYO', '', 'NULL'),
    ('SANCHEZ VERA JORGE ANDRES', 'SALOMON', '', 'DOCENTE APOYO', '', 'NULL'),
    ('GONZALEZ DÍAZ', 'JAVIER', '', 'DOCENTE APOYO', '', 'NULL'),
    ('TRIVIÑO MIRANDA', 'GABRIELA', '', 'APOYO ADMINISTRATIVO', '', 'NULL'),
    ('ESPINOSA BRAVO ROXANA', 'VIVIANA', '', 'DOCENTE CIENCIAS', '', 'NULL'),
    ('GALLARDO BARRÍA PATRICIO', 'DOMINGO', '', 'COORDINADOR EPJA', '', 'NULL'),
    ('GAETE SALAZAR LORETO', 'IRENE', '', 'COORD. TÉCNICO PROFESIONAL', 'loreto.gaete@slepllanquihue.cl', 'NULL'),
    ('CASTILLO SANCHEZ JAVIERA', 'ALEJANDRA', '', 'Reemplazo DOCENTE CASANOVA', '', 'NULL'),
    ('TRIVIÑO MIRANDA GABRIELA DEL', 'CARMEN', '', 'Reemplazo Inspectora', 'gabriela.trivino@slepllanquihue.cl', 'NULL'),
    ('VILLANUEVA VILLA DIEGO', 'ALEXIS', '', 'DOCENTE CIENCIAS', '', 'NULL'),
    ('MOREIRA IMILMAQUI CLAUDIO', 'ALEXIS', '', 'Encargado Informática', 'claudio.moreira@slepllanquihue.cl', 'NULL'),
    ('MONTENEGRO ORTIZ CARLOS', 'EUSEBIO', '', 'INSPECTOR', 'carlos.montenegro@slepllanquihue.cl', 'NULL'),
    ('FERNÁNDEZ ELGUETA JACQUELINE', 'CATALINA', '', 'INSPECTORA', 'jacqueline.fernandez@slepllanquihue.cl', 'NULL'),

    # ... aquí agregamos el resto
]

for f in funcionarios:
    try:
        cursor.execute('''
            INSERT INTO trabajadores (nombre, apellido, rut, profesion, correo, cumpleanos)
            VALUES (?, ?, ?, ?, ?, ?)
        ''', f)
    except sqlite3.IntegrityError:
        print(f"⚠️ Usuario con RUT {f[2]} ya existe. Saltando...")

conexion.commit()
conexion.close()

print("✅ Inserción completada.")
