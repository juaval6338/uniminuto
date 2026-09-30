# Registro de Asistencia — Bienestar Institucional (UNIMINUTO)

Libro para registrar la asistencia a las actividades de Bienestar. Se escribe la
cédula o el ID del participante y el archivo trae sus datos desde las bases.

- `libro/Registro_Asistencia_2026-1_V23.xlsx` — el libro listo para usar.
- `herramientas/` — scripts que generan el libro (`build9.py`), las tablas
  dinámicas del buscador (`pivots.py`), el conversor de cadenas compartidas
  (`sharedstr.py`), las fórmulas compartidas (`sharedfml.py`), el validador
  (`validar.py`), los lectores del `.xlsb` original y las pruebas con LibreOffice
  (`pruebas/`).

## Punto de partida

La versión 17 se construye **sobre el archivo original V9**: mismas hojas en el
mismo orden, mismos colores y fuentes, mismas letras de columna y la tabla con
encabezados en la fila 18 y datos desde la fila 19. Solo cambian las fórmulas y lo
que se pidió expresamente.

## Hoja Registros

| Columnas | Quién la llena | Contenido |
|---|---|---|
| E, F | usted (sin relleno) | Fecha y documento (C.C. o ID) |
| G – K | el archivo (gris) | Sede, Dependencia, ID, Apellidos y nombres, Correo institucional |
| L – O | usted, solo si sale INEXISTENTE | Correo, Tipo, Sede, Dependencia |
| P | el archivo | Tipo de participante: ESTUDIANTE, PROFESOR, ADMINISTRATIVO o EXTERNO |
| Q | usted | Nombre de espacio y observaciones |
| R – U | el archivo | Teléfono, teléfono adicional, correo adicional, número de cédula |

- Participantes: una persona cuenta una sola vez aunque en una fila se escriba su
  cédula y en otra su ID. Solo cuentan las filas completas (fecha y documento),
  igual que las participaciones. En los datos de ejemplo había 29 personas
  escritas de las dos formas: eran 93 participantes, no 121.
- Los códigos 1, 27 y 28 siguen en las bases; en Registros se muestran como texto.
  Quien está en BD EST sale como ESTUDIANTE aunque su fila no traiga el código 1;
  en colaboradores, si falta el código 27/28 se usa la columna J (DOC/ADM).
- ID repetido: la letra se pone en azul fuerte (regla de duplicados del original).
- Persona que no está en la base: el nombre dice INEXISTENTE en negrita y la fila
  se pinta de naranja (Énfasis 6 del tema Office, el mismo tema del archivo); el
  correo que se escriba en L aparece en K.
- La columna U muestra siempre la cédula, aunque en F se haya escrito el ID.
- El documento se reconoce aunque venga pegado como texto, con puntos o con
  espacios (también el espacio invisible que traen las páginas web y Forms).
- Buscador (filas 13 a 15) con **tablas dinámicas**, como en el original: el
  filtro de J14 (estudiantes), J15 (colaboradores) y L14 (correo de estudiante)
  trae los nombres en orden alfabético, sin repetidos y con cuadro «Buscar». Las
  tablas dinámicas están en I16, I17 y K16 (las mismas posiciones del original)
  y leen columnas completas de las bases (`'BD EST'!C:C`, `'BD EST'!Y:Y`,
  `'BD ADM-DOC'!D:D`). Al elegir un nombre, E14:H15 muestran sede, ID,
  dependencia y cédula.
- Resumen de participación arriba a la derecha (L1:P8) con filtro de sede en M2.
  El original ocultaba las filas 1 a 6; ahora se ven para que el cuadro aparezca.
- 10.000 filas de registro (19 a 10018). Q-Part, Día y Mes van ocultas.
- Los encabezados quedan inmovilizados hasta la fila 18, como en el original.

## Otras hojas

- **Resultados por tipo**: eliminada; su cuadro está en Registros.
- **Resultados por programa**: tres bloques horizontales (estudiantes,
  profesores, administrativos) ordenados de mayor a menor, con 100 filas por
  bloque (las bases tienen 66, 58 y 70 dependencias; antes cabían 30).
- **Evaluación**: la tabla de la derecha refleja también lo que se pega, y la
  escala sigue el formato FR-BM-DFB-03 (E, N, A, N/M, N/A).
- **Gráficos Evaluación**: una sola gráfica de barras en lugar de cuatro.

## Mantenimiento

Las bases conservan las columnas del export, así que se actualizan pegando encima
bajo el encabezado de la fila 1, o agregando filas al final. Las fórmulas apuntan
a columnas completas, sin tope de filas. La cédula y el ID se encuentran aunque
vengan pegados como texto (también el ID con ceros a la izquierda).

## Protección

Registros, Resultados por programa, Evaluación y Gráficos Evaluación están
protegidas **sin contraseña**, con las mismas opciones del original (se puede usar
el autofiltro y el filtro de las tablas dinámicas). Solo quedan libres:

- Registros: E (fecha), F (C.C. o ID), L a O (datos del INEXISTENTE),
  Q (observaciones), H8:K8 (nombre de la actividad) y M2 (sede del resumen).
- Evaluación: B10:K509 (respuestas).

Las bases (BD EST y BD ADM-DOC) no están protegidas: se alimentan pegando.

Las tablas dinámicas no se pueden actualizar con la hoja protegida. Después de
alimentar las bases: **Revisar → Desproteger hoja** (en Registros), **Datos →
Actualizar todo**, y **Revisar → Proteger hoja** marcando «Usar Autofiltro» y
«Usar tabla dinámica y gráfico dinámico».

## Validación antes de entregar

`herramientas/validar.py` revisa el archivo con las reglas que Excel exige y que
provocan el aviso «Hemos encontrado un problema con el contenido»: vistas con
selecciones en paneles inexistentes, columnas solapadas, filas o celdas fuera de
orden, estilos o cadenas inexistentes, celdas combinadas solapadas, validaciones
mal escritas y dibujos vacíos. Se probó contra las versiones anteriores: detecta
el fallo de la V13 y el de la V17, y da por buena la V16, que abrió sin aviso.

```
python3 herramientas/validar.py libro/Registro_Asistencia_2026-1_V23.xlsx
```

Las fórmulas que se repiten fila a fila se guardan como fórmulas compartidas de
Excel (`sharedfml.py`): una columna se comparte solo si cada fila es exactamente
la fórmula de la primera fila desplazada. Así el libro de 10.000 filas pesa lo
mismo que el de 2.000 (8,8 MB).

## Pruebas con LibreOffice

Con LibreOffice escuchando en el puerto 2085
(`soffice --headless --accept="socket,host=localhost,port=2085;urp;"`):

- `pruebas/prueba_registros.py LIBRO CASOS.json`: escribe cédulas e ID tecleando,
  pegando desde otro libro (texto con puntos, espacios invisibles o ceros a la
  izquierda) y con pegado especial de valores. Compara cada fila con la base y
  recalcula por su cuenta el resumen y Resultados por programa (todas las sedes,
  CIN y BUG). En la V23 salieron bien las 15 filas y todos los conteos.
- `pruebas/tablas_dinamicas.py LIBRO` y `pruebas/prueba_buscador.py LIBRO`:
  cargan las tablas dinámicas y prueban el buscador (en una copia sin
  protección, porque LibreOffice no las carga en hojas protegidas).
