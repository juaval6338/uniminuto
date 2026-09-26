# Registro de Asistencia — Bienestar Institucional (UNIMINUTO)

Libro para registrar la asistencia a las actividades de Bienestar. Se escribe la
cédula o el ID del participante y el archivo trae sus datos desde las bases.

- `libro/Registro_Asistencia_2026-1_V19.xlsx` — el libro listo para usar.
- `herramientas/` — scripts que generan el libro (`build8.py`), el conversor de
  cadenas compartidas (`sharedstr.py`), el validador (`validar.py`) y los lectores
  del `.xlsb` original.

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

- Los códigos 1, 27 y 28 siguen en las bases; en Registros se muestran como texto.
- ID repetido: la letra se pone en azul fuerte (regla de duplicados del original).
- Persona que no está en la base: el nombre dice INEXISTENTE en negrita y la fila
  se pinta de naranja (Énfasis 6 del tema Office, el mismo tema del archivo); el
  correo que se escriba en L aparece en K.
- La columna U muestra siempre la cédula, aunque en F se haya escrito el ID.
- El documento se reconoce aunque venga pegado como texto, con puntos o con
  espacios (también el espacio invisible que traen las páginas web y Forms).
- Buscador (filas 13 a 15), como en el original: J14 (estudiantes) y J15
  (colaboradores) tienen una lista desplegable con todos los nombres de la base,
  sin renglones en blanco y hasta la última fila, aunque la base crezca. También
  se puede escribir parte del nombre, o el correo en L14.
- Resumen de participación arriba a la derecha (L1:P8) con filtro de sede en M2.
- 2000 filas de registro (19 a 2018). Q-Part, Día y Mes van ocultas.
- Los encabezados quedan inmovilizados hasta la fila 18, como en el original.

## Otras hojas

- **Resultados por tipo**: eliminada; su cuadro está en Registros.
- **Resultados por programa**: tres bloques horizontales (estudiantes,
  profesores, administrativos) ordenados de mayor a menor.
- **Evaluación**: la tabla de la derecha refleja también lo que se pega, y la
  escala sigue el formato FR-BM-DFB-03 (E, N, A, N/M, N/A).
- **Gráficos Evaluación**: una sola gráfica de barras en lugar de cuatro.

## Mantenimiento

Las bases conservan las columnas del export, así que se actualizan pegando encima
bajo el encabezado de la fila 1. Las fórmulas apuntan a columnas completas, sin
tope de filas.

## Validación antes de entregar

`herramientas/validar.py` revisa el archivo con las reglas que Excel exige y que
provocan el aviso «Hemos encontrado un problema con el contenido»: vistas con
selecciones en paneles inexistentes, columnas solapadas, filas o celdas fuera de
orden, estilos o cadenas inexistentes, celdas combinadas solapadas, validaciones
mal escritas y dibujos vacíos. Se probó contra las versiones anteriores: detecta
el fallo de la V13 y el de la V17, y da por buena la V16, que abrió sin aviso.

```
python3 herramientas/validar.py libro/Registro_Asistencia_2026-1_V19.xlsx
```
