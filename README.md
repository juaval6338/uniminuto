# Registro de Asistencia — Bienestar Institucional (UNIMINUTO)

Libro para registrar la asistencia a las actividades de Bienestar. Se escribe la
cédula o el ID del participante y el archivo trae sus datos desde las bases.

- `libro/Registro_Asistencia_2026-1_V14.xlsx` — el libro listo para usar.
- `herramientas/` — scripts que generan el libro y que leen el `.xlsb` original.

## Hojas

| Hoja | Para qué |
|---|---|
| `Registros` | Registro diario, buscador por nombre y resumen de participación |
| `Resultados por programa` | Cuántos participaron por programa académico y por área |
| `Evaluación` | Captura de las encuestas de evaluación |
| `Gráficos Evaluación` | Tabla y gráfica de los resultados de la encuesta |
| `BD ADM-DOC` / `BD EST` | Bases de datos (se actualizan pegando el export encima) |

## Hoja Registros

La tabla empieza en la fila 13 (encabezados) y tiene tres zonas:

- **A – K, automáticas.** Documento, SEDE, APELLIDOS Y NOMBRES, PROGRAMA / ÁREA,
  CORREO INSTITUCIONAL, CORREO ADICIONAL, TELÉFONO, TELÉFONO ADICIONAL y
  TIPO DE PARTICIPANTE.
- **L – O, manuales.** Solo se llenan cuando el nombre sale como `INEXISTENTE`
  (la fila se pinta de naranja): correo, tipo, sede y programa.
- **P – Q.** Actividad / espacio y observaciones.

El tipo de participante se muestra como texto en mayúscula —`ESTUDIANTE`,
`PROFESOR`, `ADMINISTRATIVO`— traduciendo los códigos 1, 27 y 28 que siguen
guardados tal cual en las bases.

Arriba hay dos bloques: el **buscador por nombre** (se escribe parte del nombre y
se elige de una lista desplegable) y el **resumen de participación** con
participantes, participaciones y porcentaje por tipo, con filtro de sede.

## Colores

| Color | Significado |
|---|---|
| Amarillo | lo escribe usted |
| Blanco | lo calcula el archivo (no escribir) |
| ID en azul fuerte | ese documento ya está registrado en otra fila |
| Salmón en la fila | la persona no está en la base de datos |

La columna `Q-Part` (el contador que alimenta el resumen) va oculta, igual que las
columnas auxiliares T a AO.

## Correos y teléfonos## Correos y teléfonos

El correo que manda es el institucional (`@uniminuto.edu`). El correo adicional
del estudiante sale de `C_ESTUDIANTE2`, tomando la parte anterior al `#`, y se
deja en blanco si coincide con el institucional. El teléfono principal es
`TEL_CEL` y el adicional el primero de `TEL_RE` / `TEL_TR` que sea distinto.

En `BD ADM-DOC` las columnas de correo adicional y teléfono están vacías en las
1.585 filas, así que profesores y administrativos no traen esos datos.

## Evaluación

Las preguntas y la escala siguen el formato institucional **FR-BM-DFB-03,
Versión 1, Enero 28 de 2021**: `E` Excelente (4), `N` Notable (3), `A` Aceptable (2),
`N/M` Necesita Mejoramiento (1) y `N/A` No aplica, que no promedia.

## Mantenimiento## Mantenimiento

Las bases conservan las columnas del export original, así que se siguen
actualizando pegando encima: se pega bajo el encabezado de la fila 1 y todo lo
demás se recalcula solo. Los rangos con nombre apuntan a **columnas completas**
(`'BD EST'!$C:$C`), así que no hay ningún tope de filas: da igual cuántos
estudiantes o colaboradores lleguen en cada periodo.

El libro se guarda como `.xlsx`. Si se quiere el tamaño del `.xlsb` original,
basta con *Archivo → Guardar como → Libro binario de Excel* dentro de Excel.
