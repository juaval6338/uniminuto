# Registro de Asistencia — UNIMINUTO

Libro de registro de asistencia optimizado (versión 10), reconstruido a partir de
`Registro_Asistencia_2000_2026-1_V9.0.1.xlsb`.

## Contenido

- `libro/Registro_Asistencia_2026-1_V10.xlsx` — el libro listo para usar.
- `herramientas/` — scripts con los que se generó el libro y con los que se leyó
  el archivo `.xlsb` original (Python + openpyxl).

## Qué cambió frente a la versión 9

**Hoja `Registros`**
- Se escribe la cédula o el ID y el libro trae **Sede, Nombre, Programa, Correo,
  Teléfono y Tipo de participante**. Se retiraron los campos que ya no se usan.
- Bloque *«Complete solo si el documento no existe en la base»*: el correo escrito
  ahí aparece en la columna **Correo electrónico** (antes quedaba en `-`).
- Buscador por nombre sin tablas dinámicas, que no repite personas aunque la base
  traiga varias filas por estudiante.
- Columna **Estado** con semáforo para ver qué filas están incompletas.

**Hoja `Resultados por tipo`**
- Solo las dos tablas (Participantes y Participaciones, con cantidad y porcentaje)
  y el filtro de sede. Se eliminó el bloque auxiliar de ~60.000 fórmulas.

**Hoja `Resultados por programa`**
- Mismo diseño de tres bloques, pero alimentado directamente desde `Registros`.

**Hoja `Evaluación`**
- La segunda tabla ahora refleja los valores **pegados**, no solo los escritos:
  acepta números, texto, letras en minúscula y espacios sobrantes.
- Las entradas no válidas se marcan en rojo.

## Mantenimiento

Las bases `BD ADM-DOC` y `BD EST` conservan las mismas columnas del export, así que
se siguen actualizando pegando encima. Los rangos con nombre cubren hasta 6.000
filas de colaboradores y 60.000 de estudiantes; si el export crece más, hay que
ampliarlos en *Fórmulas → Administrador de nombres*.
