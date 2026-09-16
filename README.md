# Registro de Asistencia — UNIMINUTO

Libro de registro de asistencia simplificado (versión 11), reconstruido a partir de
`Registro_Asistencia_2000_2026-1_V9.0.1.xlsb`.

## Contenido

- `libro/Registro_Asistencia_2026-1_V11.xlsx` — el libro listo para usar.
- `herramientas/` — scripts con los que se generó el libro y con los que se leyó
  el archivo `.xlsb` original (Python + openpyxl).

## Qué cambió frente a la versión 9

**Hoja `Registros`** — una sola tabla, sin cuadros auxiliares.
- Los datos empiezan en la fila 6 (antes en la 19) y las columnas son doce:
  Q-Part, Fecha, Documento, Sede, Nombre, Programa, Correo, Teléfono, Tipo,
  Correo manual, Actividad y Observaciones.
- Se escribe la cédula o el ID en **Documento** y el resto aparece solo.
- Si la persona no está en la base, el nombre dice **Inexistente** y la fila se
  pinta de naranja. En ese caso se escribe el correo en la columna J y aparece
  en **Correo electrónico**.
- Buscador de una línea en la fila 3: se escribe parte del nombre y devuelve el
  documento y el nombre completo. Busca primero en estudiantes y luego en
  colaboradores.

**Hoja `Resultados por tipo`**
- Solo las dos tablas (Participantes y Participaciones, con cantidad y
  porcentaje) y el filtro de sede. Se eliminó el bloque auxiliar de unas
  60.000 fórmulas.

**Hoja `Resultados por programa`**
- Mismo diseño de tres bloques, alimentado directamente desde `Registros`.

**Hoja `Evaluación`**
- La segunda tabla refleja los valores **pegados**, no solo los escritos:
  acepta números, texto, letras en minúscula y espacios sobrantes.
- Las entradas no válidas se marcan en rojo.

**Hoja `Gráficos Evaluación`**
- Una tabla con las diez preguntas (nombre corto, porcentaje y texto completo)
  y **un solo gráfico** de barras horizontales, en vez de cuatro gráficos
  superpuestos.

## Mantenimiento

Las bases `BD ADM-DOC` y `BD EST` conservan las mismas columnas del export, así que
se siguen actualizando pegando encima. Los rangos con nombre cubren hasta 6.000
filas de colaboradores y 60.000 de estudiantes; si el export crece más, hay que
ampliarlos en *Fórmulas → Administrador de nombres*.
