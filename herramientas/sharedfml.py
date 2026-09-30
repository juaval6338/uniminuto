# -*- coding: utf-8 -*-
"""Convierte fórmulas repetidas fila a fila en fórmulas compartidas de Excel (t="shared").

Una columna (o un tramo de filas seguidas) se comparte solo si la fórmula de cada fila es
exactamente la de la primera fila desplazada a esa fila; si no, se deja como está.
"""
import re, zipfile
from xml.sax.saxutils import escape, unescape

CELL = re.compile(r'<c r="([A-Z]+)(\d+)"([^>]*?)(/>|>(.*?)</c>)', re.S)
FORM = re.compile(r'<f>(.*?)</f>', re.S)
# referencia A1 (con o sin $); no forma parte de un nombre ni de una función
REF = re.compile(r'(?<![A-Za-z0-9_.])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![0-9A-Za-z_(])')


def desplazar(formula, d):
    """La fórmula como quedaría d filas más abajo (solo cambian las filas relativas)."""
    partes = formula.split('"')
    for i in range(0, len(partes), 2):          # los pares van fuera de comillas
        partes[i] = REF.sub(lambda m: m.group(0) if m.group(3) else
                            '%s%s%s' % (m.group(1), m.group(2), int(m.group(4)) + d), partes[i])
    return '"'.join(partes)


def compartir_xml(s):
    formulas = {}
    for m in CELL.finditer(s):
        if m.group(5) and '<f>' in m.group(5):
            formulas.setdefault(m.group(1), []).append((int(m.group(2)), unescape(FORM.search(m.group(5)).group(1))))
    miembros, si, total = {}, 0, 0
    for col, filas in formulas.items():
        filas.sort()
        i = 0
        while i < len(filas):
            r0, f0 = filas[i]
            j = i + 1
            while j < len(filas) and filas[j][0] == filas[j - 1][0] + 1 and filas[j][1] == desplazar(f0, filas[j][0] - r0):
                j += 1
            if j - i >= 2:
                ref = '%s%d:%s%d' % (col, r0, col, filas[j - 1][0])
                miembros[(col, r0)] = '<f t="shared" ref="%s" si="%d">%s</f>' % (ref, si, escape(f0))
                for r, _ in filas[i + 1:j]:
                    miembros[(col, r)] = '<f t="shared" si="%d"/>' % si
                si += 1; total += j - i
            i = j

    def cambiar(m):
        nuevo = miembros.get((m.group(1), int(m.group(2))))
        if nuevo is None:
            return m.group(0)
        return '<c r="%s%s"%s>%s</c>' % (m.group(1), m.group(2), m.group(3), FORM.sub(lambda _: nuevo, m.group(5), 1))
    return CELL.sub(cambiar, s), si, total


def compartir(src, dst, partes):
    """partes: rutas de hoja dentro del zip (p. ej. 'xl/worksheets/sheet3.xml')."""
    zin = zipfile.ZipFile(src)
    info = []
    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED, compresslevel=9) as z:
        for it in zin.infolist():
            d = zin.read(it.filename)
            if it.filename in partes:
                x, grupos, celdas = compartir_xml(d.decode('utf8'))
                d = x.encode('utf8'); info.append((it.filename, grupos, celdas))
            z.writestr(it.filename, d)
    return info
