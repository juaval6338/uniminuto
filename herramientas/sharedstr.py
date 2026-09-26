# -*- coding: utf-8 -*-
"""Convierte las cadenas en línea de un .xlsx a cadenas compartidas (reduce el tamaño)."""
import re, shutil, zipfile, os

RX = re.compile(rb'<c([^>]*?)t="inlineStr"([^>]*?)><is><t[^>]*>(.*?)</t></is></c>', re.S)
RX_EMPTY = re.compile(rb'<c([^>]*?)t="inlineStr"([^>]*?)><is><t[^>]*/></is></c>', re.S)

def convert(src, dst):
    zin = zipfile.ZipFile(src)
    names = zin.namelist()
    if 'xl/sharedStrings.xml' in names:
        shutil.copy(src, dst); return 0, 0
    table, index = [], {}
    def sub(m):
        pre, post, txt = m.group(1), m.group(2), m.group(3)
        i = index.get(txt)
        if i is None:
            i = len(table); index[txt] = i; table.append(txt)
        return b'<c' + pre + b't="s"' + post + b'><v>' + str(i).encode() + b'</v></c>'
    out = {}
    for n in names:
        data = zin.read(n)
        if n.startswith('xl/worksheets/sheet') and n.endswith('.xml'):
            data = RX_EMPTY.sub(lambda m: b'<c' + m.group(1) + m.group(2).rstrip() + b'/>', data)
            data = RX.sub(sub, data)
        out[n] = data
    # sharedStrings.xml
    parts = [b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
             b'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
             b'count="', str(len(table)).encode(), b'" uniqueCount="', str(len(table)).encode(), b'">']
    for t in table:
        parts.append(b'<si><t xml:space="preserve">' + t + b'</t></si>')
    parts.append(b'</sst>')
    out['xl/sharedStrings.xml'] = b''.join(parts)
    # content types
    ct = out['[Content_Types].xml']
    if b'sharedStrings.xml' not in ct:
        ct = ct.replace(b'</Types>',
            b'<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-'
            b'officedocument.spreadsheetml.sharedStrings+xml"/></Types>')
        out['[Content_Types].xml'] = ct
    # workbook rels
    rl = out['xl/_rels/workbook.xml.rels']
    if b'sharedStrings.xml' not in rl:
        ids = [int(x) for x in re.findall(rb'Id="rId(\d+)"', rl)] or [0]
        nid = max(ids) + 1
        rl = rl.replace(b'</Relationships>',
            b'<Relationship Id="rId' + str(nid).encode() + b'" Type="http://schemas.openxmlformats.org'
            b'/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
            b'</Relationships>')
        out['xl/_rels/workbook.xml.rels'] = rl
    order = [n for n in names if n != '[Content_Types].xml']
    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED, compresslevel=9) as z:
        z.writestr('[Content_Types].xml', out['[Content_Types].xml'])
        for n in order: z.writestr(n, out[n])
        z.writestr('xl/sharedStrings.xml', out['xl/sharedStrings.xml'])
    zin.close()
    return len(table), os.path.getsize(dst)


def trim_empty_cells(src, dst, sheet_names):
    """Quita, en las hojas indicadas, las celdas vacías que solo llevan estilo y las filas
    que quedan vacías. No cambia nada visible: esas celdas no tienen valor, relleno ni borde."""
    zin = zipfile.ZipFile(src)
    names = zin.namelist()
    wbx = zin.read('xl/workbook.xml').decode('utf8')
    rels = zin.read('xl/_rels/workbook.xml.rels').decode('utf8')
    rid = {m.group(2): m.group(1) for m in re.finditer(r'<sheet [^>]*?name="([^"]+)"[^>]*?r:id="([^"]+)"', wbx)}
    rid.update({m.group(2): m.group(1) for m in re.finditer(r'<sheet [^>]*?r:id="([^"]+)"[^>]*?name="([^"]+)"', wbx)})
    target = {}
    for m in re.finditer(r'<Relationship [^>]*?Id="([^"]+)"[^>]*?Target="([^"]+)"', rels):
        target[m.group(1)] = m.group(2)
    for m in re.finditer(r'<Relationship [^>]*?Target="([^"]+)"[^>]*?Id="([^"]+)"', rels):
        target[m.group(2)] = m.group(1)
    parts = set()
    for r_id, nm in rid.items():
        if nm in sheet_names:
            t = target[r_id].lstrip('/')
            parts.add(t if t.startswith('xl/') else 'xl/' + t)
    empty_cell = re.compile(rb'<c r="[A-Z]+\d+"(?: s="\d+")?(?: t="n")?\s*/>')
    empty_row = re.compile(rb'<row [^>]*?/>|<row [^>]*?>\s*</row>')
    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED, compresslevel=9) as z:
        for n in names:
            data = zin.read(n)
            if n in parts:
                data = empty_cell.sub(b'', data)
                data = empty_row.sub(b'', data)
            z.writestr(n, data)
    zin.close()
    return sorted(parts)
