# -*- coding: utf-8 -*-
"""Inserta en un .xlsx tablas dinámicas de un solo filtro (como el buscador del original)."""
import re, zipfile, unicodedata
from xml.sax.saxutils import escape

NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
RT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
CT = 'application/vnd.openxmlformats-officedocument.spreadsheetml.'

def _attr(s):
    return escape(s, {'"': '&quot;'})

def _clave(s):
    s = unicodedata.normalize('NFKD', s)
    return ''.join(ch for ch in s if not unicodedata.combining(ch)).casefold()

def cache_xml(field_name, source, values):
    """source: (hoja, rango), p. ej. ('BD EST', 'C1:C1048576') = columna completa, como el original.
    values: valores de la columna en orden (sin el encabezado); None = vacío."""
    uniq, idx = [], {}
    for v in values:
        if v is None: continue
        k = v.casefold()
        if k not in idx:
            idx[k] = None; uniq.append(v)
    uniq.sort(key=_clave)
    pos = {v.casefold(): i for i, v in enumerate(uniq)}
    blank = any(v is None for v in values)
    nitems = len(uniq) + (1 if blank else 0)
    si = ''.join('<s v="%s"/>' % _attr(v) for v in uniq) + ('<m/>' if blank else '')
    definition = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<pivotCacheDefinition xmlns="%s" xmlns:r="%s" r:id="rId1" '
        'createdVersion="6" refreshedVersion="6" minRefreshableVersion="3" recordCount="%d">'
        '<cacheSource type="worksheet"><worksheetSource ref="%s" sheet="%s"/></cacheSource>'
        '<cacheFields count="1"><cacheField name="%s" numFmtId="0">'
        '<sharedItems%s count="%d">%s</sharedItems></cacheField></cacheFields>'
        '</pivotCacheDefinition>'
    ) % (NS_MAIN, NS_R, len(values), source[1], _attr(source[0]), _attr(field_name),
         ' containsBlank="1"' if blank else '', nitems, si)
    recs = ''.join('<r><x v="%d"/></r>' % (len(uniq) if v is None else pos[v.casefold()]) for v in values)
    records = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
               '<pivotCacheRecords xmlns="%s" xmlns:r="%s" count="%d">%s</pivotCacheRecords>'
               ) % (NS_MAIN, NS_R, len(values), recs)
    return definition, records, nitems

def table_xml(name, cache_id, location, caption, nitems):
    items = ''.join('<item x="%d"/>' % i for i in range(nitems))
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<pivotTableDefinition xmlns="%s" name="%s" cacheId="%d" applyNumberFormats="0" '
        'applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" '
        'applyAlignmentFormats="0" applyWidthHeightFormats="0" dataCaption="Valores" '
        'updatedVersion="6" minRefreshableVersion="3" useAutoFormatting="0" itemPrintTitles="1" '
        'createdVersion="6" indent="0" outline="1" outlineData="1" multipleFieldFilters="0">'
        '<location ref="%s" firstHeaderRow="1" firstDataRow="1" firstDataCol="0" '
        'rowPageCount="1" colPageCount="1"/>'
        '<pivotFields count="1"><pivotField name="%s" axis="axisPage" showAll="0" '
        'sortType="ascending" defaultSubtotal="0"><items count="%d">%s</items></pivotField></pivotFields>'
        '<pageFields count="1"><pageField fld="0" hier="-1"/></pageFields>'
        '<pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" '
        'showRowStripes="0" showColStripes="0" showLastColumn="1"/>'
        '</pivotTableDefinition>'
    ) % (NS_MAIN, _attr(name), cache_id, location, _attr(caption), nitems, items)

def _rels(entries):
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">%s'
            '</Relationships>') % ''.join('<Relationship Id="%s" Type="%s" Target="%s"/>' % e for e in entries)

def add_pivots(src, dst, sheet_name, specs):
    """specs: dicts con name, location, caption, field, source, values."""
    zin = zipfile.ZipFile(src)
    out = {n: zin.read(n) for n in zin.namelist()}
    order = zin.namelist(); zin.close()
    wb = out['xl/workbook.xml'].decode('utf8')
    wbrels = out['xl/_rels/workbook.xml.rels'].decode('utf8')
    ct = out['[Content_Types].xml'].decode('utf8')
    # parte de la hoja
    rid = re.search(r'<sheet [^>]*name="%s"[^>]*r:id="([^"]+)"' % re.escape(sheet_name), wb).group(1)
    tgt = re.search(r'<Relationship [^>]*Id="%s"[^>]*Target="([^"]+)"' % rid, wbrels) or \
          re.search(r'<Relationship [^>]*Target="([^"]+)"[^>]*Id="%s"' % rid, wbrels)
    sheet_part = 'xl/' + tgt.group(1).lstrip('/').replace('xl/', '', 1)
    sheet_rels_part = sheet_part.replace('worksheets/', 'worksheets/_rels/') + '.rels'
    sheet_rels = []
    if sheet_rels_part in out:
        for tag in re.findall(r'<Relationship [^>]*>', out[sheet_rels_part].decode('utf8')):
            g = lambda a: re.search(a + r'="([^"]+)"', tag).group(1)
            sheet_rels.append((g('Id'), g('Type'), g('Target')))
    used = [int(x) for x in re.findall(r'Id="rId(\d+)"', wbrels)] or [0]
    next_wb = max(used) + 1
    caches = []
    for i, sp in enumerate(specs, start=1):
        cid = i
        d, r, n = cache_xml(sp['field'], sp['source'], sp['values'])
        out['xl/pivotCache/pivotCacheDefinition%d.xml' % i] = d.encode('utf8')
        out['xl/pivotCache/pivotCacheRecords%d.xml' % i] = r.encode('utf8')
        out['xl/pivotCache/_rels/pivotCacheDefinition%d.xml.rels' % i] = _rels(
            [('rId1', RT + 'pivotCacheRecords', 'pivotCacheRecords%d.xml' % i)]).encode('utf8')
        out['xl/pivotTables/pivotTable%d.xml' % i] = table_xml(
            sp['name'], cid, sp['location'], sp['caption'], n).encode('utf8')
        out['xl/pivotTables/_rels/pivotTable%d.xml.rels' % i] = _rels(
            [('rId1', RT + 'pivotCacheDefinition', '../pivotCache/pivotCacheDefinition%d.xml' % i)]).encode('utf8')
        sheet_rels.append(('rIdPvt%d' % i, RT + 'pivotTable', '../pivotTables/pivotTable%d.xml' % i))
        wrid = 'rId%d' % next_wb; next_wb += 1
        wbrels = wbrels.replace('</Relationships>',
            '<Relationship Id="%s" Type="%s" Target="pivotCache/pivotCacheDefinition%d.xml"/></Relationships>'
            % (wrid, RT + 'pivotCacheDefinition', i))
        caches.append('<pivotCache cacheId="%d" r:id="%s"/>' % (cid, wrid))
        ct = ct.replace('</Types>',
            '<Override PartName="/xl/pivotCache/pivotCacheDefinition%d.xml" ContentType="%spivotCacheDefinition+xml"/>'
            '<Override PartName="/xl/pivotCache/pivotCacheRecords%d.xml" ContentType="%spivotCacheRecords+xml"/>'
            '<Override PartName="/xl/pivotTables/pivotTable%d.xml" ContentType="%spivotTable+xml"/></Types>'
            % (i, CT, i, CT, i, CT))
    out[sheet_rels_part] = _rels(sheet_rels).encode('utf8')
    # <pivotCaches> va justo después de <calcPr> (orden del esquema CT_Workbook)
    m = re.search(r'<calcPr[^>]*/>', wb)
    wb = wb[:m.end()] + '<pivotCaches>%s</pivotCaches>' % ''.join(caches) + wb[m.end():]
    out['xl/workbook.xml'] = wb.encode('utf8')
    out['xl/_rels/workbook.xml.rels'] = wbrels.encode('utf8')
    out['[Content_Types].xml'] = ct.encode('utf8')
    names = ['[Content_Types].xml'] + [n for n in order if n != '[Content_Types].xml'] + \
            sorted(n for n in out if n not in order)
    with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED, compresslevel=9) as z:
        for n in names: z.writestr(n, out[n])
