# -*- coding: utf-8 -*-
"""Revisa un .xlsx buscando lo que Excel marca como 'problema con el contenido'."""
import sys, re, zipfile, posixpath
import xml.etree.ElementTree as ET
NS={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
def colnum(s):
    n=0
    for ch in s: n=n*26+ord(ch)-64
    return n
def split(ref):
    m=re.match(r'\$?([A-Z]+)\$?(\d+)$',ref); return colnum(m.group(1)),int(m.group(2))
def rng(a):
    p=a.split(':'); c1,r1=split(p[0]); c2,r2=split(p[-1]); return c1,r1,c2,r2

def validar(path):
    errs=[]; z=zipfile.ZipFile(path); names=set(z.namelist())
    if z.testzip() is not None: errs.append('zip corrupto')
    for n in names:
        if n.endswith(('.xml','.rels')):
            try: ET.fromstring(z.read(n))
            except Exception as e: errs.append('XML mal formado %s: %s'%(n,e))
    # relaciones
    for rn in [x for x in names if x.endswith('.rels')]:
        base=posixpath.dirname(posixpath.dirname(rn))
        for m in re.finditer(r'<Relationship [^>]*>', z.read(rn).decode('utf8')):
            t=re.search(r'Target="([^"]+)"',m.group(0)).group(1)
            if 'External' in m.group(0) or t.startswith('http'): continue
            p=posixpath.normpath(posixpath.join(base,t)) if not t.startswith('/') else t[1:]
            if p not in names: errs.append('relación rota %s -> %s'%(rn,t))
    # estilos
    st=ET.fromstring(z.read('xl/styles.xml'))
    xfs=st.find('m:cellXfs',NS); nxf=len(xfs)
    if int(xfs.get('count',nxf))!=nxf: errs.append('styles: cellXfs count %s != %d'%(xfs.get('count'),nxf))
    dx=st.find('m:dxfs',NS); ndxf=len(dx) if dx is not None else 0
    for tag in ('fonts','fills','borders','cellStyleXfs'):
        e=st.find('m:'+tag,NS)
        if e is not None and int(e.get('count',len(e)))!=len(e): errs.append('styles: %s count'%tag)
    # cadenas compartidas
    nss=0
    if 'xl/sharedStrings.xml' in names:
        nss=len(ET.fromstring(z.read('xl/sharedStrings.xml')))
    # nombres definidos
    wb=ET.fromstring(z.read('xl/workbook.xml'))
    seen=set()
    dn=wb.find('m:definedNames',NS)
    for d in (dn if dn is not None else []):
        k=(d.get('name').lower(),d.get('localSheetId'))
        if k in seen: errs.append('nombre definido duplicado %s'%(k,))
        seen.add(k)
        if not (d.text or '').strip(): errs.append('nombre definido vacío %s'%d.get('name'))
    # hojas
    rels=z.read('xl/_rels/workbook.xml.rels').decode()
    tgt={re.search(r'Id="([^"]+)"',m).group(1):re.search(r'Target="([^"]+)"',m).group(1)
         for m in re.findall(r'<Relationship [^>]*>',rels)}
    for sh in wb.find('m:sheets',NS):
        nm=sh.get('name'); rid=sh.get('{%s}id'%NS['r'])
        part=tgt[rid].lstrip('/'); part=part if part.startswith('xl/') else 'xl/'+part
        root=ET.fromstring(z.read(part)); pre='[%s] '%nm
        # vistas
        for v in root.iter('{%s}sheetView'%NS['m']):
            pane=v.find('m:pane',NS)
            if pane is None: ok={'topLeft',None}
            else:
                xs=float(pane.get('xSplit',0)); ys=float(pane.get('ySplit',0))
                ok={'topLeft',None}
                if xs and ys: ok|={'topRight','bottomLeft','bottomRight'}
                elif ys: ok|={'bottomLeft'}
                elif xs: ok|={'topRight'}
                if pane.get('activePane','topLeft') not in ok: errs.append(pre+'vista: activePane inválido')
            ps=[s.get('pane') for s in v.findall('m:selection',NS)]
            for p in ps:
                if p not in ok: errs.append(pre+'vista: selección en panel inexistente "%s"'%p)
            if len(ps)!=len(set(ps)): errs.append(pre+'vista: selecciones repetidas %s'%ps)
        # columnas
        last=0
        for c in root.iter('{%s}col'%NS['m']):
            if int(c.get('style',0))>=nxf: errs.append(pre+'columna con estilo inexistente')
            a,b=int(c.get('min')),int(c.get('max'))
            if a>b or a<=last: errs.append(pre+'columnas solapadas o desordenadas en %d-%d'%(a,b))
            last=max(last,b)
        # filas y celdas
        lr=0
        for row in root.iter('{%s}row'%NS['m']):
            r=int(row.get('r'))
            if int(row.get('s',0))>=nxf: errs.append(pre+'fila con estilo inexistente %d'%r); break
            if r<=lr: errs.append(pre+'fila fuera de orden %d'%r); break
            lr=r; lc=0
            for c in row.findall('m:c',NS):
                cc,rr=split(c.get('r'))
                if rr!=r or cc<=lc: errs.append(pre+'celda fuera de orden %s'%c.get('r')); break
                lc=cc
                s=int(c.get('s',0))
                if s>=nxf: errs.append(pre+'estilo inexistente en %s'%c.get('r')); break
                if c.get('t')=='s':
                    v=c.find('m:v',NS)
                    if v is None or int(v.text)>=nss: errs.append(pre+'cadena compartida inválida %s'%c.get('r')); break
        # celdas combinadas
        mc=[m.get('ref') for m in root.iter('{%s}mergeCell'%NS['m'])]
        bx=[rng(x) for x in mc]
        for i in range(len(bx)):
            for j in range(i+1,len(bx)):
                a,b=bx[i],bx[j]
                if a[0]<=b[2] and b[0]<=a[2] and a[1]<=b[3] and b[1]<=a[3]:
                    errs.append(pre+'combinadas solapadas %s / %s'%(mc[i],mc[j]))
        # validaciones y formato condicional
        for dv in root.iter('{%s}dataValidation'%NS['m']):
            f=dv.find('m:formula1',NS)
            if f is not None and f.text and (f.text.startswith('=') or len(f.text)>255):
                errs.append(pre+'validación con fórmula inválida %s'%dv.get('sqref'))
        for r in root.iter('{%s}cfRule'%NS['m']):
            if r.get('dxfId') is not None and int(r.get('dxfId'))>=ndxf:
                errs.append(pre+'formato condicional con dxfId inexistente')
    # dibujos vacíos
    for n in names:
        if n.startswith('xl/drawings/drawing') and n.endswith('.xml'):
            if 'Anchor' not in z.read(n).decode('utf8'): errs.append('dibujo vacío %s'%n)
    return errs

if __name__=='__main__':
    for p in sys.argv[1:]:
        e=validar(p)
        print('%s: %s'%(p.split('/')[-1], 'SIN PROBLEMAS' if not e else '%d problema(s)'%len(e)))
        for x in e[:25]: print('   -',x)
