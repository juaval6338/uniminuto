import uno, time, sys, json
from collections import Counter, defaultdict
from com.sun.star.beans import PropertyValue
def P(n,v):
    p=PropertyValue(); p.Name=n; p.Value=v; return p
local=uno.getComponentContext()
res=local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver",local)
for _ in range(60):
    try: ctx=res.resolve("uno:socket,host=localhost,port=2085;urp;StarOffice.ComponentContext"); break
    except Exception: time.sleep(0.5)
smgr=ctx.ServiceManager
desk=smgr.createInstanceWithContext("com.sun.star.frame.Desktop",ctx)
disp=smgr.createInstanceWithContext("com.sun.star.frame.DispatchHelper",ctx)
t0=time.time()
doc=desk.loadComponentFromURL(uno.systemPathToFileUrl(sys.argv[1]),"_blank",0,(P("Hidden",True),))
print('abierto en %.0fs'%(time.time()-t0))
casos=json.load(open(sys.argv[2]))
reg=doc.Sheets.getByName("Registros"); rp=doc.Sheets.getByName("Resultados por programa")
C=lambda a: reg.getCellRangeByName(a)
print('hojas protegidas:',[(doc.Sheets.getByIndex(i).Name,doc.Sheets.getByIndex(i).isProtected()) for i in range(doc.Sheets.Count)])
lock=lambda a: C(a).CellProtection.IsLocked
print('bloqueo  E19 F19 L19 M19 N19 O19 Q19 | G19 I19 J19 P19 U19 | H8 M2 | E10018 F10018 G10018:',
      [lock(a) for a in 'E19 F19 L19 M19 N19 O19 Q19 G19 I19 J19 P19 U19 H8 M2 E10018 F10018 G10018'.split()])
# 1) escribiendo
for tipo,r,fecha,docu,exp in casos:
    if tipo!='tecleo': continue
    C('E%d'%r).setFormula(fecha); C('F%d'%r).setFormula(docu)
    if exp['J']=='INEXISTENTE':
        C('L%d'%r).setFormula(exp['K']); C('M%d'%r).setFormula('EXTERNO'); C('N%d'%r).setFormula('BUG'); C('O%d'%r).setFormula('Psicología')
# 2) pegado normal (Ctrl+V) desde otro libro con los documentos como texto
src=desk.loadComponentFromURL("private:factory/scalc","_blank",0,(P("Hidden",True),))
sh=src.Sheets.getByIndex(0)
pg=[c for c in casos if c[0]=='pegado']
for i,(tipo,r,fecha,docu,exp) in enumerate(pg):
    sh.getCellByPosition(0,i).setValue(fecha); sh.getCellByPosition(1,i).setString(docu)
sc=src.getCurrentController(); sc.select(sh.getCellRangeByName('A1:B%d'%len(pg)))
tr=sc.getTransferable()
ctrl=doc.getCurrentController()
ctrl.select(reg.getCellRangeByName('E%d:F%d'%(pg[0][1],pg[-1][1]))); ctrl.insertTransferable(tr)
print('pegado normal ->',[(C('E%d'%c[1]).getValue(),C('F%d'%c[1]).getString()) for c in pg][:3],'...')
print('  ¿quedaron libres tras pegar?',[not lock('F%d'%c[1]) for c in pg])
# 3) pegado especial de valores (Copiar + Pegado especial > solo valores)
vl=[c for c in casos if c[0]=='valores']
for i,(tipo,r,fecha,docu,exp) in enumerate(vl):
    sh.getCellByPosition(3,i).setValue(fecha); sh.getCellByPosition(4,i).setValue(docu)
sc.select(sh.getCellRangeByName('D1:E%d'%len(vl)))
disp.executeDispatch(src.getCurrentController().getFrame(),".uno:Copy","",0,())
ctrl.select(reg.getCellRangeByName('E%d:F%d'%(vl[0][1],vl[-1][1])))
disp.executeDispatch(ctrl.getFrame(),".uno:InsertContents","",0,(P("Flags","SVD"),P("FormulaCommand",0),P("SkipEmptyCells",False),P("Transpose",False),P("AsLink",False),P("MoveMode",4)))
got=[C('F%d'%c[1]).getValue() for c in vl]
print('pegado de valores ->',got, '(portapapeles OK)' if got[0] else '(sin portapapeles: se escribe con setDataArray)')
if not got[0]:
    reg.getCellRangeByName('E%d:F%d'%(vl[0][1],vl[-1][1])).setDataArray(tuple((c[2],c[3]) for c in vl))
# 4) intentar pegar sobre una columna con fórmula (bloqueada)
antes=C('G6010').getFormula()
try:
    ctrl.select(reg.getCellRangeByName('G6010:H6010')); ctrl.insertTransferable(tr); r4='aceptó'
except Exception as e: r4='rechazado: '+type(e).__name__
print('pegar en G6010 (bloqueada):',r4,'| fórmula intacta:',C('G6010').getFormula()==antes)
doc.calculateAll()
# comprobar filas
malos=0
for tipo,r,fecha,docu,exp in casos:
    got=dict(G=C('G%d'%r).getString().strip(),H=C('H%d'%r).getString().strip(),I=int(C('I%d'%r).getValue()) if C('I%d'%r).getString().strip() else '',
             J=C('J%d'%r).getString(),K=C('K%d'%r).getString(),P=C('P%d'%r).getString(),U=int(C('U%d'%r).getValue()))
    dif={k:(got[k],v) for k,v in exp.items() if got.get(k)!=v}
    malos+=bool(dif)
    print('%-8s fila %5d %-16r -> %-35s %-14s %s'%(tipo,r,docu,got['J'][:35],got['P'],'OK' if not dif else 'DIFERENCIA %s'%dif))
print('filas con diferencias:',malos,'de',len(casos))
# resumen y resultados por programa comparados con un conteo independiente de las filas
cols=reg.getCellRangeByName('A19:AA10018').getDataArray()
A=[r[0] for r in cols]; G=[r[6] for r in cols]; H=[r[7] for r in cols]; Pt=[r[15] for r in cols]; AA=[r[26] for r in cols]
def esperado(sede):
    ok=lambda i: sede=='TODAS' or G[i]==sede
    pp=Counter(Pt[i] for i in range(len(A)) if AA[i]==1 and ok(i)); pc=Counter(Pt[i] for i in range(len(A)) if A[i]==1 and ok(i))
    return [[pp[t],pc[t]] for t in ('ESTUDIANTE','PROFESOR','ADMINISTRATIVO','EXTERNO')]
for sede in ('TODAS','CIN','BUG'):
    C('M2').setString(sede); doc.calculateAll()
    hoja=[[int(C('M%d'%r).getValue()),int(C('O%d'%r).getValue())] for r in range(4,8)]
    tot=[int(C('M8').getValue()),int(C('O8').getValue())]
    exp=esperado(sede)
    print('resumen sede %-5s'%sede,hoja,'total',tot,'OK' if hoja==exp and tot==[sum(x[0] for x in exp),sum(x[1] for x in exp)] else 'DIFERENCIA esperado %s'%exp)
    tabla=rp.getCellRangeByName('B8:L108').getDataArray()
    for k,(tp,c0) in enumerate((('ESTUDIANTE',0),('PROFESOR',4),('ADMINISTRATIVO',8))):
        hojad={row[c0]:(int(row[c0+1]),int(row[c0+2])) for row in tabla[:-1] if row[c0]}
        ok=lambda i: sede=='TODAS' or G[i]==sede
        e=defaultdict(lambda:[0,0])
        for i in range(len(A)):
            if Pt[i]==tp and H[i] and ok(i):
                if AA[i]==1: e[H[i]][0]+=1
                if A[i]==1: e[H[i]][1]+=1
        e={k2:tuple(v) for k2,v in e.items() if v!=[0,0]}
        hojad={k2:v for k2,v in hojad.items() if v!=(0,0)}
        orden=[int(row[c0+1]) for row in tabla[:-1] if row[c0]]
        print('   programa %-14s %2d dependencias'%(tp,len(hojad)),'OK' if hojad==e else 'DIFERENCIA', '| total fila 108:',int(tabla[-1][c0+1]),int(tabla[-1][c0+2]),'| orden mayor a menor:',orden==sorted(orden,reverse=True))
C('M2').setString('TODAS')
doc.close(True); src.close(True)
