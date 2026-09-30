import uno, time, sys
from com.sun.star.beans import PropertyValue
def P(n,v):
    p=PropertyValue(); p.Name=n; p.Value=v; return p
local=uno.getComponentContext()
res=local.ServiceManager.createInstanceWithContext("com.sun.star.bridge.UnoUrlResolver",local)
for _ in range(60):
    try: ctx=res.resolve("uno:socket,host=localhost,port=2085;urp;StarOffice.ComponentContext"); break
    except Exception: time.sleep(0.5)
desk=ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop",ctx)
doc=desk.loadComponentFromURL(uno.systemPathToFileUrl(sys.argv[1]),"_blank",0,(P("Hidden",True),))
reg=doc.Sheets.getByName("Registros")
C=lambda a: reg.getCellRangeByName(a)
casos=[('(Todas)','(Todas)','(Todas)'),('- all -','- all -','- all -'),
 ('CHAVES ORDOÑEZ WILSON ','(Todas)','(Todas)'),
 ('CA?AS ECHEVERRY MARIA JOS?','(Todas)','(Todas)'),
 ('(Todas)','jhon.lopez-go@uniminuto.edu.co','AGUDELO QUINTERO LUIS GUILLERMO'),
 ('NO EXISTE ESTE NOMBRE','(Todas)','(Varios elementos)')]
for j14,l14,j15 in casos:
    C('J14').setString(j14); C('L14').setString(l14); C('J15').setString(j15)
    doc.calculateAll()
    print('J14=%r L14=%r J15=%r'%(j14,l14,j15))
    for r in (14,15): print('   fila',r,[C('%s%d'%(c,r)).getString() for c in 'EFGH'])
doc.close(True)
