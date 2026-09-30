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
for f in sys.argv[1:]:
    doc=desk.loadComponentFromURL(uno.systemPathToFileUrl(f),"_blank",0,(P("Hidden",True),))
    reg=doc.Sheets.getByName("Registros"); dpt=reg.getDataPilotTables()
    info=[]
    for n in dpt.ElementNames:
        t=dpt.getByName(n); fs=t.getDataPilotFields()
        info.append((n,[(fs.getByIndex(i).Name,fs.getByIndex(i).Orientation.value,len(fs.getByIndex(i).getItems().ElementNames)) for i in range(fs.Count)]))
    print(f.split('/')[-1],'->',info if info else 'NINGUNA')
    doc.close(True)
