import sys,pickle,time
SP='/tmp/claude-0/-home-user-uniminuto/b364fe32-f339-5501-a525-09b94e80e92f/scratchpad'
sys.path.insert(0,SP)
from dump2 import cells
from xlsbfml import colname

def grab(sheet, ncols):
    rows={}
    for row,col,val,fml in cells(sheet):
        if col>=ncols: continue
        if val in (None,''): continue
        rows.setdefault(row,{})[col]=val
    if not rows: return []
    last=max(rows)
    out=[]
    for r in range(last+1):
        d=rows.get(r,{})
        out.append([d.get(c) for c in range(ncols)])
    return out

t=time.time()
data={}
data['adm']=grab('BD ADM-DOC',16); print('adm',len(data['adm']),time.time()-t)
data['est']=grab('BD EST',47);     print('est',len(data['est']),time.time()-t)
data['reg']=grab('Registros',18);  print('reg',len(data['reg']),time.time()-t)
data['eva']=grab('Evaluación',12); print('eva',len(data['eva']),time.time()-t)
pickle.dump(data,open(SP+'/data.pkl','wb'),protocol=4)
print('saved', time.time()-t)
