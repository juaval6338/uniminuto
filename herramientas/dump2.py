import zipfile,sys,struct,re
SP='/tmp/claude-0/-home-user-uniminuto/b364fe32-f339-5501-a525-09b94e80e92f/scratchpad'
sys.path.insert(0,SP)
from xlsbfml import records, Reader, parse_rgce, colname
Z=zipfile.ZipFile(SP+'/original.xlsb')

rels={}
rx=Z.read('xl/_rels/workbook.bin.rels').decode('utf8')
for m in re.finditer(r'<Relationship[^>]*>', rx):
    t=m.group(0)
    i=re.search(r'Id="([^"]+)"',t); g=re.search(r'Target="([^"]+)"',t)
    if i and g: rels[i.group(1)]=g.group(1)

wb=Z.read('xl/workbook.bin')
sheets=[]; XT=[]; names={}; name_list=[]; ni=0
for rid,p in records(wb):
    if rid==156:
        r=Reader(p); r.u32(); r.u32(); rid_=r.wstr(); nm=r.wstr(); sheets.append((nm,rels.get(rid_,'?')))
    elif rid==362:
        r=Reader(p); n=r.u32()
        for _ in range(n): XT.append((r.u32(),r.u32(),r.u32()))
    elif rid==39:
        r=Reader(p); r.u32(); r.u8(); itab=r.u32(); nm=r.wstr()
        cce=r.u32(); rgce=p[r.p:r.p+cce]; r.skip(cce); cb=r.u32(); rgcb=p[r.p:r.p+cb]
        names[ni]=nm; name_list.append((ni,nm,itab,rgce,rgcb)); ni+=1
SN=[s[0] for s in sheets]

# shared strings
SST=[]
try:
    for rid,p in records(Z.read('xl/sharedStrings.bin')):
        if rid==19:
            r=Reader(p); r.u8(); SST.append(r.wstr())
except Exception as e: print('SST err',e)

def part_of(name):
    for nm,tg in sheets:
        if nm==name: return tg if tg.startswith('xl/') else 'xl/'+tg
    raise KeyError(name)

def rkval(v):
    iv=struct.unpack('<i',struct.pack('<I',v))[0]
    if v&2: x=float(iv>>2)
    else: x=struct.unpack('<d',struct.pack('<Q',(v&0xFFFFFFFC)<<32))[0]
    return x/100.0 if (v&1) else x

def cells(sheetname):
    d=Z.read(part_of(sheetname)); row=-1
    shared={}
    for rid,p in records(d):
        if rid==0: row=struct.unpack_from('<I',p,0)[0]
        elif rid in (1,2,3,4,5,6,7,8,9,10,11):
            col=struct.unpack_from('<I',p,0)[0]
            r=Reader(p); r.skip(8); val=None; rgce=rgcb=None
            if rid==1: val=None
            elif rid==2: val=rkval(r.u32())
            elif rid==3: val='#ERR%d'%r.u8()
            elif rid==4: val=bool(r.u8())
            elif rid==5: val=r.dbl()
            elif rid==6: val=r.wstr()
            elif rid==7:
                i=r.u32(); val=SST[i] if i<len(SST) else 'SST#%d'%i
            elif rid in (8,9,10,11):
                if rid==8: val=r.wstr()
                elif rid==9: val=r.dbl()
                elif rid==10: val=bool(r.u8())
                else: val='#ERR%d'%r.u8()
                r.u16(); cce=r.u32(); rgce=p[r.p:r.p+cce]; r.skip(cce); cb=r.u32(); rgcb=p[r.p:r.p+cb]
            fml=None
            if rgce is not None:
                try: fml=parse_rgce(rgce,rgcb,SN,XT,names)
                except Exception as e: fml='<ERR %s>'%e
            yield row,col,val,fml
        elif rid==426:  # BrtShrFmla
            pass
