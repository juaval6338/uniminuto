"""Minimal BIFF12 (.xlsb) record + formula (rgce) decoder."""
import struct, zipfile, io

FUNCS = {
0:'COUNT',1:'IF',2:'ISNA',3:'ISERROR',4:'SUM',5:'AVERAGE',6:'MIN',7:'MAX',8:'ROW',9:'COLUMN',10:'NA',11:'NPV',12:'STDEV',13:'DOLLAR',14:'FIXED',15:'SIN',16:'COS',17:'TAN',18:'ATAN',19:'PI',20:'SQRT',21:'EXP',22:'LN',23:'LOG10',24:'ABS',25:'INT',26:'SIGN',27:'ROUND',28:'LOOKUP',29:'INDEX',30:'REPT',31:'MID',32:'LEN',33:'VALUE',34:'TRUE',35:'FALSE',36:'AND',37:'OR',38:'NOT',39:'MOD',40:'DCOUNT',41:'DSUM',42:'DAVERAGE',43:'DMIN',44:'DMAX',45:'DSTDEV',46:'VAR',47:'DVAR',48:'TEXT',49:'LINEST',50:'TREND',51:'LOGEST',52:'GROWTH',56:'PV',57:'FV',58:'NPER',59:'PMT',60:'RATE',61:'MIRR',62:'IRR',63:'RAND',64:'MATCH',65:'DATE',66:'TIME',67:'DAY',68:'MONTH',69:'YEAR',70:'WEEKDAY',71:'HOUR',72:'MINUTE',73:'SECOND',74:'NOW',75:'AREAS',76:'ROWS',77:'COLUMNS',78:'OFFSET',82:'SEARCH',83:'TRANSPOSE',86:'TYPE',97:'ATAN2',98:'ASIN',99:'ACOS',100:'CHOOSE',101:'HLOOKUP',102:'VLOOKUP',105:'ISREF',109:'LOG',111:'CHAR',112:'LOWER',113:'UPPER',114:'PROPER',115:'LEFT',116:'RIGHT',117:'EXACT',118:'TRIM',119:'REPLACE',120:'SUBSTITUTE',121:'CODE',124:'FIND',125:'CELL',126:'ISERR',127:'ISTEXT',128:'ISNUMBER',129:'ISBLANK',130:'T',131:'N',140:'DATEVALUE',141:'TIMEVALUE',142:'SLN',143:'SYD',144:'DDB',148:'INDIRECT',162:'CLEAN',163:'MDETERM',164:'MINVERSE',165:'MMULT',167:'IPMT',168:'PPMT',169:'COUNTA',183:'PRODUCT',184:'FACT',189:'DPRODUCT',190:'ISNONTEXT',193:'STDEVP',194:'VARP',195:'DSTDEVP',196:'DVARP',197:'TRUNC',198:'ISLOGICAL',199:'DCOUNTA',212:'ROUNDUP',213:'ROUNDDOWN',216:'RANK',219:'ADDRESS',220:'DAYS360',221:'TODAY',227:'MEDIAN',228:'SUMPRODUCT',229:'SINH',230:'COSH',231:'TANH',235:'DGET',244:'INFO',247:'DB',252:'FREQUENCY',261:'ERROR.TYPE',269:'AVEDEV',270:'BETADIST',271:'GAMMALN',272:'BETAINV',273:'BINOMDIST',274:'CHIDIST',275:'CHIINV',276:'COMBIN',277:'CONFIDENCE',278:'CRITBINOM',279:'EVEN',280:'EXPONDIST',281:'FDIST',282:'FINV',283:'FISHER',284:'FISHERINV',285:'FLOOR',286:'GAMMADIST',287:'GAMMAINV',288:'CEILING',289:'HYPGEOMDIST',290:'LOGNORMDIST',291:'LOGINV',292:'NEGBINOMDIST',293:'NORMDIST',294:'NORMSDIST',295:'NORMINV',296:'NORMSINV',297:'STANDARDIZE',298:'ODD',299:'PERMUT',300:'POISSON',301:'TDIST',302:'WEIBULL',303:'SUMXMY2',304:'SUMX2MY2',305:'SUMX2PY2',306:'CHITEST',307:'CORREL',308:'COVAR',309:'FORECAST',310:'FTEST',311:'INTERCEPT',312:'PEARSON',313:'RSQ',314:'STEYX',315:'SLOPE',316:'TTEST',317:'PROB',318:'DEVSQ',319:'GEOMEAN',320:'HARMEAN',321:'SUMSQ',322:'KURT',323:'SKEW',324:'ZTEST',325:'LARGE',326:'SMALL',327:'QUARTILE',328:'PERCENTILE',329:'PERCENTRANK',330:'MODE',331:'TRIMMEAN',332:'TINV',336:'CONCATENATE',337:'POWER',342:'RADIANS',343:'DEGREES',344:'SUBTOTAL',345:'SUMIF',346:'COUNTIF',347:'COUNTBLANK',350:'ISPMT',351:'DATEDIF',352:'DATESTRING',353:'NUMBERSTRING',354:'ROMAN',358:'GETPIVOTDATA',359:'HYPERLINK',360:'PHONETIC',361:'AVERAGEA',362:'MAXA',363:'MINA',364:'STDEVPA',365:'VARPA',366:'STDEVA',367:'VARA',368:'BAHTTEXT',369:'THAIDAYOFWEEK',370:'THAIDIGIT',371:'THAIMONTHOFYEAR',372:'THAINUMSOUND',373:'THAINUMSTRING',374:'THAISTRINGLENGTH',375:'ISTHAIDIGIT',376:'ROUNDBAHTDOWN',377:'ROUNDBAHTUP',378:'THAIYEAR',379:'RTD',380:'CUBEVALUE',381:'CUBEMEMBER',382:'CUBEMEMBERPROPERTY',383:'CUBERANKEDMEMBER',384:'HEX2BIN',385:'HEX2DEC',386:'HEX2OCT',387:'DEC2BIN',388:'DEC2HEX',389:'DEC2OCT',390:'OCT2BIN',391:'OCT2HEX',392:'OCT2DEC',393:'BIN2DEC',394:'BIN2OCT',395:'BIN2HEX',396:'IMSUB',397:'IMDIV',398:'IMPOWER',399:'IMABS',400:'IMSQRT',401:'IMLN',402:'IMLOG2',403:'IMLOG10',404:'IMSIN',405:'IMCOS',406:'IMEXP',407:'IMARGUMENT',408:'IMCONJUGATE',409:'IMAGINARY',410:'IMREAL',411:'COMPLEX',412:'IMSUM',413:'IMPRODUCT',414:'SERIESSUM',415:'FACTDOUBLE',416:'SQRTPI',417:'QUOTIENT',418:'DELTA',419:'GESTEP',420:'ISEVEN',421:'ISODD',422:'MROUND',423:'ERF',424:'ERFC',425:'BESSELJ',426:'BESSELK',427:'BESSELY',428:'BESSELI',429:'XIRR',430:'XNPV',431:'PRICEMAT',432:'YIELDMAT',433:'INTRATE',434:'RECEIVED',435:'DISC',436:'PRICEDISC',437:'YIELDDISC',438:'TBILLEQ',439:'TBILLPRICE',440:'TBILLYIELD',441:'PRICE',442:'YIELD',443:'DOLLARDE',444:'DOLLARFR',445:'NOMINAL',446:'EFFECT',447:'CUMPRINC',448:'CUMIPMT',449:'EDATE',450:'EOMONTH',451:'YEARFRAC',452:'COUPDAYBS',453:'COUPDAYS',454:'COUPDAYSNC',455:'COUPNCD',456:'COUPNUM',457:'COUPPCD',458:'DURATION',459:'MDURATION',460:'ODDLPRICE',461:'ODDLYIELD',462:'ODDFPRICE',463:'ODDFYIELD',464:'RANDBETWEEN',465:'WEEKNUM',466:'AMORDEGRC',467:'AMORLINC',468:'CONVERT',469:'ACCRINT',470:'ACCRINTM',471:'WORKDAY',472:'NETWORKDAYS',473:'GCD',474:'MULTINOMIAL',475:'LCM',476:'FVSCHEDULE',477:'CUBEKPIMEMBER',478:'CUBESET',479:'CUBESETCOUNT',480:'IFERROR',481:'COUNTIFS',482:'SUMIFS',483:'AVERAGEIF',484:'AVERAGEIFS',485:'AGGREGATE',486:'BINOM.DIST',
}

def colname(c):
    s=''
    c=int(c)
    while True:
        s=chr(ord('A')+c%26)+s
        c=c//26-1
        if c<0: break
    return s

def a1(rw,col,relrow=False,relcol=False,abs_=True):
    cs=('' if relcol else '$')+colname(col)
    rs=('' if relrow else '$')+str(rw+1)
    return cs+rs

class Reader:
    def __init__(self,data):
        self.d=data; self.p=0
    def u8(self):
        v=self.d[self.p]; self.p+=1; return v
    def u16(self):
        v=struct.unpack_from('<H',self.d,self.p)[0]; self.p+=2; return v
    def i16(self):
        v=struct.unpack_from('<h',self.d,self.p)[0]; self.p+=2; return v
    def u32(self):
        v=struct.unpack_from('<I',self.d,self.p)[0]; self.p+=4; return v
    def i32(self):
        v=struct.unpack_from('<i',self.d,self.p)[0]; self.p+=4; return v
    def dbl(self):
        v=struct.unpack_from('<d',self.d,self.p)[0]; self.p+=8; return v
    def skip(self,n): self.p+=n
    def wstr(self):
        n=self.u32()
        if n==0xFFFFFFFF: return ''
        s=self.d[self.p:self.p+2*n].decode('utf-16-le','replace'); self.p+=2*n; return s
    def eof(self): return self.p>=len(self.d)

def records(data):
    """Yield (recid, payload) for each BIFF12 record."""
    p=0; n=len(data)
    while p<n:
        b=data[p]; p+=1
        if b & 0x80:
            b2=data[p]; p+=1
            rid=(b & 0x7F) | ((b2 & 0x7F)<<7)
        else:
            rid=b
        sz=0; shift=0
        for _ in range(4):
            c=data[p]; p+=1
            sz |= (c & 0x7F)<<shift; shift+=7
            if not (c & 0x80): break
        yield rid, data[p:p+sz]
        p+=sz

def parse_loc(r, base_r=0, base_c=0, rel=False):
    rw=r.u32(); cc=r.u16()
    col=cc & 0x3FFF
    relcol=bool(cc & 0x4000); relrow=bool(cc & 0x8000)
    if rel:
        if relrow: rw=(base_r+struct.unpack('<i',struct.pack('<I',rw))[0]) if rw<0x80000000 else base_r-(0x100000000-rw)
        if relcol:
            col=col if col<0x2000 else col-0x4000
            col=base_c+col
    return rw,col,relrow,relcol

def loc_str(rw,col,relrow,relcol):
    return ('' if relcol else '$')+colname(col)+('' if relrow else '$')+str(rw+1)

def parse_rgce(rgce, rgcb, sheetnames, xtis, names, cur_sheet=''):
    """Return an infix formula string."""
    r=Reader(rgce)
    st=[]
    def sheetref(ixti):
        try:
            sup,f,l=xtis[ixti]
        except Exception:
            return "[?%d]"%ixti
        try:
            a=sheetnames[f]; b=sheetnames[l]
        except Exception:
            return "[?]"
        nm = a if a==b else a+':'+b
        if any(ch in nm for ch in " -()'"): nm="'"+nm.replace("'","''")+"'"
        return nm+'!'
    while not r.eof():
        t=r.u8()
        base = t if t<0x20 else (0x20 | (t & 0x1F))
        if t==0x01:   # PtgExp -> shared/array parent row
            rw=r.u32(); st.append('{SHARED_ROW:%d}'%rw); continue
        if t==0x02:
            r.u32(); r.u32(); st.append('{table}'); continue
        if 0x03<=t<=0x11:
            op={3:'+',4:'-',5:'*',6:'/',7:'^',8:'&',9:'<',10:'<=',11:'=',12:'>=',13:'>',14:'<>',15:' ',16:',',17:':'}[t]
            b=st.pop() if st else '?'; a=st.pop() if st else '?'
            st.append(a+op+b); continue
        if t==0x12: st.append('+'+(st.pop() if st else '')); continue
        if t==0x13: st.append('-'+(st.pop() if st else '')); continue
        if t==0x14: st.append((st.pop() if st else '')+'%'); continue
        if t==0x15: st.append('('+(st.pop() if st else '')+')'); continue
        if t==0x16: st.append(''); continue
        if t==0x17:
            n=r.u16(); sv=r.d[r.p:r.p+2*n].decode('utf-16-le','replace'); r.p+=2*n
            st.append('"'+sv.replace('"','""')+'"'); continue
        if t==0x18:
            g=r.u8()
            if g==0x01: r.u16(); r.u32()
            else: r.skip(4)
            continue
        if t==0x19:
            g=r.u8()
            if g & 0x04:
                n=r.u16(); r.skip((n+1)*2)
            else:
                r.skip(2)
            continue
        if t==0x1C: st.append({0:'#NULL!',7:'#DIV/0!',15:'#VALUE!',23:'#REF!',29:'#NAME?',36:'#NUM!',42:'#N/A'}.get(r.u8(),'#ERR')); continue
        if t==0x1D: st.append('TRUE' if r.u8() else 'FALSE'); continue
        if t==0x1E: st.append(str(r.u16())); continue
        if t==0x1F:
            v=r.dbl(); st.append(('%r'%v).rstrip('0').rstrip('.') if v!=int(v) else str(int(v))); continue
        if base==0x20:  # PtgArray
            r.skip(14); st.append('{array}'); continue
        if base==0x21:  # PtgFunc
            ift=r.u16(); nm=FUNCS.get(ift,'FUNC%d'%ift)
            # fixed arg count unknown; pop greedily by known arity table fallback 1
            ar=FIXED_ARITY.get(ift, 1)
            args=[st.pop() if st else '' for _ in range(ar)][::-1]
            st.append('%s(%s)'%(nm,','.join(args))); continue
        if base==0x22:  # PtgFuncVar
            cp=r.u8() & 0x7F; ift=r.u16() & 0x7FFF
            args=[st.pop() if st else '' for _ in range(cp)][::-1]
            if ift==255:
                nm=args[0] if args else '?'
                st.append('%s(%s)'%(nm.strip('"'),','.join(args[1:])))
            else:
                st.append('%s(%s)'%(FUNCS.get(ift,'FUNC%d'%ift),','.join(args)))
            continue
        if base==0x23:  # PtgName
            idx=r.u32(); st.append(names.get(idx,'NAME%d'%idx)); continue
        if base==0x24:  # PtgRef
            rw,col,rr,rc=parse_loc(r); st.append(loc_str(rw,col,rr,rc)); continue
        if base==0x25:  # PtgArea
            r1=r.u32(); r2=r.u32(); c1=r.u16(); c2=r.u16()
            st.append(loc_str(r1,c1&0x3FFF,bool(c1&0x8000),bool(c1&0x4000))+':'+loc_str(r2,c2&0x3FFF,bool(c2&0x8000),bool(c2&0x4000))); continue
        if base in (0x26,0x27,0x28):  # PtgMemArea/Err/NoMem
            r.skip(4); r.u16(); continue
        if base==0x29:  # PtgMemFunc
            r.u16(); continue
        if base==0x2A: r.skip(6); st.append('#REF!'); continue
        if base==0x2B: r.skip(12); st.append('#REF!'); continue
        if base==0x2C:
            rw,col,rr,rc=parse_loc(r); st.append(loc_str(rw,col,rr,rc)); continue
        if base==0x2D:
            r1=r.u32(); r2=r.u32(); c1=r.u16(); c2=r.u16()
            st.append(loc_str(r1,c1&0x3FFF,bool(c1&0x8000),bool(c1&0x4000))+':'+loc_str(r2,c2&0x3FFF,bool(c2&0x8000),bool(c2&0x4000))); continue
        if base==0x39:  # PtgNameX
            ix=r.u16(); idx=r.u32(); st.append(names.get(idx,'NAME%d'%idx)); continue
        if base==0x3A:  # PtgRef3d
            ix=r.u16(); rw,col,rr,rc=parse_loc(r)
            st.append(sheetref(ix)+loc_str(rw,col,rr,rc)); continue
        if base==0x3B:  # PtgArea3d
            ix=r.u16(); r1=r.u32(); r2=r.u32(); c1=r.u16(); c2=r.u16()
            st.append(sheetref(ix)+loc_str(r1,c1&0x3FFF,bool(c1&0x8000),bool(c1&0x4000))+':'+loc_str(r2,c2&0x3FFF,bool(c2&0x8000),bool(c2&0x4000))); continue
        if base==0x3C: r.skip(8); st.append('#REF!'); continue
        if base==0x3D: r.skip(14); st.append('#REF!'); continue
        st.append('<?ptg%02X>'%t)
        break
    return st[-1] if st else ''

FIXED_ARITY={2:1,3:1,8:1,9:1,10:0,15:1,16:1,17:1,18:1,19:0,20:1,21:1,22:1,23:1,24:1,25:1,26:1,30:2,32:1,33:1,34:0,35:0,38:1,39:2,48:2,63:0,67:1,68:1,69:1,71:1,72:1,73:1,74:0,75:1,76:1,77:1,82:2,83:1,86:1,97:2,98:1,99:1,111:1,112:1,113:1,114:1,117:2,118:1,121:1,126:1,127:1,128:1,129:1,130:1,131:1,140:1,141:1,163:1,164:1,165:2,169:1,184:1,190:1,197:1,198:1,205:1,212:2,213:2,221:0,229:1,230:1,231:1,261:1,271:1,276:2,279:1,283:1,284:1,285:2,288:2,298:1,303:2,304:2,305:2,307:2,308:2,310:2,311:2,312:2,313:2,314:2,315:2,318:1,319:1,320:1,321:1,322:1,323:1,336:-1,337:2,342:1,343:1,347:1,373:1,384:1,385:1,415:1,416:1,417:2,418:2,420:1,421:1,422:2,423:1,424:1,480:2}
