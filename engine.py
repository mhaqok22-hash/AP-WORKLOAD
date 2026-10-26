"""AP reporting engine: preserves original Excel objects via targeted OOXML editing.
No external Excel executable required. Input must have the 19 standard AP columns.
"""
from __future__ import annotations
import io, zipfile, re, math, statistics, collections, calendar
from datetime import datetime, date, timedelta
from copy import deepcopy
from lxml import etree as ET

S='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
C='http://schemas.openxmlformats.org/drawingml/2006/chart'
NS={'m':S,'c':C}
Q=lambda x:'{%s}%s'%(S,x)
CQ=lambda x:'{%s}%s'%(C,x)
BASE=date(1899,12,30)
INPUT_FIELDS=['ID','Status','Invoice Number','Vendor','Invoice Date','Date Received','Amount','Aging Received','Urgent','Category','Fiscal Year','PO Number','Department','Fund','Invoice Approver','Invoice Approver 2','AP Payer','PS Entry Date','Voucher Number']

def serial(v):
 if isinstance(v,datetime):v=v.date()
 if isinstance(v,date):return (v-BASE).days
 if isinstance(v,(int,float)):return v
 return None

def as_date(v):
 if isinstance(v,datetime):return v.date()
 if isinstance(v,date):return v
 if isinstance(v,(int,float)) and 1<=v<100000:return BASE+timedelta(days=int(v))
 if isinstance(v,str):
  for fmt in ('%Y-%m-%d','%m/%d/%Y','%m/%d/%y','%Y/%m/%d','%m-%d-%Y'):
   try:return datetime.strptime(v.strip(),fmt).date()
   except ValueError:pass
 return None

def number(v):
 try:return float(str(v).replace('$','').replace(',','')) if v not in (None,'') else 0.0
 except (TypeError,ValueError):return 0.0

def read_upload(raw:bytes):
 """Parse raw input with stdlib ZIP/XML (no pandas/openpyxl conversion)."""
 with zipfile.ZipFile(io.BytesIO(raw)) as z:
  shared=[]
  if 'xl/sharedStrings.xml' in z.namelist():
   root=ET.fromstring(z.read('xl/sharedStrings.xml'))
   shared=[''.join(e.itertext()) for e in root.findall(Q('si'))]
  sheet='xl/worksheets/sheet1.xml'
  rows=[]
  with z.open(sheet) as f:
   for _,e in ET.iterparse(f,events=('end',),tag=Q('row')):
    values={}
    for cell in e.findall(Q('c')):
     match=re.match(r'([A-Z]+)',cell.get('r',''))
     if not match:continue
     col=match.group(1)
     value=cell.find(Q('v'))
     inline=cell.find(Q('is'))
     if cell.get('t')=='s' and value is not None:val=shared[int(value.text)]
     elif cell.get('t')=='inlineStr':val=''.join(inline.itertext()) if inline is not None else ''
     elif value is not None:
      try: val=float(value.text);val=int(val) if val.is_integer() else val
      except (ValueError,AttributeError):val=value.text
     else:val=None
     values[col]=val
    rows.append(values)
    e.clear()
    while e.getprevious() is not None:del e.getparent()[0]
 if not rows:raise ValueError('No input records.')
 def ncol(n):
  col=''
  while n:n,r=divmod(n-1,26);col=chr(65+r)+col
  return col
 first=rows[0]
 headers=[str(first.get(ncol(k),'') or '').strip() for k in range(1,20)]
 normalized=lambda x:re.sub(r'[^a-z0-9]','',str(x).lower())
 expected=[normalized(x) for x in INPUT_FIELDS]
 if [normalized(h) for h in headers]!=expected:
  # tolerate reordered columns where all names are known
  if not all(x in [normalized(h) for h in headers] for x in expected):
   raise ValueError('Input requires the 19 source fields: '+', '.join(INPUT_FIELDS)+'. Found: '+str(headers))
  mapping=[ncol([normalized(h) for h in headers].index(x)+1) for x in expected]
 else:mapping=[ncol(k) for k in range(1,20)]
 result=[]
 for row in rows[1:]:
  vals=[row.get(col) for col in mapping]
  if not any(v not in (None,'') for v in vals):continue
  for i in (4,5,17):vals[i]=as_date(vals[i])
  vals[6]=number(vals[6])
  result.append(vals)
 if not result:raise ValueError('No invoice records beneath header row.')
 return result

def analyze(records,report_date):
 cycles=[];included=[];pending=[];statuses=collections.Counter();payer=set();exceptions=collections.Counter()
 monthly=collections.defaultdict(lambda:collections.Counter())
 for i,r in enumerate(records):
  st=str(r[1] or '').strip();statuses[st]+=1
  if r[16] not in (None,''):payer.add(str(r[16]))
  recv,paid=r[5],r[17];ispaid=st.casefold()=='paid'
  valid=ispaid and recv is not None and paid is not None and paid>=recv
  included.append(valid)
  if valid:cycles.append((paid-recv).days)
  elif ispaid:
   if paid is None or recv is None:exceptions['missing_date']+=1
   else:exceptions['negative_cycle']+=1
  if not ispaid:
   if recv is not None:pending.append((i,(report_date-recv).days,r))
   else:exceptions['pending_missing_date']+=1
  if recv:monthly[(recv.year,recv.month)]['received']+=1
  if ispaid and paid:monthly[(paid.year,paid.month)]['paid']+=1
 latest=max((r[5] for r in records if r[5]),default=report_date)
 cycles.sort()
 return {'count':len(records),'paid':statuses.get('Paid',0),'pending':len(records)-statuses.get('Paid',0),'cycles':cycles,'eligible':len(cycles),'cycle_total':sum(cycles),'average':statistics.mean(cycles) if cycles else 0,'median':statistics.median(cycles) if cycles else 0,'pending_records':sorted(pending,key=lambda t:(-t[1],t[0])),'payer_count':len(payer),'statuses':statuses,'exceptions':exceptions,'monthly':monthly,'latest_receipt':latest,'included':included}

def set_cell(parent,ref,value=None,formula=None,kind=None,style=None):
 c=parent.find(f"{Q('c')}[@r='{ref}']")
 if c is None:
  c=ET.SubElement(parent,Q('c'),r=ref)
 for child in list(c):c.remove(child)
 c.attrib.pop('t',None)
 if style is not None:c.set('s',str(style))
 if formula is not None:
  ET.SubElement(c,Q('f')).text=formula.lstrip('=')
  if isinstance(value,(int,float)) and math.isfinite(value):ET.SubElement(c,Q('v')).text=str(value)
  elif isinstance(value,str):c.set('t','str');ET.SubElement(c,Q('v')).text=value
  else:ET.SubElement(c,Q('v'))
 elif value is None:pass
 elif isinstance(value,(int,float)) and math.isfinite(value):ET.SubElement(c,Q('v')).text=str(value)
 else:
  c.set('t','inlineStr');ET.SubElement(ET.SubElement(c,Q('is')),Q('t')).text=str(value)
 return c

def cellmap(root):
 data=root.find(Q('sheetData'))
 return {int(r.get('r')):r for r in data.findall(Q('row'))}

def ensure_row(root,mapping,n):
 if n in mapping:return mapping[n]
 data=root.find(Q('sheetData'))
 new=ET.Element(Q('row'),r=str(n))
 for idx,row in enumerate(data):
  if int(row.get('r'))>n:data.insert(idx,new);break
 else:data.append(new)
 mapping[n]=new;return new

def overwrite_master(root,records,report_date,stats):
 maps=cellmap(root)
 # Keep the existing 20,000 prepared row styles/formulas, replace actual data rows only.
 max_rows=20000
 if len(records)>max_rows:raise ValueError(f'Template supports up to {max_rows:,} invoices; received {len(records):,}. Expand the Master Data template capacity before processing.')
 for idx in range(max_rows):
  row=maps.get(idx+2)
  if row is None:continue
  vals=records[idx] if idx<len(records) else None
  for c in row.findall(Q('c')):
   ref=c.get('r','');col=re.match(r'[A-Z]+',ref).group()
   def colnum(s):
    n=0
    for ch in s:n=n*26+ord(ch)-64
    return n
   k=colnum(col)
   if 1<=k<=19:
    oldstyle=c.get('s')
    v=vals[k-1] if vals is not None else None
    set_cell(row,ref,serial(v) if isinstance(v,(date,datetime)) else v,style=oldstyle)
   elif 20<=k<=36:
    # Retain preexisting helper formulas, but clear cached values to avoid stale results.
    f=c.find(Q('f'))
    if f is not None:
     v=c.find(Q('v'))
     if v is not None:c.remove(v)
 # Date and fixed reporting baseline control
 row=ensure_row(root,maps,2)
 set_cell(row,'AE2',serial(report_date))
 # Replace master rank key: only pending, with deterministic tie break.
 # For Excel older-version compatibility uses row-level formula, but many ranks may calculate slowly.
 for idx in range(len(records)):
  row=maps.get(idx+2)
  if row is None:continue
  for col,f in {
   'T':f'IF(OR(R{idx+2}="",F{idx+2}=""),"",R{idx+2}-F{idx+2})',
   'U':f'IF(AND(B{idx+2}="Paid",ISNUMBER(T{idx+2}),T{idx+2}>=0),1,0)',
   'V':f'IF(B{idx+2}="Paid",1,0)',
   'W':f'IF(AND(B{idx+2}<>"Paid",B{idx+2}<>""),1,0)',
   'X':f'IF(F{idx+2}="","",$AE$2-F{idx+2})',
   'AF':f'IF(AND(W{idx+2}=1,ISNUMBER(X{idx+2})),X{idx+2}+ROW()/10000000,0)',
  }.items():set_cell(row,f'{col}{idx+2}',formula=f)
 # Empty row U/W needs zero caches and must not count as invoice
 return root

def chart_update(xml,chart_idx,months,monthly_stats,report_date):
 root=ET.fromstring(xml)
 if chart_idx not in (1,2,17,18,19):return xml
 # Chart 1 and 17: received & paid, 2/19: backlog, 18: cycle.
 series=root.findall('.//'+CQ('ser'))
 desired=([('B','received'),('C','paid')] if chart_idx in (1,17) else [('I','backlog')] if chart_idx in (2,19) else [('F','cycle')])
 for ser,(col,metric) in zip(series,desired):
  cat=ser.find(CQ('cat'));val=ser.find(CQ('val'))
  if cat is None or val is None:continue
  # Use internal static ranges for universal Excel chart display; regenerated on each upload.
  for node,form in ((cat,f"'Monthly Summary'!$A$5:$A${4+months}"),(val,f"'Monthly Summary'!${col}$5:${col}${4+months}")):
   f=node.find('.//'+CQ('f'))
   if f is not None:f.text=form
  # Rebuild cache to display instantly without desktop recalculation.
  for parent,series_values,is_date in ((cat,[serial(date(y,m,1)) for y,m in monthly_stats['periods']],True),(val,monthly_stats[metric],False)):
   for cache in list(parent):
    if ET.QName(cache).localname in ('numRef','strRef'):
     ref=cache;oldcache=ref.find(CQ('numCache'))
     if oldcache is not None:ref.remove(oldcache)
     if ref.find(CQ('f')) is None:ET.SubElement(ref,CQ('f'))
     cache=ET.SubElement(ref,CQ('numCache'))
     ET.SubElement(cache,CQ('formatCode')).text='mmm yyyy' if is_date else 'General'
     ET.SubElement(cache,CQ('ptCount'),val=str(len(series_values)))
     for i,v in enumerate(series_values):
      p=ET.SubElement(cache,CQ('pt'),idx=str(i));ET.SubElement(p,CQ('v')).text=str(v or 0)
     break
 return ET.tostring(root,encoding='UTF-8',xml_declaration=True)

def month_series(records,stats,asof,first=(2025,12),limit=84):
 # Use Master Data received and paid dates; avoid 1900 and trailing months.
 start=date(*first,1);end=date(asof.year,asof.month,1)
 periods=[];cur=start
 while cur<=end and len(periods)<limit:
  periods.append((cur.year,cur.month));cur=date(cur.year+cur.month//12,cur.month%12+1,1)
 received=[];paid=[];cycle=[];backlog=[]
 for year,month in periods:
  endmonth=date(year,month,calendar.monthrange(year,month)[1]);endmonth=min(endmonth,asof)
  rcount=sum(1 for r in records if r[5] and r[5].year==year and r[5].month==month)
  pcount=sum(1 for r in records if str(r[1]).strip()=='Paid' and r[17] and r[17].year==year and r[17].month==month)
  valid=[(r[17]-r[5]).days for r in records if str(r[1]).strip()=='Paid' and r[5] and r[17] and r[17]>=r[5] and r[17].year==year and r[17].month==month]
  pending_count=sum(1 for r in records if r[5] and r[5]<=endmonth and (str(r[1]).strip()!='Paid' or not r[17] or r[17]>endmonth))
  received.append(rcount);paid.append(pcount);cycle.append(sum(valid)/len(valid) if valid else 0);backlog.append(pending_count)
 return {'periods':periods,'received':received,'paid':paid,'cycle':cycle,'backlog':backlog}

def update_month_sheet(xml,records,asof,ms):
 root=ET.fromstring(xml);mapping=cellmap(root)
 for i,(year,month) in enumerate(ms['periods']):
  n=i+5; row=ensure_row(root,mapping,n)
  set_cell(row,'A'+str(n),serial(date(year,month,1)))
  # Keep original formula logic for automatic edits in Excel and copy numeric caches for immediate visibility.
  for col,key in [('B','received'),('C','paid'),('F','cycle'),('I','backlog')]:
   c=row.find(f"{Q('c')}[@r='{col}{n}']")
   formula=c.find(Q('f')).text if c is not None and c.find(Q('f')) is not None else None
   if formula is None:
    formula={'B':f'IF(A{n}>EOMONTH(\'Master Data\'!$AE$2,0),"",COUNTIF(\'Master Data\'!$Y$2:$Y$20001,A{n}))','C':f'IF(A{n}>EOMONTH(\'Master Data\'!$AE$2,0),"",COUNTIF(\'Master Data\'!$Z$2:$Z$20001,A{n}))','F':f'IF(A{n}>EOMONTH(\'Master Data\'!$AE$2,0),"",IFERROR(AVERAGEIFS(\'Master Data\'!$T$2:$T$20001,\'Master Data\'!$Z$2:$Z$20001,A{n},\'Master Data\'!$U$2:$U$20001,1),0))','I':f'IF(A{n}>EOMONTH(\'Master Data\'!$AE$2,0),"",COUNTIFS(\'Master Data\'!$F$2:$F$20001,"<="&EOMONTH(A{n},0),\'Master Data\'!$W$2:$W$20001,1))'}[col]
   set_cell(row,f'{col}{n}',ms[key][i],formula=formula)
 # existing future months: formulas remain blank after cutoff
 return ET.tostring(root,encoding='UTF-8',xml_declaration=True)

def repair_chart_workbook(root):
 # Clear stale external filename prefixes from all workbook-local chart defined names, keeping the names themselves available for future Excel edits.
 for elt in root.findall('.//'+Q('definedName')):
  if elt.text and elt.get('name','').startswith('AP_'):
   elt.text=re.sub(r"'[^']*\.xlsx'!",'',elt.text)
 return root

def build(template:bytes,records,asof:date):
 stats=analyze(records,asof)
 with zipfile.ZipFile(io.BytesIO(template)) as source:
  paths=set(source.namelist())
  sheet10=ET.fromstring(source.read('xl/worksheets/sheet10.xml'))
  overwrite_master(sheet10,records,asof,stats)
  ms=month_series(records,stats,asof)
  if not ms['periods']:raise ValueError('Reporting date precedes the template monthly timeline (December 2025).')
  changes={'xl/worksheets/sheet10.xml':ET.tostring(sheet10,encoding='UTF-8',xml_declaration=True)}
  changes['xl/worksheets/sheet8.xml']=update_month_sheet(source.read('xl/worksheets/sheet8.xml'),records,asof,ms)
  # Update cycle/audit figure caches, leaving linked formulas active.
  for sn in (1,3,6,7):
   path=f'xl/worksheets/sheet{sn}.xml';root=ET.fromstring(source.read(path));mapping=cellmap(root)
   if sn==6:
    for cell,val in [('B5',stats['average']),('B6',stats['median']),('B7',min(stats['cycles']) if stats['cycles'] else 0),('B8',max(stats['cycles']) if stats['cycles'] else 0)]:
     row=mapping[int(re.search(r'\d+',cell).group())]; c=row.find(f"{Q('c')}[@r='{cell}']")
     if c is not None:
      f=c.find(Q('f'));set_cell(row,cell,val,formula=f.text if f is not None else None)
   if sn==1:
    for cell,val in [('D15',stats['median']),('D16',stats['median'])]:
     row=mapping.get(int(re.search(r'\d+',cell).group()))
     if row is None:continue
     c=row.find(f"{Q('c')}[@r='{cell}']")
     if c is not None and c.find(Q('f')) is not None:set_cell(row,cell,val,formula=c.find(Q('f')).text)
   if sn==3:
    for ref,value in [('F20',sum(1 for _,age,_ in stats['pending_records'] if age>=91)),('F23',stats['payer_count'])]:
     row=mapping.get(int(re.search(r'\d+',ref).group()))
     if row is not None:
      c=row.find(f"{Q('c')}[@r='{ref}']")
      if c is not None:
       f=c.find(Q('f'));set_cell(row,ref,value,formula=f.text if f is not None else None)
   changes[path]=ET.tostring(root,encoding='UTF-8',xml_declaration=True)
  workbook=repair_chart_workbook(ET.fromstring(source.read('xl/workbook.xml')))
  calc=workbook.find(Q('calcPr'))
  if calc is None:calc=ET.SubElement(workbook,Q('calcPr'))
  calc.set('calcMode','auto');calc.set('fullCalcOnLoad','1');calc.set('forceFullCalc','1')
  changes['xl/workbook.xml']=ET.tostring(workbook,encoding='UTF-8',xml_declaration=True)
  for k in (1,2,17,18,19):
   path=f'xl/charts/chart{k}.xml';changes[path]=chart_update(source.read(path),k,len(ms['periods']),ms,asof)
  out=io.BytesIO()
  with zipfile.ZipFile(out,'w',compression=zipfile.ZIP_DEFLATED,compresslevel=6) as dest:
   for item in source.infolist():
    if item.filename=='xl/calcChain.xml':continue
    payload=changes.get(item.filename)
    if payload is None:payload=source.read(item.filename)
    if item.filename=='xl/_rels/workbook.xml.rels':
     rel=ET.fromstring(payload)
     for child in list(rel):
      if child.get('Type','').endswith('/calcChain'):rel.remove(child)
     payload=ET.tostring(rel,encoding='UTF-8',xml_declaration=True)
    if item.filename=='[Content_Types].xml':
     ct=ET.fromstring(payload)
     for child in list(ct):
      if child.get('PartName')=='/xl/calcChain.xml':ct.remove(child)
     payload=ET.tostring(ct,encoding='UTF-8',xml_declaration=True)
    dest.writestr(item,payload)
  result=out.getvalue()
  with zipfile.ZipFile(io.BytesIO(result)) as test:
   assert test.testzip() is None
   for p in test.namelist():
    if p.endswith(('.xml','.rels')):ET.fromstring(test.read(p))
  return result,stats,ms
