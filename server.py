from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler
from pathlib import Path
import openpyxl, json, datetime, io, collections, threading
from urllib.parse import urlparse, parse_qs
ROOT=Path(__file__).parent
TEMPLATE=ROOT/'data/template.xlsx'
STORE=ROOT/'data/entries.json'
wb=openpyxl.load_workbook(TEMPLATE,data_only=True)
s=wb['log']
rows=list(s.values)
def fmt(v):
    if isinstance(v,datetime.datetime): return v.strftime('%Y-%m-%d')
    if isinstance(v,datetime.time): return v.strftime('%H:%M:%S')
    if isinstance(v,float) and v.is_integer(): return str(int(v))
    return str(v).strip() if v is not None else ''
form=wb['RFI Form']
# Only values used by the RFI form belong in the entry workflow.
FORM_MAP={'B':['F2'],'C':['N2','AE2'],'E':['I3'],'H':['A5'],'I':['I5'],
 'O':['W5'],'R':['Z6'],'S':['A8'],'U':['I9'],'V':['G13'],'W':['O13'],
 'X':['Z13'],'Y':['G15','O15','Z15'],'Z':['G19'],'AA':['K19'],
 'AB':['Q19'],'AC':['AA19'],'AD':['G21','K21','Q21','AA21']}
FIXED={'B':fmt(form['F2'].value),'C':'CIV','I':'12:00','O':'1',
 'R':fmt(form['Z6'].value),'V':'**','W':'Haedr Ahmaed','X':'Dhiya Mahdi',
 'Z':'','AA':'','AB':'Haedr Ahmaed','AC':'Engr. Ga Zang Dong Zhi'}
FORM_SNAPSHOT={c.coordinate:c.value for row in form for c in row if c.value is not None}
fields=[]
def add(key,label,group,kind='fixed'):
 idx=openpyxl.utils.column_index_from_string(key)-1
 options=list(dict.fromkeys(fmt(r[idx]) for r in rows[6:] if r[0] is not None and fmt(r[idx]))) if kind=='select' else []
 fields.append(dict(id=key,label=label,group=group,type=kind,options=options,
  fixed=kind=='fixed',default=FIXED.get(key,'')))
for key,label,kind in [('B','Subcontractor','fixed'),('C','Discipline','fixed'),
 ('E','Area / Location','text'),('H','Date Of Inspection','date'),
 ('I','Time Of Inspection','fixed'),('O','ITP Item No.','fixed'),
 ('R','ITR No.','fixed'),('S','Inspection Activity / Description','select'),
 ('U','Drawing No. & Rev No.','select')]: add(key,label,'Request for inspection',kind)
for key,label in [('V','Sub-Contractor Construction'),('W','Al-Yamama Construction'),('X','Project Manager')]:
 add(key,label,'Submitted by · Checked & ready for inspection')
add('Y','Submission date — same for all three names','Submitted by · Checked & ready for inspection','date')
for key,label in [('Z','Subcontractor QC · Received by'),('AA','Subcontractor QC · Checked by'),
 ('AB','Al-Yamama QC · Received by'),('AC','Sinoma QC · Received by')]:
 add(key,label,'For quality control department use only')
add('AD','Receipt date — same for all four columns','For quality control department use only','date')
max_id=max(int(fmt(r[0])) for r in rows[6:] if fmt(r[0]).isdigit())
append_row=max(i+1 for i,r in enumerate(rows) if i>=6 and r[0] is not None)+1
del wb,s,rows

def entries(): return json.loads(STORE.read_text()) if STORE.exists() else []
def validate(d):
 out={}
 for f in fields:
    key=f['id'];v=str(d.get(key,''))
    if f['fixed']: v=f['default']
    elif f['type']=='date':
     try: datetime.date.fromisoformat(v)
     except ValueError: raise ValueError(f['label']+' requires a valid calendar date')
    elif f['type']=='select' and v not in f['options']: raise ValueError('Choose an option: '+f['label'])
    if f['type']=='text' and not v.strip(): raise ValueError(f['label']+' is required')
    out[key]=v
 return out
LOCK=threading.Lock()
def export_workbook(all,number=None):
 w=openpyxl.load_workbook(TEMPLATE)
 for sheet in list(w):
  if sheet.title not in ['RFI Form','log']: w.remove(sheet)
 sheet=w['log']
 for n,e in enumerate(all,start=append_row):
  sheet.cell(n,1,e['number'])
  vals=validate(e['values'])
  for key,val in vals.items():
   if key in ['H','Y','AD']: val=datetime.datetime.fromisoformat(val)
   elif key=='O': val=1
   cell=sheet.cell(n,openpyxl.utils.column_index_from_string(key))
   cell.value=val
   if isinstance(val,str): cell.data_type='s'
 selected=next((e for e in all if e['number']==number),None) if number else (all[-1] if all else None)
 if number and selected is None: raise ValueError('Record not found')
 f=w['RFI Form']
 for row in f:
  for cell in row:
   if cell.data_type=='f': cell.value=FORM_SNAPSHOT.get(cell.coordinate)
 # Freeze the unchanged lower form from its cached template values.
 for coordinate,value in FORM_SNAPSHOT.items():
  cell=f[coordinate];cell.value=value
  if isinstance(value,str):cell.data_type='s'
 if selected:
  vals=validate(selected['values']); f['AG2']=selected['number']
  for key,coords in FORM_MAP.items():
   value=vals[key]
   if key in ['H','Y','AD']: value=datetime.datetime.fromisoformat(value)
   elif key=='O':value=1
   for coord in coords:
    cell=f[coord];cell.value=value
    if key in ['H','Y','AD']:cell.number_format='yyyy/mm/dd'
    if isinstance(value,str):cell.data_type='s'
 b=io.BytesIO();w.save(b);return b.getvalue()
class Handler(SimpleHTTPRequestHandler):
 def __init__(self,*args,**kw): super().__init__(*args,directory=str(ROOT/'static'),**kw)
 def send(self,data,status=200,kind='application/json'):
    body=json.dumps(data).encode() if kind=='application/json' else data
    self.send_response(status);self.send_header('Content-Type',kind);self.send_header('Content-Length',str(len(body)));self.end_headers();self.wfile.write(body)
 def do_GET(self):
    if self.path=='/api/schema': return self.send(dict(fields=fields,nextId=max_id+1,project='AL-YAMAMA CEMENT PLANT 6000 TPD CLINKER PROJECT'))
    if self.path=='/api/entries': return self.send(entries())
    if urlparse(self.path).path=='/api/export':
      try:
       number=parse_qs(urlparse(self.path).query).get('number',[None])[0]
       payload=export_workbook(entries(),number)
      except ValueError as e: return self.send({'error':str(e)},400)
      self.send_response(200);self.send_header('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');self.send_header('Content-Disposition','attachment; filename="Yamama-RFI.xlsx"');self.end_headers();self.wfile.write(payload);return
    return super().do_GET()
 def do_POST(self):
    try:
     d=json.loads(self.rfile.read(int(self.headers.get('Content-Length',0)))); vals=validate(d['values'])
     with LOCK:
      all=entries()
      if d.get('number'):
       item=next(e for e in all if e['number']==d['number']);item['values']=vals
      else: all.append(dict(number=str(max_id+len(all)+1).zfill(4),values=vals))
      tmp=STORE.with_suffix('.tmp');tmp.write_text(json.dumps(all,indent=2));tmp.replace(STORE)
     self.send(all)
    except (ValueError,KeyError,StopIteration) as e:self.send({'error':str(e) or 'Record not found'},400)
if __name__=='__main__':
 print('RFI workspace running on port 8000',flush=True)
 ThreadingHTTPServer(('0.0.0.0',8000),Handler).serve_forever()
