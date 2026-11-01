from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler
from pathlib import Path
import openpyxl, json, datetime, io, collections
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
groups=[('Request details', [('B','SCTR'),('C','DISC'),('D','PR N°'),('E','AREA / LOCATION'),('G','SUB-SYSTEM N°'),('H','DATE OF INSPECTION'),('I','TIME'),('N','ITP N°'),('O','ITP ITEM N°'),('P','QC CODE (If applicable)'),('Q','FORM N°'),('R','ITR N°'),('S','INSPECTION ACTIVITY / DESCRIPTION'),('T','ITEM TAG N°'),('U','DOCS REFERENCE & REV N°')]),('Intervention point',[(c,l) for c,l in zip('JKLM',['ABG','BH','SGC','TP'])]),('Ready for inspection',[(c,l) for c,l in zip('VWXY',['Sub. CONST','Construction Engineer','Construction Manager','Ready date'])]),('Quality control',[(c,l) for c,l in zip(['Z','AA','AB','AC','AD'],['Sub. QC — Received','Sub. QC — Checked','CTR QC — Received','CPY QC — Received','QC date'])]),('Results & closure',[(c,l) for c,l in zip(['AE','AF','AG','AH','AI','AJ','AK','AL','AM','AN','AO','AP','AQ','AR'],['RFI ON HOLD REPORT','Hold date','ACTION','Action date','RESULT','REASON FOR CANCELLATION / NOT READY / REJECTION','COMMENTS / REMARKS','PL N°','Inspected by — Sub. QC','Inspected by — Al-Yamama','Inspected by — Sinoma','Inspection confirmation date','RFI CLOSURE DATE','NCR N° (If any)'])])]
fields=[]
for group,cols in groups:
 for col,label in cols:
    idx=openpyxl.utils.column_index_from_string(col)-1
    vals=[fmt(r[idx]) for r in rows[6:] if r[0] is not None and r[idx] is not None]
    counts=collections.Counter(v for v in vals if v)
    options=[v for v,_ in counts.most_common()]
    if col in ['AE','AG']: options=['YES','NO']
    if col=='AI': options=['A','PA','H','C','R','NR']
    if col=='AJ': options=[fmt(wb['RFI Form'][c].value) for c in ['A33','M33','Y33','A34','M34','Y34','A35','M35','Y35','A36','M36','Y36','A37','M37']]
    fields.append(dict(id=col,label=label,group=group,options=options,fixed=len(options)==1 and col not in ['E','AI','AJ'],default=options[0] if options else ''))
max_id=max(int(fmt(r[0])) for r in rows[6:] if fmt(r[0]).isdigit())
append_row=max(i+1 for i,r in enumerate(rows) if i>=6 and r[0] is not None)+1
del wb,s,rows

def entries(): return json.loads(STORE.read_text()) if STORE.exists() else []
def validate(d):
 out={}
 for f in fields:
    v=str(d.get(f['id'],''))
    if f['id']!='E' and v and v not in f['options']: raise ValueError('Invalid option: '+f['label'])
    if f['fixed']: v=f['default']
    out[f['id']]=v
 if not out['E'].strip(): raise ValueError('AREA / LOCATION is required')
 if not out['H']: raise ValueError('DATE OF INSPECTION is required')
 return out
class Handler(SimpleHTTPRequestHandler):
 def __init__(self,*args,**kw): super().__init__(*args,directory=str(ROOT/'static'),**kw)
 def send(self,data,status=200,kind='application/json'):
    body=json.dumps(data).encode() if kind=='application/json' else data
    self.send_response(status);self.send_header('Content-Type',kind);self.send_header('Content-Length',str(len(body)));self.end_headers();self.wfile.write(body)
 def do_GET(self):
    if self.path=='/api/schema': return self.send(dict(fields=fields,nextId=max_id+1,project='AL-YAMAMA CEMENT PLANT 6000 TPD CLINKER PROJECT'))
    if self.path=='/api/entries': return self.send(entries())
    if self.path=='/api/export':
      w=openpyxl.load_workbook(TEMPLATE); sheet=w['log']
      for n,e in enumerate(entries(),start=append_row):
       sheet.cell(n,1,e['number'])
       for f in fields:
        val=e['values'].get(f['id'],'')
        if val and 'date' in f['label'].lower():
         try: val=datetime.datetime.strptime(val,'%Y-%m-%d')
         except ValueError: pass
        cell=sheet.cell(n,openpyxl.utils.column_index_from_string(f['id']))
        cell.value=val
        if isinstance(val,str): cell.data_type='s'
      b=io.BytesIO();w.save(b)
      self.send_response(200);self.send_header('Content-Type','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');self.send_header('Content-Disposition','attachment; filename="Yamama-RFI-Log.xlsx"');self.end_headers();self.wfile.write(b.getvalue());return
    return super().do_GET()
 def do_POST(self):
    try:
     d=json.loads(self.rfile.read(int(self.headers.get('Content-Length',0)))); vals=validate(d['values']); all=entries()
     if d.get('number'):
      item=next(e for e in all if e['number']==d['number']);item['values']=vals
     else: all.append(dict(number=str(max_id+len(all)+1).zfill(4),values=vals))
     tmp=STORE.with_suffix('.tmp');tmp.write_text(json.dumps(all,indent=2));tmp.replace(STORE)
     self.send(all)
    except (ValueError,KeyError,StopIteration) as e:self.send({'error':str(e) or 'Record not found'},400)
if __name__=='__main__':
 print('RFI workspace running on port 8000',flush=True)
 ThreadingHTTPServer(('0.0.0.0',8000),Handler).serve_forever()
