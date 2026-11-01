import unittest, datetime, io, sys
from pathlib import Path
sys.path.insert(0,str(Path(__file__).resolve().parents[1]))
import server, openpyxl
class RfiTests(unittest.TestCase):
 def values(self):
  d={f['id']:f['default'] for f in server.fields}
  d.update(E='Silo 2',H='2027-01-12',Y='2027-01-11',AD='2027-01-13',S=next(f['options'][0] for f in server.fields if f['id']=='S'),U=next(f['options'][0] for f in server.fields if f['id']=='U'))
  return d
 def test_fixed_values_and_dates(self):
  d=self.values();d.update(C='MECH',I='06:00',O='7',W='wrong',AB='wrong')
  v=server.validate(d)
  for k,x in {'C':'CIV','I':'12:00','O':'1','W':'Haedr Ahmaed','AB':'Haedr Ahmaed','AC':'Engr. Ga Zang Dong Zhi'}.items():self.assertEqual(v[k],x)
  self.assertEqual(v['H'],'2027-01-12')
  for key in ['H','Y','AD']:
   d=self.values();d[key]='2027-02-30'
   with self.assertRaises(ValueError):server.validate(d)
 def test_only_rfi_fields_and_closed_drawing_list(self):
  self.assertEqual(set(f['id'] for f in server.fields),set(server.FORM_MAP))
  d=self.values();d['U']='invented drawing'
  with self.assertRaises(ValueError):server.validate(d)
 def test_filled_form_and_only_two_sheets(self):
  d=self.values();b=server.export_workbook([{'number':'0510','values':d}],'0510')
  w=openpyxl.load_workbook(io.BytesIO(b));f=w['RFI Form']
  self.assertEqual(w.sheetnames,['RFI Form','log'])
  for cell,value in {'AG2':'0510','I3':'Silo 2','N2':'CIV','I5':'12:00','W5':1,'G13':'**','O13':'Haedr Ahmaed','Z13':'Dhiya Mahdi','Q19':'Haedr Ahmaed','AA19':'Engr. Ga Zang Dong Zhi'}.items():self.assertEqual(f[cell].value,value)
  for c in ['G15','O15','Z15']:self.assertEqual(f[c].value,datetime.datetime(2027,1,11))
  for c in ['G21','K21','Q21','AA21']:self.assertEqual(f[c].value,datetime.datetime(2027,1,13))
  self.assertEqual(f['I9'].value,d['U'])
  self.assertEqual(f['A8'].value,d['S'])
  self.assertFalse(any(c.data_type=='f' for row in f for c in row))
  self.assertEqual(w['log'].cell(server.append_row,15).value,1)
if __name__=='__main__':unittest.main()
