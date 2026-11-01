let schema, saved=[], editing=null;const $=s=>document.querySelector(s);const el=(tag,cls,text)=>{let e=document.createElement(tag);if(cls)e.className=cls;if(text!==undefined)e.textContent=text;return e};
function localToday(){const d=new Date();return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;}
function renderForm(values={}){
 const root=$('#fields');root.replaceChildren();let group;
 for(const f of schema.fields){
  if(f.group!==group){group=f.group;root.append(el('h3','section-heading',group));root.append(el('div','grid'));}
  const box=el('div','field');if(['E','S','U'].includes(f.id))box.classList.add('wide');
  const label=el('label','',f.label+(f.fixed?'  ▣':''));label.htmlFor='f'+f.id;box.append(label);
  let input;
  if(f.type==='select'){
   input=el('select');const empty=el('option','','Select an option');empty.value='';input.append(empty);
   for(const v of f.options){const o=el('option','',v);o.value=v;input.append(o);}input.required=true;
  }else{
   input=el('input');input.type=f.type==='date'?'date':'text';
   if(f.fixed){input.readOnly=true;box.classList.add('fixed');}
   else input.required=true;
   if(f.type==='text'){input.placeholder='Enter area or location';input.maxLength=1000;}
  }
  input.id='f'+f.id;input.name=f.id;
  input.value=f.fixed?f.default:(values[f.id]??(f.type==='date'?localToday():''));
  box.append(input);root.lastChild.append(box);
 }
 const note=el('div','notice','The remaining RFI form sections retain their template values. Signatures remain blank.');root.append(note);
 $('#number').textContent='86-RFI-CIV-'+(editing||String(schema.nextId+saved.length).padStart(4,'0'));
 $('#recordTitle').textContent=editing?'Edit inspection request':'New inspection request';
}
function tab(records){$('#editor').hidden=records;$('#records').hidden=!records;$('#formTab').classList.toggle('selected',!records);$('#logTab').classList.toggle('selected',records);if(records)renderRecords();}
function renderRecords(){let root=$('#table');root.replaceChildren();let query=$('#search').value.toLowerCase();let rows=saved.filter(e=>(e.number+' '+e.values.E).toLowerCase().includes(query));if(!rows.length){root.append(el('div','empty',saved.length?'No matching requests.':'No saved requests yet. Create your first inspection request.'));return;}let table=el('table'),head=el('tr');['RFI N°','AREA / LOCATION','INSPECTION DATE','EXPORT',''].forEach(t=>head.append(el('th','',t)));table.append(head);for(let e of rows){let r=el('tr');[e.number,e.values.E,e.values.H,''].forEach(v=>r.append(el('td','',v)));const download=el('a','edit','↓ RFI Excel');download.href='/api/export?number='+encodeURIComponent(e.number);r.lastChild.append(download);let td=el('td'),b=el('button','edit','Edit ↗');b.onclick=()=>{editing=e.number;renderForm(e.values);tab(false);window.scrollTo(0,0);};td.append(b);r.append(td);table.append(r);}root.append(table);}
$('#formTab').onclick=()=>tab(false);$('#logTab').onclick=()=>tab(true);$('#search').oninput=renderRecords;$('#newNav').onclick=()=>tab(false);$('#reset').onclick=()=>{editing=null;renderForm();$('#message').textContent='Form reset. Ready for a new request.';};
$('#form').onsubmit=async e=>{e.preventDefault();let values={};schema.fields.forEach(f=>values[f.id]=$('#f'+f.id).value);let button=$('#form button[type=submit]');button.disabled=true;try{let r=await fetch('/api/entries',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({number:editing,values})});let data=await r.json();if(!r.ok)throw Error(data.error);saved=data;$('#count').textContent=$('#badge').textContent=saved.length;editing=null;renderForm();$('#message').textContent='✓ Inspection request saved successfully.';}catch(err){$('#message').textContent='Unable to save: '+err.message;}finally{button.disabled=false;}};
(async()=>{try{let responses=await Promise.all([fetch('/api/schema'),fetch('/api/entries')]);[schema,saved]=await Promise.all(responses.map(r=>{if(!r.ok)throw Error('Server unavailable');return r.json();}));$('#count').textContent=$('#badge').textContent=saved.length;renderForm();}catch(e){$('#message').textContent='Unable to load workspace. Restart the server and refresh.';}})();
