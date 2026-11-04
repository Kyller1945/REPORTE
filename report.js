export const DEFAULT_EXCLUSIONS = ['APM TERMINALS CALLAO SOCIED','BANCO DE CREDITO DEL PERU','CALL SHIPPING','INTEGRATEL PERU','INTERNATIONAL MOTORS LLC','LA CASA DEL RETEN','LUZ DEL SUR','OPERADOR LOGISTICO PUMA','PEX PERU','PLUZ ENERGIA PERU','RESTAURANTE Y PICANTERIA','RIMAC','SERV DE AGUA POTABLE Y ALCANTARILLADO','SERVICIO DE AGUA POTABLE Y ALCANTARILLADO','SINBA SURA','TALMA SERVICIOS AEROPOR','TRAMA PACIFICO','VILLAS OQUENDO'];
export const INSURANCE = ['SEGUROS','ASEGURADORA','MAPFRE','LA POSITIVA','PACIFICO SALUD','PACIFICO EPS','CHUBB'];
export const MONTHS=['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];
export const norm=v=>String(v??'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toUpperCase().replace(/[^A-Z0-9]/g,'');
const accounts=new Set(['42120001','42120002','42410001','43120001','43120002']);
export function excludedReason(name, exclusions=DEFAULT_EXCLUSIONS, insurance=INSURANCE){
 const n=norm(name);if(n.includes('ONCOSALUD'))return '';
 const hit=exclusions.find(p=>norm(p)&&n.includes(norm(p)));
 if(hit)return `Proveedor excluido: ${hit}`;
 const ins=insurance.find(p=>norm(p)&&n.includes(norm(p)));
 return ins?`Seguros: ${ins}`:'';
}
export function asDate(v){
 if(v instanceof Date&&Number.isFinite(v.getTime()))return v;
 if(typeof v==='number'&&v>0&&v<100000)return new Date(Date.UTC(1899,11,30)+Math.floor(v)*86400000);
 if(typeof v==='string'){
  let m=v.trim().match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})(?:\s.*)?$/);let y,mo,d;
  if(m){[,d,mo,y]=m.map(Number);}else{m=v.trim().match(/^(\d{4})-(\d{2})-(\d{2})(?:[T\s].*)?$/);if(m)[,y,mo,d]=m.map(Number);}
  if(m){const dt=new Date(Date.UTC(y,mo-1,d));if(dt.getUTCFullYear()===y&&dt.getUTCMonth()===mo-1&&dt.getUTCDate()===d)return dt;}
 }return null;
}
export function cleanReport(headers, source, exclusions=DEFAULT_EXCLUSIONS, insurance=INSURANCE){
 const find=n=>{const matches=headers.map((v,i)=>norm(v)===norm(n)?i:-1).filter(i=>i>=0);if(matches.length!==1)throw Error(`Falta o se repite la columna ${n}.`);return matches[0];};
 const p=find('Proveedor'),a=find('Cuenta principal'),d=find('Fecha comprobante'),pen=find('Saldo PEN'),usd=find('Saldo USD'),end=find('Saldo EUR');
 const keep=headers.map((_,i)=>i).filter(i=>i<end&&norm(headers[i])!=='IMPORTEEUR');
 if(![p,a,d,pen,usd].every(i=>keep.includes(i)))throw Error('Las columnas principales deben aparecer antes de Saldo EUR.');
 const out={headers:keep.map(i=>headers[i]),rows:[],excluded:[],ignored:0,eligible:0,sourceCount:0};
 source.forEach((r,j)=>{
  if(r.every(v=>v===null||v===undefined||v===''))return;out.sourceCount++;
  if(!accounts.has(String(r[a]??'').trim())){out.ignored++;return;}
  const date=asDate(r[d]);if(!date)throw Error(`Fila ${j+2}: fecha de comprobante inválida.`);
  if(date.getUTCFullYear()!==2026){out.ignored++;return;}out.eligible++;
  if(!String(r[p]??'').trim())throw Error(`Fila ${j+2}: proveedor vacío.`);
  const copy=keep.map(i=>r[i]??null);copy[keep.indexOf(d)]=date;
  for(const i of keep){if(/^(SALDO|IMPORTE)(PEN|USD)$/.test(norm(headers[i]))){const v=r[i];if(typeof v!=='number'||!Number.isFinite(v))throw Error(`Fila ${j+2}: ${headers[i]} debe ser un número, no texto ni vacío.`);}}
  const reason=excludedReason(r[p],exclusions,insurance);
  if(reason)out.excluded.push([...copy,reason]);else out.rows.push(copy);
 });
 out.provider=keep.indexOf(p);out.date=keep.indexOf(d);out.pen=keep.indexOf(pen);out.usd=keep.indexOf(usd);
 out.rows.sort((x,y)=>x[out.date]-y[out.date]||String(x[out.provider]).localeCompare(String(y[out.provider]),'es'));
 return out;
}

function value(cell){const v=cell.value;if(v&&typeof v==='object'&&!(v instanceof Date)){if('formula'in v||'sharedFormula'in v){if(v.result===undefined)throw Error('La base tiene fórmulas sin resultado guardado. Abre y guarda el archivo en Excel primero.');return v.result;}if(v.richText)return v.richText.map(x=>x.text).join('');if(v.error)throw Error('La base contiene una celda con error de Excel.');if(v.text)return v.text;}return v;}
export async function readReport(buffer,Excel,mode='new'){
 const wb=new Excel.Workbook();await wb.xlsx.load(buffer);
 if(mode==='update'){
  const total=wb.getWorksheet('TOTAL');if(!total)throw Error('No existe la hoja TOTAL. Selecciona el reporte generado y editado.');
  const base=readSheet(total,wb.properties.date1904),headers=base.headers;
  const find=n=>{const hits=headers.map((h,i)=>norm(h)===norm(n)?i:-1).filter(i=>i>=0);if(hits.length!==1)throw Error(`TOTAL: falta o se repite ${n}.`);return hits[0];};
  const provider=find('Proveedor'),date=find('Fecha comprobante'),pen=find('Saldo PEN'),usd=find('Saldo USD');find('Cuenta principal');
  const rows=base.rows.filter(r=>r.some(v=>v!=null&&v!==''));
  rows.forEach((r,i)=>{const d=asDate(r[date]);if(!d||d.getUTCFullYear()!==2026)throw Error(`TOTAL, registro ${i+1}: Fecha comprobante debe ser válida y de 2026.`);r[date]=d;if(!String(r[provider]??'').trim())throw Error(`TOTAL, registro ${i+1}: proveedor vacío.`);headers.forEach((h,c)=>{if(/^(SALDO|IMPORTE)(PEN|USD)$/.test(norm(h))&&(typeof r[c]!=='number'||!Number.isFinite(r[c])))throw Error(`TOTAL, registro ${i+1}: ${h} debe ser un número.`);});});
  let excluded=[];const es=wb.getWorksheet('EXCLUIDOS');let restored=0;
  if(es){const ex=readSheet(es,wb.properties.date1904);const indexes=headers.map(h=>ex.headers.findIndex(e=>norm(e)===norm(h)));if(indexes.some(i=>i<0))throw Error('EXCLUIDOS debe conservar las mismas columnas de TOTAL.');const reason=ex.headers.findIndex(h=>norm(h)==='MOTIVODEEXCLUSION');
   const key=r=>JSON.stringify(r.map(v=>v??null)),matches=new Map();for(const r of rows){const k=key(r);matches.set(k,(matches.get(k)||0)+1);}
   for(const row of ex.rows.filter(r=>r.some(v=>v!=null&&v!==''))){const r=indexes.map(i=>row[i]??null),k=key(r);if(matches.get(k)>0){matches.set(k,matches.get(k)-1);restored++;}else excluded.push([...r,reason>=0?row[reason]:'Excluido previamente']);}
  }
  return {data:{headers,rows,excluded,provider,date,pen,usd,ignored:0,restored}};
 }
 const candidates=wb.worksheets.filter(s=>s.getRow(1).values.some(v=>norm(v)==='CUENTAPRINCIPAL'));
 if(candidates.length!==1)throw Error('Debe existir una sola hoja con los encabezados en la primera fila, incluida Cuenta principal.');
 return readSheet(candidates[0],wb.properties.date1904);
}
function readSheet(s,date1904){
 if(s.rowCount>100000)throw Error('Máximo 100 000 filas por archivo.');
 const headers=Array.from({length:s.columnCount},(_,i)=>String(value(s.getCell(1,i+1))??''));
 // Pasting a full EXCLUIDOS row can leave its reason beyond TOTAL's headers.
 while(headers.length&&!headers.at(-1).trim())headers.pop();
 const rows=[];for(let r=2;r<=s.rowCount;r++)rows.push(headers.map((_,i)=>value(s.getCell(r,i+1))));
 // Convert unformatted Excel date serials using the workbook's actual date system.
 const dateColumns=headers.map((h,i)=>norm(h).startsWith('FECHA')?i:-1).filter(i=>i>=0);
 for(const r of rows)for(const c of dateColumns)if(typeof r[c]==='number')r[c]=new Date(Date.UTC(date1904?1904:1899,date1904?0:11,date1904?1:30)+Math.floor(r[c])*86400000);
 return {headers,rows};
}
const x=v=>String(v??'').replace(/[<>&"']/g,c=>({'<':'&lt;','>':'&gt;','&':'&amp;','"':'&quot;',"'":'&apos;'}[c]));
const ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const rel='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const xml=s=>'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+s;
const col=n=>{let s='';for(n++;n;n=Math.floor((n-1)/26))s=String.fromCharCode(65+(n-1)%26)+s;return s;};

export async function buildReport(data,Excel,Zip){
 if(!data.rows.length)throw Error('No quedan registros para el reporte. Revisa la lista de exclusiones.');
 const wb=new Excel.Workbook();wb.creator='Herramientas de trabajo';
 function sheet(name,headers,rows){const s=wb.addWorksheet(name,{views:[{state:'frozen',ySplit:1}]});s.addRow(headers);s.addRows(rows);s.getRow(1).font={bold:true,color:{argb:'FFFFFFFF'}};s.getRow(1).fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF163F4A'}};s.getRow(1).height=30;
  s.autoFilter={from:{row:1,column:1},to:{row:rows.length+1,column:headers.length}};
  headers.forEach((h,i)=>{const c=s.getColumn(i+1);c.width=norm(h)==='PROVEEDOR'?48:norm(h).includes('MOTIVO')?48:Math.min(26,Math.max(17,String(h).length+2));if(/^(SALDO|IMPORTE)/.test(norm(h)))c.numFmt='#,##0;[Red]-#,##0';if(norm(h).startsWith('FECHA'))c.numFmt='dd/mm/yyyy';});return s;}
 sheet('TOTAL',data.headers,data.rows);
 sheet('EXCLUIDOS',[...data.headers,'Motivo de exclusión'],data.excluded);
 const groups=new Map(),months=[...new Set(data.rows.map(r=>r[data.date].getUTCMonth()))].sort((a,b)=>a-b);
 for(const r of data.rows){const name=String(r[data.provider]).trim();if(!groups.has(name))groups.set(name,{pen:0,usd:0,months:new Map()});const g=groups.get(name),m=r[data.date].getUTCMonth();g.pen+=r[data.pen];g.usd+=r[data.usd];if(!g.months.has(m))g.months.set(m,[0,0]);g.months.get(m)[0]+=r[data.pen];g.months.get(m)[1]+=r[data.usd];}
 const names=[...groups.keys()].sort((a,b)=>a.localeCompare(b,'es'));
 const summary=sheet('RESUMEN',['Proveedor','Saldo PEN','Saldo USD'],names.map(n=>[n,groups.get(n).pen,groups.get(n).usd]));
 summary.addRow(['TOTAL',{formula:`SUM(B2:B${names.length+1})`,result:names.reduce((s,n)=>s+groups.get(n).pen,0)},{formula:`SUM(C2:C${names.length+1})`,result:names.reduce((s,n)=>s+groups.get(n).usd,0)}]).font={bold:true};
 const monthDates=data.rows.map(r=>new Date(Date.UTC(2026,r[data.date].getUTCMonth(),1)));
 const ds=sheet('DATA',[...data.headers,'Mes'],data.rows.map((r,i)=>[...r,monthDates[i]]));ds.getColumn(data.headers.length+1).numFmt='[$-es-PE]mmmm - yy';
 const ps=wb.addWorksheet('VCTO',{views:[{state:'frozen',xSplit:1,ySplit:2}]});ps.addRow(['Proveedor',...months.flatMap(m=>[`${MONTHS[m]} - 26`,null])]);ps.addRow(['',...months.flatMap(()=>['Saldo PEN','Saldo USD'])]);
 months.forEach((_,i)=>ps.mergeCells(1,2+i*2,1,3+i*2));
 for(const n of names)ps.addRow([n,...months.flatMap(m=>groups.get(n).months.get(m)||[0,0])]);
 ps.addRow(['Total general',...months.flatMap(m=>[0,1].map(j=>names.reduce((sum,n)=>sum+(groups.get(n).months.get(m)?.[j]||0),0)))]).font={bold:true};
 ps.getColumn(1).width=48;for(let i=2;i<=1+months.length*2;i++){ps.getColumn(i).width=19;ps.getColumn(i).numFmt='#,##0;[Red]-#,##0';}
 for(const i of [1,2]){ps.getRow(i).font={bold:true,color:{argb:'FFFFFFFF'}};ps.getRow(i).fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF163F4A'}};ps.getRow(i).height=26;}
 const zip=await Zip.loadAsync(await wb.xlsx.writeBuffer());
 // Native OOXML PivotTable, with a cache sourced from DATA and two sum fields.
 const headers=[...data.headers,'Mes'],mi=headers.length-1,unique=headers.map(()=>[]),maps=headers.map(()=>new Map());
 const cached=data.rows.map((r,i)=>[...r,monthDates[i]]);
 // Force chronological month ordering and supplier ordering in shared items.
 for(const n of names){maps[data.provider].set(n,unique[data.provider].length);unique[data.provider].push(n);}
 for(const m of months){const d=new Date(Date.UTC(2026,m,1));maps[mi].set(d.toISOString(),unique[mi].length);unique[mi].push(d);}
 const records=cached.map(row=>'<r>'+row.map((v,i)=>{if(i===data.provider)v=String(v).trim();if(typeof v==='number')return `<n v="${v}"/>`;if(v===null||v===undefined)return '<m/>';const key=v instanceof Date?v.toISOString():String(v);if(!maps[i].has(key)){maps[i].set(key,unique[i].length);unique[i].push(v);}return `<x v="${maps[i].get(key)}"/>`;}).join('')+'</r>');
 const fields=headers.map((h,i)=>{
  const values=cached.map(r=>r[i]),date=values.some(v=>v instanceof Date),number=values.some(v=>typeof v==='number'),blank=values.some(v=>v==null),string=values.some(v=>v!=null&&!(v instanceof Date)&&typeof v!=='number');
  const flags=` containsDate="${+date}" containsNumber="${+number}" containsString="${+string}" containsBlank="${+blank}" containsNonDate="${+(number||string||blank)}" containsMixedTypes="${+(+date+ +number+ +string>1)}" containsSemiMixedTypes="${+(blank||string&&(date||number))}"`;
  return `<cacheField name="${x(h)}" numFmtId="${i===mi?165:0}"><sharedItems count="${unique[i].length}"${flags}>${unique[i].map(v=>v instanceof Date?`<d v="${v.toISOString().slice(0,19)}"/>`:`<s v="${x(v)}"/>`).join('')}</sharedItems></cacheField>`;
 }).join('');
 zip.file('xl/pivotCache/pivotCacheDefinition1.xml',xml(`<pivotCacheDefinition xmlns="${ns}" xmlns:r="${rel}" r:id="rId1" recordCount="${cached.length}" refreshOnLoad="0" createdVersion="6" refreshedVersion="6" minRefreshableVersion="3"><cacheSource type="worksheet"><worksheetSource ref="A1:${col(mi)}${cached.length+1}" sheet="DATA"/></cacheSource><cacheFields count="${headers.length}">${fields}</cacheFields></pivotCacheDefinition>`));
 zip.file('xl/pivotCache/pivotCacheRecords1.xml',xml(`<pivotCacheRecords xmlns="${ns}" count="${cached.length}">${records.join('')}</pivotCacheRecords>`));
 zip.file('xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels',xml(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/pivotCacheRecords" Target="pivotCacheRecords1.xml"/></Relationships>`));
 const pf=headers.map((_,i)=>{const axis=i===data.provider?'axisRow':i===mi?'axisCol':null;return axis?`<pivotField axis="${axis}" showAll="0" defaultSubtotal="0"${i===mi?' numFmtId="165"':''}><items count="${unique[i].length}">${unique[i].map((_,j)=>`<item x="${j}"/>`).join('')}</items></pivotField>`:`<pivotField${[data.pen,data.usd].includes(i)?' dataField="1"':''} showAll="0"/>`;}).join('');
 zip.file('xl/pivotTables/pivotTable1.xml',xml(`<pivotTableDefinition xmlns="${ns}" name="ReporteMes" cacheId="1" dataCaption="Saldos" rowGrandTotals="0" colGrandTotals="1" compact="0" compactData="0" gridDropZones="0" multipleFieldFilters="0" createdVersion="6" updatedVersion="6" minRefreshableVersion="3"><location ref="A1:${col(months.length*2)}${names.length+3}" firstHeaderRow="1" firstDataRow="2" firstDataCol="1"/><pivotFields count="${headers.length}">${pf}</pivotFields><rowFields count="1"><field x="${data.provider}"/></rowFields><rowItems count="${names.length+1}">${names.map((_,i)=>`<i><x v="${i}"/></i>`).join('')}<i t="grand"><x/></i></rowItems><colFields count="2"><field x="${mi}"/><field x="-2"/></colFields><colItems count="${months.length*2}">${months.flatMap((_,i)=>[0,1].map(j=>`<i i="${j}"><x v="${i}"/><x v="${j}"/></i>`)).join('')}</colItems><dataFields count="2"><dataField name="Suma de Saldo PEN" fld="${data.pen}" subtotal="sum" numFmtId="3"/><dataField name="Suma de Saldo USD" fld="${data.usd}" subtotal="sum" numFmtId="3"/></dataFields><pivotTableStyleInfo name="PivotStyleMedium9" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/></pivotTableDefinition>`));
 zip.file('xl/pivotTables/_rels/pivotTable1.xml.rels',xml(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${rel}/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition1.xml"/></Relationships>`));
 const patch=async(path,before,insert)=>zip.file(path,(await zip.file(path).async('string')).replace(before,insert+before));
 await patch('xl/workbook.xml','</workbook>','<pivotCaches><pivotCache cacheId="1" r:id="rIdPivotCache"/></pivotCaches>');
 await patch('xl/_rels/workbook.xml.rels','</Relationships>',`<Relationship Id="rIdPivotCache" Type="${rel}/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition1.xml"/>`);
 await patch('xl/worksheets/sheet5.xml','</worksheet>','<pivotParts count="1"><pivotPart r:id="rIdPivot"/></pivotParts>');
 zip.file('xl/worksheets/_rels/sheet5.xml.rels',xml(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdPivot" Type="${rel}/pivotTable" Target="../pivotTables/pivotTable1.xml"/></Relationships>`));
 await patch('[Content_Types].xml','</Types>',['pivotCache/pivotCacheDefinition1','pivotCache/pivotCacheRecords1','pivotTables/pivotTable1'].map((p,i)=>`<Override PartName="/xl/${p}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.${['pivotCacheDefinition','pivotCacheRecords','pivotTable'][i]}+xml"/>`).join(''));
 // Match the native pivot month field to the actual custom date format ID.
 const styles=await zip.file('xl/styles.xml').async('string');const monthFormat=styles.match(/<numFmt numFmtId="(\d+)" formatCode="\[\$-es-PE\]mmmm - yy"/);if(!monthFormat)throw Error('No se pudo verificar el formato mensual.');
 for(const p of ['xl/pivotCache/pivotCacheDefinition1.xml','xl/pivotTables/pivotTable1.xml'])zip.file(p,(await zip.file(p).async('string')).replaceAll('numFmtId="165"',`numFmtId="${monthFormat[1]}"`));
 return zip.generateAsync({type:'uint8array',compression:'DEFLATE'});
}
