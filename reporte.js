import {readReport,cleanReport,buildReport,DEFAULT_EXCLUSIONS,INSURANCE} from './lib/report.js';
const $=id=>document.getElementById(id);
let selected=null,url=null,generation=0;
const show=t=>{$('report-status').textContent=t;};
$('excluded-names').value=DEFAULT_EXCLUSIONS.join('\n');$('insurance-names').value=INSURANCE.join('\n');
function release(){if(url)URL.revokeObjectURL(url);url=null;$('report-download').removeAttribute('href');$('report-result').hidden=true;}
function select(file){generation++;release();selected=null;$('generate-report').disabled=true;if(!file)return;if(!/\.xlsx$/i.test(file.name)||file.size>20*1024*1024){show('Selecciona un archivo .xlsx de hasta 20 MiB.');return;}selected=file;$('chosen-file').textContent=file.name;$('generate-report').disabled=false;show('Base seleccionada. Puedes revisar las exclusiones antes de generar.');}
$('report-file').addEventListener('change',e=>{select(e.target.files[0]);e.target.value='';});
for(const type of ['dragenter','dragover'])$('report-drop').addEventListener(type,e=>{e.preventDefault();$('report-drop').classList.add('is-dragging');});
$('report-drop').addEventListener('dragleave',()=>$('report-drop').classList.remove('is-dragging'));
$('report-drop').addEventListener('drop',e=>{e.preventDefault();$('report-drop').classList.remove('is-dragging');if(e.dataTransfer.files.length!==1){show('Selecciona una sola base a la vez.');return;}select(e.dataTransfer.files[0]);});
document.addEventListener('dragover',e=>e.preventDefault());document.addEventListener('drop',e=>e.preventDefault());
function clear(){generation++;selected=null;release();$('report-file').value='';$('chosen-file').textContent='Ningún archivo seleccionado.';$('generate-report').disabled=true;show('Sesión limpia. Selecciona una base para comenzar.');}
$('clear-report').addEventListener('click',clear);window.addEventListener('pagehide',clear);
$('generate-report').addEventListener('click',async()=>{
 if(!selected)return;const ticket=++generation;const file=selected;const mode=$('report-mode').value;const rules=$('excluded-names').value.split('\n').map(s=>s.trim()).filter(Boolean);const insurers=$('insurance-names').value.split('\n').map(s=>s.trim()).filter(Boolean);
 release();$('generate-report').disabled=true;show('Leyendo y filtrando la base…');
 try{
  const raw=await file.arrayBuffer();if(ticket!==generation)return;
  // Check declared expanded sizes before handing the workbook to its reader.
  const zip=await JSZip.loadAsync(raw);let expanded=0;for(const entry of Object.values(zip.files)){expanded+=entry._data?.uncompressedSize||0;if(expanded>120*1024*1024)throw Error('El Excel descomprimido supera 120 MiB. Reduce el tamaño de la base.');}
  const source=await readReport(raw,ExcelJS,mode);if(ticket!==generation)return;
  const data=mode==='update'?source.data:cleanReport(source.headers,source.rows,rules,insurers);show('Generando las cinco hojas y la tabla dinámica…');
  const bytes=await buildReport(data,ExcelJS,JSZip);if(ticket!==generation)return;
  url=URL.createObjectURL(new Blob([bytes],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}));
  $('report-download').href=url;$('report-download').download=`${file.name.replace(/\.xlsx$/i,'')}_${mode==='update'?'actualizado':'reporte_2026'}.xlsx`;
  $('report-counts').textContent=mode==='update'?`${data.rows.length} registros de TOTAL · ${data.excluded.length} en EXCLUIDOS · ${data.restored} coincidencias retiradas de EXCLUIDOS.`:`${data.rows.length} registros incluidos · ${data.excluded.length} excluidos por proveedor · ${data.ignored} de otras cuentas o años.`;
  $('report-result').hidden=false;show('Reporte listo. Revisa EXCLUIDOS para comprobar las coincidencias de proveedores.');
 }catch(e){if(ticket===generation)show(`No se generó el reporte: ${e.message}`);}finally{if(ticket===generation)$('generate-report').disabled=!selected;}
});

$('report-mode').addEventListener('change',()=>{generation++;release();const update=$('report-mode').value==='update';$('report-rules').hidden=update;$('generate-report').textContent=update?'Actualizar desde TOTAL':'Generar reporte';$('generate-report').disabled=!selected;$('mode-help').textContent=update?'Edita TOTAL y guarda tu Excel. Se reconstruyen RESUMEN, DATA y VCTO sin aplicar exclusiones. Se retiran de EXCLUIDOS las filas idénticas copiadas a TOTAL. Las filas eliminadas de TOTAL no se archivan automáticamente. Se genera un nuevo archivo con las cinco hojas; no se conservan hojas adicionales ni formatos personalizados.':'Aplica las cuentas y exclusiones configuradas.';show('Modo seleccionado. Selecciona el archivo correspondiente y pulsa el botón para continuar.');});
