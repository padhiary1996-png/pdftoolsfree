(() => {
  const $ = (id) => document.getElementById(id);
  const grid = $('toolGrid'), converter = $('converter'), dropzone = $('dropzone'), input = $('fileInput');
  const state = { tool: null, file: null };
  if (window.pdfjsLib) pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
  $('year').textContent = new Date().getFullYear();

  const configs = {
    'pdf-excel': { title:'PDF to Excel', description:'Select a text-based PDF to extract its text into an Excel workbook.', accept:'.pdf,application/pdf', label:'Drop your PDF here', sub:'or click to browse your device', types:'PDF files · Up to 20 MB', button:'Convert to Excel', option:true },
    'excel-pdf': { title:'Excel to PDF', description:'Export an Excel or CSV sheet as a simple printable PDF.', accept:'.xlsx,.xls,.csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel,text/csv', label:'Drop your spreadsheet here', sub:'or click to browse your device', types:'XLSX, XLS or CSV · Up to 20 MB', button:'Convert to PDF', option:false },
    'csv-excel': { title:'CSV to Excel', description:'Convert a comma-separated values file into an Excel workbook.', accept:'.csv,text/csv', label:'Drop your CSV here', sub:'or click to browse your device', types:'CSV files · Up to 20 MB', button:'Convert to Excel', option:false }
  };

  function openTool(tool) {
    state.tool = tool; state.file = null;
    const c = configs[tool]; if (!c) return;
    $('converterTitle').textContent = c.title;
    $('converterDescription').textContent = c.description;
    $('dropTitle').textContent = c.label; $('dropDescription').textContent = c.sub;
    $('acceptedTypes').textContent = c.types; input.accept = c.accept; input.value = '';
    $('converterOptions').style.display = c.option ? 'block' : 'none';
    $('convertButton').innerHTML = c.button + ' <span>→</span>';
    $('convertButton').disabled = true; $('fileSelected').hidden = true; dropzone.style.display = 'flex';
    setStatus('');
    converter.classList.add('open'); converter.scrollIntoView({behavior:'smooth', block:'start'});
  }
  function setStatus(message, type='') { const el=$('statusMessage'); el.textContent=message; el.className='status-message '+type; }
  function selectFile(file) {
    if (!file || !state.tool) return;
    if (file.size > 20*1024*1024) { setStatus('Please choose a file smaller than 20 MB.','error'); return; }
    state.file = file; $('selectedName').textContent=file.name; $('selectedSize').textContent=(file.size/1024).toFixed(1)+' KB';
    $('fileSelected').hidden=false; dropzone.style.display='none'; $('convertButton').disabled=false; setStatus('');
  }
  grid.addEventListener('click', e => { const b=e.target.closest('[data-tool]'); if(b) openTool(b.dataset.tool); });
  $('closeConverter').addEventListener('click',()=>{converter.classList.remove('open'); state.file=null;});
  dropzone.addEventListener('click',()=>input.click());
  dropzone.addEventListener('keydown',e=>{if(e.key==='Enter'||e.key===' '){e.preventDefault();input.click();}});
  input.addEventListener('change',()=>selectFile(input.files[0]));
  $('removeFile').addEventListener('click',()=>{state.file=null;input.value='';$('fileSelected').hidden=true;dropzone.style.display='flex';$('convertButton').disabled=true;setStatus('');});
  ['dragenter','dragover'].forEach(ev=>dropzone.addEventListener(ev,e=>{e.preventDefault();dropzone.classList.add('dragging');}));
  ['dragleave','drop'].forEach(ev=>dropzone.addEventListener(ev,e=>{e.preventDefault();dropzone.classList.remove('dragging');}));
  dropzone.addEventListener('drop',e=>selectFile(e.dataTransfer.files[0]));
  $('toolSearch').addEventListener('input',e=>{const q=e.target.value.toLowerCase().trim();grid.querySelectorAll('.tool-card').forEach(card=>card.style.display=card.dataset.search.includes(q)?'':'none');});

  function downloadBlob(blob, filename) {
    const url=URL.createObjectURL(blob), a=document.createElement('a'); a.href=url; a.download=filename; document.body.appendChild(a); a.click(); a.remove(); setTimeout(()=>URL.revokeObjectURL(url),1500);
  }
  function safeName(name){return name.replace(/\.[^.]+$/,'').replace(/[^\w-]+/g,'_')||'converted_file';}
  async function pdfToExcel(file) {
    if (!window.pdfjsLib || !window.XLSX) throw new Error('Conversion libraries did not load. Check your internet connection and reload this page.');
    const pdf=await pdfjsLib.getDocument({data:await file.arrayBuffer()}).promise;
    const rows=[];
    for(let pageNo=1; pageNo<=pdf.numPages; pageNo++){
      setStatus(`Reading page ${pageNo} of ${pdf.numPages}…`);
      const page=await pdf.getPage(pageNo), content=await page.getTextContent();
      const items=content.items.filter(x=>x.str && x.str.trim()).map(x=>({text:x.str.trim(),x:x.transform[4],y:x.transform[5]}));
      if(!items.length) continue;
      const lineMap=new Map();
      for(const item of items){const key=Math.round(item.y/3)*3;if(!lineMap.has(key))lineMap.set(key,[]);lineMap.get(key).push(item);}
      const lines=[...lineMap.entries()].sort((a,b)=>b[0]-a[0]);
      for(const [,line] of lines){
        line.sort((a,b)=>a.x-b.x);
        const cells=[]; let lastX=null;
        for(const item of line){
          if(lastX!==null && item.x-lastX>38) cells.push('');
          cells.push(item.text); lastX=item.x+item.text.length*4;
        }
        rows.push([`Page ${pageNo}`,...cells]);
      }
      if(rows.length>10000) throw new Error('This PDF contains too much text for this demo. Try fewer pages.');
    }
    if(!rows.length) throw new Error('No selectable text was found. This may be a scanned PDF; OCR is not included in this demo.');
    const ws=XLSX.utils.aoa_to_sheet([['Page','Extracted text / columns'],...rows]);
    ws['!cols']=[{wch:12},{wch:28},{wch:28},{wch:28},{wch:28},{wch:28}];
    const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,($('sheetName').value.trim().slice(0,31)||'Converted Data'));
    const out=XLSX.write(wb,{bookType:'xlsx',type:'array'});
    downloadBlob(new Blob([out],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}),safeName(file.name)+'.xlsx');
    return `Done — exported ${rows.length} text rows. Please review the spreadsheet; complex PDF tables may need cleanup.`;
  }
  async function loadWorkbook(file) {
    if(!window.XLSX) throw new Error('Spreadsheet library did not load. Check your internet connection and reload.');
    const buf=await file.arrayBuffer();
    return XLSX.read(buf,{type:'array'});
  }
  async function excelToPdf(file) {
    const wb=await loadWorkbook(file);
    const sheet=wb.Sheets[wb.SheetNames[0]];
    const rows=XLSX.utils.sheet_to_json(sheet,{header:1,defval:''});
    if(!rows.length) throw new Error('No spreadsheet data found.');
    const maxCols=Math.min(Math.max(...rows.map(r=>r.length)),12);
    const visible=rows.slice(0,300).map(r=>r.slice(0,maxCols));
    const htmlDoc='<!doctype html><html><head><meta charset="utf-8"><title>Spreadsheet export</title><style>@page{size:landscape;margin:12mm}body{font:10px Arial,sans-serif;color:#222}h2{font-size:15px}table{border-collapse:collapse;width:100%;table-layout:auto}td,th{border:1px solid #bbb;padding:5px;overflow-wrap:anywhere}tr:first-child{background:#e9edff;font-weight:bold}small{color:#666}</style></head><body><h2>'+escapeHtml(wb.SheetNames[0])+'</h2><small>Exported from SheetPDF. In the print dialog choose “Save as PDF”.</small><table>'+visible.map((r,i)=>'<tr>'+Array.from({length:maxCols},(_,j)=>'<'+(i===0?'th':'td')+'>'+escapeHtml(r[j]??'')+'</'+(i===0?'th':'td')+'>').join('')+'</tr>').join('')+'</table></body></html>';
    const blob=new Blob([htmlDoc],{type:'text/html;charset=utf-8'}); const url=URL.createObjectURL(blob); const w=window.open(url,'_blank');
    if(!w){URL.revokeObjectURL(url);throw new Error('Your browser blocked the print preview. Allow pop-ups for this site and try again.');}
    setStatus('Print preview opened in a new tab. Choose Print → Save as PDF.','success');
    setTimeout(()=>URL.revokeObjectURL(url),60000);
    return null;
  }
  function escapeHtml(value){return String(value).replace(/[&<>"']/g,ch=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[ch]));}
  async function csvToExcel(file){
    if(!window.XLSX) throw new Error('Spreadsheet library did not load. Check your internet connection and reload.');
    const text=await file.text(); const wb=XLSX.read(text,{type:'string'}); const out=XLSX.write(wb,{bookType:'xlsx',type:'array'});
    downloadBlob(new Blob([out],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}),safeName(file.name)+'.xlsx');
    return 'Done — your Excel workbook has been downloaded.';
  }
  $('convertButton').addEventListener('click',async()=>{
    if(!state.file||!state.tool)return;
    $('convertButton').disabled=true;setStatus('Working on your file…');
    try{
      let result;
      if(state.tool==='pdf-excel')result=await pdfToExcel(state.file);
      else if(state.tool==='excel-pdf')result=await excelToPdf(state.file);
      else if(state.tool==='csv-excel')result=await csvToExcel(state.file);
      if(result)setStatus(result,'success');
    }catch(err){setStatus(err.message||'Something went wrong. Please try another file.','error');}
    finally{$('convertButton').disabled=false;}
  });
})();