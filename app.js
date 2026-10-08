(() => {
"use strict";
const input=document.getElementById("pdfFile"), zone=document.getElementById("dropZone"), nameEl=document.getElementById("fileName");
const button=document.getElementById("convertBtn"), ocr=document.getElementById("enableOcr"), ocrAll=document.getElementById("ocrAll");
const progress=document.getElementById("progressWrap"), bar=document.getElementById("progressBar"), status=document.getElementById("status"), result=document.getElementById("result");
let selected=null;
if(window.pdfjsLib) pdfjsLib.GlobalWorkerOptions.workerSrc="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
input.addEventListener("change",()=>pick(input.files[0]));
["dragenter","dragover"].forEach(t=>zone.addEventListener(t,e=>{e.preventDefault();zone.style.background="#eef4ff"}));
["dragleave","drop"].forEach(t=>zone.addEventListener(t,e=>{e.preventDefault();zone.style.background="#fbfdff"}));
zone.addEventListener("drop",e=>pick(e.dataTransfer.files[0]));
function pick(f){if(!f)return;result.hidden=true;if(!(f.type==="application/pdf"||f.name.toLowerCase().endsWith(".pdf"))){error("Please choose a PDF file.");return}selected=f;nameEl.textContent=f.name+" · "+(f.size/1048576).toFixed(2)+" MB";button.disabled=false}
function say(s,p){status.textContent=s;bar.style.width=Math.min(100,Math.max(0,p))+"%"}
function error(s){result.hidden=false;result.className="error";result.textContent=s}
function download(blob,n){result.hidden=false;result.className="";result.replaceChildren();let b=document.createElement("b");b.textContent="Your Excel file is ready.";let a=document.createElement("a");a.className="download";a.href=URL.createObjectURL(blob);a.download=n;a.textContent="Download Excel file (.xlsx)";let p=document.createElement("p");p.textContent="Review the workbook after downloading; complex layouts and OCR may need corrections.";result.append(b,document.createElement("br"),a,p)}
button.addEventListener("click",async()=>{if(!selected)return;if(!window.pdfjsLib||!window.XLSX){error("A required library did not load. Check your internet connection and refresh.");return}
button.disabled=true;progress.hidden=false;result.hidden=true;
try{say("Opening PDF…",3);let pdf=await pdfjsLib.getDocument({data:new Uint8Array(await selected.arrayBuffer())}).promise;let wb=XLSX.utils.book_new();
for(let n=1;n<=pdf.numPages;n++){say(`Reading page ${n} of ${pdf.numPages}…`,(n-1)/pdf.numPages*90+5);let page=await pdf.getPage(n), content=await page.getTextContent();
let items=content.items.filter(i=>i.str&&i.str.trim()).map(i=>({text:i.str.trim(),x:i.transform[4],y:i.transform[5],width:i.width||0}));let rows;
if(ocr.checked&&(ocrAll.checked||items.length<4)){let recognized=await ocrPage(page,n,pdf.numPages);rows=(ocrAll.checked||!items.length)&&recognized.length?recognized:groupItems(items)}else rows=groupItems(items);
if(!rows.length)rows=[["No readable text detected"]];let ws=XLSX.utils.aoa_to_sheet(rows);ws["!cols"]=Array.from({length:Math.max(1,...rows.map(r=>r.length))},(_,c)=>({wch:Math.min(42,Math.max(10,...rows.map(r=>String(r[c]||"").length).slice(0,300)))}));XLSX.utils.book_append_sheet(wb,ws,`Page ${n}`)}
say("Creating Excel workbook…",96);let bytes=XLSX.write(wb,{bookType:"xlsx",type:"array"});let safe=selected.name.replace(/\.pdf$/i,"").replace(/[\\/:*?""<>|]/g,"_")||"converted";download(new Blob([bytes],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}),safe+".xlsx");say(`Done — ${pdf.numPages} page(s) exported.`,100)
}catch(e){console.error(e);error("Conversion failed: "+(e.message||"unknown error")+". Try a smaller or unlocked PDF, or refresh and retry.");say("Conversion stopped.",0)}finally{button.disabled=!selected}});
async function ocrPage(page,n,total){if(!window.Tesseract)throw Error("OCR library did not load. Check your connection.");say(`OCR on page ${n}/${total}; this can take a while…`,n/total*85);let vp=page.getViewport({scale:1.7}),canvas=document.createElement("canvas");canvas.width=Math.ceil(vp.width);canvas.height=Math.ceil(vp.height);await page.render({canvasContext:canvas.getContext("2d",{willReadFrequently:true}),viewport:vp}).promise;
let {data}=await Tesseract.recognize(canvas,"eng",{logger:m=>{if(m.status==="recognizing text")say(`OCR page ${n}/${total}: ${Math.round(m.progress*100)}%`,Math.round((n-1+m.progress)/total*90))}});canvas.width=canvas.height=0;return wordsToRows(data.words||[])}
function wordsToRows(words){words=words.filter(w=>(w.text||"").trim());if(!words.length)return[];let hs=words.map(w=>w.bbox.y1-w.bbox.y0).filter(h=>h>0).sort((a,b)=>a-b),tol=Math.max(7,(hs[Math.floor(hs.length/2)]||12)*.65);words.sort((a,b)=>(a.bbox.y0+a.bbox.y1)-(b.bbox.y0+b.bbox.y1)||a.bbox.x0-b.bbox.x0);let lines=[];
for(let w of words){let cy=(w.bbox.y0+w.bbox.y1)/2,line=lines.find(l=>Math.abs(l.cy-cy)<=tol);if(!line){line={cy,words:[]};lines.push(line)}line.words.push(w);line.cy=line.words.reduce((s,x)=>s+(x.bbox.y0+x.bbox.y1)/2,0)/line.words.length}
return lines.sort((a,b)=>a.cy-b.cy).map(l=>{l.words.sort((a,b)=>a.bbox.x0-b.bbox.x0);let cells=[],cur="",end=null;for(let w of l.words){let gap=end===null?0:w.bbox.x0-end;if(gap>Math.max(18,(w.bbox.y1-w.bbox.y0)*1.6)){if(cur)cells.push(cur.trim());cur=w.text}else cur+=(cur?" ":"")+w.text;end=w.bbox.x1}if(cur)cells.push(cur.trim());return cells.length?cells:[""]})}
function groupItems(items){if(!items.length)return[];items.sort((a,b)=>b.y-a.y||a.x-b.x);let lines=[];for(let it of items){let l=lines.find(x=>Math.abs(x.y-it.y)<=3.2);if(!l){l={y:it.y,items:[]};lines.push(l)}l.items.push(it);l.y=l.items.reduce((s,x)=>s+x.y,0)/l.items.length}
return lines.sort((a,b)=>b.y-a.y).map(l=>{l.items.sort((a,b)=>a.x-b.x);let cells=[],cur="",right=null;for(let it of l.items){let gap=right===null?0:it.x-right;if(gap>24){if(cur.trim())cells.push(cur.trim());cur=it.text}else cur+=(cur?" ":"")+it.text;right=it.x+it.width}if(cur.trim())cells.push(cur.trim());return cells.length?cells:[""]})}
})();