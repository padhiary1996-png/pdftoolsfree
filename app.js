(()=>{"use strict";
const $=id=>document.getElementById(id),input=$("pdfFile"),zone=$("dropZone"),nameEl=$("fileName"),button=$("convertBtn"),ocr=$("enableOcr"),ocrAll=$("ocrAll"),includeImages=$("includeImages"),progress=$("progressWrap"),bar=$("progressBar"),status=$("status"),result=$("result");let selected=null;
if(window.pdfjsLib)pdfjsLib.GlobalWorkerOptions.workerSrc="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
input.addEventListener("change",()=>pick(input.files[0]));["dragenter","dragover"].forEach(t=>zone.addEventListener(t,e=>{e.preventDefault();zone.style.background="#eef4ff"}));["dragleave","drop"].forEach(t=>zone.addEventListener(t,e=>{e.preventDefault();zone.style.background="#fbfdff"}));zone.addEventListener("drop",e=>pick(e.dataTransfer.files[0]));
function pick(f){if(!f)return;result.hidden=true;if(!(f.type==="application/pdf"||f.name.toLowerCase().endsWith(".pdf"))){err("Please choose a PDF file.");return}selected=f;nameEl.textContent=f.name+" · "+(f.size/1048576).toFixed(2)+" MB";button.disabled=false}
function say(s,p){status.textContent=s;bar.style.width=Math.min(100,Math.max(0,p))+"%"}function err(s){result.hidden=false;result.className="error";result.textContent=s}
function finish(blob,n){result.hidden=false;result.className="";result.replaceChildren();const b=document.createElement("b");b.textContent="Your Excel file is ready.";const a=document.createElement("a");a.className="download";a.href=URL.createObjectURL(blob);a.download=n;a.textContent="Download Excel file (.xlsx)";const p=document.createElement("p");p.textContent="Data worksheets contain extracted text/OCR; PDF Page worksheets contain visual copies. Verify extracted data before relying on it.";result.append(b,document.createElement("br"),a,p)}
button.addEventListener("click",async()=>{if(!selected)return;if(!window.pdfjsLib||!window.ExcelJS){err("A required library did not load. Check your internet connection and refresh.");return}button.disabled=true;progress.hidden=false;result.hidden=true;
try{say("Opening PDF…",2);const pdf=await pdfjsLib.getDocument({data:new Uint8Array(await selected.arrayBuffer())}).promise;const wb=new ExcelJS.Workbook();wb.creator="SheetPDF";wb.subject="PDF converted to Excel";
for(let n=1;n<=pdf.numPages;n++){const page=await pdf.getPage(n);say(`Extracting page ${n} of ${pdf.numPages}…`,(n-1)/pdf.numPages*80+3);const content=await page.getTextContent();const items=content.items.filter(i=>i.str&&i.str.trim()).map(i=>({text:i.str.trim(),x:i.transform[4],y:i.transform[5],width:i.width||0}));let rows;if(ocr.checked&&(ocrAll.checked||items.length<4)){const found=await ocrPage(page,n,pdf.numPages);rows=(ocrAll.checked||!items.length)&&found.length?found:groupItems(items)}else rows=groupItems(items);if(!rows.length)rows=[["No readable text detected"]];
const ws=wb.addWorksheet(`Data Page ${n}`);ws.addRows(rows);ws.views=[{state:"frozen",ySplit:1}];ws.getRow(1).font={bold:true};ws.columns=Array.from({length:Math.max(1,...rows.map(r=>r.length))},(_,c)=>({width:Math.min(42,Math.max(10,...rows.slice(0,300).map(r=>String(r[c]??"").length)))}));
if(includeImages.checked){say(`Adding visual page ${n} of ${pdf.numPages}…`,((n-1)/pdf.numPages)*80+7);const original=page.getViewport({scale:1});const scale=Math.min(1.35,1500/Math.max(original.width,original.height));const vp=page.getViewport({scale});const canvas=document.createElement("canvas");canvas.width=Math.ceil(vp.width);canvas.height=Math.ceil(vp.height);await page.render({canvasContext:canvas.getContext("2d",{alpha:false}),viewport:vp}).promise;const data=canvas.toDataURL("image/jpeg",0.82).split(",")[1];const imageId=wb.addImage({base64:data,extension:"jpeg"});const visual=wb.addWorksheet(`PDF Page ${n}`);visual.getCell("A1").value=`Visual copy of PDF page ${n} — use Data Page ${n} for extracted cells.`;visual.getCell("A1").font={bold:true};visual.addImage(imageId,{tl:{col:0,row:2},ext:{width:Math.round(vp.width*.72),height:Math.round(vp.height*.72)}});visual.getColumn(1).width=24;canvas.width=canvas.height=0;try{await addEmbeddedImages(page,wb,n)}catch(imageError){console.warn("Some embedded PDF images could not be extracted on page",n,imageError)}}}
say("Building Excel workbook…",92);const buffer=await wb.xlsx.writeBuffer();const safe=selected.name.replace(/\.pdf$/i,"").replace(/[\\/:*?"<>|]/g,"_")||"converted";finish(new Blob([buffer],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}),safe+".xlsx");say(`Done — ${pdf.numPages} page(s) exported.`,100)
}catch(e){console.error(e);err("Conversion failed: "+(e.message||"unknown error")+". Try a smaller or unlocked PDF, or refresh and retry.");say("Conversion stopped.",0)}finally{button.disabled=!selected}});
async function addEmbeddedImages(page,wb,pageNumber){
  // Best-effort extraction of raster image XObjects. PDF.js does not expose every
  // vector/inline/background graphic as a standalone image, so the full-page copy remains available.
  const ops=await page.getOperatorList();
  const paintOps=new Set([pdfjsLib.OPS.paintImageXObject,pdfjsLib.OPS.paintJpegXObject,pdfjsLib.OPS.paintImageMaskXObject]);
  const ids=[];
  for(let i=0;i<ops.fnArray.length;i++){
    if(paintOps.has(ops.fnArray[i])){
      const a=ops.argsArray[i]||[];const id=a[0];
      if(typeof id==="string"&&!ids.includes(id))ids.push(id);
    }
  }
  if(!ids.length)return;
  const ws=wb.addWorksheet(`Images P${pageNumber}`);ws.getCell("A1").value=`Extracted embedded images from PDF page ${pageNumber}`;ws.getCell("A1").font={bold:true};
  let row=2,added=0;
  for(const id of ids){
    const obj=await new Promise(resolve=>{try{page.objs.get(id,resolve)}catch(e){resolve(null)}});
    if(!obj)continue;
    try{
      const canvas=document.createElement("canvas");let ctx=canvas.getContext("2d");
      const w=Number(obj.width||obj.bitmap?.width||0),h=Number(obj.height||obj.bitmap?.height||0);
      if(!w||!h||w*h>16000000)continue;
      canvas.width=w;canvas.height=h;
      if(obj.bitmap){ctx.drawImage(obj.bitmap,0,0,w,h)}
      else if(obj.data&&obj.data.length>=w*h*4){ctx.putImageData(new ImageData(new Uint8ClampedArray(obj.data),w,h),0,0)}
      else if(obj instanceof HTMLImageElement||obj instanceof HTMLCanvasElement){canvas.width=obj.width;canvas.height=obj.height;ctx.drawImage(obj,0,0)}
      else continue;
      const base64=canvas.toDataURL("image/png").split(",")[1];
      const imageId=wb.addImage({base64,extension:"png"});
      ws.addImage(imageId,{tl:{col:0,row},ext:{width:Math.min(480,w),height:Math.min(360,h,Math.min(360,h)*Math.min(1,480/w))}});
      ws.getCell(`A${row}`).value=`Image ${++added} (${w} × ${h})`;
      row+=Math.ceil(Math.min(360,h)/20)+3;
      canvas.width=canvas.height=0;
    }catch(e){console.warn("Skipping an unsupported embedded image",e)}
  }
  if(!added)wb.removeWorksheet(ws.id);
}
async function ocrPage(page,n,total){if(!window.Tesseract)throw Error("OCR library did not load. Check your connection.");say(`OCR on page ${n}/${total}…`,n/total*80);const vp=page.getViewport({scale:1.7}),canvas=document.createElement("canvas");canvas.width=Math.ceil(vp.width);canvas.height=Math.ceil(vp.height);await page.render({canvasContext:canvas.getContext("2d",{willReadFrequently:true}),viewport:vp}).promise;const {data}=await Tesseract.recognize(canvas,"eng",{logger:m=>{if(m.status==="recognizing text")say(`OCR page ${n}/${total}: ${Math.round(m.progress*100)}%`,Math.round((n-1+m.progress)/total*80))}});canvas.width=canvas.height=0;return wordsToRows(data.words||[])}
function wordsToRows(words){words=words.filter(w=>(w.text||"").trim());if(!words.length)return[];const hs=words.map(w=>w.bbox.y1-w.bbox.y0).filter(h=>h>0).sort((a,b)=>a-b),tol=Math.max(7,(hs[Math.floor(hs.length/2)]||12)*.65);words.sort((a,b)=>(a.bbox.y0+a.bbox.y1)-(b.bbox.y0+b.bbox.y1)||a.bbox.x0-b.bbox.x0);const lines=[];for(const w of words){const cy=(w.bbox.y0+w.bbox.y1)/2;let l=lines.find(x=>Math.abs(x.cy-cy)<=tol);if(!l){l={cy,words:[]};lines.push(l)}l.words.push(w);l.cy=l.words.reduce((s,x)=>s+(x.bbox.y0+x.bbox.y1)/2,0)/l.words.length}return lines.sort((a,b)=>a.cy-b.cy).map(l=>{l.words.sort((a,b)=>a.bbox.x0-b.bbox.x0);const cells=[];let cur="",end=null;for(const w of l.words){const gap=end===null?0:w.bbox.x0-end;if(gap>Math.max(18,(w.bbox.y1-w.bbox.y0)*1.6)){if(cur)cells.push(cur.trim());cur=w.text}else cur+=(cur?" ":"")+w.text;end=w.bbox.x1}if(cur)cells.push(cur.trim());return cells.length?cells:[""]})}
function groupItems(items){if(!items.length)return[];items.sort((a,b)=>b.y-a.y||a.x-b.x);const lines=[];for(const it of items){let l=lines.find(x=>Math.abs(x.y-it.y)<=3.2);if(!l){l={y:it.y,items:[]};lines.push(l)}l.items.push(it);l.y=l.items.reduce((s,x)=>s+x.y,0)/l.items.length}return lines.sort((a,b)=>b.y-a.y).map(l=>{l.items.sort((a,b)=>a.x-b.x);const cells=[];let cur="",right=null;for(const it of l.items){const gap=right===null?0:it.x-right;if(gap>24){if(cur.trim())cells.push(cur.trim());cur=it.text}else cur+=(cur?" ":"")+it.text;right=it.x+it.width}if(cur.trim())cells.push(cur.trim());return cells.length?cells:[""]})}
})();