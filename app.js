const {PDFDocument, degrees} = PDFLib;
if(window.pdfjsLib) pdfjsLib.GlobalWorkerOptions.workerSrc="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
const tools={
 merge:["Mesclar PDF","Escolha dois ou mais PDFs e combine-os em um único documento."],
 split:["Dividir PDF","Extraia páginas específicas e salve-as em um novo PDF."],
 png:["PDF para PNG","Converta as páginas de um PDF em imagens PNG."],
 images:["Imagens para PDF","Transforme imagens JPG ou PNG em um documento PDF."],
 rotate:["Girar páginas","Gire todas as páginas do documento em 90 graus."],
 compress:["Otimizar PDF","Regrave o PDF. A redução de tamanho depende do conteúdo do arquivo."],
 word:["PDF para Word","Extraia o texto do PDF para um arquivo DOCX. A diagramação original pode não ser preservada."],
 about:["Sobre o PDF Fácil","Ferramentas gratuitas com processamento local no navegador."]
};
let current="merge", files=[];
const $=id=>document.getElementById(id);
function setTool(t){current=t;document.querySelectorAll(".tool").forEach(b=>b.classList.toggle("active",b.dataset.tool===t));$("tool-title").textContent=tools[t][0];$("tool-description").textContent=tools[t][1];$("file-input").accept=t==="images"?"image/png,image/jpeg":".pdf,application/pdf";$("file-input").multiple=["merge","images"].includes(t);$("file-list").innerHTML="";files=[];$("status").textContent="";renderOptions();renderFiles();}
document.querySelectorAll(".tool").forEach(b=>b.addEventListener("click",()=>setTool(b.dataset.tool)));
function renderOptions(){let o=$("options");o.innerHTML="";if(current==="split")o.innerHTML='<label>Páginas para extrair (ex.: 1,3,5-7) <input id="ranges" type="text" placeholder="1-3,5"></label>';if(current==="png")o.innerHTML='<label>Escala da imagem <select id="scale"><option value="1">Normal</option><option value="1.5" selected>Alta</option><option value="2">Muito alta</option></select></label>';if(current==="rotate")o.innerHTML='<label>Ângulo <select id="angle"><option value="90">90°</option><option value="180">180°</option><option value="270">270°</option></select></label>';}
function renderFiles(){$("file-list").innerHTML=files.map((f,i)=>`<div class="file-row"><span>▤ ${escapeHtml(f.name)} · ${(f.size/1024).toFixed(0)} KB</span><button data-remove="${i}" aria-label="Remover">Remover</button></div>`).join("");document.querySelectorAll("[data-remove]").forEach(b=>b.onclick=()=>{files.splice(+b.dataset.remove,1);renderFiles()});}
function escapeHtml(s){return s.replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));}
$("file-input").addEventListener("change",e=>{files.push(...[...e.target.files]);renderFiles();e.target.value="";});
const dz=$("dropzone");dz.addEventListener("dragover",e=>{e.preventDefault();dz.style.borderColor="#245889"});dz.addEventListener("dragleave",()=>dz.style.borderColor="");dz.addEventListener("drop",e=>{e.preventDefault();dz.style.borderColor="";files.push(...[...e.dataTransfer.files].filter(f=>current==="images"?f.type.startsWith("image/"):f.type==="application/pdf"));renderFiles()});
$("clear").onclick=()=>{files=[];renderFiles();$("status").textContent="";};
function download(blob,name){const url=URL.createObjectURL(blob),a=document.createElement("a");a.href=url;a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(url),1500);}
function parsePages(s,total){let out=new Set();for(const part of s.split(",")){let p=part.trim();if(!p)continue;if(p.includes("-")){let [a,b]=p.split("-").map(Number);if(!Number.isInteger(a)||!Number.isInteger(b)||a<1||b<a||b>total)throw Error("Intervalo de páginas inválido.");for(let n=a;n<=b;n++)out.add(n-1);}else{let n=Number(p);if(!Number.isInteger(n)||n<1||n>total)throw Error("Número de página inválido.");out.add(n-1);}}return [...out].sort((a,b)=>a-b);}
async function pdfToWord(file){
 if(!window.docx) throw Error("A biblioteca de Word não carregou. Verifique sua conexão com a internet e tente novamente.");
 const pdf=await pdfjsLib.getDocument({data:await file.arrayBuffer()}).promise;
 const paragraphs=[];
 for(let n=1;n<=pdf.numPages;n++){
  const page=await pdf.getPage(n),content=await page.getTextContent();
  let line="",lastY=null;
  for(const item of content.items){
   if(!item.str) continue;
   const y=item.transform?item.transform[5]:null;
   if(lastY!==null && y!==null && Math.abs(y-lastY)>3){
    if(line.trim()) paragraphs.push(new docx.Paragraph({children:[new docx.TextRun(line.trim())]}));
    line="";
   }
   line+=item.str+(item.hasEOL?"\n":" ");
   lastY=y;
  }
  if(line.trim()) paragraphs.push(new docx.Paragraph({children:[new docx.TextRun(line.trim())]}));
  if(n<pdf.numPages) paragraphs.push(new docx.Paragraph({children:[new docx.TextRun(" ")]}));
 }
 if(!paragraphs.length) throw Error("Não foi encontrado texto selecionável. Se o PDF for digitalizado como imagem, será necessário OCR.");
 const document=new docx.Document({sections:[{properties:{},children:paragraphs}]});
 return await docx.Packer.toBlob(document);
}
async function processFiles(){if(current==="about"){$("status").textContent="PDF Fácil: os arquivos são tratados localmente. Não feche a página durante o processamento.";return;}if(!files.length){$("status").textContent="Selecione pelo menos um arquivo.";return;}$("process").disabled=true;$("status").textContent="Processando…";
try{
 if(current==="merge"){if(files.length<2)throw Error("Selecione pelo menos dois PDFs.");const out=await PDFDocument.create();for(const f of files){const doc=await PDFDocument.load(await f.arrayBuffer());(await out.copyPages(doc,doc.getPageIndices())).forEach(p=>out.addPage(p));}download(new Blob([await out.save()],{type:"application/pdf"}),"pdf-mesclado.pdf");}
 if(current==="split"){const doc=await PDFDocument.load(await files[0].arrayBuffer()),ids=parsePages($("ranges").value,doc.getPageCount());if(!ids.length)throw Error("Informe as páginas que deseja extrair.");const out=await PDFDocument.create();(await out.copyPages(doc,ids)).forEach(p=>out.addPage(p));download(new Blob([await out.save()],{type:"application/pdf"}),"paginas-extraidas.pdf");}
 if(current==="rotate"){const doc=await PDFDocument.load(await files[0].arrayBuffer()),a=Number($("angle").value);doc.getPages().forEach(p=>p.setRotation(degrees((p.getRotation().angle+a)%360)));download(new Blob([await doc.save()],{type:"application/pdf"}),"pdf-girado.pdf");}
 if(current==="compress"){const doc=await PDFDocument.load(await files[0].arrayBuffer());download(new Blob([await doc.save({useObjectStreams:true})],{type:"application/pdf"}),"pdf-otimizado.pdf");}
 if(current==="png"){const pdf=await pdfjsLib.getDocument({data:await files[0].arrayBuffer()}).promise,scale=Number($("scale").value);for(let n=1;n<=pdf.numPages;n++){const page=await pdf.getPage(n),viewport=page.getViewport({scale}),canvas=document.createElement("canvas");canvas.width=viewport.width;canvas.height=viewport.height;await page.render({canvasContext:canvas.getContext("2d"),viewport}).promise;const blob=await new Promise(r=>canvas.toBlob(r,"image/png"));download(blob,`pagina-${String(n).padStart(3,"0")}.png`);}}
 if(current==="images"){const out=await PDFDocument.create();for(const f of files){const bytes=await f.arrayBuffer(),img=f.type==="image/png"?await out.embedPng(bytes):await out.embedJpg(bytes),page=out.addPage([img.width,img.height]);page.drawImage(img,{x:0,y:0,width:img.width,height:img.height});}download(new Blob([await out.save()],{type:"application/pdf"}),"imagens.pdf");}
 if(current==="word"){const blob=await pdfToWord(files[0]);download(blob,files[0].name.replace(/\.pdf$/i,"")+".docx");}
 $("status").textContent="Concluído. O download deve começar automaticamente.";
}catch(e){$("status").textContent=e.message||"Não foi possível processar o arquivo.";}finally{$("process").disabled=false;}}
$("process").onclick=processFiles;renderOptions();
if("serviceWorker" in navigator && location.protocol.startsWith("http"))navigator.serviceWorker.register("./sw.js").catch(()=>{});
