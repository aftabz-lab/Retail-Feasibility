import {SHEET_ORDER,INFO_FIELDS,cellKey,formatCell,columnName,addressParts} from './calculation.mjs?v=retail-v2-20261015';
import {resolveSignatories} from './approval-rules.mjs';
import {STORE_TYPES,MANPOWER_RULES,displayCellValue} from './controls.mjs?v=retail-v2-20261015';

export function safeFilename(value){return String(value||'Project').replace(/[^A-Za-z0-9]+/g,'-').replace(/^-|-$/g,'').slice(0,65)||'Project';}
export function downloadBlob(blob,name){const url=URL.createObjectURL(blob),link=document.createElement('a');link.href=url;link.download=name;document.body.appendChild(link);link.click();link.remove();setTimeout(()=>URL.revokeObjectURL(url),2000);}
async function dataUrl(asset){if(!asset)return null;if(asset.dataUrl)return asset.dataUrl;const response=await fetch(asset.url);if(!response.ok)throw Error('Could not load signature '+asset.label);const bytes=new Uint8Array(await response.arrayBuffer());let binary='';for(let i=0;i<bytes.length;i+=8192)binary+=String.fromCharCode(...bytes.subarray(i,i+8192));return `data:image/${asset.extension||'png'};base64,${btoa(binary)}`;}
const border={top:{style:'thin',color:{argb:'FFB8C6CF'}},bottom:{style:'thin',color:{argb:'FFB8C6CF'}},left:{style:'thin',color:{argb:'FFB8C6CF'}},right:{style:'thin',color:{argb:'FFB8C6CF'}}};
let fontPromise;
async function loadPdfFonts(){
 if(globalThis.RETAIL_PDF_FONTS)return globalThis.RETAIL_PDF_FONTS;
 if(!fontPromise)fontPromise=Promise.all(['regular','bold'].map(async name=>{const response=await fetch(`./fonts/retail-sans-${name}.ttf`);if(!response.ok)throw Error('PDF font did not load. Check that the fonts folder was uploaded.');const bytes=new Uint8Array(await response.arrayBuffer());let text='';for(let i=0;i<bytes.length;i+=8192)text+=String.fromCharCode(...bytes.subarray(i,i+8192));return btoa(text);})).then(([normal,bold])=>({normal,bold})).catch(error=>{fontPromise=null;throw error;});
 return fontPromise;
}

export async function buildRulesExcel(state,engine,assets){
 const ExcelJS=globalThis.ExcelJS;if(!ExcelJS?.Workbook)throw Error('Excel library did not load. Reload the page.');
 const workbook=new ExcelJS.Workbook();workbook.creator='Shwapno';workbook.title='Retail Feasibility';workbook.subject=String(engine.value('INFORMATION','C3'));workbook.created=new Date();workbook.calcProperties={fullCalcOnLoad:true,forceFullCalc:true,calcMode:'auto'};
 // Create every referenced sheet before writing the original formulas.
 for(const name of SHEET_ORDER)workbook.addWorksheet(name);
 for(const name of SHEET_ORDER){const source=state.sheets[name],sheet=workbook.getWorksheet(name);sheet.views=[{state:'frozen',xSplit:name===' Monthly PNL'||name==='Yearly -5yrs PNL'?1:0,ySplit:name===' Monthly PNL'?2:name==='Investment'?4:1,showGridLines:false}];
   for(const [a,c]of Object.entries(source.cells)){const cell=sheet.getCell(a),key=cellKey(name,a);if(Object.hasOwn(state.overrides,key))cell.value=state.overrides[key];else if(c.f){const result=engine.value(name,a);cell.value={formula:c.f,result:typeof result==='string'&&/^#(?:NUM!|DIV\/0!)$/.test(result)?{error:result}:result};}else cell.value=c.v??null;cell.numFmt=c.z||'General';cell.font={name:'Calibri',size:11,bold:Boolean(c.bold),color:{argb:'FF000000'}};cell.alignment={vertical:'middle',wrapText:true,horizontal:typeof engine.value(name,a)==='number'?'right':'left'};cell.border=border;}
   // Include intentionally blank manual fields in the downloaded Information sheet.
   if(name==='INFORMATION')for(const [a]of INFO_FIELDS){const key=cellKey(name,a);if(Object.hasOwn(state.overrides,key))sheet.getCell(a).value=state.overrides[key];}
   for(const [a,value]of Object.entries(state.overrides)){const prefix=name+'!';if(a.startsWith(prefix)&&!source.cells[a.slice(prefix.length)])sheet.getCell(a.slice(prefix.length)).value=value;}
   for(const [col,width]of Object.entries(source.widths||{}))sheet.getColumn(col).width=Math.min(Math.max(Number(width)||12,8),70);
   for(const [row,height]of Object.entries(source.heights||{}))sheet.getRow(Number(row)).height=Math.max(height||18,18);
   for(const merge of source.merges||[])sheet.mergeCells(merge);
   sheet.pageSetup={paperSize:9,orientation:name===' Monthly PNL'?'landscape':'portrait',fitToPage:true,fitToWidth:1,fitToHeight:name==='Investment'?0:1,margins:{left:.25,right:.25,top:.35,bottom:.35,header:.15,footer:.15}};
   sheet.headerFooter={oddHeader:'&C&B Retail Feasibility',oddFooter:'&L Shwapno &R Page &P of &N'};
 }
 const info=workbook.getWorksheet('INFORMATION');info.getColumn('B').width=38;info.getColumn('C').width=28;info.getColumn('E').width=29;['F','G','H'].forEach(c=>info.getColumn(c).width=c==='F'?10:18);info.getRow(3).height=50;info.getRow(22).height=36;
 const staffing=workbook.addWorksheet('Manpower Rules',{state:'hidden'});staffing.addRow(['Manpower Retail Feasibility']);staffing.addRow([]);staffing.addRow(['Role',...STORE_TYPES]);MANPOWER_RULES.forEach(row=>staffing.addRow(row));staffing.addRow(['Total',...STORE_TYPES.map((_,i)=>MANPOWER_RULES.reduce((sum,r)=>sum+r[i+1],0))]);
 info.getCell('E24').value='Select Store Type';info.getCell('F24').value=state.project.storeType||'';info.getCell('F24').dataValidation={type:'list',allowBlank:true,formulae:['"'+STORE_TYPES.join(',')+'"']};info.getCell('E25').value='Manpower mode';info.getCell('F25').value=state.manpowerMode==='automatic'?'Automatic':'Manual';info.getCell('F25').dataValidation={type:'list',allowBlank:false,formulae:['"Automatic,Manual"']};
 info.mergeCells('F24:H24');info.mergeCells('F25:H25');info.getRow(24).height=30;info.getCell('F24').alignment={wrapText:true,vertical:'middle'};
 for(let r=6;r<=20;r++){const value=engine.value('INFORMATION','F'+r),fallback=typeof value==='number'?String(value):'""',matrixRow=r-2;info.getCell('F'+r).value={formula:`IF($F$25="Automatic",IF($F$24="Super Store",'Manpower Rules'!B${matrixRow},IF($F$24="Mega Store",'Manpower Rules'!C${matrixRow},IF($F$24="PNP Convenience Store",'Manpower Rules'!D${matrixRow},""))),${fallback})`,result:value??''};}
 const monthly=workbook.getWorksheet(' Monthly PNL');monthly.getColumn('A').width=39;for(let c=2;c<=14;c++)monthly.getColumn(c).width=18;monthly.getRow(2).height=48;
 const yearly=workbook.getWorksheet('Yearly -5yrs PNL');yearly.getColumn('A').width=43;for(let c=2;c<=13;c++)yearly.getColumn(c).width=20;yearly.getRow(10).height=48;
 const budget=workbook.getWorksheet('Investment');budget.getColumn('A').width=25;budget.getColumn('B').width=70;for(let c=3;c<=7;c++)budget.getColumn(c).width=18;
 const loc=workbook.getWorksheet('location');loc.getRow(3).height=56;loc.getRow(5).height=85;loc.getColumn('B').width=38;loc.getColumn('C').width=38;
 if(state.project.includeExistingOutlets){const nearby=workbook.addWorksheet('Existing outlets within 1 KM');nearby.addRow(['Existing Outlet No. within 1 KM',state.project.existingOutlets??'']);nearby.addRow(['District',state.project.district]);nearby.addRow(['Google Map Location',state.project.googleMapLocation]);nearby.addRow(['Coordinates',state.locationAssessment?`${state.locationAssessment.latitude}, ${state.locationAssessment.longitude}`:'']);nearby.addRow(['Assessment',state.locationStatus]);nearby.addRow([]);nearby.addRow(['Outlet code','Outlet name','RHO','Zonal','Distance (KM)','Latitude','Longitude']);for(const o of state.locationAssessment?.existingOutletMatches||[])nearby.addRow([o.code,o.name,o.rho,o.zonal,o.distanceKm,o.lat,o.lon]);nearby.columns.forEach((c,i)=>c.width=i===0?40:i===1?48:25);nearby.getCell('B4').alignment={wrapText:true};nearby.getRow(3).height=44;nearby.getRow(5).height=44;nearby.views=[{state:'frozen',ySplit:7}];}
 const approval=workbook.addWorksheet('Approval');approval.columns=[{width:23},{width:38},{width:47},{width:22},{width:18}];approval.addRow(['Role','Full name','Designation','Signature image','Include in PDF']);
 for(const person of resolveSignatories(state)){const row=approval.addRow([person.role,person.name,person.designation,person.signatureId,person.includeInPdf?'Yes':'No']);row.height=50;row.eachCell(c=>{c.alignment={wrapText:true,vertical:'middle'};c.border=border;});}
 for(const [i,person]of resolveSignatories(state).entries()){if(!person.includeInPdf||!person.signatureId)continue;const asset=assets.find(a=>a.id===person.signatureId);if(!asset)continue;const image=await dataUrl(asset);const id=workbook.addImage({base64:image,extension:asset.extension==='jpeg'?'jpeg':'png'});approval.addImage(id,{tl:{col:3.08,row:i+1.1},br:{col:3.92,row:i+1.85},editAs:'oneCell'});}
 for(const sheet of workbook.worksheets){sheet.getRow(1).font={name:'Calibri',size:11,bold:true};sheet.properties.defaultRowHeight=20;}
 return workbook.xlsx.writeBuffer();
}

const pdfNumber=(value,format)=>typeof value==='number'?formatCell(value,format||'#,##0'):String(value??'');
function clean(value){return String(value??'').replace(/[–—]/g,'-').replace(/[’‘]/g,"'").replace(/[“”]/g,'"').replace(/\r/g,'');}
function drawText(doc,value,x,y,width,{size=8,bold=false,align='left',lines=1}={}){doc.setFont('RetailSans',bold?'bold':'normal');doc.setFontSize(size);doc.setTextColor(0,0,0);const content=clean(value);const split=doc.splitTextToSize(content,Math.max(width-6,5));if(split.length>lines)throw Error('PDF text is too long to fit: '+content.slice(0,75));const dx=align==='right'?x+width-3:align==='center'?x+width/2:x+3;doc.text(split,dx,y,{align,lineHeightFactor:1.05});}
function cellLayout(doc,value,width,size,bold){let fontSize=size;doc.setFont('RetailSans',bold?'bold':'normal');doc.setFontSize(fontSize);if(/^[\d,().+\-\s%]+$/.test(clean(value)))while(doc.getTextWidth(clean(value))>width-6&&fontSize>5.5){fontSize=Math.max(5.5,fontSize-.1);doc.setFontSize(fontSize);}return{fontSize,lines:doc.splitTextToSize(clean(value),width-6).length};}
function tableRow(doc,values,x,y,widths,height,{size=8,bold=false,aligns=[],maxLines=2}={}){let cursor=x;doc.setDrawColor(0,0,0);doc.setLineWidth(.35);values.forEach((v,i)=>{doc.rect(cursor,y,widths[i],height);const layout=cellLayout(doc,v,widths[i],size,bold),cellSize=layout.fontSize,lines=Math.min(maxLines,layout.lines);drawText(doc,v,cursor,y+height/2-(lines-1)*cellSize*.52+cellSize*.32,widths[i],{size:cellSize,bold,align:aligns[i]||'left',lines:maxLines});cursor+=widths[i];});return y+height;}
function header(doc,title,project,page){const width=doc.internal.pageSize.getWidth();drawText(doc,'RETAIL FEASIBILITY',28,25,width-56,{size:15,bold:true});drawText(doc,title,28,41,width-56,{size:10,bold:true});const date=new Intl.DateTimeFormat('en-GB',{day:'2-digit',month:'short',year:'numeric',timeZone:'Asia/Dhaka'}).format(new Date());drawText(doc,date,width-158,24,130,{size:8,align:'right'});if(project){drawText(doc,project,28,56,width-56,{size:8,lines:2});}doc.setFontSize(7);doc.text('Shwapno',28,doc.internal.pageSize.getHeight()-16);doc.text(`${page} / 3`,width-28,doc.internal.pageSize.getHeight()-16,{align:'right'});}
async function imageAssets(assets){const results=await Promise.all(assets.map(async a=>[a.id,await dataUrl(a)]));return new Map(results);}
function putImage(doc,url,x,y,width,height){if(!url)return;const properties=doc.getImageProperties(url),scale=Math.min(width/properties.width,height/properties.height),w=properties.width*scale,h=properties.height*scale;doc.addImage(url,/image\/jpe?g/.test(url)?'JPEG':'PNG',x+(width-w)/2,y+(height-h)/2,w,h,undefined,'FAST');}
function reviewSignature(doc,people,images){const person=people.find(p=>p.signatureId==='saiful-alam-rasel-signature'&&p.includeInPdf);if(!person)return;const width=doc.internal.pageSize.getWidth(),height=doc.internal.pageSize.getHeight();putImage(doc,images.get(person.signatureId),width-160,height-80,100,38);drawText(doc,person.name,width-195,height-34,150,{size:7,align:'center'});}
function signatureBlocks(doc,people,images,y){const pageWidth=doc.internal.pageSize.getWidth();const groups=[people.slice(0,4),people.slice(4,7)];for(let start=7;start<people.length;start+=3)groups.push(people.slice(start,start+3));const rowHeight=75;
 if(y+groups.filter(g=>g.length).length*rowHeight>doc.internal.pageSize.getHeight()-28)throw Error('Too many signatories for the three-page PDF. Reduce signatories before downloading.');
 for(const group of groups){if(!group.length)continue;const width=(pageWidth-56)/group.length;for(let i=0;i<group.length;i++){const p=group[i],x=28+i*width,lineWidth=Math.min(width-18,170);if(p.includeInPdf)putImage(doc,images.get(p.signatureId),x+(width-104)/2,y-4,104,40);doc.setLineWidth(.6);doc.setLineDashPattern([1.6,1.8],0);doc.line(x+(width-lineWidth)/2,y+23,x+(width+lineWidth)/2,y+23);doc.setLineDashPattern([],0);drawText(doc,p.role,x,y+39,width,{size:6.5,align:'center'});drawText(doc,p.name,x,y+50,width,{size:7,bold:true,align:'center',lines:2});drawText(doc,p.designation,x,y+64,width,{size:6.1,align:'center',lines:2});}y+=rowHeight;}
}

export async function buildPdf(state,engine,assets,{management=false}={}){
 const JsPDF=globalThis.jspdf?.jsPDF;if(!JsPDF)throw Error('PDF library did not load. Reload the page.');
 const doc=new JsPDF({unit:'pt',format:'a4',orientation:'portrait',compress:true});const fonts=await loadPdfFonts();doc.addFileToVFS('RetailSans.ttf',fonts.normal);doc.addFont('RetailSans.ttf','RetailSans','normal');doc.addFileToVFS('RetailSansBold.ttf',fonts.bold);doc.addFont('RetailSansBold.ttf','RetailSans','bold');doc.setFont('RetailSans','normal');const people=resolveSignatories(state),enabledIds=new Set(people.filter(p=>p.includeInPdf).map(p=>p.signatureId));const images=management?new Map():await imageAssets(assets.filter(a=>enabledIds.has(a.id)));
 const project=String(displayCellValue(state,engine,'INFORMATION','C3')??'');doc.setProperties({title:'Retail Feasibility - '+project,author:'Shwapno',subject:management?'Management feasibility':'Retail feasibility with approvals'});
 // Page 1: Information and manpower only.
 header(doc,'Information',project,1);let y=80;
 const leftX=28,leftWidths=[169,107],rightX=318,rightWidths=[114,30,48,57];
 tableRow(doc,['BUSINESS FEASIBILITY INFORMATION','Value'],leftX,y,leftWidths,25,{size:8,bold:true});let leftY=y+25;
 for(const [a,label]of INFO_FIELDS){const source=state.sheets.INFORMATION.cells[a],v=!source&&!Object.hasOwn(state.overrides,cellKey('INFORMATION',a))?'':engine.value('INFORMATION',a),text=pdfNumber(v,source?.z);const lines=Math.max(doc.splitTextToSize(clean(label),leftWidths[0]-6).length,doc.splitTextToSize(clean(text),leftWidths[1]-6).length);const height=Math.max(22,lines*9+9);leftY=tableRow(doc,[label,text],leftX,leftY,leftWidths,height,{size:8,aligns:['left',typeof v==='number'?'right':'left'],maxLines:6});}
 tableRow(doc,['MANPOWER','Qty','Salary','Total'],rightX,y,rightWidths,25,{size:8,bold:true,maxLines:2});let rightY=y+25;
 for(let r=6;r<=20;r++){rightY=tableRow(doc,[state.sheets.INFORMATION.cells['E'+r]?.v,...['F','G','H'].map(c=>pdfNumber(engine.value('INFORMATION',c+r)))],rightX,rightY,rightWidths,22,{size:8,aligns:['left','right','right','right']});}
 rightY=tableRow(doc,['Total manpower',pdfNumber(displayCellValue(state,engine,'INFORMATION','F22')),'',''],rightX,rightY,rightWidths,24,{size:8,bold:true,aligns:['left','right','right','right']});
 if(state.project.includeExistingOutlets){let nearbyY=Math.max(leftY,rightY)+22;drawText(doc,'Existing Outlet No. within 1 KM: '+(state.project.existingOutlets??'Pending map check'),28,nearbyY,539,{size:8,bold:true});drawText(doc,'District: '+state.project.district,28,nearbyY+15,539,{size:8});drawText(doc,'Google Map Location: '+state.project.googleMapLocation,28,nearbyY+30,539,{size:7,lines:4});}
 if(!management)reviewSignature(doc,people,images);
 // Page 2: Monthly PNL, including the original annual and cost calculations.
 doc.addPage('a4','landscape');header(doc,'Monthly PNL',project,2);const page2Width=doc.internal.pageSize.getWidth(),widths=[176,...Array(13).fill((page2Width-56-176)/13)];let monthY=75;
 monthY=tableRow(doc,['Metric',...Array.from({length:12},(_,i)=>`Month ${i+1}`),'Annual'],28,monthY,widths,19,{size:6.6,bold:true,maxLines:1});
 for(let r=3;r<=60;r++){const label=state.sheets[' Monthly PNL'].cells['A'+r]?.v;if(!label)continue;const values=[label,...Array.from({length:13},(_,i)=>{const a=columnName(i+2)+r,c=state.sheets[' Monthly PNL'].cells[a];return c?pdfNumber(engine.value(' Monthly PNL',a),c.z):'';})];monthY=tableRow(doc,values,28,monthY,widths,9,{size:6.25,bold:[12,42,43,46,48,49].includes(r),aligns:['left',...Array(13).fill('right')],maxLines:1});}
 if(monthY>doc.internal.pageSize.getHeight()-85)throw Error('Monthly report exceeds the three-page PDF area.');if(!management)reviewSignature(doc,people,images);
 // Page 3: Five-year PNL and the reference dashboard's 4 + 3 approval arrangement.
 doc.addPage('a4','portrait');header(doc,'Yearly -5yrs PNL',project,3);let yearY=75;const yearWidths=[150,...Array(7).fill((doc.internal.pageSize.getWidth()-56-150)/7)];
 for(let r=1;r<=70;r++){const label=state.sheets['Yearly -5yrs PNL'].cells['A'+r]?.v;if(!label&&![8,60,63].includes(r))continue;
   if(r===10){yearY=tableRow(doc,[label,project],28,yearY,[yearWidths[0],yearWidths.slice(1).reduce((a,b)=>a+b,0)],25,{size:7,maxLines:2});continue;}
   const values=Array.from({length:8},(_,i)=>{const a=columnName(i+1)+r,c=state.sheets['Yearly -5yrs PNL'].cells[a];if(!c)return '';return pdfNumber(engine.value('Yearly -5yrs PNL',a),c.z);});
   // D1 repeats the project name already shown in the page header.
   if(r===1)values[3]='';
   const fontSize=6.55,bold=[5,8,49,51,54,56,60,61,63,64,66,67,68,69,70].includes(r);const lineCount=Math.max(...values.map((v,i)=>cellLayout(doc,v,yearWidths[i],fontSize,bold).lines));const rowHeight=Math.max(9.2,lineCount*fontSize+2.7);
   yearY=tableRow(doc,values,28,yearY,yearWidths,rowHeight,{size:fontSize,bold,aligns:['left',...Array(7).fill('right')],maxLines:2});
 }
 if(!management){yearY+=16;drawText(doc,'APPROVAL & SIGNATURES',28,yearY,539,{size:8,bold:true,align:'center'});signatureBlocks(doc,people,images,yearY+10);}
 if(doc.getNumberOfPages()!==3)throw Error('PDF must contain exactly three pages.');return doc;
}

