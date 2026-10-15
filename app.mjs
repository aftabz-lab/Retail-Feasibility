import {importWorkbook} from './uploads.mjs';
import {WorkbookEngine,makeDefaultState,copy,cellKey,setCategory,WORKBOOK_RULES,structuralErrors} from './calculation.mjs';
import {DEFAULT_SIGNATORIES,resolveSignatories,getSignatoryAutoLink,getOpenedDesignation} from './approval-rules.mjs';
import {shellHtml,VIEWS} from './view.mjs';
import {readSaved,saveProject} from './storage.mjs';
import {geocodeLocationArea,loadOutletLocations,outletsWithinRadius} from './geo-outlets.js';
import {resolveGoogleMapsLink} from './map-link-resolver.js';
import {buildRulesExcel,buildPdf,downloadBlob,safeFilename} from './exports.mjs';

let state=makeDefaultState();state.signatories=copy(DEFAULT_SIGNATORIES);
let view='entry',assets=[],outlets=[],engine,status={kind:'loading',message:'Loading saved project…'},saveTimer,mapToken=0,busy=false;
let storageWarning='',uploadQueue=Promise.resolve();
const root=document.querySelector('#app');
function recalculate(){engine=new WorkbookEngine(state.sheets,state.overrides);return engine.scan();}
function render(){const errors=recalculate();if(errors.length&&status.kind==='ready')status={kind:'warning',message:`${errors.length} formula result(s) unavailable. First: ${errors[0].sheet}!${errors[0].address} (${errors[0].code}).`};root.innerHTML=shellHtml(state,engine,assets,view,status);}
function report(kind,message){status={kind,message};const node=document.querySelector('#status');if(node){node.className='status status-'+kind;node.textContent=message;}}
function persist(){clearTimeout(saveTimer);saveTimer=setTimeout(async()=>{try{await saveProject(copy(state));}catch(e){storageWarning=e.message;report('warning','Browser saving failed. Use Backup project to keep your changes.');}},350);}
function changed(message='Inputs updated.'){report('ready',storageWarning||message);render();persist();}
function displayAssets(){assets=[...assets.filter(a=>!a.id.startsWith('upload-')),...(state.signatureUploads||[])];}

async function uploadSource(file,type){
 if(!file)return;if(file.size>25*1024*1024)throw Error('Please upload an Excel file smaller than 25 MB.');
 report('loading','Reading '+file.name+'…');
 const wb=globalThis.XLSX.read(await file.arrayBuffer(),{type:'array',cellFormula:true,cellNF:true,cellDates:false});
 const next=importWorkbook(wb,type,file.name,state);
 // Save the version before reporting that the upload is retained.
 try{await saveProject(next);storageWarning='';}catch(e){storageWarning='Browser saving is unavailable. Backup this project to keep uploaded versions.';}
 state=next;if(type==='location')mapToken++;changed(storageWarning||`${file.name} loaded. Linked values updated.`);if(type==='location'&&state.project.district)void checkMap();
}

async function checkMap(){
 const token=++mapToken;const input=state.project.googleMapLocation.trim(),district=state.project.district;
 state.locationAssessment=null;
 if(!input||!district){state.project.existingOutlets=null;state.locationStatus='Enter Google Map Location and select District to calculate the 1 KM count.';changed();return;}
 state.project.existingOutlets=null;state.locationStatus='Resolving location and checking the selected district…';changed();
 try{
   if(!outlets.length){const loaded=await loadOutletLocations();outlets=loaded.outlets;}
   const target=await geocodeLocationArea(input,district,fetch,resolveGoogleMapsLink);if(token!==mapToken)return;
   if(!outlets.length)throw Error('The outlet map is unavailable. Enter the count manually or retry.');
   const matches=outletsWithinRadius(outlets,target,district,1);
   state.project.existingOutlets=matches.length;state.project.existingCountMode='map';state.locationAssessment={latitude:Number(target.lat),longitude:Number(target.lon),targetLabel:target.label||input,geocoder:target.provider||'',district,existingOutletMatches:matches};
   const mapped=outlets.filter(o=>o.district===district).length;
   state.locationStatus=`${matches.length} existing outlet(s) within 1 KM. Checked ${mapped} mapped outlets in ${district}.`;
   changed('Google Map Location and existing outlet count updated.');
 }catch(error){if(token!==mapToken)return;state.project.existingOutlets=null;state.locationAssessment=null;state.locationStatus=error.message+' The count remains editable manually.';changed('Automatic map check could not complete. Enter the count manually or recheck.');}
}

root.addEventListener('change',async event=>{
 const el=event.target;
 try{
  if(el.dataset.cell){const key=cellKey(el.dataset.sheet,el.dataset.cell),type=el.dataset.type;let value=el.value;if(type!=='text'){if(value===''){delete state.overrides[key];changed('Sheet value restored.');return;}value=Number(value);if(!Number.isFinite(value))throw Error('Enter a valid number.');if(type==='percent')value/=100;if(el.min==='0'&&value<0)throw Error('Quantity and salary cannot be negative.');if(el.step==='1'&&!Number.isInteger(value))throw Error('Enter a whole quantity.');}state.overrides[key]=value;changed();return;}
  if(el.dataset.project){const key=el.dataset.project;let val=el.type==='checkbox'?el.checked:el.type==='number'?(el.value===''?null:Number(el.value)):el.value;if(el.min==='0'&&(val<0||!Number.isInteger(val)))throw Error('Enter a whole non-negative outlet count.');state.project[key]=val;
   if(key==='openedBy'){const designation=getOpenedDesignation(val);if(designation)state.project.openedDesignation=designation;}
   if(key==='existingOutlets'){mapToken++;state.project.existingCountMode='manual';state.locationStatus='Manually entered count. Recheck map to calculate automatically.';}
   changed();if(key==='googleMapLocation'||key==='district')void checkMap();return;
  }
  if(el.dataset.sign!==undefined){const p=state.signatories[Number(el.dataset.sign)];p[el.dataset.key]=el.type==='checkbox'?el.checked:el.value;changed();return;}
  if(el.dataset.action==='category-version'){setCategory(state,el.value);changed('Category '+el.value+' selected.');}
 }catch(error){report('error',error.message);}
});

root.addEventListener('click',async event=>{
 const el=event.target.closest('button,[data-view]');if(!el)return;
 if(el.dataset.view){view=el.dataset.view;render();return;}
 const action=el.dataset.action;if(!action)return;
 try{
  if(action.startsWith('upload-')){const kind=action.replace('upload-','');document.querySelector('#'+kind+'-file')?.click();return;}
  if(action==='rule'){delete state.overrides[cellKey(el.dataset.sheet,el.dataset.cell)];changed('Sheet rule restored.');}
  else if(action==='map')await checkMap();
  else if(action==='sign-mode'){const i=Number(el.dataset.index),p=state.signatories[i],r=resolveSignatories(state)[i];if(!p.manualOverride){p.name=r.name;p.designation=r.designation;}p.manualOverride=!p.manualOverride;changed();}
  else if(action==='sign-remove'){if(state.signatories.length>1){state.signatories.splice(Number(el.dataset.index),1);changed();}}
  else if(action==='sign-add'){state.signatories.push({role:'Signatory',name:'',designation:'',signatureId:'',includeInPdf:false});changed();}
  else if(action==='reset'){mapToken++;const categories=state.categories,signatureUploads=state.signatureUploads;state=makeDefaultState();state.signatories=copy(DEFAULT_SIGNATORIES);state.categories=categories;state.signatureUploads=signatureUploads;displayAssets();changed('Default inputs restored. Uploaded category versions remain available.');}
  else if(action==='backup'){downloadBlob(new Blob([JSON.stringify({application:'Retail Feasibility',version:1,state},null,2)],{type:'application/json'}),'Retail-Feasibility-project.json');report('ready','Project backup downloaded.');}
  else if(action==='restore')document.querySelector('#project-file').click();
  else if(['excel','pdf','management'].includes(action)){
   if(busy)return;busy=true;el.disabled=true;report('loading','Creating '+(action==='excel'?'Excel with Rules':'PDF')+'…');
   try{const snapshot=copy(state),calculated=new WorkbookEngine(snapshot.sheets,snapshot.overrides),allErrors=calculated.scan(),errors=structuralErrors(allErrors);if(errors.length)throw Error('Fix formula error at '+errors[0].sheet+'!'+errors[0].address+' before exporting.');const name=safeFilename(calculated.value('INFORMATION','C3'));if(action==='excel'){const bytes=await buildRulesExcel(snapshot,calculated,copy(assets));downloadBlob(new Blob([bytes],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}),'Retail-Feasibility-'+name+'-with-Rules.xlsx');}else{const pdf=await buildPdf(snapshot,calculated,copy(assets),{management:action==='management'});downloadBlob(pdf.output('blob'),'Retail-Feasibility-'+name+(action==='management'?'-Management':'')+'.pdf');}report(allErrors.length?'warning':'ready',allErrors.length?'Download created with the formula errors shown in the model.':'Download created.');}finally{busy=false;el.disabled=false;}
  }
 }catch(error){busy=false;el.disabled=false;report('error',error.message);}
});

for(const [id,type]of [['location','location'],['investment','Investment'],['category','category']])document.querySelector('#'+id+'-file').addEventListener('change',e=>{const file=e.target.files[0];e.target.value='';if(!file)return;uploadQueue=uploadQueue.then(async()=>{root.inert=true;try{await uploadSource(file,type);}catch(error){report('error',error.message);}finally{root.inert=false;}});});
document.querySelector('#signature-file').addEventListener('change',async event=>{const file=event.target.files[0];event.target.value='';if(!file)return;try{if(!['image/png','image/jpeg'].includes(file.type)||file.size>3*1024*1024)throw Error('Use a PNG or JPEG signature smaller than 3 MB.');const dataUrl=await new Promise((resolve,reject)=>{const r=new FileReader();r.onload=()=>resolve(r.result);r.onerror=()=>reject(Error('Could not read signature.'));r.readAsDataURL(file);});const asset={id:'upload-'+Date.now()+'-'+Math.random().toString(36).slice(2,7),label:file.name,dataUrl,extension:file.type==='image/png'?'png':'jpeg'};state.signatureUploads.push(asset);displayAssets();changed('Signature uploaded. Select it in the signature manager.');}catch(error){report('error',error.message);}});
document.querySelector('#project-file').addEventListener('change',async event=>{const file=event.target.files[0];event.target.value='';if(!file)return;try{if(file.size>25*1024*1024)throw Error('Backup is too large.');const payload=JSON.parse(await file.text()),next=payload.state;if(payload.application!=='Retail Feasibility'||payload.version!==1||!next?.sheets?.INFORMATION||!Array.isArray(next.categories)||!Array.isArray(next.signatories))throw Error('This is not a Retail Feasibility project backup.');const allowed=new Set(Object.keys(WORKBOOK_RULES));if([...allowed].some(n=>!next.sheets[n])||Object.keys(next.sheets).some(n=>!allowed.has(n)))throw Error('Backup contains an unsupported sheet.');for(const cat of next.categories)if(cat.sheet?.name!=='category')throw Error('Backup has an invalid category version.');const errors=structuralErrors(new WorkbookEngine(next.sheets,next.overrides).scan());if(errors.length)throw Error('Backup contains a formula error.');await saveProject(next);mapToken++;state=next;displayAssets();changed('Project restored with all category versions.');}catch(error){report('error',error.message);}});

async function boot(){
 const results=await Promise.allSettled([readSaved(),fetch('./signs/manifest.json').then(r=>{if(!r.ok)throw Error('Signature manifest is unavailable');return r.json();}),loadOutletLocations()]);
 if(results[0].status==='fulfilled'&&results[0].value){const saved=results[0].value;if(saved.schemaVersion===1&&saved.sheets?.INFORMATION)state=saved;}
 if(results[0].status==='rejected')storageWarning='Browser saving is unavailable. Use Backup project to retain changes.';
 if(results[1].status==='fulfilled')assets=results[1].value.signatures.map(a=>({...a,url:'./signs/'+a.file}));
 if(results[2].status==='fulfilled')outlets=results[2].value.outlets;
 displayAssets();const issues=recalculate();status={kind:issues.length?'error':storageWarning?'warning':'ready',message:issues.length?'Formula error: '+issues[0].sheet+'!'+issues[0].address:storageWarning||'Ready. Inputs and uploaded category versions are saved in this browser.'};render();
}
globalThis.RetailFeasibility={getState:()=>copy(state),getEngine:()=>engine,importWorkbook,checkMap,views:VIEWS};
render();void boot();
