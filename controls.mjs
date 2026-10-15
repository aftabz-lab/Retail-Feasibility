import {WorkbookEngine,WORKBOOK_RULES,INFO_FIELDS,EXPENSE_INPUTS,makeDefaultState,copy,cellKey,columnName,formulaDependencies} from './calculation.mjs?v=retail-v2-20261015';

export const STORE_TYPES=['Super Store','Mega Store','PNP Convenience Store'];
export const MANPOWER_RULES=[
 ['OM-Parmanent',1,2,2],['ICMO-Parmanent',1,1,0],['CRM',0,1,0],['IMO',0,0,0],['CG',6,10,5],['COMMODITY',2,4,2],['PROTEIN',5,12,0],['PERISHABLES',2,5,0],['GM & LIFESTYLE',2,4,2],['POS',5,10,4],['PORTER',2,6,2],['BSM',1,2,0],['BK STR',4,7,3],['SECURITY',4,6,1],['CLEANER',4,9,4]
];
const splitKey=key=>[key.slice(0,key.lastIndexOf('!')),key.slice(key.lastIndexOf('!')+1)];
const blank=value=>value===null||value===undefined||value==='';
export function missingInputs(state,sheet,address,{ignoreSales=false}={}){
 const visited=new Set(),missing=new Set();
 function visit(key){if(visited.has(key))return;visited.add(key);if(ignoreSales&&key===cellKey('INFORMATION','C14'))return;
  if(Object.hasOwn(state.overrides,key)){if(blank(state.overrides[key]))missing.add(key);return;}
  const [s,a]=splitKey(key);for(const dep of formulaDependencies(state.sheets,s,a))visit(dep);
 }
 visit(cellKey(sheet,address));return [...missing];
}
export function displayCellValue(state,engine,sheet,address){
 const key=cellKey(sheet,address);
 if(sheet==='INFORMATION'&&address==='C16'&&!Object.hasOwn(state.overrides,key)&&state.activeCategory==='Default'&&blank(engine.value('INFORMATION','C14')))return '';
 if(missingInputs(state,sheet,address).length)return '';
 return engine.value(sheet,address);
}
export function valueForInput(value,type='number'){
 if(blank(value))return '';if(typeof value!=='number')return value;
 return type==='percent'?Number((value*100).toFixed(9)):Number(value.toFixed(2));
}
export function invalidateCalculatedOverrides(state,changedKeys){
 const changed=new Set(changedKeys);let again=true;
 while(again){again=false;for(const [name,sheet]of Object.entries(state.sheets))for(const [a,c]of Object.entries(sheet.cells)){if(!c.f)continue;const key=cellKey(name,a);if(changed.has(key))continue;if(formulaDependencies(state.sheets,name,a).some(dep=>changed.has(dep))){changed.add(key);delete state.overrides[key];again=true;}}}
}
export function setEntryCell(state,sheet,address,value){
 invalidateCalculatedOverrides(state,[cellKey(sheet,address)]);state.overrides[cellKey(sheet,address)]=value;
 if(sheet==='INFORMATION'&&/^[FG](?:[6-9]|1\d|20)$/.test(address))state.manpowerMode='manual';
}
export function restoreSheetRule(state,sheet,address){
 const key=cellKey(sheet,address);invalidateCalculatedOverrides(state,[key]);
 if(sheet==='INFORMATION'&&['C3','C5','C7','C10','C11'].includes(address)&&!state.uploadSources.location)state.overrides[key]=null;
 else if(sheet==='INFORMATION'&&address==='C12'&&!state.uploadSources.Investment)state.overrides[key]=null;
 else delete state.overrides[key];
}
export function selectStoreType(state,type){
 if(!STORE_TYPES.includes(type)){state.project.storeType='';state.manpowerMode='manual';return;}
 const index=STORE_TYPES.indexOf(type)+1,keys=[];
 for(let r=6;r<=20;r++){const f=cellKey('INFORMATION','F'+r),g=cellKey('INFORMATION','G'+r);state.overrides[f]=MANPOWER_RULES[r-6][index];keys.push(f);if(blank(state.overrides[g])&&Object.hasOwn(state.overrides,g)){state.overrides[g]=WORKBOOK_RULES.INFORMATION.cells['G'+r].v;keys.push(g);}}
 invalidateCalculatedOverrides(state,keys);state.project.storeType=type;state.manpowerMode='automatic';
 state.sheets.INFORMATION.cells.F22.f='SUM(F6:F20)';
}
export function row43Status(state,engine=new WorkbookEngine(state.sheets,state.overrides),{ignoreSales=false}={}){
 const missing=new Set();for(let i=2;i<=14;i++)for(const k of missingInputs(state,' Monthly PNL',columnName(i)+'43',{ignoreSales}))missing.add(k);
 const values=Array.from({length:13},(_,i)=>engine.value(' Monthly PNL',columnName(i+2)+'43'));
 const ready=missing.size===0,positive=v=>typeof v==='number'&&Number.isFinite(v)&&Math.round(v*100)>0;
 return {ready,passes:ready&&values.every(positive),values,missing:[...missing]};
}
export function missingInputLabel(key){const[s,a]=splitKey(key);if(s==='INFORMATION'){const f=INFO_FIELDS.find(x=>x[0]===a);if(f)return f[1];if(/^[FG]\d+$/.test(a)){const r=Number(a.slice(1));return (WORKBOOK_RULES.INFORMATION.cells['E'+r]?.v||a)+(a[0]==='F'?' quantity':' salary');}}if(s===' Monthly PNL')return WORKBOOK_RULES[s].cells['A'+a.slice(1)]?.v||key;return key;}
export function autoCorrectSales(state){
 const template=copy(state);invalidateCalculatedOverrides(template,[cellKey('INFORMATION','C14')]);
 const before=new WorkbookEngine(template.sheets,template.overrides),prerequisites=row43Status(template,before,{ignoreSales:true});
 if(!prerequisites.ready)return {success:false,message:'Complete '+prerequisites.missing.slice(0,4).map(missingInputLabel).join(', ')+(prerequisites.missing.length>4?' and the remaining row 43 inputs.':'.')};
 const initial=Number(before.value('INFORMATION','C14'))||0,start=Math.max(0,Math.ceil(initial/1000));
 const evaluate=units=>{const overrides={...template.overrides,[cellKey('INFORMATION','C14')]:units*1000};const engine=new WorkbookEngine(template.sheets,overrides);return row43Status({...template,overrides},engine);};
 if(evaluate(start).passes)return {success:true,sales:start*1000,unchanged:start*1000===initial};
 const a=evaluate(start),b=evaluate(start+1);
 if(a.values.some((v,i)=>typeof v!=='number'||typeof b.values[i]!=='number'||(Math.round(v*100)<=0&&b.values[i]<=v)))return {success:false,message:'The current GP and expense rules cannot make every month positive by increasing Sales per day. Other inputs were kept unchanged.'};
 let low=start,high=Math.max(start+1,1);while(high<=1e9&&!evaluate(high).passes)high*=2;
 if(high>1e9)return {success:false,message:'A positive result was not found within the sales search range. Inputs were kept unchanged.'};
 while(high-low>1){const mid=Math.floor((high+low)/2);if(evaluate(mid).passes)high=mid;else low=mid;}
 return {success:true,sales:high*1000,unchanged:false};
}
export function activateSource(state,type){
 const fields=type==='location'?['C3','C5','C7','C8','C9','C10','C11','C13']:['C12','C13'];
 fields.forEach(a=>delete state.overrides[cellKey('INFORMATION',a)]);
 invalidateCalculatedOverrides(state,fields.map(a=>cellKey('INFORMATION',a)));
 if(type==='location'){
  const s=state.sheets.location,fill=(a,value)=>{if(!blank(value))state.overrides[cellKey('INFORMATION',a)]=value;};
  fill('C6',s.cells.E5?.v);fill('C20',s.cells.D5?.v);const text=String(s.cells.K5?.v||''),rate=text.match(/([\d.]+)\s*%/),years=text.match(/after\s*(\d+)/i);if(rate)fill('C21',Number(rate[1])/100);if(years)fill('C22',Number(years[1]));
 }
}
export function fillCompanionInputs(state,workbook){
 const find=name=>workbook.Sheets[workbook.SheetNames.find(n=>n.trim().toLowerCase()===name.toLowerCase())];
 const info=find('information'),monthly=find('Monthly PNL'),yearly=find('Yearly -5yrs PNL');
 const fill=(sheet,address,cell)=>{if(cell?.f||blank(cell?.v))return;state.overrides[cellKey(sheet,address)]=cell.v;};
 if(info){for(const a of ['C6','C14','C17','C18','C19','C20','C21','C22'])fill('INFORMATION',a,info[a]);for(let r=6;r<=20;r++)for(const c of ['F','G'])if(c!=='F'||state.manpowerMode!=='automatic')fill('INFORMATION',c+r,info[c+r]);}
 if(monthly)for(const a of EXPENSE_INPUTS)fill(' Monthly PNL',a,monthly[a]||{v:0});
 if(yearly)fill('Yearly -5yrs PNL','B66',yearly.B66);
}
export function migrateState(saved){
 if(!saved?.sheets?.INFORMATION)return makeDefaultState();if(saved.schemaVersion===2){saved.sheets.INFORMATION.cells.F22.f='SUM(F6:F20)';return saved;}
 const next=makeDefaultState();next.sheets=copy(saved.sheets);next.sheets.INFORMATION.cells.F22.f='SUM(F6:F20)';next.categories=copy(saved.categories||[]);next.activeCategory=saved.activeCategory||'Default';next.signatories=copy(saved.signatories||[]);next.signatureUploads=copy(saved.signatureUploads||[]);
 for(const type of ['location','Investment'])if(saved.uploadSources?.[type]&&saved.uploadSources[type]!=='Default'){next.uploadSources[type]=saved.uploadSources[type];activateSource(next,type);}
 Object.assign(next.overrides,copy(saved.overrides||{}));
 if(Object.keys(saved.overrides||{}).length||Object.values(next.uploadSources).some(Boolean)){
  for(const key of ['district','includeExistingOutlets'])if(saved.project?.[key]!==undefined)next.project[key]=saved.project[key];
  const p=saved.project||{};if(p.googleMapLocation!==WORKBOOK_RULES.location.cells.C5.v||next.uploadSources.location)next.project.googleMapLocation=p.googleMapLocation||'';
  if(p.salesGivenBy&&p.salesGivenBy!=='Mr. Mamun')next.project.salesGivenBy=p.salesGivenBy;
  if(p.openedBy&&p.openedBy!=='Mr. Abu Sayem Mohammad Al Azad'){next.project.openedBy=p.openedBy;next.project.openedDesignation=p.openedDesignation||'';}
  if(saved.locationAssessment||p.existingCountMode==='manual'){next.project.existingOutlets=p.existingOutlets;next.locationAssessment=saved.locationAssessment;next.locationStatus=saved.locationStatus;}
 }
 return next;
}
