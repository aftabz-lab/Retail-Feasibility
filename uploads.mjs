import {WorkbookEngine,copy,cellKey,sheetFromXlsx,validateSourceSheet,normalizeCategory,setCategory,structuralErrors} from './calculation.mjs?v=retail-v2-20261015';
import {activateSource,fillCompanionInputs,invalidateCalculatedOverrides} from './controls.mjs?v=retail-v2-20261015';
export function importWorkbook(workbook,type,fileName,current) {
  const names=workbook.SheetNames.filter(n=>n.trim().toLowerCase()!=='sheet5');
  const match=names.find(n=>n.trim().toLowerCase()===type.toLowerCase());
  const candidates=match?[match]:names;let sheet,lastError,selectedName;
  for(const name of candidates){try{const s=sheetFromXlsx(workbook.Sheets[name],type,globalThis.XLSX);if(type==='category'){sheet=normalizeCategory(s);}else{validateSourceSheet(s,type);sheet=s;}selectedName=name;break;}catch(error){lastError=error;}}
  if(!sheet)throw lastError||Error('No compatible '+type+' sheet found.');
  // Keep canonical cross-sheet names even when the uploaded tab has another name.
  if(type!=='category')for(const cell of Object.values(sheet.cells)){if(!cell.f)continue;const name=selectedName,escaped=name.replace(/[.*+?^${}()|[\]\\]/g,'\\$&');cell.f=cell.f.replace(new RegExp("'"+escaped.replaceAll("'","''")+"'!",'g'),"'"+type+"'!");if(/^[A-Za-z_]\w*$/.test(name))cell.f=cell.f.replace(new RegExp('\\b'+escaped+'!','g'),type+'!');}
  const next=copy(current);
  if(type==='category'){
    const n=next.categories.reduce((max,c)=>Math.max(max,Number(c.id.match(/\d+$/)?.[0])||0),0)+1;
    const id='Upload '+n;next.categories.push({id,fileName,uploadedAt:new Date().toISOString(),sheet});setCategory(next,id);invalidateCalculatedOverrides(next,[cellKey('INFORMATION','C16')]);
  }else{
    next.sheets[type]=sheet;next.uploadSources[type]=fileName;
    for(const k of Object.keys(next.overrides))if(k.startsWith(type+'!'))delete next.overrides[k];
    const reset=type==='location'?['C3','C5','C7','C8','C9','C10','C11','C13']:['C12','C13'];
    reset.forEach(a=>delete next.overrides[cellKey('INFORMATION',a)]);
    if(type==='location'){
      next.project.googleMapLocation=String(sheet.cells.C5?.v||'');next.project.existingOutlets=null;next.locationAssessment=null;next.locationStatus='Location file updated. Select District and recheck the map.';
    }
    activateSource(next,type);fillCompanionInputs(next,workbook);
  }
  const test=new WorkbookEngine(next.sheets,next.overrides),errors=structuralErrors(test.scan());if(errors.length)throw Error('Upload has a formula that cannot calculate: '+errors[0].sheet+'!'+errors[0].address+' — '+errors[0].error);
  return next;
}


