import { WORKBOOK_RULES, SHEET_ORDER } from './workbook-rules.mjs';

export const copy = value => JSON.parse(JSON.stringify(value));
export const cellKey = (sheet, address) => `${sheet}!${address.replaceAll('$','').toUpperCase()}`;
export const structuralErrors = errors => errors.filter(e=>!['#NUM!','#DIV/0!'].includes(e.code));
export function columnNumber(name) { return [...name.toUpperCase()].reduce((a,c)=>a*26+c.charCodeAt(0)-64,0); }
export function columnName(n) { let s=''; for(;n;n=Math.floor((n-1)/26)) s=String.fromCharCode(65+(n-1)%26)+s; return s; }
export function addressParts(a) { const m=a.replaceAll('$','').match(/^([A-Z]+)(\d+)$/i); if(!m)throw Error('Invalid cell '+a); return [columnNumber(m[1]),Number(m[2])]; }
export class FormulaError extends Error { constructor(code,message) {super(message||code); this.code=code;} }
const numeric=v=>{ if(v===''||v==null)return 0; if(typeof v==='number')return v; if(typeof v==='boolean')return Number(v); const n=Number(v); if(!Number.isFinite(n))throw new FormulaError('#VALUE!'); return n; };
const flatten=a=>a.flat(Infinity);

function tokenize(formula) {
  const tokens=[]; let at=0;
  while(at<formula.length) {
    const rest=formula.slice(at); let m;
    if((m=rest.match(/^\s+/))){at+=m[0].length;continue;}
    if((m=rest.match(/^(?:'((?:[^']|'')+)'!|([A-Za-z_][\w .-]*)!)?(\$?[A-Za-z]{1,3}\$?\d+)(?::(\$?[A-Za-z]{1,3}\$?\d+))?/))) {
      tokens.push({type:'ref',sheet:m[1]?.replaceAll("''", "'")??m[2],address:m[3].replaceAll('$','').toUpperCase(),end:m[4]?.replaceAll('$','').toUpperCase()});at+=m[0].length;continue;
    }
    if((m=rest.match(/^(\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?/))){tokens.push({type:'literal',value:Number(m[0])});at+=m[0].length;continue;}
    if((m=rest.match(/^"((?:[^"]|"")*)"/))){tokens.push({type:'literal',value:m[1].replaceAll('""','"')});at+=m[0].length;continue;}
    if((m=rest.match(/^[A-Za-z_][A-Za-z_0-9.]*/))){tokens.push({type:'name',value:m[0].toUpperCase()});at+=m[0].length;continue;}
    if((m=rest.match(/^(<=|>=|<>|[+\-*/^%(),=&<>])/))){tokens.push({type:m[0]});at+=m[0].length;continue;}
    throw new FormulaError('#NAME?','Unsupported formula near '+rest.slice(0,35));
  }
  return tokens;
}
const precedence={'=':1,'<>':1,'<':1,'>':1,'<=':1,'>=':1,'&':2,'+':3,'-':3,'*':4,'/':4,'^':5};
const parsed=new Map();
function parse(formula) {
  if(parsed.has(formula))return parsed.get(formula);
  const t=tokenize(formula);let i=0;
  const consume=type=>{if(t[i]?.type!==type)throw new FormulaError('#NAME?','Expected '+type+' in '+formula);return t[i++];};
  function expression(min=0) {
    let left; const current=t[i++];
    if(!current)throw new FormulaError('#NAME?','Incomplete formula');
    if(current.type==='+'||current.type==='-')left={type:'unary',op:current.type,child:expression(5)};
    else if(current.type==='('){left=expression();consume(')');}
    else if(current.type==='name') {
      if(current.value==='TRUE'||current.value==='FALSE')left={type:'literal',value:current.value==='TRUE'};
      else {consume('(');const args=[];if(t[i]?.type!==')'){do{args.push(expression());if(t[i]?.type!==',')break;i++;}while(true);}consume(')');left={type:'call',name:current.value,args};}
    } else if(current.type==='ref'||current.type==='literal')left=current;
    else throw new FormulaError('#NAME?','Unexpected '+current.type);
    while(i<t.length){const op=t[i].type;if(op==='%'){i++;left={type:'unary',op:'%',child:left};continue;}const p=precedence[op];if(p===undefined||p<min)break;i++;left={type:'binary',op,left,right:expression(p+(op==='^'?0:1))};}
    return left;
  }
  const result=expression(); if(i!==t.length)throw new FormulaError('#NAME?','Unparsed formula '+formula);parsed.set(formula,result);return result;
}

export function irr(values,guess=.1) {
  const v=flatten(values).map(numeric); if(!v.some(n=>n>0)||!v.some(n=>n<0))throw new FormulaError('#NUM!','IRR needs positive and negative cash flows');
  const npv=r=>v.reduce((s,n,i)=>s+n/(1+r)**i,0);
  let r=guess;
  for(let i=0;i<100;i++){const f=npv(r),d=v.reduce((s,n,k)=>k?s-k*n/(1+r)**(k+1):s,0);const next=r-f/d;if(!Number.isFinite(next)||next<=-1)break;if(Math.abs(next-r)<1e-12)return next;r=next;}
  let lo=-.999999,hi=1;while(npv(lo)*npv(hi)>0&&hi<1e6)hi=hi*2+1;
  if(npv(lo)*npv(hi)>0)throw new FormulaError('#NUM!','IRR does not converge');
  for(let i=0;i<200;i++){r=(lo+hi)/2;if(npv(lo)*npv(r)<=0)hi=r;else lo=r;}return r;
}

export class WorkbookEngine {
  constructor(sheets=WORKBOOK_RULES,overrides={}) {this.sheets=sheets;this.overrides=overrides;this.cache=new Map();this.active=new Set();this.errors=[];}
  get(sheet,address) {
    address=address.replaceAll('$','').toUpperCase(); const key=cellKey(sheet,address);
    if(this.cache.has(key))return this.cache.get(key);
    if(Object.hasOwn(this.overrides,key))return this.overrides[key];
    const cell=this.sheets[sheet]?.cells[address]; if(!this.sheets[sheet])throw new FormulaError('#REF!','Missing sheet '+sheet);if(!cell)return 0;
    if(!cell.f)return cell.v??0;
    if(this.active.has(key))throw new FormulaError('#REF!','Circular reference '+key);
    this.active.add(key);
    try{const value=this.evaluate(parse(cell.f),sheet);if(typeof value==='number'&&!Number.isFinite(value))throw new FormulaError('#NUM!');this.cache.set(key,value);return value;}finally{this.active.delete(key);}
  }
  evaluate(node,sheet) {
    const ev=n=>this.evaluate(n,sheet);
    if(node.type==='literal')return node.value;
    if(node.type==='ref') {
      const target=node.sheet??sheet;
      if(!node.end)return this.get(target,node.address);
      const [a,b]=addressParts(node.address),[c,d]=addressParts(node.end),values=[];
      for(let r=b;r<=d;r++)for(let col=a;col<=c;col++)values.push(this.get(target,columnName(col)+r));return values;
    }
    if(node.type==='unary'){const n=numeric(ev(node.child));return node.op==='-'?-n:node.op==='%'?n/100:n;}
    if(node.type==='binary') {
      const a=ev(node.left),b=ev(node.right);if(node.op==='&')return String(a)+String(b);
      if(node.op==='=')return a===b;if(node.op==='<>')return a!==b;if(node.op==='<')return a<b;if(node.op==='>')return a>b;if(node.op==='<=')return a<=b;if(node.op==='>=')return a>=b;
      const x=numeric(a),y=numeric(b);
      if(node.op==='+')return x+y;if(node.op==='-')return x-y;if(node.op==='*')return x*y;if(node.op==='^')return x**y;if(node.op==='/'){if(y===0)throw new FormulaError('#DIV/0!');return x/y;}
    }
    if(node.type==='call') {
      if(node.name==='IFERROR'){try{return ev(node.args[0]);}catch{return ev(node.args[1]);}}
      if(node.name==='IF')return ev(node.args[0])?ev(node.args[1]):node.args[2]?ev(node.args[2]):false;
      const args=node.args.map(ev),all=flatten(args);
      if(node.name==='SUM')return all.filter(v=>typeof v==='number').reduce((a,b)=>a+b,0);
      if(node.name==='SUMPRODUCT'){const arrays=args.map(a=>Array.isArray(a)?flatten(a):[a]);if(arrays.some(a=>a.length!==arrays[0].length))throw new FormulaError('#VALUE!');return arrays[0].reduce((sum,_,i)=>sum+arrays.reduce((p,a)=>p*numeric(a[i]),1),0);}
      if(node.name==='ABS')return Math.abs(numeric(args[0]));
      if(node.name==='NPV'){const rate=numeric(args[0]);return flatten(args.slice(1)).filter(v=>typeof v==='number').reduce((s,n,i)=>s+n/(1+rate)**(i+1),0);}
      if(node.name==='IRR')return irr(args[0],args[1]??.1);
      if(node.name==='ROUND')return Math.round(numeric(args[0])*10**numeric(args[1]??0))/10**numeric(args[1]??0);
      if(node.name==='MIN')return Math.min(...all.map(numeric));if(node.name==='MAX')return Math.max(...all.map(numeric));
      throw new FormulaError('#NAME?','Unsupported function '+node.name);
    }
    throw new FormulaError('#NAME?');
  }
  value(sheet,address) {try{return this.get(sheet,address);}catch(error){return error.code||'#VALUE!';}}
  scan() {this.errors=[];for(const [name,sheet]of Object.entries(this.sheets))for(const [a,c]of Object.entries(sheet.cells))if(c.f){try{this.get(name,a);}catch(e){this.errors.push({sheet:name,address:a,formula:c.f,error:e.message,code:e.code});}}return this.errors;}
}

export const INFO_FIELDS = [
 ['C3','Project name','text'],['C5','SFT','number'],['C6','Ground floor','text'],['C7','Per SFT rent','number'],['C8','Total rent','number'],['C9','Rent with VAT','number'],['C10','Security','number'],['C11','Advance','number'],['C12','Fixed asset investment','number'],['C13','Total investment','number'],['C14','Sales per day','number'],['C15','Month sales','number'],['C16','GP%','percent'],['C17','Basket size','number'],['C18','Comparing outlet name','text'],['C19','Other income %','percent'],['C20','Business agreement year','number'],['C21','Rent increase %','percent'],['C22','Rent increase after (years later)','number']
];
export function makeDefaultState() {return {schemaVersion:1,sheets:copy(WORKBOOK_RULES),overrides:{},categories:[],activeCategory:'Default',uploadSources:{location:'Default',Investment:'Default'},project:{district:'',googleMapLocation:WORKBOOK_RULES.location.cells.C5.v,existingOutlets:0,includeExistingOutlets:true,salesGivenBy:'Mr. Mamun',openedBy:'Mr. Abu Sayem Mohammad Al Azad',openedDesignation:'Head Of Business Expansion'},signatories:[],signatureUploads:[],locationAssessment:null,locationStatus:'Enter Google Map Location and select District to calculate the 1 KM count.'};}

export function formatCell(value,format='General',digits) {
  if(value==null||value==='')return '';
  if(typeof value!=='number')return String(value);
  const decimals=format.split(';')[0].match(/\.([0#?]+)/)?.[1],isPercent=format.includes('%');
  const max=digits??(decimals?decimals.length:format==='General'?(Number.isInteger(value)?0:2):0),min=digits===undefined?(decimals?.replace(/[^0]/g,'').length||0):0;
  const n=isPercent?value*100:value;
  return n.toLocaleString('en-US',{minimumFractionDigits:Math.min(min,max),maximumFractionDigits:max})+(isPercent?'%':'');
}

export function sheetFromXlsx(sheet,name,XLSX) {
  const result={name,cells:{},rows:0,cols:0,merges:(sheet['!merges']||[]).map(m=>XLSX.utils.encode_range(m)),widths:{},heights:{}};
  for(const [address,c]of Object.entries(sheet)){if(address.startsWith('!')||c.t==='z')continue;if(!/^[A-Z]+\d+$/.test(address))continue;const [col,row]=addressParts(address);result.rows=Math.max(result.rows,row);result.cols=Math.max(result.cols,col);const value=c.v instanceof Date?c.v.toISOString():c.v;result.cells[address]={v:value??null,z:c.z||'General',...(c.f?{f:c.f.replace(/^=/,'')}:{})};}
  (sheet['!cols']||[]).forEach((c,i)=>{if(c?.wch)result.widths[columnName(i+1)]=c.wch;});return result;
}
export function validateSourceSheet(sheet,type) {
  const cell=a=>String(sheet.cells[a]?.v??'').toLowerCase();
  if(type==='location'&&(!/proposed|outlet/.test(cell('B3'))||!/(sft|area)/.test(cell('F3'))||!sheet.cells.B5||!sheet.cells.F5))throw Error('Location upload must keep the location layout: header row 3, first outlet row 5, and total SFT in F9.');
  if(type==='Investment'&&(!/work|type/.test(cell('B4'))||!/(rate)/.test(cell('E4'))||!sheet.cells.G3))throw Error('Investment upload must keep the Investment layout, with the total fixed asset investment in G3.');
}
export function normalizeCategory(source) {
  const rows=[];let header=0;
  for(let r=1;r<=Math.min(source.rows,25);r++)if(/category/.test(String(source.cells['A'+r]?.v??'').toLowerCase())&&/contr|mix/.test(String(source.cells['B'+r]?.v??'').toLowerCase())){header=r;break;}
  if(!header)throw Error('Category upload needs columns A: Category, B: Contr %, C: Daily Sales, D: GP%.');
  const num=v=>{if(typeof v==='number')return v;const s=String(v??'').trim();if(!s)throw Error('Category contribution and GP% cannot be blank.');const n=Number(s.replaceAll(',','').replace('%',''));if(!Number.isFinite(n))throw Error('Invalid category percentage: '+s);return s.includes('%')?n/100:n;};
  for(let r=header+1;r<=source.rows;r++){const label=String(source.cells['A'+r]?.v??'').trim();if(!label)continue;if(/^total\b/i.test(label))break;const mix=num(source.cells['B'+r]?.v),gp=num(source.cells['D'+r]?.v);if(mix<0||mix>1||gp<0||gp>1)throw Error('Store percentage values as Excel percentages, for example 58.5% and 19%.');rows.push({label,mix,gp});}
  if(!rows.length)throw Error('No category rows found.');
  const s={name:'category',cells:{A1:{v:'Category'},B1:{v:'Contr %'},C1:{v:'Daily Sales'},D1:{v:'GP%'}},rows:rows.length+2,cols:4,merges:[],widths:{A:29,B:15,C:20,D:15},heights:{}};
  const total=s.rows;rows.forEach((r,i)=>{const n=i+2;s.cells['A'+n]={v:r.label};s.cells['B'+n]={v:r.mix,z:'0.0%'};s.cells['C'+n]={f:`B${n}*C$${total}`,z:'#,##0.00'};s.cells['D'+n]={v:r.gp,z:'0.0%'};});
  s.cells['A'+total]={v:'Total',bold:true};s.cells['B'+total]={f:`SUM(B2:B${total-1})`,z:'0.0%'};s.cells['C'+total]={f:'INFORMATION!C14',z:'#,##0.00'};s.cells['D'+total]={f:`SUMPRODUCT(D2:D${total-1},B2:B${total-1})`,z:'0.0%'};return s;
}
export function setCategory(state,id) {const version=id==='Default'?WORKBOOK_RULES.category:state.categories.find(c=>c.id===id)?.sheet;if(!version)throw Error('Category version not found');state.sheets.category=copy(version);state.activeCategory=id;state.sheets.INFORMATION.cells.C16.f=`category!D${version.rows}`;delete state.overrides[cellKey('INFORMATION','C16')];for(const k of Object.keys(state.overrides))if(k.startsWith('category!'))delete state.overrides[k];}
export { WORKBOOK_RULES, SHEET_ORDER };
