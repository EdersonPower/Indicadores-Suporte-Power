const MONTHS={jan:1,fev:2,mar:3,abr:4,mai:5,jun:6,jul:7,ago:8,set:9,out:10,nov:11,dez:12};
const MONTH_LABEL={1:'Jan',2:'Fev',3:'Mar',4:'Abr',5:'Mai',6:'Jun',7:'Jul',8:'Ago',9:'Set',10:'Out',11:'Nov',12:'Dez'};
const norm=s=>String(s??'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').trim().toLowerCase();
const num=v=>{const n=Number(v);return Number.isFinite(n)?n:0};

function detectPeriod(fileName, months){
 const year=(String(fileName).match(/20\d{2}/)||[''])[0]||new Date().getFullYear();
 const first=Math.min(...months.map(m=>m.month));
 const q=Math.ceil(first/3);
 return {year:Number(year),quarter:q,label:`${q}º Trimestre ${year}`,id:`${year}-T${q}`};
}
function parseNpsSheet(ws,name){
 const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:null,raw:true});
 if(!rows.length)return null;
 const head=rows[0].map(norm);
 const ix={agent:head.findIndex(x=>x.includes('atendente')),calls:head.findIndex(x=>x.includes('qtd atendimentos')),ratings:head.findIndex(x=>x.includes('qtd avaliados')),rate:head.findIndex(x=>x.includes('percentual')),score:head.findIndex(x=>x.includes('media notas final'))};
 if(Object.values(ix).some(v=>v<0))return null;
 const agents=[];let total=null;
 for(const r of rows.slice(1)){
   const label=String(r[ix.agent]??'').trim(); if(!label)continue;
   const item={agent:label,calls:num(r[ix.calls]),ratings:num(r[ix.ratings]),rate:num(r[ix.rate]),score:num(r[ix.score])};
   if(norm(label)==='total')total=item; else if(/\d{4}\s*-/.test(label))agents.push(item);
 }
 if(!total){const calls=agents.reduce((a,b)=>a+b.calls,0),ratings=agents.reduce((a,b)=>a+b.ratings,0);total={agent:'TOTAL',calls,ratings,rate:calls?ratings/calls:0,score:agents.length?agents.reduce((a,b)=>a+b.score*b.ratings,0)/(ratings||1):0};}
 return {name,agents,total};
}
export function parseWorkbook(buffer,fileName='planilha.xlsx'){
 const wb=XLSX.read(buffer,{type:'array',cellDates:true});
 const monthly=[];let quarterly=null;const diagnostics=[];
 for(const name of wb.SheetNames){
   const n=norm(name); if(!n.startsWith('nps'))continue;
   const parsed=parseNpsSheet(wb.Sheets[name],name);
   if(!parsed){diagnostics.push({name,status:'warn',message:'Aba NPS encontrada, mas cabeçalhos esperados não foram localizados.'});continue;}
   if(n.includes('trim')){quarterly=parsed;diagnostics.push({name,status:'ok',message:`Consolidado válido com ${parsed.agents.length} atendentes.`});continue;}
   const key=Object.keys(MONTHS).find(k=>n.includes(k));
   if(key){monthly.push({...parsed,month:MONTHS[key],monthLabel:MONTH_LABEL[MONTHS[key]]});diagnostics.push({name,status:'ok',message:`Mês reconhecido: ${MONTH_LABEL[MONTHS[key]]}.`});}
 }
 monthly.sort((a,b)=>a.month-b.month);
 if(!monthly.length)throw new Error('Nenhuma aba mensal no padrão NPS Jan, NPS Fev... foi localizada.');
 if(!quarterly){
   const map=new Map(); monthly.forEach(m=>m.agents.forEach(a=>{const x=map.get(a.agent)||{agent:a.agent,calls:0,ratings:0,weighted:0};x.calls+=a.calls;x.ratings+=a.ratings;x.weighted+=a.score*a.ratings;map.set(a.agent,x)}));
   const agents=[...map.values()].map(x=>({...x,rate:x.calls?x.ratings/x.calls:0,score:x.ratings?x.weighted/x.ratings:0}));
   const calls=agents.reduce((s,a)=>s+a.calls,0),ratings=agents.reduce((s,a)=>s+a.ratings,0);
   quarterly={name:'Consolidado calculado',agents,total:{agent:'TOTAL',calls,ratings,rate:calls?ratings/calls:0,score:ratings?agents.reduce((s,a)=>s+a.score*a.ratings,0)/ratings:0}};
   diagnostics.push({name:'NPS Trimestral',status:'warn',message:'Aba trimestral ausente; consolidado calculado a partir dos meses.'});
 }
 const period=detectPeriod(fileName,monthly);
 return {fileName,period,monthly,quarterly,diagnostics,sheetNames:wb.SheetNames};
}
