const MONTHS={jan:1,fev:2,mar:3,abr:4,mai:5,jun:6,jul:7,ago:8,set:9,out:10,nov:11,dez:12};
const MONTH_LABEL={1:'Jan',2:'Fev',3:'Mar',4:'Abr',5:'Mai',6:'Jun',7:'Jul',8:'Ago',9:'Set',10:'Out',11:'Nov',12:'Dez'};
const norm=s=>String(s??'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').trim().toLowerCase();
const num=v=>{if(v===null||v===undefined||v==='')return 0;const n=Number(v);return Number.isFinite(n)?n:0};
const isAgent=s=>/^\d{4}\s*-/.test(String(s??'').trim());

function detectPeriod(fileName,months){
 const year=(String(fileName).match(/20\d{2}/)||[''])[0]||new Date().getFullYear();
 const first=Math.min(...months.map(m=>m.month)); const q=Math.ceil(first/3);
 return {year:Number(year),quarter:q,label:`${q}º Trimestre ${year}`,id:`${year}-T${q}`};
}
function rowsOf(ws){return XLSX.utils.sheet_to_json(ws,{header:1,defval:null,raw:true});}
function parseTopTable(rows,name,isQuarterly=false){
 const header=(rows[0]||[]).map(norm);
 const ix={
  agent:header.findIndex(x=>x.includes('atendente')),
  calls:header.findIndex(x=>x.includes('qtd atendimentos')),
  ratings:header.findIndex(x=>x.includes('qtd avaliados')),
  rate:header.findIndex(x=>x.includes('percentual')),
  avg:header.findIndex(x=>x==='media notas'),
  discount:header.findIndex(x=>x.includes('descontos')),
  final:header.findIndex(x=>x.includes('media notas final')),
  box:header.findIndex(x=>x.includes('pontuacao caixinha')),
  bonus:header.findIndex(x=>x.includes('bonificacao'))
 };
 const required=['agent','calls','ratings','rate','avg','discount','final'];
 if(required.some(k=>ix[k]<0))return null;
 const agents=[];let total=null;
 for(let r=1;r<Math.min(rows.length,14);r++){
  const label=String(rows[r]?.[ix.agent]??'').trim(); if(!label)continue;
  const item={agent:label,calls:num(rows[r][ix.calls]),ratings:num(rows[r][ix.ratings]),rate:num(rows[r][ix.rate]),average:num(rows[r][ix.avg]),discount:num(rows[r][ix.discount]),score:num(rows[r][ix.final]),box:ix.box>=0?num(rows[r][ix.box]):0,bonus:ix.bonus>=0?num(rows[r][ix.bonus]):0,row:r+1};
  item.finalCheck=item.average-item.discount;
  if(norm(label)==='total')total=item; else if(isAgent(label))agents.push(item);
 }
 if(!total){
  const calls=agents.reduce((s,a)=>s+a.calls,0),ratings=agents.reduce((s,a)=>s+a.ratings,0),bonus=agents.reduce((s,a)=>s+a.bonus,0);
  total={agent:'TOTAL',calls,ratings,rate:calls?ratings/calls:0,average:agents.length?agents.reduce((s,a)=>s+a.average,0)/agents.length:0,discount:agents.reduce((s,a)=>s+a.discount,0),score:agents.length?agents.reduce((s,a)=>s+a.score,0)/agents.length:0,box:0,bonus};
 }
 return {name,agents,total,isQuarterly};
}
function parseBaseTable(rows){
 const headerRow=rows.findIndex((r,i)=>i>=10&&norm(r?.[0])==='atendente'&&norm(r?.[1])==='qsa');
 if(headerRow<0)return {rows:[],headerRow:-1};
 const header=(rows[headerRow]||[]).map(norm);
 const ix={agent:0,qsa:header.findIndex(x=>x==='qsa'),qea:header.findIndex(x=>x==='qea'),ts:header.findIndex(x=>x.includes('ts fone')),te:header.findIndex(x=>x.includes('te fone')),whats:header.findIndex(x=>x.includes('te whats')),whatsRated:header.findIndex(x=>x.includes('whats avaliados'))};
 const out=[];
 for(let r=headerRow+1;r<rows.length;r++){
  const label=String(rows[r]?.[0]??'').trim(); if(!isAgent(label))continue;
  const item={agent:label,qsa:num(rows[r][ix.qsa]),qea:num(rows[r][ix.qea]),tsPhone:num(rows[r][ix.ts]),tePhone:num(rows[r][ix.te]),whats:num(rows[r][ix.whats]),whatsRated:num(rows[r][ix.whatsRated]),row:r+1};
  item.callsCalculated=item.tsPhone+item.tePhone+item.whats;
  item.ratingsCalculated=item.qsa+item.qea;
  out.push(item);
 }
 return {rows:out,headerRow:headerRow+1};
}
function diagnosticsFor(parsed,base){
 const items=[];
 parsed.agents.forEach(a=>{
  const b=base.rows.find(x=>x.agent===a.agent);
  if(b){
   if(Math.abs(a.calls-b.callsCalculated)>.001)items.push({status:'warn',message:`${a.agent}: atendimentos A1=${a.calls} e base=${b.callsCalculated}.`});
   if(Math.abs(a.ratings-b.ratingsCalculated)>.001)items.push({status:'warn',message:`${a.agent}: avaliações A1=${a.ratings} e base=${b.ratingsCalculated}.`});
  }
  if(Math.abs(a.score-a.finalCheck)>.005)items.push({status:'warn',message:`${a.agent}: média final ${a.score.toFixed(2)} difere de média − desconto ${a.finalCheck.toFixed(2)}.`});
 });
 if(!items.length)items.push({status:'ok',message:'Tabela superior, dados-base e média final estão consistentes.'});
 return items;
}
export function parseWorkbook(buffer,fileName='planilha.xlsx'){
 const wb=XLSX.read(buffer,{type:'array',cellDates:true});
 const monthly=[];let quarterly=null;const diagnostics=[];
 for(const name of wb.SheetNames){
  const n=norm(name); if(!n.startsWith('nps'))continue;
  const rows=rowsOf(wb.Sheets[name]); const isQuarter=n.includes('trim');
  const parsed=parseTopTable(rows,name,isQuarter); if(!parsed){diagnostics.push({name,status:'warn',message:'Cabeçalhos esperados entre A1:J1 não foram encontrados.'});continue;}
  const base=parseBaseTable(rows); const checks=diagnosticsFor(parsed,base);
  parsed.base=base; parsed.checks=checks;
  if(isQuarter){quarterly=parsed;diagnostics.push({name,status:checks.some(x=>x.status==='warn')?'warn':'ok',message:`Consolidado com ${parsed.agents.length} atendentes e bonificação total de ${parsed.total.bonus.toFixed(2)}.`});continue;}
  const key=Object.keys(MONTHS).find(k=>n.includes(k));
  if(key){monthly.push({...parsed,month:MONTHS[key],monthLabel:MONTH_LABEL[MONTHS[key]]});diagnostics.push({name,status:checks.some(x=>x.status==='warn')?'warn':'ok',message:`${MONTH_LABEL[MONTHS[key]]}: ${parsed.agents.length} atendentes, caixinha lida da coluna I.`});}
 }
 monthly.sort((a,b)=>a.month-b.month);
 if(!monthly.length)throw new Error('Nenhuma aba mensal no padrão NPS Jan, NPS Fev... foi localizada.');
 if(!quarterly)throw new Error('A aba NPS Trim ou NPS Trimestral não foi localizada.');
 const period=detectPeriod(fileName,monthly);
 return {fileName,period,monthly,quarterly,diagnostics,sheetNames:wb.SheetNames};
}
