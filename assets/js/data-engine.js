const monthOrder={Jan:1,Fev:2,Mar:3,Abr:4,Mai:5,Jun:6,Jul:7,Ago:8,Set:9,Out:10,Nov:11,Dez:12};
export const cleanName=v=>String(v??'').replace(/^\s*\d+\s*-\s*/,'').trim();
const num=v=>Number.isFinite(Number(v))?Number(v):0;
const monthToken=name=>Object.keys(monthOrder).find(m=>new RegExp(`\\b${m}\\b`,'i').test(name));
function rowsOf(wb,sheet){return XLSX.utils.sheet_to_json(wb.Sheets[sheet],{header:1,defval:null,raw:true});}
function parseTop(rows){const out=[];for(let r=1;r<9;r++){const row=rows[r]||[];const name=cleanName(row[0]);if(!name||/^TOTAL$/i.test(name)) continue;if(row[1]==null&&row[2]==null) continue;out.push({name,att:num(row[1]),rated:num(row[2]),rate:num(row[3])*100,avg:num(row[4]),discount:num(row[5]),final:num(row[6]),box:num(row[8]),bonus:num(row[9])});}return out;}
function parseBase(rows){const out=[];for(let r=14;r<20;r++){const row=rows[r]||[];const name=cleanName(row[0]);if(!name||/^QSA/i.test(name))continue;if(row[1]==null&&row[2]==null&&row[3]==null)continue;out.push({name,qsa:num(row[1]),qea:num(row[2]),ts:num(row[3]),te:num(row[4]),whats:num(row[5]),whatsRated:num(row[6])});}return out;}
function mergeAudit(top,base){const map=new Map(base.map(x=>[x.name,x]));return top.map(x=>{const b=map.get(x.name);if(!b)return {name:x.name,status:'warn',detail:'Sem linha correspondente na tabela-base A14'};const calcAtt=b.ts+b.te+b.whats;const calcRated=b.qsa+b.qea;const issues=[];if(Math.abs(calcAtt-x.att)>0.01)issues.push(`Atendimentos: tabela ${x.att}, base ${calcAtt}`);if(Math.abs(calcRated-x.rated)>0.01)issues.push(`Avaliações: tabela ${x.rated}, base ${calcRated}`);if(Math.abs((x.avg-x.discount)-x.final)>0.01)issues.push('Média final divergente de média − desconto');return {name:x.name,status:issues.length?'warn':'ok',detail:issues.length?issues.join(' · '):'Cálculos conferidos'};});}
export function parseWorkbook(buffer,fileName,labelHint=''){
 const wb=XLSX.read(buffer,{type:'array',cellFormula:true,cellDates:true});
 const npsSheets=wb.SheetNames.filter(n=>/^nps\s/i.test(n));
 const monthly=npsSheets.filter(n=>!/(trim|trimestral)/i.test(n)).map(sheet=>{const rows=rowsOf(wb,sheet),token=monthToken(sheet)||sheet.replace(/^nps\s*/i,'').trim();return {sheet,month:token,monthIndex:monthOrder[token]||99,people:parseTop(rows),base:parseBase(rows),audit:mergeAudit(parseTop(rows),parseBase(rows))};}).sort((a,b)=>a.monthIndex-b.monthIndex);
 const trimSheet=npsSheets.find(n=>/(trim|trimestral)/i.test(n));
 const trimRows=trimSheet?rowsOf(wb,trimSheet):[];
 const trimPeople=trimSheet?parseTop(trimRows):[];
 const totalRow=trimRows.find(r=>String(r?.[0]??'').trim().toUpperCase()==='TOTAL')||[];
 const inferred=fileName.match(/([1-4])\s*trim.*?(20\d{2})/i);
 const quarter=inferred?Number(inferred[1]):Math.ceil(((monthly[0]?.monthIndex)||1)/3);
 const year=inferred?Number(inferred[2]):Number((labelHint.match(/20\d{2}/)||['2026'])[0]);
 return {id:`${year}-Q${quarter}-${fileName}`,fileName,label:labelHint||`${quarter}º Trimestre ${year}`,year,quarter,months:monthly,trim:{sheet:trimSheet,people:trimPeople,total:{att:num(totalRow[1]),rated:num(totalRow[2]),rate:num(totalRow[3])*100,avg:num(totalRow[4]),discount:num(totalRow[5]),final:num(totalRow[6]),bonus:num(totalRow[9])}},sheets:wb.SheetNames};
}
export function aggregatePeriod(p){return p.trim?.total?.att?{...p.trim.total,people:p.trim.people}:aggregateMonths(p.months);}
export function aggregateMonths(months){const map=new Map();for(const m of months)for(const p of m.people){const x=map.get(p.name)||{name:p.name,att:0,rated:0,avgWeighted:0,finalWeighted:0,discountWeighted:0,box:0};x.att+=p.att;x.rated+=p.rated;x.avgWeighted+=p.avg*p.rated;x.finalWeighted+=p.final*p.rated;x.discountWeighted+=p.discount*p.rated;x.box+=p.box;map.set(p.name,x)}const people=[...map.values()].map(x=>({...x,rate:x.att?x.rated/x.att*100:0,avg:x.rated?x.avgWeighted/x.rated:0,final:x.rated?x.finalWeighted/x.rated:0,discount:x.rated?x.discountWeighted/x.rated:0}));const att=people.reduce((s,x)=>s+x.att,0),rated=people.reduce((s,x)=>s+x.rated,0);return {att,rated,rate:att?rated/att*100:0,final:rated?people.reduce((s,x)=>s+x.final*x.rated,0)/rated:0,avg:rated?people.reduce((s,x)=>s+x.avg*x.rated,0)/rated:0,discount:rated?people.reduce((s,x)=>s+x.discount*x.rated,0)/rated:0,people};}
