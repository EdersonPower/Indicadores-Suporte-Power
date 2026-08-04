(function(){
'use strict';
const monthOrder={Jan:1,Fev:2,Mar:3,Abr:4,Mai:5,Jun:6,Jul:7,Ago:8,Set:9,Out:10,Nov:11,Dez:12};
const state={periods:[],period:null,month:null,charts:{}};
const $=id=>document.getElementById(id);
const fmt={int:v=>new Intl.NumberFormat('pt-BR',{maximumFractionDigits:0}).format(v||0),dec:(v,d=2)=>new Intl.NumberFormat('pt-BR',{minimumFractionDigits:d,maximumFractionDigits:d}).format(v||0),pct:v=>`${new Intl.NumberFormat('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2}).format(v||0)}%`,money:v=>new Intl.NumberFormat('pt-BR',{style:'currency',currency:'BRL'}).format(v||0)};
const titles={dashboard:['Dashboard Executivo','Visão consolidada da operação'],team:['Equipe','Desempenho individual'],hall:['Hall da Fama','Destaques e reconhecimentos'],box:['Caixinha','Ranking mensal da dinâmica'],bonus:['Bonificação','Acompanhamento financeiro'],voice:['Voz do Cliente','Comentários e percepção dos clientes'],history:['Evolução','Comparativos históricos'],audit:['Auditoria','Conferência das regras de negócio']};
function slug(s){return String(s||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/[^a-z0-9]+/g,'-').replace(/(^-|-$)/g,'')}
function avatar(name,size=''){const initial=(name||'?').trim().charAt(0).toUpperCase();const file=`assets/img/team/${slug(name)}.jpg`;return `<div class="avatar ${size}"><img src="${file}" alt="${name}" onerror="this.style.display='none';this.parentElement.dataset.fallback='${initial}'"><span class="avatar-fallback">${initial}</span></div>`}
function metric(label,value,delta=''){return `<div class="metric"><div class="label">${label}</div><div class="value">${value}</div><div class="delta">${delta||'Período selecionado'}</div></div>`}
function trend(current,previous,invert=false){if(!previous)return 'Sem comparação anterior';const d=(current-previous)/Math.abs(previous)*100;const good=invert?d<=0:d>=0;return `<span style="color:${good?'var(--green)':'var(--red)'}">${d>=0?'▲':'▼'} ${Math.abs(d).toFixed(1)}% vs anterior</span>`}
function showStatus(message,type='error'){const el=$('systemStatus');if(!el)return;el.hidden=false;el.className=`system-status ${type}`;el.innerHTML=message}
function clearStatus(){const el=$('systemStatus');if(el)el.hidden=true}
function num(v){return Number.isFinite(Number(v))?Number(v):0}
function cleanName(v){return String(v??'').replace(/^\s*\d+\s*-\s*/,'').trim()}
function monthToken(name){return Object.keys(monthOrder).find(m=>new RegExp(`\\b${m}\\b`,'i').test(name))}
function rowsOf(wb,sheet){return XLSX.utils.sheet_to_json(wb.Sheets[sheet],{header:1,defval:null,raw:true})}


function parseReviews(wb){
  const sheets=wb.SheetNames.filter(n=>/^avalia[cç][oõ]es\s/i.test(n));
  const reviews=[];
  for(const sheet of sheets){
    const token=monthToken(sheet)||sheet.replace(/^avalia[cç][oõ]es\s*/i,'').trim();
    const rows=rowsOf(wb,sheet);
    for(let r=1;r<rows.length;r++){
      const row=rows[r]||[];
      const person=cleanName(row[0]);
      const comment=String(row[1]??'').trim();
      if(person&&comment)reviews.push({month:token,monthIndex:monthOrder[token]||99,sheet,person,comment});
    }
  }
  return reviews.sort((a,b)=>a.monthIndex-b.monthIndex);
}

function parseTop(rows){const out=[];for(let r=1;r<9;r++){const row=rows[r]||[];const name=cleanName(row[0]);if(!name||/^TOTAL$/i.test(name))continue;if(row[1]==null&&row[2]==null)continue;out.push({name,att:num(row[1]),rated:num(row[2]),rate:num(row[3])*100,avg:num(row[4]),discount:num(row[5]),final:num(row[6]),box:num(row[8]),bonus:num(row[9])})}return out}
function parseBase(rows){const out=[];for(let r=14;r<30;r++){const row=rows[r]||[];const name=cleanName(row[0]);if(!name||/^ATENDENTE$/i.test(name)||/^TOTAL$/i.test(name))continue;if(row[1]==null&&row[2]==null&&row[3]==null)continue;out.push({name,qsa:num(row[1]),qea:num(row[2]),ts:num(row[3]),te:num(row[4]),whats:num(row[5]),whatsRated:num(row[6])})}return out}
function mergeAudit(top,base){const map=new Map(base.map(x=>[x.name,x]));return top.map(x=>{const b=map.get(x.name);if(!b)return{name:x.name,status:'warn',detail:'Sem linha correspondente na tabela-base A14'};const calcAtt=b.ts+b.te+b.whats,calcRated=b.qsa+b.qea,issues=[];if(Math.abs(calcAtt-x.att)>.01)issues.push(`Atendimentos: tabela ${x.att}, base ${calcAtt}`);if(Math.abs(calcRated-x.rated)>.01)issues.push(`Avaliações: tabela ${x.rated}, base ${calcRated}`);if(Math.abs((x.avg-x.discount)-x.final)>.01)issues.push('Média final divergente de média − desconto');return{name:x.name,status:issues.length?'warn':'ok',detail:issues.length?issues.join(' · '):'Cálculos conferidos'}})}
function parseWorkbook(buffer,fileName,labelHint=''){if(!window.XLSX)throw new Error('Biblioteca de leitura do Excel indisponível.');const wb=XLSX.read(buffer,{type:'array',cellFormula:true,cellDates:true});const npsSheets=wb.SheetNames.filter(n=>/^nps\s/i.test(n));const monthly=npsSheets.filter(n=>!/(trim|trimestral)/i.test(n)).map(sheet=>{const rows=rowsOf(wb,sheet),token=monthToken(sheet)||sheet.replace(/^nps\s*/i,'').trim(),top=parseTop(rows),base=parseBase(rows);return{sheet,month:token,monthIndex:monthOrder[token]||99,people:top,base,audit:mergeAudit(top,base)}}).sort((a,b)=>a.monthIndex-b.monthIndex);const reviews=parseReviews(wb);const trimSheet=npsSheets.find(n=>/(trim|trimestral)/i.test(n)),trimRows=trimSheet?rowsOf(wb,trimSheet):[],trimPeople=trimSheet?parseTop(trimRows):[],totalRow=trimRows.find(r=>String(r?.[0]??'').trim().toUpperCase()==='TOTAL')||[],inferred=fileName.match(/([1-4])\s*trim.*?(20\d{2})/i),quarter=inferred?Number(inferred[1]):Math.ceil(((monthly[0]?.monthIndex)||1)/3),year=inferred?Number(inferred[2]):Number((labelHint.match(/20\d{2}/)||['2026'])[0]);return{id:`${year}-Q${quarter}-${fileName}`,fileName,label:labelHint||`${quarter}º Trimestre ${year}`,year,quarter,months:monthly,reviews,trim:{sheet:trimSheet,people:trimPeople,total:{att:num(totalRow[1]),rated:num(totalRow[2]),rate:num(totalRow[3])*100,avg:num(totalRow[4]),discount:num(totalRow[5]),final:num(totalRow[6]),bonus:num(totalRow[9])}},sheets:wb.SheetNames}}
function aggregateMonths(months){const map=new Map();for(const m of months)for(const p of m.people){const x=map.get(p.name)||{name:p.name,att:0,rated:0,avgWeighted:0,finalWeighted:0,discountWeighted:0,box:0};x.att+=p.att;x.rated+=p.rated;x.avgWeighted+=p.avg*p.rated;x.finalWeighted+=p.final*p.rated;x.discountWeighted+=p.discount*p.rated;x.box+=p.box;map.set(p.name,x)}const people=[...map.values()].map(x=>({...x,rate:x.att?x.rated/x.att*100:0,avg:x.rated?x.avgWeighted/x.rated:0,final:x.rated?x.finalWeighted/x.rated:0,discount:x.rated?x.discountWeighted/x.rated:0}));const att=people.reduce((s,x)=>s+x.att,0),rated=people.reduce((s,x)=>s+x.rated,0);return{att,rated,rate:att?rated/att*100:0,final:rated?people.reduce((s,x)=>s+x.final*x.rated,0)/rated:0,avg:rated?people.reduce((s,x)=>s+x.avg*x.rated,0)/rated:0,discount:rated?people.reduce((s,x)=>s+x.discount*x.rated,0)/rated:0,people}}
function aggregatePeriod(p){return p.trim?.total?.att?{...p.trim.total,people:p.trim.people}:aggregateMonths(p.months)}
function destroyChart(key){if(state.charts[key]&&typeof state.charts[key].destroy==='function')state.charts[key].destroy();state.charts[key]=null}
function chart(key,el,type,data,options={}){destroyChart(key);if(!el||!window.Chart){if(!window.Chart)showStatus('<b>Os dados foram carregados, mas a biblioteca de gráficos não respondeu.</b> Atualize a página ou verifique o bloqueio de scripts externos.','warn');return}state.charts[key]=new Chart(el,{type,data,options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{labels:{color:'#b8c5d9'}},...(options.plugins||{})},scales:type==='doughnut'?undefined:{x:{ticks:{color:'#9aabc2'},grid:{color:'#20314e55'}},y:{ticks:{color:'#9aabc2'},grid:{color:'#20314e55'},beginAtZero:true},...(options.scales||{})},...options}})}
async function loadDefault(){clearStatus();try{const r=await fetch('data/dashboard-data.json?v=3.7.0',{cache:'no-store'});if(!r.ok)throw new Error(`Dados HTTP ${r.status}`);state.periods=await r.json();if(!Array.isArray(state.periods)||!state.periods.length)throw new Error('Arquivo de dados vazio');setupPeriods()}catch(err){showStatus(`<b>Não foi possível carregar os dados consolidados.</b> Use “Importar Excel”. Detalhe: ${err.message}`)}}
function setupPeriods(){const ps=$('periodSelect');if(!state.periods.length)return;ps.innerHTML=state.periods.map((p,i)=>`<option value="${i}">${p.label}</option>`).join('');ps.value=String(Math.max(0,state.periods.length-1));selectPeriod(Number(ps.value))}
function selectPeriod(index){state.period=state.periods[index];if(!state.period)return;const ms=$('monthSelect');ms.innerHTML='<option value="all">Trimestre completo</option>'+state.period.months.map((m,i)=>`<option value="${i}">${m.month}</option>`).join('');ms.value='all';state.month=null;renderAll()}
function currentData(){return state.month==null?aggregatePeriod(state.period):aggregateMonths([state.period.months[state.month]])}
function previousData(){const idx=state.periods.indexOf(state.period);if(state.month!=null){const m=state.month-1;return m>=0?aggregateMonths([state.period.months[m]]):null}return idx>0?aggregatePeriod(state.periods[idx-1]):null}
function renderKpis(){const d=currentData(),p=previousData();$('kpis').innerHTML=[metric('Total de atendimentos',fmt.int(d.att),p?trend(d.att,p.att):''),metric('Avaliações recebidas',fmt.int(d.rated),p?trend(d.rated,p.rated):''),metric('Taxa de avaliação',fmt.pct(d.rate),p?trend(d.rate,p.rate):''),metric('Média final',fmt.dec(d.final),p?trend(d.final,p.final):''),metric('Bonificação total',fmt.money(state.period.trim.total.bonus),'Consolidado do trimestre')].join('')}
function renderInsights(){const d=currentData(),p=previousData(),people=[...(d.people||[])],topAtt=[...people].sort((a,b)=>b.att-a.att)[0],topFinal=[...people].sort((a,b)=>b.final-a.final)[0],msgs=[];if(p)msgs.push(`<b>${d.att>=p.att?'Crescimento':'Variação'} de volume:</b> ${trend(d.att,p.att)}`);if(topAtt)msgs.push(`<b>${topAtt.name}</b> lidera em atendimentos com ${fmt.int(topAtt.att)}.`);if(topFinal)msgs.push(`<b>${topFinal.name}</b> possui a maior média final: ${fmt.dec(topFinal.final)}.`);$('insights').innerHTML=msgs.map(x=>`<div class="insight">${x}</div>`).join('')}
function renderDashboardCharts(){const months=state.period.months;chart('monthly',$('monthlyChart'),'line',{labels:months.map(x=>x.month),datasets:[{label:'Atendimentos',data:months.map(x=>x.people.reduce((s,p)=>s+p.att,0)),borderColor:'#8b5cf6',backgroundColor:'#8b5cf633',tension:.35,fill:true},{label:'Avaliações',data:months.map(x=>x.people.reduce((s,p)=>s+p.rated,0)),borderColor:'#22d3ee',tension:.35}]});const d=currentData(),rank=[...(d.people||[])].sort((a,b)=>b.att-a.att);chart('ranking',$('rankingChart'),'bar',{labels:rank.map(x=>x.name),datasets:[{label:'Atendimentos',data:rank.map(x=>x.att),backgroundColor:'#8b5cf6'}]},{indexAxis:'y',plugins:{legend:{display:false}}})}
function renderTeam(){const d=currentData(),people=[...(d.people||[])].sort((a,b)=>b.final-a.final);$('personSelect').innerHTML=people.map(x=>`<option>${x.name}</option>`).join('');$('teamGrid').innerHTML=people.map(x=>`<div class="card person-card" data-person="${x.name}"><div class="person-head">${avatar(x.name)}<div><div class="person-name">${x.name}</div><div class="muted">Média final ${fmt.dec(x.final)}</div></div></div><div class="mini-stats"><div class="mini"><span class="muted">Atendimentos</span><b>${fmt.int(x.att)}</b></div><div class="mini"><span class="muted">Avaliação</span><b>${fmt.pct(x.rate)}</b></div><div class="mini"><span class="muted">Avaliações</span><b>${fmt.int(x.rated)}</b></div><div class="mini"><span class="muted">Caixinha</span><b>${fmt.dec(x.box||0,1)}</b></div></div></div>`).join('');renderPerson(people[0]?.name);document.querySelectorAll('[data-person]').forEach(el=>el.addEventListener('click',()=>{$('personSelect').value=el.dataset.person;renderPerson(el.dataset.person)}))}
function renderPerson(name){if(!name)return;const labels=state.period.months.map(x=>x.month),vals=state.period.months.map(m=>m.people.find(p=>p.name===name)?.att??null),finals=state.period.months.map(m=>m.people.find(p=>p.name===name)?.final??null);$('profileTitle').textContent=`Evolução de ${name}`;chart('person',$('personChart'),'line',{labels,datasets:[{label:'Atendimentos',data:vals,borderColor:'#8b5cf6',tension:.35,yAxisID:'y'},{label:'Média final',data:finals,borderColor:'#34d399',tension:.35,yAxisID:'y1'}]},{scales:{y1:{position:'right',min:0,max:5,ticks:{color:'#9aabc2'},grid:{drawOnChartArea:false}}}})}
function renderHall(){const d=currentData(),people=d.people||[],defs=[['📞','Maior atendimento',[...people].sort((a,b)=>b.att-a.att)[0],x=>fmt.int(x.att)],['⭐','Melhor nota',[...people].sort((a,b)=>b.final-a.final)[0],x=>fmt.dec(x.final)],['📈','Melhor taxa',[...people].sort((a,b)=>b.rate-a.rate)[0],x=>fmt.pct(x.rate)],['🎁','Campeão da caixinha',[...people].sort((a,b)=>(b.box||0)-(a.box||0))[0],x=>fmt.dec(x.box||0,1)+' pts']];$('hallGrid').innerHTML=defs.map(([ico,title,p,fn])=>p?`<div class="card hall"><div class="trophy">${ico}</div><div class="muted">${title}</div>${avatar(p.name)}<div class="winner">${p.name}</div><div class="score">${fn(p)}</div></div>`:'').join('')}
function renderBox(){const m=state.month==null?state.period.months[state.period.months.length-1]:state.period.months[state.month],rank=[...(m?.people||[])].sort((a,b)=>b.box-a.box),order=[rank[1],rank[0],rank[2]],classes=['two','one','three'],medals=['🥈','🥇','🥉'];$('podium').innerHTML=order.map((p,i)=>p?`<div class="podium-item">${avatar(p.name)}<b>${p.name}</b><div class="muted">${fmt.dec(p.box,1)} pontos</div><div class="step ${classes[i]}"><div style="font-size:28px">${medals[i]}</div><b>${i===1?'1º':i===0?'2º':'3º'}</b></div></div>`:'').join('');$('boxRanking').innerHTML=rank.map((p,i)=>`<div class="rank-row"><div class="rank-pos">${i+1}º</div><div><b>${p.name}</b><div class="muted">${m.month}</div></div><b>${fmt.dec(p.box,1)}</b></div>`).join('')}
function renderBonus(){const p=state.period,people=[...(p.trim.people||[])].sort((a,b)=>b.bonus-a.bonus);$('bonusKpis').innerHTML=[metric('Total distribuído',fmt.money(p.trim.total.bonus)),metric('Média por colaborador',fmt.money(people.length?p.trim.total.bonus/people.length:0)),metric('Maior bonificação',fmt.money(people[0]?.bonus||0)),metric('Colaboradores',fmt.int(people.filter(x=>x.bonus>0).length)),metric('Período',p.label)].join('');chart('bonus',$('bonusChart'),'bar',{labels:people.map(x=>x.name),datasets:[{label:'Bonificação',data:people.map(x=>x.bonus),backgroundColor:'#34d399'}]},{indexAxis:'y',plugins:{legend:{display:false}}});chart('bonusHistory',$('bonusHistoryChart'),'line',{labels:state.periods.map(x=>x.label),datasets:[{label:'Total distribuído',data:state.periods.map(x=>x.trim.total.bonus),borderColor:'#fbbf24',backgroundColor:'#fbbf2433',fill:true,tension:.3}]})}


function selectedReviews(){
  let reviews=[...(state.period?.reviews||[])];
  if(state.month!=null){
    const month=state.period.months[state.month]?.month;
    reviews=reviews.filter(x=>x.month===month);
  }
  const person=$('voicePerson')?.value||'all';
  if(person!=='all')reviews=reviews.filter(x=>x.person===person);
  return reviews;
}
function renderVoiceFilters(){
  const all=[...(state.period?.reviews||[])];
  const people=[...new Set(all.map(x=>x.person))].sort();
  const current=$('voicePerson')?.value||'all';
  $('voicePerson').innerHTML='<option value="all">Todos</option>'+people.map(x=>`<option value="${x}">${x}</option>`).join('');
  if(people.includes(current))$('voicePerson').value=current;
}
function renderVoice(){
  renderVoiceFilters();
  const all=[...(state.period?.reviews||[])];
  const reviews=selectedReviews();
  const counts=reviews.reduce((m,x)=>(m[x.person]=(m[x.person]||0)+1,m),{});
  const top=Object.entries(counts).sort((a,b)=>b[1]-a[1])[0];
  $('voiceKpis').innerHTML=[
    metric('Avaliações descritas',fmt.int(reviews.length)),
    metric('Colaboradores citados',fmt.int(Object.keys(counts).length)),
    metric('Mais elogiado',top?top[0]:'—')
  ].join('');
  $('voiceEmpty').style.display=all.length?'none':'block';
  $('voiceGrid').innerHTML=reviews.map(x=>`
    <article class="review-card simple-review">
      <div class="review-head">
        <div class="review-person">${avatar(x.person)}<div><b>${x.person}</b><span>${x.month}</span></div></div>
        <span class="review-type">💬 Avaliação</span>
      </div>
      <blockquote>“${String(x.comment).replace(/</g,'&lt;').replace(/>/g,'&gt;')}”</blockquote>
    </article>`).join('');
}


function evolutionMetricConfig(metric){
  const configs={
    att:{label:'Atendimentos',format:v=>fmt.int(v),axis:'Quantidade',color:'#8b5cf6'},
    rated:{label:'Avaliações',format:v=>fmt.int(v),axis:'Quantidade',color:'#22d3ee'},
    rate:{label:'Taxa de avaliação',format:v=>fmt.pct(v),axis:'Percentual',color:'#34d399'},
    final:{label:'Média final',format:v=>fmt.dec(v,2),axis:'Nota',color:'#f472b6'},
    box:{label:'Pontuação da Caixinha',format:v=>fmt.dec(v,2),axis:'Pontos',color:'#fbbf24'}
  };
  return configs[metric]||configs.att;
}
function allEvolutionPeople(){
  const names=new Set();
  for(const period of state.periods){
    for(const month of period.months||[])for(const person of month.people||[])names.add(person.name);
    for(const person of period.trim?.people||[])names.add(person.name);
  }
  return [...names].sort((a,b)=>a.localeCompare(b,'pt-BR'));
}
function renderEvolutionFilters(){
  const select=$('evolutionPerson');
  if(!select)return;
  const current=select.value||'all';
  const names=allEvolutionPeople();
  select.innerHTML='<option value="all">Equipe completa</option>'+names.map(name=>`<option value="${name}">${name}</option>`).join('');
  if(current==='all'||names.includes(current))select.value=current;
}
function personMetricFromMonth(month,personName,metric){
  if(personName==='all'){
    const aggregate=aggregateMonths([month]);
    if(metric==='box')return (month.people||[]).reduce((sum,p)=>sum+num(p.box),0);
    return num(aggregate[metric]);
  }
  const person=(month.people||[]).find(p=>p.name===personName);
  return person?num(person[metric]):null;
}
function personMetricFromPeriod(period,personName,metric){
  if(personName==='all'){
    const aggregate=aggregatePeriod(period);
    if(metric==='box')return (aggregate.people||[]).reduce((sum,p)=>sum+num(p.box),0);
    return num(aggregate[metric]);
  }
  const aggregate=aggregatePeriod(period);
  const person=(aggregate.people||[]).find(p=>p.name===personName);
  return person?num(person[metric]):null;
}
function evolutionSeries(){
  const mode=$('evolutionMode')?.value||'monthly';
  const metric=$('evolutionMetric')?.value||'att';
  const person=$('evolutionPerson')?.value||'all';
  const points=[];
  if(mode==='monthly'){
    for(const period of state.periods){
      for(const month of period.months||[]){
        if(!(month.people||[]).length)continue;
        points.push({
          label:`${month.month}/${period.year}`,
          value:personMetricFromMonth(month,person,metric),
          year:period.year,
          order:month.monthIndex
        });
      }
    }
    points.sort((a,b)=>a.year-b.year||a.order-b.order);
  }else{
    for(const period of state.periods){
      points.push({
        label:period.label.replace(' (parcial)',''),
        value:personMetricFromPeriod(period,person,metric),
        year:period.year,
        order:period.quarter
      });
    }
  }
  return {mode,metric,person,points};
}
function evolutionVariation(current,previous){
  if(current==null||previous==null||previous===0)return '—';
  const change=(current-previous)/Math.abs(previous)*100;
  return `${change>=0?'▲':'▼'} ${Math.abs(change).toFixed(1)}%`;
}
function renderEvolutionTable(series,config){
  $('evolutionValueHeader').textContent=config.label;
  $('evolutionBody').innerHTML=series.points.map((point,index)=>{
    const previous=index?series.points[index-1].value:null;
    const variation=evolutionVariation(point.value,previous);
    const changeClass=variation.startsWith('▲')?'diag-ok':variation.startsWith('▼')?'diag-warn':'';
    return `<tr><td>${point.label}</td><td>${point.value==null?'—':config.format(point.value)}</td><td class="${changeClass}">${variation}</td></tr>`;
  }).join('');
}
function renderEvolutionKpis(series,config){
  const valid=series.points.filter(point=>point.value!=null);
  const values=valid.map(point=>point.value);
  const latest=valid.at(-1);
  const previous=valid.at(-2);
  let best=null;
  if(valid.length)best=[...valid].sort((a,b)=>b.value-a.value)[0];
  $('evolutionKpis').innerHTML=[
    metric('Último período',latest?config.format(latest.value):'—',latest?.label||'Sem dados'),
    metric('Período anterior',previous?config.format(previous.value):'—',previous?.label||'Sem comparação'),
    metric('Melhor resultado',best?config.format(best.value):'—',best?.label||'Sem dados'),
    metric('Variação recente',latest&&previous?evolutionVariation(latest.value,previous.value):'—','Em relação ao período anterior')
  ].join('');
}
function renderHistory(){
  renderEvolutionFilters();
  const series=evolutionSeries();
  const config=evolutionMetricConfig(series.metric);
  const personLabel=series.person==='all'?'Equipe completa':series.person;
  const modeLabel=series.mode==='monthly'?'mensal':'trimestral';
  $('evolutionChartTitle').textContent=`Evolução ${modeLabel}`;
  $('evolutionChartSubtitle').textContent=`${personLabel} · ${config.label}`;
  renderEvolutionKpis(series,config);
  renderEvolutionTable(series,config);
  chart('history',$('historyChart'),'line',{
    labels:series.points.map(point=>point.label),
    datasets:[{
      label:config.label,
      data:series.points.map(point=>point.value),
      borderColor:config.color,
      backgroundColor:`${config.color}22`,
      pointBackgroundColor:config.color,
      pointRadius:4,
      pointHoverRadius:6,
      spanGaps:false,
      fill:true,
      tension:.3
    }]
  },{
    interaction:{mode:'index',intersect:false},
    plugins:{
      tooltip:{
        callbacks:{
          label:context=>`${config.label}: ${context.raw==null?'Sem dados':config.format(context.raw)}`
        }
      }
    },
    scales:{
      y:{
        beginAtZero:series.metric!=='final',
        suggestedMin:series.metric==='final'?4.5:undefined,
        suggestedMax:series.metric==='final'?5:undefined,
        title:{display:true,text:config.axis,color:'#9aabc2'}
      }
    }
  });
}

function renderAudit(){const rows=[];for(const p of state.periods){for(const m of p.months){for(const a of m.audit||[])rows.push(`<tr><td>${p.fileName}</td><td>${m.sheet} · ${a.name}</td><td class="${a.status==='ok'?'diag-ok':'diag-warn'}">${a.status==='ok'?'OK':'Atenção'}</td><td>${a.detail}</td></tr>`);if(!(m.audit||[]).length)rows.push(`<tr><td>${p.fileName}</td><td>${m.sheet}</td><td class="diag-warn">Atenção</td><td>Nenhum atendente identificado</td></tr>`)}}$('auditBody').innerHTML=rows.join('')}
function renderAll(){try{renderKpis();renderInsights();renderDashboardCharts();renderTeam();renderHall();renderBox();renderBonus();renderVoice();renderHistory();renderAudit();clearStatus()}catch(err){console.error(err);showStatus(`<b>Erro ao montar o dashboard.</b> ${err.message}`)}}
function showPage(name){document.querySelectorAll('.page').forEach(x=>x.classList.toggle('active',x.id===`page-${name}`));document.querySelectorAll('.nav-btn').forEach(x=>x.classList.toggle('active',x.dataset.page===name));$('pageTitle').textContent=titles[name][0];$('pageSubtitle').textContent=titles[name][1];$('sidebar').classList.remove('open')}
function bind(){ $('periodSelect').addEventListener('change',e=>selectPeriod(Number(e.target.value)));$('monthSelect').addEventListener('change',e=>{state.month=e.target.value==='all'?null:Number(e.target.value);renderAll()});$('personSelect').addEventListener('change',e=>renderPerson(e.target.value));$('voicePerson').addEventListener('change',renderVoice);$('evolutionMode').addEventListener('change',renderHistory);$('evolutionMetric').addEventListener('change',renderHistory);$('evolutionPerson').addEventListener('change',renderHistory);$('nav').addEventListener('click',e=>{const b=e.target.closest('[data-page]');if(b)showPage(b.dataset.page)});$('menuBtn').addEventListener('click',()=>$('sidebar').classList.toggle('open'));$('fileInput').addEventListener('change',async e=>{try{for(const f of e.target.files)state.periods.push(parseWorkbook(await f.arrayBuffer(),f.name));state.periods.sort((a,b)=>a.year-b.year||a.quarter-b.quarter);setupPeriods();clearStatus()}catch(err){showStatus(`<b>Não foi possível importar a planilha.</b> ${err.message}`)}})}
document.addEventListener('DOMContentLoaded',()=>{bind();loadDefault()});
})();
