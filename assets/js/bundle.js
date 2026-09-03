(function(){
'use strict';
const monthOrder={Jan:1,Fev:2,Mar:3,Abr:4,Mai:5,Jun:6,Jul:7,Ago:8,Set:9,Out:10,Nov:11,Dez:12};
const state={periods:[],period:null,month:null,charts:{},teamConfig:{employees:[]},settings:{}};

const MANAGER_PASSWORD='Power@2026@';
const MANAGER_SESSION_KEY='powerAnalyticsManagerAuth';
const PROTECTED_PAGES=new Set(['manager','settings']);
let pendingProtectedPage=null;
function managerIsAuthenticated(){return sessionStorage.getItem(MANAGER_SESSION_KEY)==='true'}
function updateManagerAccessUI(){
  const logged=managerIsAuthenticated();
  const btn=$('managerSessionBtn');
  if(btn){btn.textContent=logged?'🔓 Sair do gestor':'🔒 Área do gestor';btn.classList.toggle('authenticated',logged)}
  document.querySelectorAll('.protected-nav').forEach(item=>{
    const page=item.dataset.page;
    item.textContent=`${logged?'🔓':'🔒'} ${page==='manager'?'Central do Gestor':'Gestão da Equipe'}`;
    item.classList.toggle('unlocked',logged);
  });
}
function openManagerAuth(page=null){
  pendingProtectedPage=page;
  $('managerAuthError').textContent='';
  $('managerPassword').value='';
  $('managerPassword').type='password';
  $('toggleManagerPassword').textContent='👁';
  $('managerAuthOverlay').classList.add('open');
  $('managerAuthOverlay').setAttribute('aria-hidden','false');
  setTimeout(()=>$('managerPassword').focus(),50);
}
function closeManagerAuth(){
  $('managerAuthOverlay').classList.remove('open');
  $('managerAuthOverlay').setAttribute('aria-hidden','true');
  pendingProtectedPage=null;
}
function managerLogout(){
  sessionStorage.removeItem(MANAGER_SESSION_KEY);
  updateManagerAccessUI();
  const active=document.querySelector('.nav-btn.active')?.dataset.page;
  if(PROTECTED_PAGES.has(active))showPage('dashboard');
  showStatus('<b>Sessão do gestor encerrada.</b>','success');
  setTimeout(clearStatus,2200);
}
function requestPage(name){
  if(PROTECTED_PAGES.has(name)&&!managerIsAuthenticated()){openManagerAuth(name);return}
  showPage(name);
}
function bindManagerAuth(){
  $('managerAuthClose').addEventListener('click',closeManagerAuth);
  $('managerAuthCancel').addEventListener('click',closeManagerAuth);
  $('managerAuthOverlay').addEventListener('click',e=>{if(e.target===e.currentTarget)closeManagerAuth()});
  $('toggleManagerPassword').addEventListener('click',()=>{const input=$('managerPassword');input.type=input.type==='password'?'text':'password';$('toggleManagerPassword').textContent=input.type==='password'?'👁':'🙈'});
  $('managerAuthForm').addEventListener('submit',e=>{
    e.preventDefault();
    if($('managerPassword').value!==MANAGER_PASSWORD){$('managerAuthError').textContent='Senha incorreta. Tente novamente.';$('managerPassword').select();return}
    sessionStorage.setItem(MANAGER_SESSION_KEY,'true');
    const target=pendingProtectedPage;
    closeManagerAuth();
    updateManagerAccessUI();
    if(target)showPage(target);
  });
  $('managerSessionBtn').addEventListener('click',()=>managerIsAuthenticated()?managerLogout():openManagerAuth());
  document.addEventListener('keydown',e=>{if(e.key==='Escape'&&$('managerAuthOverlay').classList.contains('open'))closeManagerAuth()});
  updateManagerAccessUI();
}

const $=id=>document.getElementById(id);
const fmt={int:v=>new Intl.NumberFormat('pt-BR',{maximumFractionDigits:0}).format(v||0),dec:(v,d=2)=>new Intl.NumberFormat('pt-BR',{minimumFractionDigits:d,maximumFractionDigits:d}).format(v||0),pct:v=>`${new Intl.NumberFormat('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2}).format(v||0)}%`,money:v=>new Intl.NumberFormat('pt-BR',{style:'currency',currency:'BRL'}).format(v||0)};
const titles={dashboard:['Dashboard Executivo','Visão consolidada da operação'],team:['Equipe','Desempenho individual'],profiles:['Perfil & Conquistas','Pessoas, tempo de casa e reconhecimentos'],hall:['Hall da Fama','Destaques e reconhecimentos'],box:['Caixinha','Ranking mensal da dinâmica'],bonus:['Bonificação','Acompanhamento financeiro'],voice:['Voz do Cliente','Comentários e percepção dos clientes'],history:['Evolução','Comparativos históricos'],manager:['Central do Gestor','Metas, alertas e destaques'],settings:['Gestão da Equipe','Cadastro, status e metas'],audit:['Auditoria','Conferência das regras de negócio']};
function slug(s){return String(s||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/[^a-z0-9]+/g,'-').replace(/(^-|-$)/g,'')}
function avatar(name,size=''){const emp=employeeByName(name);const display=emp?.displayName||name;const initial=(display||'?').trim().charAt(0).toUpperCase();const file=emp?.photo||`assets/img/team/${slug(name)}.jpg`;return `<div class="avatar ${size}"><img src="${file}" alt="${display}" onerror="this.style.display='none';this.parentElement.dataset.fallback='${initial}'"><span class="avatar-fallback">${initial}</span></div>`}
function metric(label,value,delta=''){return `<div class="metric"><div class="label">${label}</div><div class="value">${value}</div><div class="delta">${delta||'Período selecionado'}</div></div>`}
function trend(current,previous,invert=false){if(!previous)return 'Sem comparação anterior';const d=(current-previous)/Math.abs(previous)*100;const good=invert?d<=0:d>=0;return `<span style="color:${good?'var(--green)':'var(--red)'}">${d>=0?'▲':'▼'} ${Math.abs(d).toFixed(1)}% vs anterior</span>`}
function showStatus(message,type='error'){const el=$('systemStatus');if(!el)return;el.hidden=false;el.className=`system-status ${type}`;el.innerHTML=message}
function clearStatus(){const el=$('systemStatus');if(el)el.hidden=true}
function num(v){return Number.isFinite(Number(v))?Number(v):0}
function cleanName(v){return String(v??'').replace(/^\s*\d+\s*-\s*/,'').trim()}
function monthToken(name){return Object.keys(monthOrder).find(m=>new RegExp(`\\b${m}\\b`,'i').test(name))}
function rowsOf(wb,sheet){return XLSX.utils.sheet_to_json(wb.Sheets[sheet],{header:1,defval:null,raw:true})}



function normalizeEmployeeName(name){return String(name||'').trim().toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'')}
const IGNORED_EMPLOYEE_NAMES=new Set(['livre','murilo','sem colaborador','sem atendente','vago','disponivel']);
function isIgnoredEmployeeName(name){return IGNORED_EMPLOYEE_NAMES.has(normalizeEmployeeName(name))}
function canonicalEmployeeName(name){
  const clean=cleanName(name);const n=normalizeEmployeeName(clean);
  if(n==='thamies')return 'Thamires';
  return clean;
}
function employeeByName(name){
  const n=normalizeEmployeeName(canonicalEmployeeName(name));
  return (state.teamConfig.employees||[]).find(e=>normalizeEmployeeName(e.name)===n||normalizeEmployeeName(e.displayName)===n);
}
function employeeProfile(name){
  const emp=employeeByName(name)||{};
  return {...emp,name:emp.name||name,achievements:[...(emp.achievements||[])]};
}
function hasProfileData(name){const p=employeeProfile(name);return !!(p.admissionDate||p.birthDayMonth||p.achievements?.length)}
function latestRosterNames(){
  const names=new Set();
  const period=state.periods[state.periods.length-1];
  if(!period)return names;
  const populated=[...(period.months||[])].filter(m=>(m.people||[]).length);
  const latest=populated[populated.length-1];
  for(const p of (latest?.people||period.trim?.people||[]))if(!isIgnoredEmployeeName(p.name))names.add(normalizeEmployeeName(canonicalEmployeeName(p.name)));
  return names;
}
function profileRoster(){
  const recent=latestRosterNames();
  return (state.teamConfig.employees||[]).filter(e=>!isIgnoredEmployeeName(e.name)&&e.status!=='inactive'&&(hasProfileData(e.name)||recent.has(normalizeEmployeeName(e.name)))).sort((a,b)=>(a.displayName||a.name).localeCompare(b.displayName||b.name,'pt-BR'));
}
function parseIsoDate(value){
  const m=String(value||'').match(/^(\d{4})-(\d{2})-(\d{2})$/); if(!m)return null;
  return new Date(Number(m[1]),Number(m[2])-1,Number(m[3]),12,0,0);
}
function formatDateBR(value){const d=parseIsoDate(value);return d?new Intl.DateTimeFormat('pt-BR').format(d):'Não cadastrada'}
function parseBirthDayMonth(value){const m=String(value||'').trim().match(/^(\d{1,2})\/(\d{1,2})$/);if(!m)return null;const day=Number(m[1]),month=Number(m[2]);if(day<1||day>31||month<1||month>12)return null;return {day,month}}
function tenureText(value,now=new Date()){
  const d=parseIsoDate(value);if(!d||d>now)return 'Admissão não cadastrada';
  let years=now.getFullYear()-d.getFullYear(),months=now.getMonth()-d.getMonth(),days=now.getDate()-d.getDate();
  if(days<0){months--;const prev=new Date(now.getFullYear(),now.getMonth(),0).getDate();days+=prev}
  if(months<0){years--;months+=12}
  const parts=[];if(years)parts.push(`${years} ${years===1?'ano':'anos'}`);if(months)parts.push(`${months} ${months===1?'mês':'meses'}`);if(!years&&!months)parts.push(`${Math.max(0,days)} ${days===1?'dia':'dias'}`);return parts.join(' e ');
}
function yearsAtCompany(value,at=new Date()){const d=parseIsoDate(value);if(!d)return 0;let y=at.getFullYear()-d.getFullYear();if(at.getMonth()<d.getMonth()||(at.getMonth()===d.getMonth()&&at.getDate()<d.getDate()))y--;return Math.max(0,y)}
function daysUntilMonthDay(value,now=new Date()){
  const b=parseBirthDayMonth(value);if(!b)return null;
  let next=new Date(now.getFullYear(),b.month-1,b.day,12);const today=new Date(now.getFullYear(),now.getMonth(),now.getDate(),12);
  if(next<today)next=new Date(now.getFullYear()+1,b.month-1,b.day,12);
  return {days:Math.round((next-today)/86400000),date:next};
}
function daysUntilAnniversary(value,now=new Date()){
  const d=parseIsoDate(value);if(!d)return null;
  let next=new Date(now.getFullYear(),d.getMonth(),d.getDate(),12);const today=new Date(now.getFullYear(),now.getMonth(),now.getDate(),12);
  if(next<today)next=new Date(now.getFullYear()+1,d.getMonth(),d.getDate(),12);
  return {days:Math.round((next-today)/86400000),date:next,years:next.getFullYear()-d.getFullYear()};
}
function activePeople(people){return (people||[]).filter(p=>!isIgnoredEmployeeName(p.name)&&employeeByName(p.name)?.status!=='inactive')}
function loadSavedTeam(){
  try{
    const saved=localStorage.getItem('powerAnalyticsTeamConfig');
    if(!saved)return;
    const local=JSON.parse(saved);
    const official=state.teamConfig?.employees||[];
    const savedEmployees=local?.employees||[];
    state.teamConfig={
      ...state.teamConfig,
      ...local,
      employees:official.map(base=>{
        const previous=savedEmployees.find(item=>normalizeEmployeeName(item.name)===normalizeEmployeeName(base.name));
        if(!previous)return base;
        return {
          ...base,
          ...previous,
          photo:base.photo,
          ramal:base.ramal||previous.ramal,
          admissionDate:previous.admissionDate||base.admissionDate||'',
          birthDayMonth:previous.birthDayMonth||base.birthDayMonth||'',
          terminationDate:previous.terminationDate||base.terminationDate||null,
          achievements:[...(base.achievements||[]),...(previous.achievements||[])].filter((item,index,arr)=>arr.findIndex(x=>JSON.stringify(x)===JSON.stringify(item))===index),
          goals:{...(base.goals||{}),...(previous.goals||{})}
        };
      })
    };
  }catch(e){console.warn('Configuração local indisponível',e)}
}
function saveTeamConfig(){
  try{
    localStorage.setItem('powerAnalyticsTeamConfig',JSON.stringify(state.teamConfig));
    showStatus('<b>Configurações salvas.</b> Metas e status foram atualizados neste navegador.','success');
    setTimeout(clearStatus,2500);
  }catch(e){showStatus('<b>Não foi possível salvar as configurações.</b> '+e.message)}
}
function autoRegisterEmployees(){
  const names=new Set();
  for(const period of state.periods){
    for(const month of period.months||[])for(const p of month.people||[])if(!isIgnoredEmployeeName(p.name))names.add(canonicalEmployeeName(p.name));
    for(const p of period.trim?.people||[])if(!isIgnoredEmployeeName(p.name))names.add(canonicalEmployeeName(p.name));
  }
  // Perfis cadastrados devem aparecer mesmo antes do primeiro atendimento.
  for(const p of state.teamConfig.employees||[])if(!isIgnoredEmployeeName(p.name))names.add(canonicalEmployeeName(p.name));
  const defaults=state.settings.defaultGoals||{att:500,rated:100,rate:30,final:4.9,box:500};
  for(const name of names){
    if(!name||isIgnoredEmployeeName(name)||employeeByName(name))continue;
    state.teamConfig.employees.push({
      id:slug(name),ramal:'',name,displayName:name,role:'Analista de Suporte',status:'active',
      photo:`assets/img/team/${slug(name)}.jpg`,goals:{...defaults}
    });
  }
}

function goalProgress(value,target,metric){
  if(!target)return 0;
  if(metric==='final')return Math.min(100,value/target*100);
  return Math.min(100,value/target*100);
}
function employeeGoals(name){return employeeByName(name)?.goals||state.settings.defaultGoals||{}}

function parseReviews(wb){
  const sheets=wb.SheetNames.filter(n=>/^avalia[cç][oõ]es\s/i.test(n));
  const reviews=[];
  for(const sheet of sheets){
    const token=monthToken(sheet)||sheet.replace(/^avalia[cç][oõ]es\s*/i,'').trim();
    const rows=rowsOf(wb,sheet);
    for(let r=1;r<rows.length;r++){
      const row=rows[r]||[];
      const person=canonicalEmployeeName(row[0]);
      const comment=String(row[1]??'').trim();
      if(person&&!isIgnoredEmployeeName(person)&&comment)reviews.push({month:token,monthIndex:monthOrder[token]||99,sheet,person,comment});
    }
  }
  return reviews.sort((a,b)=>a.monthIndex-b.monthIndex);
}

function parseTop(rows){const out=[];for(let r=1;r<9;r++){const row=rows[r]||[];const name=canonicalEmployeeName(row[0]);if(!name||isIgnoredEmployeeName(name)||/^TOTAL$/i.test(name))continue;if(row[1]==null&&row[2]==null)continue;out.push({name,att:num(row[1]),rated:num(row[2]),rate:num(row[3])*100,avg:num(row[4]),discount:num(row[5]),final:num(row[6]),box:num(row[8]),bonus:num(row[9])})}return out}
function parseBase(rows){const out=[];for(let r=14;r<30;r++){const row=rows[r]||[];const name=canonicalEmployeeName(row[0]);if(!name||isIgnoredEmployeeName(name)||/^ATENDENTE$/i.test(name)||/^TOTAL$/i.test(name))continue;if(row[1]==null&&row[2]==null&&row[3]==null)continue;out.push({name,qsa:num(row[1]),qea:num(row[2]),ts:num(row[3]),te:num(row[4]),whats:num(row[5]),whatsRated:num(row[6])})}return out}
function mergeAudit(top,base){const map=new Map(base.map(x=>[x.name,x]));return top.map(x=>{const b=map.get(x.name);if(!b)return{name:x.name,status:'warn',detail:'Sem linha correspondente na tabela-base A14'};const calcAtt=b.ts+b.te+b.whats,calcRated=b.qsa+b.qea,issues=[];if(Math.abs(calcAtt-x.att)>.01)issues.push(`Atendimentos: tabela ${x.att}, base ${calcAtt}`);if(Math.abs(calcRated-x.rated)>.01)issues.push(`Avaliações: tabela ${x.rated}, base ${calcRated}`);if(Math.abs((x.avg-x.discount)-x.final)>.01)issues.push('Média final divergente de média − desconto');return{name:x.name,status:issues.length?'warn':'ok',detail:issues.length?issues.join(' · '):'Cálculos conferidos'}})}

function periodKey(period){
  return `${Number(period?.year)||0}-Q${Number(period?.quarter)||0}`;
}
function normalizePeriods(periods){
  const unique=new Map();
  for(const period of periods||[]){
    if(!period)continue;
    unique.set(periodKey(period),period);
  }
  return [...unique.values()].sort((a,b)=>a.year-b.year||a.quarter-b.quarter);
}
function upsertPeriod(period){
  const key=periodKey(period);
  const index=state.periods.findIndex(item=>periodKey(item)===key);
  if(index>=0){
    const previous=state.periods[index];
    state.periods[index]={
      ...previous,
      ...period,
      id:previous.id||period.id,
      label:period.label||previous.label
    };
    return {action:'updated',index};
  }
  state.periods.push(period);
  state.periods=normalizePeriods(state.periods);
  return {action:'added',index:state.periods.findIndex(item=>periodKey(item)===key)};
}

function summarizePeople(people){
  const rows=(people||[]).filter(p=>!isIgnoredEmployeeName(p.name));
  const att=rows.reduce((sum,p)=>sum+num(p.att),0),rated=rows.reduce((sum,p)=>sum+num(p.rated),0);
  return {
    att,rated,rate:att?rated/att*100:0,
    avg:rated?rows.reduce((sum,p)=>sum+num(p.avg)*num(p.rated),0)/rated:0,
    discount:rated?rows.reduce((sum,p)=>sum+num(p.discount)*num(p.rated),0)/rated:0,
    final:rated?rows.reduce((sum,p)=>sum+num(p.final)*num(p.rated),0)/rated:0,
    bonus:rows.reduce((sum,p)=>sum+num(p.bonus),0)
  };
}
function parseWorkbook(buffer,fileName,labelHint=''){if(!window.XLSX)throw new Error('Biblioteca de leitura do Excel indisponível.');const wb=XLSX.read(buffer,{type:'array',cellFormula:true,cellDates:true});const npsSheets=wb.SheetNames.filter(n=>/^nps\s/i.test(n));const monthly=npsSheets.filter(n=>!/(trim|trimestral)/i.test(n)).map(sheet=>{const rows=rowsOf(wb,sheet),token=monthToken(sheet)||sheet.replace(/^nps\s*/i,'').trim(),top=parseTop(rows),base=parseBase(rows);return{sheet,month:token,monthIndex:monthOrder[token]||99,people:top,base,audit:mergeAudit(top,base)}}).sort((a,b)=>a.monthIndex-b.monthIndex);const reviews=parseReviews(wb);const trimSheet=npsSheets.find(n=>/(trim|trimestral)/i.test(n)),trimRows=trimSheet?rowsOf(wb,trimSheet):[],trimPeople=trimSheet?parseTop(trimRows):[],totalRow=trimRows.find(r=>String(r?.[0]??'').trim().toUpperCase()==='TOTAL')||[],inferred=fileName.match(/([1-4])\s*trim.*?(20\d{2})/i),quarter=inferred?Number(inferred[1]):Math.ceil(((monthly[0]?.monthIndex)||1)/3),year=inferred?Number(inferred[2]):Number((labelHint.match(/20\d{2}/)||['2026'])[0]);const rawTotal={att:num(totalRow[1]),rated:num(totalRow[2]),rate:num(totalRow[3])*100,avg:num(totalRow[4]),discount:num(totalRow[5]),final:num(totalRow[6]),bonus:num(totalRow[9])},total=trimPeople.length?summarizePeople(trimPeople):rawTotal;const isPartial=!total.att&&monthly.some(m=>(m.people||[]).length);const label=labelHint||`${quarter}º Trimestre ${year}${isPartial?' (parcial)':''}`;return{id:`${year}-Q${quarter}`,fileName,label,year,quarter,status:isPartial?'partial':'closed',months:monthly,reviews,trim:{sheet:trimSheet,people:trimPeople,total},sheets:wb.SheetNames}}
function aggregateMonths(months){const map=new Map();for(const m of months)for(const p of m.people){if(isIgnoredEmployeeName(p.name))continue;const personName=canonicalEmployeeName(p.name);const x=map.get(personName)||{name:personName,att:0,rated:0,avgWeighted:0,finalWeighted:0,discountWeighted:0,box:0};x.att+=p.att;x.rated+=p.rated;x.avgWeighted+=p.avg*p.rated;x.finalWeighted+=p.final*p.rated;x.discountWeighted+=p.discount*p.rated;x.box+=p.box;map.set(personName,x)}const people=[...map.values()].map(x=>({...x,rate:x.att?x.rated/x.att*100:0,avg:x.rated?x.avgWeighted/x.rated:0,final:x.rated?x.finalWeighted/x.rated:0,discount:x.rated?x.discountWeighted/x.rated:0}));const att=people.reduce((s,x)=>s+x.att,0),rated=people.reduce((s,x)=>s+x.rated,0);return{att,rated,rate:att?rated/att*100:0,final:rated?people.reduce((s,x)=>s+x.final*x.rated,0)/rated:0,avg:rated?people.reduce((s,x)=>s+x.avg*x.rated,0)/rated:0,discount:rated?people.reduce((s,x)=>s+x.discount*x.rated,0)/rated:0,people}}
function aggregatePeriod(p){const people=(p.trim?.people||[]).filter(x=>!isIgnoredEmployeeName(x.name));return people.length?{...summarizePeople(people),people}:aggregateMonths(p.months)}
function destroyChart(key){if(state.charts[key]&&typeof state.charts[key].destroy==='function')state.charts[key].destroy();state.charts[key]=null}
function chart(key,el,type,data,options={}){destroyChart(key);if(!el||!window.Chart){if(!window.Chart)showStatus('<b>Os dados foram carregados, mas a biblioteca de gráficos não respondeu.</b> Atualize a página ou verifique o bloqueio de scripts externos.','warn');return}state.charts[key]=new Chart(el,{type,data,options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{labels:{color:'#b8c5d9'}},...(options.plugins||{})},scales:type==='doughnut'?undefined:{x:{ticks:{color:'#9aabc2'},grid:{color:'#20314e55'}},y:{ticks:{color:'#9aabc2'},grid:{color:'#20314e55'},beginAtZero:true},...(options.scales||{})},...options}})}
async function loadPeriodsFromManifest(){
  const manifestResponse=await fetch(`data/manifest.json?v=${Date.now()}`,{cache:'no-store'});
  if(!manifestResponse.ok)throw new Error(`Manifesto HTTP ${manifestResponse.status}`);
  const manifest=await manifestResponse.json();
  if(!Array.isArray(manifest)||!manifest.length)throw new Error('Manifesto de planilhas vazio');
  const periods=[];
  const failures=[];
  for(const item of manifest){
    try{
      const separator=String(item.path).includes('?')?'&':'?';
      const response=await fetch(`${item.path}${separator}v=${Date.now()}`,{cache:'no-store'});
      if(!response.ok)throw new Error(`HTTP ${response.status}`);
      const buffer=await response.arrayBuffer();
      periods.push(parseWorkbook(buffer,item.path.split('/').pop(),item.label||''));
    }catch(error){
      failures.push(`${item.label||item.path}: ${error.message}`);
    }
  }
  if(!periods.length)throw new Error(`Nenhuma planilha pôde ser carregada. ${failures.join(' | ')}`);
  if(failures.length)console.warn('Algumas planilhas não foram carregadas:',failures);
  return normalizePeriods(periods);
}
async function loadPeriodsFromJsonFallback(){
  const response=await fetch(`data/dashboard-data.json?v=${Date.now()}`,{cache:'no-store'});
  if(!response.ok)throw new Error(`Dados HTTP ${response.status}`);
  return normalizePeriods(await response.json());
}
async function loadDefault(){clearStatus();try{
  const [teamResponse,settingsResponse]=await Promise.all([
    fetch(`config/team.json?v=${Date.now()}`,{cache:'no-store'}),
    fetch(`config/settings.json?v=${Date.now()}`,{cache:'no-store'})
  ]);
  if(teamResponse.ok)state.teamConfig=await teamResponse.json();
  if(settingsResponse.ok)state.settings=await settingsResponse.json();
  loadSavedTeam();
  try{
    state.periods=await loadPeriodsFromManifest();
  }catch(excelError){
    console.warn('Leitura automática do Excel falhou; usando JSON de contingência.',excelError);
    state.periods=await loadPeriodsFromJsonFallback();
    showStatus(`<b>Os dados foram carregados pelo arquivo de contingência.</b> A leitura automática das planilhas falhou: ${excelError.message}`,'warn');
  }
  if(!Array.isArray(state.periods)||!state.periods.length)throw new Error('Nenhum período encontrado');
  autoRegisterEmployees();
  setupPeriods();
}catch(err){
  showStatus(`<b>Não foi possível carregar os dados.</b> Detalhe: ${err.message}`);
}}
function setupPeriods(){const ps=$('periodSelect');if(!state.periods.length)return;ps.innerHTML=state.periods.map((p,i)=>`<option value="${i}">${p.label}</option>`).join('');ps.value=String(Math.max(0,state.periods.length-1));selectPeriod(Number(ps.value))}
function selectPeriod(index){state.period=state.periods[index];if(!state.period)return;const ms=$('monthSelect');ms.innerHTML='<option value="all">Trimestre completo</option>'+state.period.months.map((m,i)=>`<option value="${i}">${m.month}</option>`).join('');ms.value='all';state.month=null;renderAll()}
function currentData(){return state.month==null?aggregatePeriod(state.period):aggregateMonths([state.period.months[state.month]])}
function previousData(){const idx=state.periods.indexOf(state.period);if(state.month!=null){const m=state.month-1;return m>=0?aggregateMonths([state.period.months[m]]):null}return idx>0?aggregatePeriod(state.periods[idx-1]):null}
function renderKpis(){const d=currentData(),p=previousData();$('kpis').innerHTML=[metric('Atendimentos',fmt.int(d.att),p?trend(d.att,p.att):'Volume do período'),metric('Avaliações',fmt.int(d.rated),p?trend(d.rated,p.rated):'Retornos recebidos'),metric('Taxa de avaliação',fmt.pct(d.rate),p?trend(d.rate,p.rate):'Participação dos clientes'),metric('Média final',fmt.dec(d.final),p?trend(d.final,p.final):'Qualidade percebida'),metric('Bonificação',fmt.money(state.period.trim.total.bonus),state.period.status==='partial'?'Trimestre em andamento':'Consolidado do trimestre')].join('')}
function dashboardReviews(){let r=[...(state.period?.reviews||[])];if(state.month!=null){const m=state.period.months[state.month]?.month;r=r.filter(x=>x.month===m)}return r}
function employeeOfMonthData(){const d=currentData(),people=activePeople([...(d.people||[])]),reviews=dashboardReviews();if(!people.length)return null;const max=k=>Math.max(1,...people.map(x=>num(x[k]))),rc=n=>reviews.filter(x=>normalizeEmployeeName(x.person)===normalizeEmployeeName(n)).length,mr=Math.max(1,...people.map(x=>rc(x.name)));return people.map(x=>({...x,reviewCount:rc(x.name),score:Math.min(1,x.final/5)*.38+x.rate/max('rate')*.20+(x.box||0)/max('box')*.17+x.rated/max('rated')*.15+rc(x.name)/mr*.10})).sort((a,b)=>b.score-a.score)[0]}
function renderExecutiveHero(){const d=currentData(),p=previousData(),w=employeeOfMonthData(),reviews=dashboardReviews();$('executivePeriod').textContent=state.month!=null?`${state.period.months[state.month]?.month}/${state.period.year}`:state.period.label;$('heroFinal').textContent=fmt.dec(d.final);$('heroFinalDelta').innerHTML=p?trend(d.final,p.final):'Período selecionado';$('executiveSummary').innerHTML=`A equipe realizou <b>${fmt.int(d.att)} atendimentos</b>, recebeu <b>${fmt.int(d.rated)} avaliações</b> e alcançou média final de <b>${fmt.dec(d.final)}</b>. ${w?`<b>${w.name}</b> foi o principal destaque do período.`:''}`;$('executiveBadges').innerHTML=`<span>◉ ${fmt.pct(d.rate)} avaliados</span><span>💬 ${reviews.length} comentários</span><span>${state.period.status==='partial'?'◷ Período parcial':'✓ Período consolidado'}</span>`}
function executivePeriodLabel(){return state.month!=null?`${state.period.months[state.month]?.month}/${state.period.year}`:(state.period?.label||'Período selecionado')}
function renderEmployeeOfMonth(){const w=employeeOfMonthData();if(!w){$('employeeOfMonth').innerHTML='<div class="empty">Sem dados.</div>';return}const display=employeeByName(w.name)?.displayName||w.name;$('employeeOfMonth').innerHTML=`<div class="employee-of-month-glow"></div><div class="employee-period-title"><span class="employee-period-icon">🏆</span><div><span class="eyebrow">Colaborador do período</span><small>${executivePeriodLabel()}</small></div></div><div class="employee-period-content"><div class="employee-period-profile">${avatar(w.name,'executive-avatar')}<div class="employee-period-copy"><h2>${display}</h2><p>Desempenho combinado de qualidade, volume, avaliações e Caixinha.</p></div></div><div class="employee-feature-stats"><div class="employee-stat purple"><span>Média</span><b>${fmt.dec(w.final)}</b></div><div class="employee-stat blue"><span>Atendimentos</span><b>${fmt.int(w.att)}</b></div><div class="employee-stat green"><span>Taxa de avaliação</span><b>${fmt.pct(w.rate)}</b></div><div class="employee-stat amber"><span>Caixinha</span><b>${fmt.dec(w.box||0,1)}</b></div></div></div>`}
function renderInsights(){const d=currentData(),p=previousData(),people=activePeople(d.people||[]),reviews=dashboardReviews(),items=[],ta=[...people].sort((a,b)=>b.att-a.att)[0],tf=[...people].sort((a,b)=>b.final-a.final)[0],tr=[...people].sort((a,b)=>b.rate-a.rate)[0],vc=reviews.reduce((m,x)=>(m[x.person]=(m[x.person]||0)+1,m),{}),tv=Object.entries(vc).sort((a,b)=>b[1]-a[1])[0];if(p){const c=p.att?(d.att-p.att)/p.att*100:0;items.push([c>=0?'↗':'↘','Volume de atendimentos',`${c>=0?'Crescimento':'Redução'} de ${Math.abs(c).toFixed(1)}% em relação ao anterior.`,c>=0?'positive':'attention'])}if(ta)items.push(['📞','Liderança de volume',`${ta.name} realizou ${fmt.int(ta.att)} atendimentos.`,'neutral']);if(tf)items.push(['⭐','Qualidade em destaque',`${tf.name} obteve média ${fmt.dec(tf.final)}.`,'positive']);if(tr)items.push(['◎','Taxa de avaliação',`${tr.name} liderou com ${fmt.pct(tr.rate)}.`,'neutral']);if(tv)items.push(['💬','Voz do cliente',`${tv[0]} foi citado em ${tv[1]} comentários.`,'positive']);$('insights').innerHTML=items.map(x=>`<div class="executive-insight ${x[3]}"><span class="insight-icon">${x[0]}</span><div><b>${x[1]}</b><p>${x[2]}</p></div></div>`).join('')}
function renderDashboardReviews(){const r=dashboardReviews().slice(-3).reverse();$('dashboardReviews').innerHTML=r.length?r.map(x=>`<div class="dashboard-review">${avatar(x.person)}<div><b>${x.person}</b><p>“${String(x.comment).replace(/</g,'&lt;').replace(/>/g,'&gt;')}”</p><span>${x.month}</span></div></div>`).join(''):'<div class="empty">Nenhum comentário neste período.</div>'}
function renderDashboardGoals(){const d=currentData(),people=activePeople(d.people||[]),def=state.settings.defaultGoals||{att:500,rated:100,rate:25,final:4.9},t={att:people.reduce((s,p)=>s+(employeeGoals(p.name).att||def.att),0),rated:people.reduce((s,p)=>s+(employeeGoals(p.name).rated||def.rated),0),rate:people.length?people.reduce((s,p)=>s+(employeeGoals(p.name).rate||def.rate),0)/people.length:def.rate,final:people.length?people.reduce((s,p)=>s+(employeeGoals(p.name).final||def.final),0)/people.length:def.final},rows=[['Atendimentos',d.att,t.att,fmt.int],['Avaliações',d.rated,t.rated,fmt.int],['Taxa',d.rate,t.rate,fmt.pct],['Média final',d.final,t.final,v=>fmt.dec(v,2)]];$('dashboardGoals').innerHTML=rows.map(([l,v,g,f])=>`<div class="dashboard-goal"><div><span>${l}</span><b>${f(v)} <small>/ ${f(g)}</small></b></div><div class="goal-track"><span style="width:${Math.min(100,g?v/g*100:0)}%"></span></div></div>`).join('')}
function renderDashboardCharts(){const months=state.period.months;chart('monthly',$('monthlyChart'),'line',{labels:months.map(x=>x.month),datasets:[{label:'Atendimentos',data:months.map(x=>x.people.reduce((s,p)=>s+p.att,0)),borderColor:'#8b5cf6',backgroundColor:'#8b5cf633',tension:.35,fill:true},{label:'Avaliações',data:months.map(x=>x.people.reduce((s,p)=>s+p.rated,0)),borderColor:'#22d3ee',tension:.35}]});const d=currentData(),rank=[...(d.people||[])].sort((a,b)=>b.att-a.att);chart('ranking',$('rankingChart'),'bar',{labels:rank.map(x=>x.name),datasets:[{label:'Atendimentos',data:rank.map(x=>x.att),backgroundColor:'#8b5cf6'}]},{indexAxis:'y',plugins:{legend:{display:false}}})}
function renderTeam(){const d=currentData(),people=activePeople([...(d.people||[])]).sort((a,b)=>b.final-a.final);$('personSelect').innerHTML=people.map(x=>`<option>${x.name}</option>`).join('');$('teamGrid').innerHTML=people.map(x=>`<div class="card person-card" data-person="${x.name}"><div class="person-head">${avatar(x.name)}<div><div class="person-name">${x.name}</div><div class="muted">Média final ${fmt.dec(x.final)}</div></div></div><div class="mini-stats"><div class="mini"><span class="muted">Atendimentos</span><b>${fmt.int(x.att)}</b></div><div class="mini"><span class="muted">Avaliação</span><b>${fmt.pct(x.rate)}</b></div><div class="mini"><span class="muted">Avaliações</span><b>${fmt.int(x.rated)}</b></div><div class="mini"><span class="muted">Caixinha</span><b>${fmt.dec(x.box||0,1)}</b></div></div></div>`).join('');renderPerson(people[0]?.name);document.querySelectorAll('[data-person]').forEach(el=>el.addEventListener('click',()=>{$('personSelect').value=el.dataset.person;renderPerson(el.dataset.person)}))}
function renderPerson(name){if(!name)return;const labels=state.period.months.map(x=>x.month),vals=state.period.months.map(m=>m.people.find(p=>p.name===name)?.att??null),finals=state.period.months.map(m=>m.people.find(p=>p.name===name)?.final??null);$('profileTitle').textContent=`Evolução de ${name}`;chart('person',$('personChart'),'line',{labels,datasets:[{label:'Atendimentos',data:vals,borderColor:'#8b5cf6',tension:.35,yAxisID:'y'},{label:'Média final',data:finals,borderColor:'#34d399',tension:.35,yAxisID:'y1'}]},{scales:{y1:{position:'right',min:0,max:5,ticks:{color:'#9aabc2'},grid:{drawOnChartArea:false}}}})}

function employeeAchievements(name){
  if(isIgnoredEmployeeName(name))return [];
  const normalized=normalizeEmployeeName(canonicalEmployeeName(name)),events=[],profile=employeeProfile(name),now=new Date();
  const admitted=parseIsoDate(profile.admissionDate);
  if(admitted){
    const maxYears=yearsAtCompany(profile.admissionDate,now);
    for(let year=1;year<=maxYears;year++){
      const date=new Date(admitted.getFullYear()+year,admitted.getMonth(),admitted.getDate(),12);
      events.push({person:name,type:'anniversary',icon:'🎉',title:`${year} ${year===1?'ano':'anos'} de Power System`,detail:`Marco de ${year} ${year===1?'ano':'anos'} de empresa`,date:date.toISOString().slice(0,10),sortDate:date});
    }
  }
  for(const period of state.periods||[]){
    for(const month of period.months||[]){
      const people=(month.people||[]).filter(p=>p.box>0);if(!people.length)continue;
      const max=Math.max(...people.map(p=>num(p.box)));if(max<=0)continue;
      for(const p of people.filter(p=>num(p.box)===max&&normalizeEmployeeName(p.name)===normalized)){
        const mi=monthOrder[month.month]||1,date=new Date(period.year,mi,0,12);
        events.push({person:name,type:'box',icon:'🏆',title:`Campeão da Caixinha · ${month.month}/${period.year}`,detail:`${fmt.dec(p.box,1)} pontos`,date:date.toISOString().slice(0,10),sortDate:date});
      }
    }
  }
  for(const item of profile.achievements||[]){
    const date=parseIsoDate(item.date)||new Date(0);
    events.push({person:name,type:'manual',icon:item.icon||'⭐',title:item.title||'Conquista',detail:item.detail||'',date:item.date||'',sortDate:date});
  }
  return events.sort((a,b)=>b.sortDate-a.sortDate);
}
function allActiveAchievements(){return profileRoster().flatMap(e=>employeeAchievements(e.name)).sort((a,b)=>b.sortDate-a.sortDate)}
function profileUpcoming(){
  const now=new Date(),events=[];
  for(const e of profileRoster()){
    const p=employeeProfile(e.name),birth=daysUntilMonthDay(p.birthDayMonth,now),ann=daysUntilAnniversary(p.admissionDate,now);
    if(birth&&birth.days<=45)events.push({person:e.name,kind:'birthday',icon:'🎂',days:birth.days,date:birth.date,title:birth.days===0?'Aniversário hoje!':`Aniversário em ${birth.days} ${birth.days===1?'dia':'dias'}`});
    if(ann&&ann.days<=45&&ann.years>0)events.push({person:e.name,kind:'anniversary',icon:'🎉',days:ann.days,date:ann.date,title:ann.days===0?`Completa ${ann.years} ${ann.years===1?'ano':'anos'} de empresa hoje!`:`Completa ${ann.years} ${ann.years===1?'ano':'anos'} de empresa em ${ann.days} ${ann.days===1?'dia':'dias'}`});
  }
  return events.sort((a,b)=>a.days-b.days);
}
function renderProfiles(){
  const roster=profileRoster(),filter=$('profilePersonFilter');if(!filter)return;
  const previous=filter.value||'all';filter.innerHTML='<option value="all">Toda a equipe</option>'+roster.map(e=>`<option value="${e.name}">${e.displayName||e.name}</option>`).join('');
  filter.value=[...filter.options].some(o=>o.value===previous)?previous:'all';const selected=filter.value;
  const visible=selected==='all'?roster:roster.filter(e=>e.name===selected),allEvents=allActiveAchievements(),events=selected==='all'?allEvents:allEvents.filter(e=>normalizeEmployeeName(e.person)===normalizeEmployeeName(selected)),upcoming=profileUpcoming();
  const wins=allEvents.filter(e=>e.type==='box').length,anniversaries=allEvents.filter(e=>e.type==='anniversary'&&e.sortDate.getFullYear()===new Date().getFullYear()).length;
  $('profileKpis').innerHTML=[metric('Colaboradores no mural',fmt.int(roster.length)),metric('Conquistas registradas',fmt.int(allEvents.length)),metric('Caixinhas conquistadas',fmt.int(wins)),metric('Anos de empresa celebrados',fmt.int(anniversaries)),metric('Celebrações próximas',fmt.int(upcoming.length),'Próximos 45 dias')].join('');
  $('profileCelebration').innerHTML=upcoming.length?`<div class="celebration-icon">${upcoming[0].icon}</div><div><span class="eyebrow">Próxima celebração</span><h3>${upcoming[0].person}</h3><p>${upcoming[0].title} · ${new Intl.DateTimeFormat('pt-BR',{day:'2-digit',month:'long'}).format(upcoming[0].date)}</p></div><div class="celebration-list">${upcoming.slice(1,4).map(x=>`<span>${x.icon} <b>${x.person}</b> · ${x.days===0?'hoje':x.days+'d'}</span>`).join('')}</div>`:`<div class="celebration-icon">✨</div><div><span class="eyebrow">Agenda da equipe</span><h3>Nenhuma celebração nos próximos 45 dias</h3><p>O painel continuará acompanhando aniversários e marcos de empresa automaticamente.</p></div>`;
  $('profileGrid').innerHTML=visible.map(e=>{const p=employeeProfile(e.name),birth=parseBirthDayMonth(p.birthDayMonth),eventsFor=employeeAchievements(e.name),years=yearsAtCompany(p.admissionDate);return `<article class="card profile-card" data-profile-person="${e.name}"><div class="profile-card-head">${avatar(e.name,'large')}<div><h3>${e.displayName||e.name}</h3><span>${e.role||'Equipe de Suporte'}</span></div></div><div class="profile-tenure"><span>Tempo de empresa</span><strong>${tenureText(p.admissionDate)}</strong><small>${p.admissionDate?'Desde '+formatDateBR(p.admissionDate):'Informe a admissão na Gestão da Equipe'}</small></div><div class="profile-facts"><div><span>🎂 Aniversário</span><b>${birth?String(birth.day).padStart(2,'0')+'/'+String(birth.month).padStart(2,'0'):'Não cadastrado'}</b></div><div><span>🎖 Conquistas</span><b>${eventsFor.length}</b></div><div><span>🏢 Anos completos</span><b>${years}</b></div></div><div class="profile-badges">${eventsFor.slice(0,3).map(x=>`<span>${x.icon} ${x.title}</span>`).join('')||'<span>✨ Novas conquistas aparecerão aqui</span>'}</div></article>`}).join('')||'<div class="card empty-state">Nenhum colaborador ativo disponível para este filtro.</div>';
  $('achievementFeed').innerHTML=events.length?events.slice(0,40).map(ev=>`<div class="achievement-item"><div class="achievement-icon ${ev.type}">${ev.icon}</div><div><div class="achievement-meta"><b>${ev.person}</b><span>${ev.date?new Intl.DateTimeFormat('pt-BR',{month:'short',year:'numeric'}).format(ev.sortDate):''}</span></div><h4>${ev.title}</h4>${ev.detail?`<p>${ev.detail}</p>`:''}</div></div>`).join(''):'<div class="empty">Ainda não há conquistas registradas para este colaborador.</div>';
  document.querySelectorAll('[data-profile-person]').forEach(card=>card.addEventListener('click',()=>{filter.value=card.dataset.profilePerson;renderProfiles()}));
}
function renderHall(){const d=currentData(),people=activePeople(d.people||[]),defs=[['📞','Maior atendimento',[...people].sort((a,b)=>b.att-a.att)[0],x=>fmt.int(x.att)],['⭐','Melhor nota',[...people].sort((a,b)=>b.final-a.final)[0],x=>fmt.dec(x.final)],['📈','Melhor taxa',[...people].sort((a,b)=>b.rate-a.rate)[0],x=>fmt.pct(x.rate)],['🎁','Campeão da caixinha',[...people].sort((a,b)=>(b.box||0)-(a.box||0))[0],x=>fmt.dec(x.box||0,1)+' pts']];$('hallGrid').innerHTML=defs.map(([ico,title,p,fn])=>p?`<div class="card hall"><div class="trophy">${ico}</div><div class="muted">${title}</div>${avatar(p.name)}<div class="winner">${p.name}</div><div class="score">${fn(p)}</div></div>`:'').join('')}
function renderBox(){const m=state.month==null?state.period.months[state.period.months.length-1]:state.period.months[state.month],rank=activePeople([...(m?.people||[])]).sort((a,b)=>b.box-a.box),order=[rank[1],rank[0],rank[2]],classes=['two','one','three'],medals=['🥈','🥇','🥉'];$('podium').innerHTML=order.map((p,i)=>p?`<div class="podium-item">${avatar(p.name)}<b>${p.name}</b><div class="muted">${fmt.dec(p.box,1)} pontos</div><div class="step ${classes[i]}"><div style="font-size:28px">${medals[i]}</div><b>${i===1?'1º':i===0?'2º':'3º'}</b></div></div>`:'').join('');$('boxRanking').innerHTML=rank.map((p,i)=>`<div class="rank-row"><div class="rank-pos">${i+1}º</div><div><b>${p.name}</b><div class="muted">${m.month}</div></div><b>${fmt.dec(p.box,1)}</b></div>`).join('')}
function renderBonus(){const p=state.period,people=activePeople([...(p.trim.people||[])]).sort((a,b)=>b.bonus-a.bonus),bonusTotal=people.reduce((sum,x)=>sum+num(x.bonus),0);$('bonusKpis').innerHTML=[metric('Total distribuído',fmt.money(bonusTotal)),metric('Média por colaborador',fmt.money(people.length?bonusTotal/people.length:0)),metric('Maior bonificação',fmt.money(people[0]?.bonus||0)),metric('Colaboradores',fmt.int(people.filter(x=>x.bonus>0).length)),metric('Período',p.label)].join('');chart('bonus',$('bonusChart'),'bar',{labels:people.map(x=>x.name),datasets:[{label:'Bonificação',data:people.map(x=>x.bonus),backgroundColor:'#34d399'}]},{indexAxis:'y',plugins:{legend:{display:false}}});chart('bonusHistory',$('bonusHistoryChart'),'line',{labels:state.periods.map(x=>x.label),datasets:[{label:'Total distribuído',data:state.periods.map(x=>(x.trim?.people||[]).filter(p=>!isIgnoredEmployeeName(p.name)).reduce((sum,p)=>sum+num(p.bonus),0)),borderColor:'#fbbf24',backgroundColor:'#fbbf2433',fill:true,tension:.3}]})}


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


function progressRow(label,value,target,formatter,metric){
  const progress=goalProgress(value,target,metric);
  return `<div class="goal-row"><div class="goal-row-head"><span>${label}</span><b>${formatter(value)} / ${formatter(target)}</b></div><div class="goal-track"><span style="width:${progress}%"></span></div></div>`;
}
function renderManager(){
  const d=currentData();
  const people=activePeople([...(d.people||[])]);
  const achieved=people.filter(p=>{
    const g=employeeGoals(p.name);
    return p.final>=g.final&&p.rate>=g.rate;
  }).length;
  const topFinal=[...people].sort((a,b)=>b.final-a.final)[0];
  const topRate=[...people].sort((a,b)=>b.rate-a.rate)[0];
  $('managerKpis').innerHTML=[
    metric('Colaboradores ativos',fmt.int(people.length)),
    metric('Metas principais atingidas',`${achieved}/${people.length}`),
    metric('Melhor média',topFinal?`${topFinal.name} · ${fmt.dec(topFinal.final)}`:'—'),
    metric('Melhor taxa',topRate?`${topRate.name} · ${fmt.pct(topRate.rate)}`:'—')
  ].join('');
  $('managerGoals').innerHTML=people.map(p=>{
    const g=employeeGoals(p.name);
    return `<div class="manager-person"><div class="manager-person-head">${avatar(p.name)}<div><b>${employeeByName(p.name)?.displayName||p.name}</b><span>${employeeByName(p.name)?.role||'Equipe de Suporte'}</span></div></div>
      ${progressRow('Atendimentos',p.att,g.att,fmt.int,'att')}
      ${progressRow('Taxa de avaliação',p.rate,g.rate,fmt.pct,'rate')}
      ${progressRow('Média final',p.final,g.final,v=>fmt.dec(v,2),'final')}
    </div>`;
  }).join('');
  const alerts=[];
  for(const p of people){
    const g=employeeGoals(p.name);
    if(p.rate<g.rate)alerts.push({type:'warn',text:`${p.name} está abaixo da meta de taxa de avaliação (${fmt.pct(p.rate)} de ${fmt.pct(g.rate)}).`});
    if(p.final>=g.final)alerts.push({type:'ok',text:`${p.name} atingiu a meta de média final com ${fmt.dec(p.final)}.`});
  }
  if(!alerts.length)alerts.push({type:'ok',text:'Nenhum alerta relevante para o período selecionado.'});
  $('managerAlerts').innerHTML=alerts.slice(0,12).map(a=>`<div class="manager-alert ${a.type}">${a.type==='ok'?'✓':'!'}<span>${a.text}</span></div>`).join('');
}
function renderTeamAdmin(){
  const data=currentData();
  const performance=new Map((data.people||[]).map(p=>[p.name,p]));
  $('teamAdminGrid').innerHTML=(state.teamConfig.employees||[]).sort((a,b)=>a.displayName.localeCompare(b.displayName,'pt-BR')).map((e,index)=>{
    const p=performance.get(e.name);
    const g=e.goals||{};
    return `<article class="team-admin-card" data-admin-name="${e.name}">
      <div class="team-admin-head">${avatar(e.name,'large')}<div><h3>${e.displayName}</h3><span>${e.ramal?'Ramal '+e.ramal:'Sem ramal identificado'}</span></div></div>
      <div class="admin-fields">
        <label>Nome de exibição<input class="control admin-display" value="${e.displayName||e.name}"></label>
        <label>Cargo<input class="control admin-role" value="${e.role||''}"></label>
        <label>Status<select class="control admin-status"><option value="active" ${e.status!=='inactive'?'selected':''}>Ativo</option><option value="inactive" ${e.status==='inactive'?'selected':''}>Desligado</option></select></label>
        <label>Data de admissão<input type="date" class="control admin-admission" value="${employeeProfile(e.name).admissionDate||''}"></label>
        <label>Aniversário (DD/MM)<input type="text" inputmode="numeric" maxlength="5" placeholder="DD/MM" class="control admin-birthday" value="${employeeProfile(e.name).birthDayMonth||''}"></label>
        <label>Data de desligamento<input type="date" class="control admin-termination" value="${employeeProfile(e.name).terminationDate||''}"></label>
        <label>Meta atendimentos<input type="number" class="control admin-goal-att" value="${g.att||500}"></label>
        <label>Meta avaliações<input type="number" class="control admin-goal-rated" value="${g.rated||100}"></label>
        <label>Meta taxa (%)<input type="number" step="0.01" class="control admin-goal-rate" value="${g.rate||25}"></label>
        <label>Meta média final<input type="number" step="0.01" class="control admin-goal-final" value="${g.final||4.9}"></label>
        <label>Meta Caixinha<input type="number" step="0.01" class="control admin-goal-box" value="${g.box||500}"></label>
      </div>
      <div class="admin-summary">${p?`Período atual: ${fmt.int(p.att)} atendimentos · ${fmt.pct(p.rate)} · média ${fmt.dec(p.final)}`:'Sem dados no período selecionado'}</div>
    </article>`;
  }).join('');
}
function collectTeamAdmin(){
  document.querySelectorAll('[data-admin-name]').forEach(card=>{
    const e=employeeByName(card.dataset.adminName);
    if(!e)return;
    e.displayName=card.querySelector('.admin-display').value.trim()||e.name;
    e.role=card.querySelector('.admin-role').value.trim();
    e.status=card.querySelector('.admin-status').value;
    e.admissionDate=card.querySelector('.admin-admission')?.value||'';
    e.birthDayMonth=card.querySelector('.admin-birthday')?.value.trim()||'';
    e.terminationDate=card.querySelector('.admin-termination')?.value||'';
    e.goals={
      att:num(card.querySelector('.admin-goal-att').value),
      rated:num(card.querySelector('.admin-goal-rated').value),
      rate:num(card.querySelector('.admin-goal-rate').value),
      final:num(card.querySelector('.admin-goal-final').value),
      box:num(card.querySelector('.admin-goal-box').value)
    };
  });
}

function renderAudit(){const rows=[];for(const p of state.periods){for(const m of p.months){for(const a of m.audit||[])rows.push(`<tr><td>${p.fileName}</td><td>${m.sheet} · ${a.name}</td><td class="${a.status==='ok'?'diag-ok':'diag-warn'}">${a.status==='ok'?'OK':'Atenção'}</td><td>${a.detail}</td></tr>`);if(!(m.audit||[]).length)rows.push(`<tr><td>${p.fileName}</td><td>${m.sheet}</td><td class="diag-warn">Atenção</td><td>Nenhum atendente identificado</td></tr>`)}}$('auditBody').innerHTML=rows.join('')}
function renderAll(){try{renderExecutiveHero();renderKpis();renderEmployeeOfMonth();renderInsights();renderDashboardReviews();renderDashboardGoals();renderDashboardCharts();renderTeam();renderProfiles();renderHall();renderBox();renderBonus();renderVoice();renderHistory();renderManager();renderTeamAdmin();renderAudit();clearStatus()}catch(err){console.error(err);showStatus(`<b>Erro ao montar o dashboard.</b> ${err.message}`)}}
function showPage(name){document.querySelectorAll('.page').forEach(x=>x.classList.toggle('active',x.id===`page-${name}`));document.querySelectorAll('.nav-btn').forEach(x=>x.classList.toggle('active',x.dataset.page===name));$('pageTitle').textContent=titles[name][0];$('pageSubtitle').textContent=titles[name][1];$('sidebar').classList.remove('open')}
function bind(){ $('periodSelect').addEventListener('change',e=>selectPeriod(Number(e.target.value)));$('monthSelect').addEventListener('change',e=>{state.month=e.target.value==='all'?null:Number(e.target.value);renderAll()});$('personSelect').addEventListener('change',e=>renderPerson(e.target.value));$('profilePersonFilter').addEventListener('change',renderProfiles);$('voicePerson').addEventListener('change',renderVoice);$('evolutionMode').addEventListener('change',renderHistory);$('evolutionMetric').addEventListener('change',renderHistory);$('evolutionPerson').addEventListener('change',renderHistory);$('saveTeamBtn').addEventListener('click',()=>{collectTeamAdmin();saveTeamConfig();renderAll()});$('nav').addEventListener('click',e=>{const b=e.target.closest('[data-page]');if(b)requestPage(b.dataset.page)});$('menuBtn').addEventListener('click',()=>$('sidebar').classList.toggle('open'));$('fileInput').addEventListener('change',async e=>{try{
  const results=[];
  for(const f of e.target.files){
    const parsed=parseWorkbook(await f.arrayBuffer(),f.name);
    const result=upsertPeriod(parsed);
    results.push({period:parsed,result});
  }
  state.periods=normalizePeriods(state.periods);
  autoRegisterEmployees();
  setupPeriods();
  const updated=results.filter(x=>x.result.action==='updated').map(x=>x.period.label);
  const added=results.filter(x=>x.result.action==='added').map(x=>x.period.label);
  const messages=[];
  if(updated.length)messages.push(`<b>Período atualizado:</b> ${updated.join(', ')}`);
  if(added.length)messages.push(`<b>Novo período adicionado:</b> ${added.join(', ')}`);
  showStatus(messages.join('<br>')||'<b>Importação concluída.</b>','success');
  e.target.value='';
  setTimeout(clearStatus,4500);
}catch(err){
  showStatus(`<b>Não foi possível importar a planilha.</b> ${err.message}`);
  e.target.value='';
}})}
document.addEventListener('DOMContentLoaded',()=>{bind();bindManagerAuth();loadDefault()});
})();
