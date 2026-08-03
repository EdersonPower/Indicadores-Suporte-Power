const charts={};
const palette=['#38bdf8','#22c55e','#a78bfa','#f59e0b','#f472b6','#14b8a6','#fb7185'];
const grid='#20334a',text='#9fb2c8';
Chart.defaults.color=text;Chart.defaults.borderColor=grid;Chart.defaults.font.family='Inter';
function reset(id){if(charts[id])charts[id].destroy()}
const short=n=>n.split(' - ')[1]||n;
export function drawDashboard(data){
 const months=data.monthly;
 reset('monthly');charts.monthly=new Chart(document.getElementById('monthlyChart'),{type:'line',data:{labels:months.map(x=>x.monthLabel),datasets:[{label:'Atendimentos',data:months.map(x=>x.total.calls),borderColor:palette[0],backgroundColor:'#38bdf822',fill:true,tension:.35},{label:'Avaliações',data:months.map(x=>x.total.ratings),borderColor:palette[1],backgroundColor:'#22c55e22',fill:true,tension:.35}]},options:{responsive:true,maintainAspectRatio:false,interaction:{mode:'index',intersect:false},plugins:{legend:{position:'bottom'}},scales:{x:{grid:{display:false}},y:{beginAtZero:true}}}});
 const agents=[...data.quarterly.agents].sort((a,b)=>b.calls-a.calls);
 reset('ranking');charts.ranking=new Chart(document.getElementById('rankingChart'),{type:'bar',data:{labels:agents.map(x=>short(x.agent)),datasets:[{label:'Atendimentos',data:agents.map(x=>x.calls),backgroundColor:palette[0],borderRadius:8}]},options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',plugins:{legend:{display:false}},scales:{x:{beginAtZero:true},y:{grid:{display:false}}}}});
 reset('score');charts.score=new Chart(document.getElementById('scoreChart'),{type:'bar',data:{labels:agents.map(x=>short(x.agent)),datasets:[{label:'Média final',data:agents.map(x=>x.score),backgroundColor:palette[2],borderRadius:8}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{y:{min:4,max:5},x:{grid:{display:false}}}}});
}
export function drawAgent(data,agent){
 const rows=data.monthly.map(m=>m.agents.find(a=>a.agent===agent)||{calls:0,ratings:0,rate:0,score:0,box:0});
 reset('agent');charts.agent=new Chart(document.getElementById('agentChart'),{type:'line',data:{labels:data.monthly.map(m=>m.monthLabel),datasets:[{label:'Atendimentos',data:rows.map(x=>x.calls),borderColor:palette[0],tension:.3,yAxisID:'y'},{label:'Avaliações',data:rows.map(x=>x.ratings),borderColor:palette[1],tension:.3,yAxisID:'y'},{label:'Taxa %',data:rows.map(x=>x.rate*100),borderColor:palette[3],tension:.3,yAxisID:'y1'},{label:'Caixinha',data:rows.map(x=>x.box),borderColor:palette[2],borderDash:[6,4],tension:.3,yAxisID:'y'}]},options:{responsive:true,maintainAspectRatio:false,interaction:{mode:'index',intersect:false},plugins:{legend:{position:'bottom'}},scales:{y:{beginAtZero:true,position:'left'},y1:{beginAtZero:true,position:'right',grid:{drawOnChartArea:false},ticks:{callback:v=>v+'%'}}}}});
}
export function drawBox(month){
 const agents=[...month.agents].sort((a,b)=>b.box-a.box);
 reset('box');charts.box=new Chart(document.getElementById('boxChart'),{type:'bar',data:{labels:agents.map(x=>short(x.agent)),datasets:[{label:'Pontuação',data:agents.map(x=>x.box),backgroundColor:agents.map((_,i)=>palette[i%palette.length]),borderRadius:9}]},options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',plugins:{legend:{display:false}},scales:{x:{beginAtZero:true},y:{grid:{display:false}}}}});
}
export function drawBonus(data,all){
 const agents=[...data.quarterly.agents].sort((a,b)=>b.bonus-a.bonus);
 reset('bonus');charts.bonus=new Chart(document.getElementById('bonusChart'),{type:'bar',data:{labels:agents.map(x=>short(x.agent)),datasets:[{label:'Bonificação',data:agents.map(x=>x.bonus),backgroundColor:palette[1],borderRadius:9}]},options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>'R$ '+c.raw.toLocaleString('pt-BR',{minimumFractionDigits:2})}}},scales:{x:{beginAtZero:true,ticks:{callback:v=>'R$ '+v}},y:{grid:{display:false}}}}});
 const sorted=[...all].sort((a,b)=>a.period.id.localeCompare(b.period.id));
 reset('bonusHistory');charts.bonusHistory=new Chart(document.getElementById('bonusHistoryChart'),{type:'line',data:{labels:sorted.map(x=>x.period.label),datasets:[{label:'Total distribuído',data:sorted.map(x=>x.quarterly.total.bonus),borderColor:palette[3],backgroundColor:'#f59e0b22',fill:true,tension:.35}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'bottom'},tooltip:{callbacks:{label:c=>'R$ '+c.raw.toLocaleString('pt-BR',{minimumFractionDigits:2})}}},scales:{y:{beginAtZero:true,ticks:{callback:v=>'R$ '+v}},x:{grid:{display:false}}}}});
}
export function drawComparison(all){
 const sorted=[...all].sort((a,b)=>a.period.id.localeCompare(b.period.id));
 reset('comparison');charts.comparison=new Chart(document.getElementById('comparisonChart'),{type:'bar',data:{labels:sorted.map(x=>x.period.label),datasets:[{label:'Atendimentos',data:sorted.map(x=>x.quarterly.total.calls),backgroundColor:palette[0],borderRadius:8},{label:'Avaliações',data:sorted.map(x=>x.quarterly.total.ratings),backgroundColor:palette[1],borderRadius:8}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'bottom'}},scales:{y:{beginAtZero:true},x:{grid:{display:false}}}}});
}
