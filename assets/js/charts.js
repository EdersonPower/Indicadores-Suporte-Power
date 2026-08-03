const charts={};
const palette=['#38bdf8','#22c55e','#a78bfa','#f59e0b','#f472b6','#14b8a6','#fb7185'];
const grid='#20334a',text='#9fb2c8';
Chart.defaults.color=text;Chart.defaults.borderColor=grid;Chart.defaults.font.family='Inter';
function reset(id){if(charts[id])charts[id].destroy()}
export function drawDashboard(data){
 const months=data.monthly;
 reset('monthly');charts.monthly=new Chart(document.getElementById('monthlyChart'),{type:'line',data:{labels:months.map(x=>x.monthLabel),datasets:[{label:'Atendimentos',data:months.map(x=>x.total.calls),borderColor:palette[0],backgroundColor:'#38bdf822',fill:true,tension:.35},{label:'Avaliações',data:months.map(x=>x.total.ratings),borderColor:palette[1],backgroundColor:'#22c55e22',fill:true,tension:.35}]},options:{responsive:true,maintainAspectRatio:false,interaction:{mode:'index',intersect:false},plugins:{legend:{position:'bottom'}},scales:{x:{grid:{display:false}},y:{beginAtZero:true}}}});
 const agents=[...data.quarterly.agents].sort((a,b)=>b.calls-a.calls);
 reset('ranking');charts.ranking=new Chart(document.getElementById('rankingChart'),{type:'bar',data:{labels:agents.map(x=>x.agent.split(' - ')[1]||x.agent),datasets:[{label:'Atendimentos',data:agents.map(x=>x.calls),backgroundColor:palette[0],borderRadius:8}]},options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',plugins:{legend:{display:false}},scales:{x:{beginAtZero:true},y:{grid:{display:false}}}}});
 reset('share');charts.share=new Chart(document.getElementById('shareChart'),{type:'doughnut',data:{labels:agents.map(x=>x.agent.split(' - ')[1]||x.agent),datasets:[{data:agents.map(x=>x.ratings),backgroundColor:palette,borderColor:'#0c1929',borderWidth:4}]},options:{responsive:true,maintainAspectRatio:false,cutout:'68%',plugins:{legend:{position:'bottom'}}}});
 reset('score');charts.score=new Chart(document.getElementById('scoreChart'),{type:'bar',data:{labels:agents.map(x=>x.agent.split(' - ')[1]||x.agent),datasets:[{label:'Nota média',data:agents.map(x=>x.score),backgroundColor:palette[2],borderRadius:8}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{y:{min:4,max:5},x:{grid:{display:false}}}}});
}
export function drawAgent(data,agent){
 const rows=data.monthly.map(m=>m.agents.find(a=>a.agent===agent)||{calls:0,ratings:0,rate:0,score:0});
 reset('agent');charts.agent=new Chart(document.getElementById('agentChart'),{type:'line',data:{labels:data.monthly.map(m=>m.monthLabel),datasets:[{label:'Atendimentos',data:rows.map(x=>x.calls),borderColor:palette[0],tension:.3,yAxisID:'y'},{label:'Avaliações',data:rows.map(x=>x.ratings),borderColor:palette[1],tension:.3,yAxisID:'y'},{label:'Taxa %',data:rows.map(x=>x.rate*100),borderColor:palette[3],tension:.3,yAxisID:'y1'}]},options:{responsive:true,maintainAspectRatio:false,interaction:{mode:'index',intersect:false},plugins:{legend:{position:'bottom'}},scales:{y:{beginAtZero:true,position:'left'},y1:{beginAtZero:true,position:'right',grid:{drawOnChartArea:false},ticks:{callback:v=>v+'%'}}}}});
}
export function drawComparison(all){
 const sorted=[...all].sort((a,b)=>a.period.id.localeCompare(b.period.id));
 reset('comparison');charts.comparison=new Chart(document.getElementById('comparisonChart'),{type:'bar',data:{labels:sorted.map(x=>x.period.label),datasets:[{label:'Atendimentos',data:sorted.map(x=>x.quarterly.total.calls),backgroundColor:palette[0],borderRadius:8},{label:'Avaliações',data:sorted.map(x=>x.quarterly.total.ratings),backgroundColor:palette[1],borderRadius:8}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'bottom'}},scales:{y:{beginAtZero:true},x:{grid:{display:false}}}}});
}
