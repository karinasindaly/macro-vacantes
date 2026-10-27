// =============================================================
// ── NAVEGACIÓN TABS ──
// =============================================================
function switchTab(name, btn) {
  document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
  document.getElementById('panel-' + name).classList.add('active');
  btn.classList.add('active');
}

// =============================================================
// ── SEMÁFORO: ESTADO GLOBAL ──
// =============================================================
let allRows = [];
let activeZona = '';
const PAGE = 25;
const comentarios = {};
if(typeof ChartDataLabels!=='undefined'){Chart.register(ChartDataLabels);Chart.defaults.set('plugins.datalabels',{display:false});}
// ── Comparativo Semanal: estado global ──
let comparativoSnapshots = [];
let compActiveZona = '';
let compChartSemaforo = null;
let compChartPieSemaforo = null;
let compChartDonaZona = null;
let compChartZonaGroup = null;
let compChartDiasTrend = null;
let compAutoLabel = 1;
const state = {
  verde:{open:false,filtered:[],page:0,q:''},
  amarillo:{open:false,filtered:[],page:0,q:''},
  rojo:{open:false,filtered:[],page:0,q:''},
};

// drag & drop
function onDragOver(e){e.preventDefault();document.getElementById('upload-section').classList.add('drag-over')}
function onDragLeave(){document.getElementById('upload-section').classList.remove('drag-over')}
function onDrop(e){e.preventDefault();onDragLeave();const f=e.dataTransfer.files[0];if(f)handleFile(f)}

function resetAllData(){
  resetAll();
  detResetAll();
}
function resetAll(){
  allRows=[];activeZona='';
  Object.keys(comentarios).forEach(k=>delete comentarios[k]);
  Object.keys(state).forEach(k=>{state[k].open=false;state[k].filtered=[];state[k].page=0;state[k].q=''});
  document.getElementById('upload-section').style.display='block';
  ['file-badge','dashboard','error-box'].forEach(id=>document.getElementById(id).style.display='none');
  document.getElementById('analytics-bar').style.display='none';
  analyticsOpen=false;
  Object.keys(chartInstances).forEach(k=>{if(chartInstances[k])chartInstances[k].destroy();delete chartInstances[k];});
  document.getElementById('fileInput').value='';
  ['verde','amarillo','rojo'].forEach(c=>{
    document.getElementById('acord-'+c).classList.remove('open');
    document.getElementById('kpi-'+c).classList.remove('open');
  });
  document.querySelectorAll('.zbtn-dyn').forEach(b=>b.remove());
  document.getElementById('zc-todas').textContent='0';
  document.getElementById('zbtn-todas').classList.add('z-active');
}

// =============================================================
// ── COMPARATIVO SEMANAL ──
// =============================================================
function toggleComparativo(){
  const content=document.getElementById('comp-content');
  const arrow=document.getElementById('comp-arrow');
  const isOpen=content.style.maxHeight&&content.style.maxHeight!=='0px';
  if(isOpen){content.style.maxHeight='0';content.style.opacity='0';arrow.style.transform='rotate(0deg)';}
  else{content.style.maxHeight='2600px';content.style.opacity='1';arrow.style.transform='rotate(180deg)';}
}

function compDragOver(e){e.preventDefault();document.getElementById('comp-upload').classList.add('drag-over')}
function compDragLeave(){document.getElementById('comp-upload').classList.remove('drag-over')}
function compDrop(e){e.preventDefault();compDragLeave();handleCompFiles(e.dataTransfer.files)}

function compShowError(msg){const el=document.getElementById('comp-error');el.textContent='Aviso: '+msg;el.style.display='block'}
function compClearError(){document.getElementById('comp-error').style.display='none'}

function handleCompFiles(fileList){
  compClearError();
  const files=Array.from(fileList||[]);
  if(!files.length)return;
  const espacioDisponible=5-comparativoSnapshots.length;
  if(espacioDisponible<=0){compShowError('Ya tienes 5 reportes cargados (el máximo). Elimina uno para agregar otro.');document.getElementById('compFileInput').value='';return;}
  const aProcesar=files.slice(0,espacioDisponible);
  if(files.length>espacioDisponible){compShowError('Solo se admite un máximo de 5 reportes. Se cargaron los primeros '+espacioDisponible+'.');}
  aProcesar.forEach(f=>processCompFile(f));
  document.getElementById('compFileInput').value='';
}

function processCompFile(file){
  const reader=new FileReader();
  reader.onload=function(e){
    try{
      const wb=XLSX.read(e.target.result,{type:'array',cellFormula:false,cellNF:false,raw:false});
      const result=parseVacantesWorkbook(wb);
      if(result.error){compShowError(file.name+': '+result.error);return;}
      if(!result.rows.length){compShowError(file.name+': no se encontraron vacantes.');return;}
      const baseName=file.name.replace(/\.[^/.]+$/,'');
      comparativoSnapshots.push({
        id:'snap_'+Date.now()+'_'+Math.random().toString(36).slice(2,7),
        label:baseName,
        fileName:file.name,
        rows:result.rows,
        totalPos:result.totalPosiciones,
        totalOcup:result.totalOcupadas,
        allPos:result.allPos||new Set(),
        posZonaMap:result.posZonaMap||new Map(),
      });
      sortCompSnapshots();
      renderComparativo();
    }catch(err){compShowError(file.name+': error al procesar — '+err.message)}
  };
  reader.readAsArrayBuffer(file);
}

function compPeriodKey(label){
  const s=String(label||'');
  const mYear=s.match(/(\d{4})/);
  const mP=s.match(/P\s*-?\s*(\d{1,3})/i);
  if(!mP)return null;
  const year=mYear?parseInt(mYear[1],10):0;
  return year*1000+parseInt(mP[1],10);
}
function sortCompSnapshots(){
  // Ordena por número de periodo (Pxx-aaaa) para que "Movimiento entre semanas"
  // siempre compare el periodo anterior contra el más reciente, sin importar
  // el orden en que se hayan subido los archivos.
  const withKey=comparativoSnapshots.map((s,i)=>({s,i,k:compPeriodKey(s.label)}));
  if(withKey.every(x=>x.k===null))return; // no se pudo detectar periodo en ningún label, se respeta el orden de carga
  withKey.sort((a,b)=>{
    if(a.k===null&&b.k===null)return a.i-b.i;
    if(a.k===null)return 1;
    if(b.k===null)return -1;
    return a.k-b.k;
  });
  comparativoSnapshots=withKey.map(x=>x.s);
}

function removeCompSnapshot(id){
  comparativoSnapshots=comparativoSnapshots.filter(s=>s.id!==id);
  renderComparativo();
}

function renameCompSnapshot(id,val){
  const snap=comparativoSnapshots.find(s=>s.id===id);
  if(snap)snap.label=val||snap.label;
  sortCompSnapshots();
  renderComparativo();
}

function compSnapCounts(snap){
  const nV=snap.rows.filter(r=>r.estatus==='Verde').length;
  const nA=snap.rows.filter(r=>r.estatus==='Amarillo').length;
  const nR=snap.rows.filter(r=>r.estatus==='Rojo').length;
  return{nV,nA,nR,total:snap.rows.length};
}

function renderComparativo(){
  // chips
  const chips=document.getElementById('comp-chips');
  chips.innerHTML=comparativoSnapshots.map(s=>{
    const c=compSnapCounts(s);
    return`<div class="comp-chip">
      <input type="text" value="${escHtml(s.label)}" onchange="renameCompSnapshot('${s.id}',this.value)" title="${escHtml(s.fileName)}">
      <span class="comp-chip-count">${c.total} vac.</span>
      <button onclick="removeCompSnapshot('${s.id}')" title="Quitar">✕</button>
    </div>`;
  }).join('');

  const results=document.getElementById('comp-results');
  if(comparativoSnapshots.length<2){
    results.style.display='none';
    if(comparativoSnapshots.length===1){compShowError('Carga al menos 2 reportes para generar el comparativo (hasta 5).');}
    else{compClearError();}
    return;
  }
  compClearError();
  results.style.display='block';

  // zonas presentes en todos los snapshots (para el filtro y el chart agrupado)
  const zonas=new Set();
  comparativoSnapshots.forEach(s=>s.rows.forEach(r=>{if(r.zona)zonas.add(r.zona)}));
  const orderZ=['Norte','Sur','Centro','Corporativo','Polanco'];
  const zonaList=[...orderZ.filter(z=>zonas.has(z)),...[...zonas].sort().filter(z=>!orderZ.includes(z))];
  if(compActiveZona&&!zonaList.includes(compActiveZona))compActiveZona='';

  // ── Barra de filtro por complejo ──
  const zbar=document.getElementById('comp-zona-bar');
  zbar.innerHTML='<span class="zona-label">Complejo</span>'+
    `<button class="zbtn${compActiveZona===''?' z-active':''}" onclick="setCompZona('',this)">General</button>`+
    zonaList.map(z=>`<button class="zbtn${compActiveZona===z?' z-active':''}" onclick="setCompZona('${escHtml(z)}',this)">${escHtml(z)}</button>`).join('');
  document.getElementById('comp-zona-tag').textContent=compActiveZona?`— Complejo: ${compActiveZona}`:'';

  // ── Snapshots filtrados por el complejo activo (o completos si es "General") ──
  const filteredSnaps=comparativoSnapshots.map(s=>({...s,rows:compActiveZona?s.rows.filter(r=>r.zona===compActiveZona):s.rows}));
  const zonaListForChart=compActiveZona?[compActiveZona]:zonaList;

  const labels=filteredSnaps.map(s=>s.label);
  const counts=filteredSnaps.map(compSnapCounts);
  const lastSnap=filteredSnaps[filteredSnaps.length-1];
  const lastCounts=counts[counts.length-1];

  document.getElementById('comp-last-label-1').textContent=lastSnap.label;
  document.getElementById('comp-last-label-2').textContent=lastSnap.label;
  document.getElementById('comp-last-label-3').textContent=lastSnap.label;

  // Paleta fija por Complejo (misma que ya se usaba en "Vacantes por Complejo
  // — Tendencia"): Norte azul, Sur rojo, Centro gris, Corporativo amarillo,
  // Polanco verde. "General" (sin complejo activo) usa el azul Pilgrim's.
  function colorComplejo(z){
    const n=norm(z);
    if(n==='norte')return '#003087';
    if(n==='sur')return '#C8102E';
    if(n==='centro')return '#5A6070';
    if(n==='corporativo')return '#D97706';
    if(n==='polanco')return '#047857';
    return '#9F1239';
  }
  const colorTotal=compActiveZona?colorComplejo(compActiveZona):'#003087';
  const round1=n=>Math.round(n*10)/10;

  // Promedio de cada categoría de semáforo a través de TODOS los reportes
  // cargados (no solo el último), para las leyendas que el usuario pidió
  // en formato "promedio de los reportes" en vez de "última semana".
  const avgCounts={
    nR:counts.reduce((a,c)=>a+c.nR,0)/counts.length,
    nA:counts.reduce((a,c)=>a+c.nA,0)/counts.length,
    nV:counts.reduce((a,c)=>a+c.nV,0)/counts.length,
    total:counts.reduce((a,c)=>a+c.total,0)/counts.length,
  };

  // ── Chart 1: combo — barras apiladas por semáforo (Rojo/Amarillo/Verde,
  // siempre el mismo color porque representan el estatus) + línea de Total,
  // que SÍ cambia de color según el Complejo activo para distinguir de un
  // vistazo qué complejo se está viendo (misma paleta que la gráfica de
  // tendencia por complejo) ──
  if(compChartSemaforo)compChartSemaforo.destroy();
  compChartSemaforo=new Chart(document.getElementById('chart-comp-semaforo'),{
    data:{labels,datasets:[
      {type:'bar',label:'Vencidas',data:counts.map(c=>c.nR),backgroundColor:'#C8102E',stack:'s',order:2},
      {type:'bar',label:'Por Vencer',data:counts.map(c=>c.nA),backgroundColor:'#D97706',stack:'s',order:2},
      {type:'bar',label:'Vigentes',data:counts.map(c=>c.nV),backgroundColor:'#047857',stack:'s',order:2},
      {type:'line',label:'Total ('+(compActiveZona||'General')+')',data:counts.map(c=>c.total),borderColor:colorTotal,backgroundColor:colorTotal,fill:false,tension:.3,pointRadius:4,pointBackgroundColor:colorTotal,borderWidth:2.5,order:1,datalabels:{display:true,align:'top',anchor:'end',offset:8,color:'#fff',backgroundColor:colorTotal,borderRadius:4,padding:{top:2,bottom:2,left:5,right:5},font:{size:10.5,weight:'700'}}},
    ]},
    options:{responsive:true,maintainAspectRatio:false,layout:{padding:{top:26}},scales:{x:{stacked:true,ticks:{font:{size:10.5},maxRotation:40,minRotation:0,autoSkip:false}},y:{stacked:true,beginAtZero:true,ticks:{precision:0},suggestedMax:Math.max(1,...counts.map(c=>c.total))*1.28}},plugins:{legend:{display:false},datalabels:{display:ctx=>ctx.dataset.type==='bar'&&ctx.dataset.data[ctx.dataIndex]>0,anchor:'center',align:'center',color:'#fff',font:{size:11,weight:'800'},formatter:v=>v}}}
  });
  buildSideStats('chart-comp-semaforo-stats',['Vencidas','Por Vencer','Vigentes','Total'],[round1(avgCounts.nR),round1(avgCounts.nA),round1(avgCounts.nV),round1(avgCounts.total)],['#C8102E','#D97706','#047857',colorTotal],avgCounts.total);

  // ── Chart 3: dona — distribución de POSICIONES (plantilla total: ocupadas
  // + vacantes, no solo vacantes) por complejo en la última semana. Usa
  // posZonaMap (trae zona de TODAS las posiciones, no solo las de snap.rows
  // que son únicamente vacantes) y respeta el filtro de Complejo activo
  // igual que ya lo hace "Movimiento entre semanas" ──
  const posZonaMapLast=lastSnap.posZonaMap||new Map();
  const posZonaCountsLast={};
  posZonaMapLast.forEach((zona,posicion)=>{
    if(compActiveZona&&zona!==compActiveZona)return;
    const z=zona||'Otro';
    posZonaCountsLast[z]=(posZonaCountsLast[z]||0)+1;
  });
  const pzLabelsLast=Object.keys(posZonaCountsLast);
  const pzDataLast=pzLabelsLast.map(z=>posZonaCountsLast[z]);
  const pzColorsLast=pzLabelsLast.map(z=>colorComplejo(z));
  const pzTotalLast=pzDataLast.reduce((a,b)=>a+b,0);
  if(compChartPieSemaforo)compChartPieSemaforo.destroy();
  compChartPieSemaforo=new Chart(document.getElementById('chart-comp-pie-semaforo'),{
    type:'doughnut',
    data:{labels:pzLabelsLast,datasets:[{data:pzDataLast,backgroundColor:pzColorsLast,borderWidth:2,borderColor:'#fff',hoverOffset:6}]},
    options:{responsive:true,maintainAspectRatio:false,cutout:'55%',layout:{padding:24},plugins:{legend:{display:false},datalabels:{display:false},tooltip:{callbacks:{label:ctx=>` ${ctx.label}: ${ctx.parsed} (${pzTotalLast?Math.round(ctx.parsed/pzTotalLast*100):0}%)`}}}}
  });
  buildSideStats('chart-comp-pie-semaforo-stats',pzLabelsLast,pzDataLast,pzColorsLast,pzTotalLast);

  // ── Chart 4: dona — distribución por complejo en la última semana ──
  const zonaCountsLast={};
  lastSnap.rows.forEach(r=>{const z=r.zona||'Otro';zonaCountsLast[z]=(zonaCountsLast[z]||0)+1;});
  const zLabelsLast=Object.keys(zonaCountsLast);
  const zDataLast=zLabelsLast.map(z=>zonaCountsLast[z]);
  const zColorsLast=makeColors(zLabelsLast.length);
  if(compChartDonaZona)compChartDonaZona.destroy();
  compChartDonaZona=new Chart(document.getElementById('chart-comp-dona-zona'),{
    type:'doughnut',
    data:{labels:zLabelsLast,datasets:[{data:zDataLast,backgroundColor:zColorsLast,borderWidth:2,borderColor:'#fff',hoverOffset:6}]},
    options:{responsive:true,maintainAspectRatio:false,cutout:'55%',layout:{padding:24},plugins:{legend:{display:false},datalabels:{display:false},tooltip:{callbacks:{label:ctx=>` ${ctx.label}: ${ctx.parsed} (${lastSnap.rows.length?Math.round(ctx.parsed/lastSnap.rows.length*100):0}%)`}}}}
  });
  buildSideStats('chart-comp-dona-zona-stats',zLabelsLast,zDataLast,zColorsLast,lastSnap.rows.length);

  // ── Chart 5: línea de tendencia — vacantes por complejo, todas las semanas (una línea por complejo, más clara que barras agrupadas) ──
  if(compChartZonaGroup)compChartZonaGroup.destroy();
  const dsZona=zonaListForChart.map((z,i)=>({label:z,data:filteredSnaps.map(s=>s.rows.filter(r=>r.zona===z).length),borderColor:colorComplejo(z),backgroundColor:colorComplejo(z),fill:false,tension:.3,pointRadius:4,borderWidth:2}));
  // Solo se muestra la etiqueta en el ÚLTIMO punto de cada línea (no en
  // cada semana) para no saturar. Cuando dos complejos consecutivos (en
  // valor) están muy cerca — como Norte y Centro — uno se pone ARRIBA de
  // su línea y el otro ABAJO, para garantizar que no se encimen sin
  // importar qué tan parecido sea el valor (un simple empujón hacia arriba
  // no bastaba cuando la diferencia era muy chica).
  const yMaxZona=Math.max(1,...dsZona.flatMap(d=>d.data))*1.3;
  const lastVals=dsZona.map(d=>d.data[d.data.length-1]);
  const ordenPorValor=dsZona.map((_,i)=>i).sort((a,b)=>lastVals[b]-lastVals[a]);
  const alignPorDataset={};
  let valorPrevio=null,abajo=false;
  ordenPorValor.forEach(idx=>{
    const v=lastVals[idx];
    abajo=(valorPrevio!==null&&Math.abs(valorPrevio-v)<yMaxZona*0.09)?!abajo:false;
    alignPorDataset[idx]=abajo?'bottom':'top';
    valorPrevio=v;
  });
  compChartZonaGroup=new Chart(document.getElementById('chart-comp-zona-group'),{
    type:'line',
    data:{labels,datasets:dsZona},
    options:{responsive:true,maintainAspectRatio:false,layout:{padding:{top:26,bottom:20}},scales:{x:{ticks:{font:{size:10.5},maxRotation:40,minRotation:0,autoSkip:false}},y:{beginAtZero:true,ticks:{precision:0},suggestedMax:yMaxZona}},plugins:{legend:{display:false},datalabels:{display:true,align:ctx=>alignPorDataset[ctx.datasetIndex]||'top',anchor:'center',offset:10,color:'#fff',backgroundColor:ctx=>ctx.dataset.borderColor,borderRadius:4,padding:{top:2,bottom:2,left:5,right:5},font:{size:10,weight:'700'}}}}
  });
  // Promedio de vacantes por complejo a través de todos los reportes
  // cargados (reutiliza los datos ya calculados por línea en dsZona).
  const zonaAvgList=dsZona.map(d=>d.data.reduce((a,b)=>a+b,0)/d.data.length);
  const zonaAvgTotal=zonaAvgList.reduce((a,b)=>a+b,0);
  buildSideStats('chart-comp-zona-group-stats',zonaListForChart,zonaAvgList.map(round1),zonaListForChart.map(colorComplejo),zonaAvgTotal);


  // ── Chart 6: línea — promedio de días sin ocupar por semana ──
  const avgDias=filteredSnaps.map(s=>{
    const tot=s.rows.reduce((sum,r)=>sum+(r.dias||0),0);
    return s.rows.length?Math.round(tot/s.rows.length):0;
  });
  if(compChartDiasTrend)compChartDiasTrend.destroy();
  compChartDiasTrend=new Chart(document.getElementById('chart-comp-dias-trend'),{
    type:'line',
    data:{labels,datasets:[{label:'Promedio de días sin ocupar',data:avgDias,borderColor:'#C8102E',backgroundColor:'rgba(200,16,46,.12)',fill:true,tension:.3,pointRadius:4,pointBackgroundColor:'#C8102E'}]},
    options:{responsive:true,maintainAspectRatio:false,layout:{padding:{top:20}},scales:{x:{ticks:{font:{size:10.5},maxRotation:40,minRotation:0,autoSkip:false}},y:{beginAtZero:true,ticks:{precision:0},suggestedMax:Math.max(1,...avgDias)*1.2}},plugins:{legend:{display:false},datalabels:{display:true,align:'top',anchor:'end',color:'#C8102E',font:{size:11,weight:'700'}}}}
  });

  // ── Tabla comparativa por zona ──
  const thead=document.getElementById('comp-table-head');
  thead.innerHTML='<th>Complejo</th>'+labels.map((l,i)=>`<th>${l}${i>0?' <span style="opacity:.7;font-weight:600">(Δ)</span>':''}</th>`).join('');

  const tbody=document.getElementById('comp-table-body');
  let rowsHtml='';
  zonaListForChart.forEach(z=>{
    const vals=filteredSnaps.map(s=>s.rows.filter(r=>r.zona===z).length);
    rowsHtml+='<tr><td>'+z+'</td>'+vals.map((v,i)=>{
      if(i===0)return`<td>${v}</td>`;
      const d=v-vals[i-1];
      const dTxt=d===0?'sin cambio':(d>0?'+'+d:d);
      const color=d>0?'#C8102E':(d<0?'#047857':'#8a8f9b');
      return`<td>${v} <span style="font-size:10.5px;color:${color}">(${dTxt})</span></td>`;
    }).join('')+'</tr>';
  });
  tbody.innerHTML=rowsHtml;

  // ── Movimiento entre las últimas dos semanas cargadas ──
  const n=filteredSnaps.length;
  const prev=filteredSnaps[n-2],curr=filteredSnaps[n-1];
  document.getElementById('comp-mov-title').textContent=`Movimiento entre "${prev.label}" y "${curr.label}"`;
  const prevPos=new Set(prev.rows.map(r=>r.posicion));
  const currPos=new Set(curr.rows.map(r=>r.posicion));
  // Universo de posiciones contra el que se compara "sigue existiendo" para
  // distinguir cubiertas de eliminadas. Si hay un Complejo activo, se limita
  // a las posiciones de ESE complejo (usando posZonaMap, que sí trae zona
  // para ocupadas y vacantes), en vez de todas las posiciones de la empresa
  // — así una posición que cambió de complejo no se cuenta como "cubierta".
  const currAllPosRaw=curr.allPos||new Set();
  const currAllPos=compActiveZona&&curr.posZonaMap
    ?new Set([...currAllPosRaw].filter(p=>curr.posZonaMap.get(p)===compActiveZona))
    :currAllPosRaw;
  const yaNoVacantes=prev.rows.filter(r=>!currPos.has(r.posicion));
  // cubiertas: la posición sigue existiendo en el archivo actual (ya no vacante, se ocupó)
  // eliminadas: la posición ya no existe en el archivo actual (se dio de baja del reporte)
  const cubiertas=yaNoVacantes.filter(r=>currAllPos.has(r.posicion));
  const eliminadas=yaNoVacantes.filter(r=>!currAllPos.has(r.posicion));
  // nuevas: SOLO posiciones de nueva creación — la posición no existía en
  // absoluto (ni ocupada ni vacante) en el reporte anterior. Antes se
  // consideraba "nueva" cualquier vacante que no estuviera vacante la
  // semana pasada, lo cual mezclaba nueva creación real con posiciones que
  // ya existían ocupadas y solo rotaron a vacante (alguien salió) — eso NO
  // es nueva creación, así que ya no se cuenta aquí.
  const prevAllPosRaw=prev.allPos||new Set();
  const prevAllPos=compActiveZona&&prev.posZonaMap
    ?new Set([...prevAllPosRaw].filter(p=>prev.posZonaMap.get(p)===compActiveZona))
    :prevAllPosRaw;
  const nuevas=curr.rows.filter(r=>!prevAllPos.has(r.posicion));
  const elResueltas=document.getElementById('comp-resueltas');
  const elEliminadas=document.getElementById('comp-eliminadas');
  const elNuevas=document.getElementById('comp-nuevas');
  function movItemHtml(r){
    const clasifNorm=String(r.clasificacion||'').trim().toUpperCase();
    const tipo=clasifNorm.includes('EVENTUAL')?'Eventual':(clasifNorm==='—'||clasifNorm==='')?'Sin clasificar':'Planta';
    const tipoColor=tipo==='Planta'?'#1D4ED8':tipo==='Eventual'?'#92400E':'#5A6070';
    const tipoBg=tipo==='Planta'?'#EEF2FF':tipo==='Eventual'?'#FFF7ED':'#F1F3F5';
    return `<div style="padding:4px 0;border-bottom:1px dashed #E2E8F0;color:#1A1A2E"><strong>${escHtml(r.posicion)}</strong> — ${escHtml(r.descripcion)} <span style="color:#8a8f9b">(${escHtml(r.zona||'—')})</span> <span style="background:${tipoBg};color:${tipoColor};border-radius:10px;padding:1px 7px;font-size:10px;font-weight:800;margin-left:4px">${tipo}</span></div>`;
  }
  elResueltas.innerHTML=cubiertas.length?cubiertas.map(movItemHtml).join(''):'<div style="color:#5A6070;font-style:italic">Sin vacantes cubiertas en este periodo</div>';
  elEliminadas.innerHTML=eliminadas.length?eliminadas.map(movItemHtml).join(''):'<div style="color:#5A6070;font-style:italic">Sin posiciones eliminadas en este periodo</div>';
  elNuevas.innerHTML=nuevas.length?nuevas.map(movItemHtml).join(''):'<div style="color:#5A6070;font-style:italic">Sin posiciones de nueva creación en este periodo</div>';
}

// =============================================================
// ── EXPORTAR PDF COMPARATIVO SEMANAL ──
// =============================================================
async function exportarComparativoPDF(){
  if(comparativoSnapshots.length<2){alert('Carga al menos 2 reportes para generar el comparativo antes de exportar.');return;}
  const btn=document.getElementById('btn-comp-export-pdf');
  btn.disabled=true;btn.textContent='⏳ Generando...';
  try{
    const hoy=new Date().toLocaleDateString('es-MX',{day:'2-digit',month:'long',year:'numeric'});
    const semanas=comparativoSnapshots.map(s=>s.label).join(' → ');

    function chartImg(id,titulo){
      const canvas=document.getElementById(id);
      if(!canvas||!canvas.width||!canvas.height)return `<div style="background:#F8FAFC;border:1px solid #DDE1EA;border-radius:10px;padding:12px;margin-bottom:16px;font-size:10px;color:#aaa;text-align:center">${titulo}<br>(gráfica no disponible)</div>`;
      let dataUrl='';
      try{dataUrl=canvas.toDataURL('image/png',1.0);}catch(e){return `<div style="background:#F8FAFC;border:1px solid #DDE1EA;border-radius:10px;padding:12px;margin-bottom:16px;font-size:10px;color:#aaa;text-align:center">${titulo}<br>(gráfica no disponible)</div>`;}
      const statsEl=document.getElementById(id+'-stats');
      const statsHtml=statsEl&&statsEl.innerHTML.trim()?`<div style="width:150px;flex-shrink:0;display:flex;flex-direction:column;gap:6px;justify-content:center">${statsEl.innerHTML}</div>`:'';
      return `<div style="background:#F8FAFC;border:1px solid #DDE1EA;border-radius:10px;padding:12px;margin-bottom:16px;page-break-inside:avoid">
        <div style="font-size:11px;font-weight:800;text-transform:uppercase;letter-spacing:.05em;color:#003087;margin-bottom:8px">${titulo}</div>
        <div style="display:flex;gap:12px;align-items:center">
          <img src="${dataUrl}" style="flex:1;min-width:0;max-height:300px;object-fit:contain">
          ${statsHtml}
        </div>
      </div>`;
    }

    const chartsHtml=
      chartImg('chart-comp-semaforo','Vacantes por Semáforo y Total — Tendencia')+
      chartImg('chart-comp-zona-group','Vacantes por Complejo — Tendencia')+
      chartImg('chart-comp-dias-trend','Promedio de Días Sin Ocupar — Tendencia')+
      `<div style="display:flex;gap:16px">
        <div style="flex:1">${chartImg('chart-comp-pie-semaforo','Distribución de Posiciones por Complejo — Última Semana')}</div>
        <div style="flex:1">${chartImg('chart-comp-dona-zona','Distribución de Vacantes por Complejo — Última Semana')}</div>
      </div>`;

    const tableHead=document.getElementById('comp-table-head').innerHTML;
    const tableBody=document.getElementById('comp-table-body').innerHTML;
    const movTitle=document.getElementById('comp-mov-title').textContent;
    const resueltasHtml=document.getElementById('comp-resueltas').innerHTML;
    const eliminadasHtml=document.getElementById('comp-eliminadas').innerHTML;
    const nuevasHtml=document.getElementById('comp-nuevas').innerHTML;

    const htmlContent=`<div style="font-family:'Abadi','Century Gothic','Trebuchet MS',Arial,sans-serif;color:#111;font-size:12px;padding:20px;background:#fff">
      <div style="background:linear-gradient(135deg,#003087,#1B4F9C,#C8102E);padding:18px 28px;border-radius:10px 10px 0 0"><div style="color:#fff;font-size:22px;font-weight:900">PILGRIMS</div><div style="color:rgba(255,255,255,.8);font-size:11px;text-transform:uppercase">Atracción de Talento</div></div>
      <div style="border:2px solid #003087;border-top:none;border-radius:0 0 10px 10px;padding:16px 22px;margin-bottom:20px"><div style="font-size:17px;font-weight:900;color:#003087;text-transform:uppercase;margin-bottom:4px">Comparativo Semanal de Vacantes</div><div style="font-size:11px;color:#777"><strong>Fecha:</strong> ${hoy} &nbsp; <strong>Complejo:</strong> ${escHtml(compActiveZona||'General (todos)')} &nbsp; <strong>Semanas comparadas:</strong> ${escHtml(semanas)}</div></div>

      <div style="font-size:11px;font-weight:900;text-transform:uppercase;letter-spacing:.06em;color:#5A6070;border-bottom:2px solid #DDE1EA;padding-bottom:6px;margin-bottom:14px">Tendencia semana a semana</div>
      ${chartsHtml}

      <div style="font-size:11px;font-weight:900;text-transform:uppercase;letter-spacing:.06em;color:#5A6070;border-bottom:2px solid #DDE1EA;padding-bottom:6px;margin:20px 0 14px">Tabla comparativa por complejo</div>
      <table style="width:100%;border-collapse:collapse;font-size:11px;margin-bottom:20px">
        <thead><tr style="background:#003087;color:#fff">${tableHead.replace(/<th>/g,'<th style="padding:7px 9px;text-align:left;font-size:10px;text-transform:uppercase">')}</tr></thead>
        <tbody>${tableBody.replace(/<td>/g,'<td style="padding:6px 9px;border-bottom:1px solid #EEF0F5">')}</tbody>
      </table>

      <div style="font-size:11px;font-weight:900;text-transform:uppercase;letter-spacing:.06em;color:#5A6070;border-bottom:2px solid #DDE1EA;padding-bottom:6px;margin-bottom:14px">${escHtml(movTitle)}</div>
      <div style="display:flex;gap:14px;margin-bottom:10px">
        <div style="flex:1;background:#F0FDF4;border:1px solid #BBF7D0;border-radius:10px;padding:12px">
          <div style="font-size:11px;font-weight:800;color:#166534;margin-bottom:8px">Vacantes cubiertas (ya no aparecen, siguen en el archivo)</div>
          <div style="font-size:10.5px">${resueltasHtml}</div>
        </div>
        <div style="flex:1;background:#FFF1F2;border:1.5px solid #FECDD3;border-radius:10px;padding:12px">
          <div style="font-size:11px;font-weight:800;color:#9F1239;margin-bottom:8px">Posiciones eliminadas (ya no existen en el archivo)</div>
          <div style="font-size:10.5px">${eliminadasHtml}</div>
        </div>
        <div style="flex:1;background:#FFF7ED;border:1px solid #FDE68A;border-radius:10px;padding:12px">
          <div style="font-size:11px;font-weight:800;color:#92400E;margin-bottom:8px">Vacantes nuevas (de nueva creación)</div>
          <div style="font-size:10.5px">${nuevasHtml}</div>
        </div>
      </div>

      <div style="margin-top:20px;text-align:center;font-size:10px;color:#bbb;border-top:1px solid #e5e7eb;padding-top:10px">Pilgrims · Atracción de Talento · Documento confidencial · ${hoy}</div>
    </div>`;

    await abrirVentanaImpresion('Comparativo_Semanal_'+(compActiveZona||'General')+'_'+hoy, htmlContent);
  }catch(err){
    console.error('Error generando el comparativo semanal:',err);
    alert('Ocurrió un error al generar el reporte: '+err.message);
  }finally{
    btn.disabled=false;btn.innerHTML='Exportar PDF';
  }
}

function norm(s){return(s||'').toString().toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').trim()}

// ── Vigencia y semáforo ──
function getVig(g){const n=parseInt(g);if(isNaN(n))return 60;if(n>=1&&n<=9)return 60;if(n>=10&&n<=13)return 90;if(n>=14&&n<=19)return 120;return 60;}
function getEst(d,v){if(v===60){if(d<=45)return'Verde';if(d<=52)return'Amarillo';return'Rojo'}if(v===90){if(d<=75)return'Verde';if(d<=82)return'Amarillo';return'Rojo'}if(v===120){if(d<=105)return'Verde';if(d<=112)return'Amarillo';return'Rojo'}return null}
function esSindicalizado(g){return isNaN(parseInt(g))&&g!==''&&g!=null}

function normZona(v){
  const s=norm(v);
  if(!s)return'';
  if(s.includes('norte'))return'Norte';
  if(s.includes('sur'))return'Sur';
  if(s.includes('centro'))return'Centro';
  if(s.includes('corporativo')||s.includes('corp'))return'Corporativo';
  if(s.includes('polanco'))return'Polanco';
  if(s==='n'||s==='nor'||s==='nte')return'Norte';
  if(s==='s'||s==='so'||s==='sr')return'Sur';
  if(s==='c'||s==='cen'||s==='ctr')return'Centro';
  return String(v||'').trim();
}
// Reclasifica Corporativo/Polanco a partir de Localidad, Unidad, Centro de Costos o Descripción de Posición,
// ya que el reporte de SAP suele traer estas posiciones marcadas como "Centro" en la columna Complejo.
// Reclasifica una posición hacia Corporativo o Polanco por su "Descripción
// localidad" — PERO solo cuando el Complejo que ya trae el Excel es
// exactamente "Centro". Para cualquier otro valor (Norte, Sur, o lo que
// sea) se respeta tal cual y NUNCA se reclasifica, sin importar la
// localidad — el usuario pidió que el filtro de reclasificación aplique
// únicamente cuando el Excel dice "Centro".
function reclassifyZona(localidad,unidad,ccosto,posDesc,complejoCrudo){
  if(normZona(complejoCrudo||'')!=='Centro')return null;
  const nl=norm(localidad||'');
  if(nl==='oficinas polanco'||nl.includes('polanco'))return'Polanco';
  if(nl==='oficinas corporativas torre capita'||nl.includes('torre capita'))return'Corporativo';
  return null;
}

// ── Parser reutilizable de la hoja de vacantes (usado por la carga principal y por el Comparativo Semanal) ──
function parseVacantesWorkbook(wb){
  const justMap={};
  const justSheetName=wb.SheetNames.find(n=>norm(n).includes('justif'));
  if(justSheetName){
    const justRows=XLSX.utils.sheet_to_json(wb.Sheets[justSheetName],{header:1,defval:''});
    for(let i=1;i<justRows.length;i++){
      const jr=justRows[i];
      const posJ=jr[0];
      if(!posJ&&posJ!==0)continue;
      const diasAsig=parseFloat(jr[8]);
      const tipo=String(jr[5]||'').trim();
      const comentario=String(jr[4]||'').trim();
      justMap[String(posJ)]={diasAsig:isNaN(diasAsig)?9999:Math.round(diasAsig),tipo,comentario};
    }
  }
  const repName=wb.SheetNames.find(n=>norm(n).includes('rep pos')||norm(n).includes('rep posicion'))||wb.SheetNames.find(n=>norm(n).includes('planta'));
  if(!repName){return{error:'No se encontró la hoja "Rep Posiciones" en el archivo.'};}
  const mx=XLSX.utils.sheet_to_json(wb.Sheets[repName],{header:1,defval:''});
  if(!mx||!mx.length){return{error:'La hoja Rep Posiciones está vacía.'};}
  let hIdx=0;
  for(let i=0;i<Math.min(10,mx.length);i++){if(mx[i].filter(c=>c!==''&&c!=null).length>=4){hIdx=i;break}}
  const hr=mx[hIdx].map(c=>norm(c));
  function col(kws){for(let i=0;i<hr.length;i++){if(hr[i]&&kws.some(k=>hr[i].includes(k)))return i}return-1}
  function ci(preferred,fallback){return preferred>=0?preferred:(fallback<hr.length?fallback:-1)}
  const CI={
    pos:ci(col(['posicion','posición']),0),desc:ci(col(['descripcion posicion','descripción posición']),1),
    ccosto:ci(col(['descripcion centro de costos','descripción centro de costos']),5),
    nivel:ci(col(['nivel']),6),sub:ci(col(['subnivel']),7),grupo:ci(col(['descripcion grupo','descripción grupo']),13),
    local:ci(col(['descripcion localidad','descripción localidad']),19),estPos:ci(col(['estado']),20),
    dias:ci(col(['dias sin ocupar','días sin ocupar']),24),jrh:ci(col(['jefe rh','jefe de rh']),27),
    vig:ci(col(['vigencia']),28),estatExcel:ci(col(['estatus']),29),dv:ci(col(['dias vencidas','días vencidas']),30),
    razon:ci(col(['razon','razón']),31),just:ci(col(['tipo de just','tipo just']),32),
    clasif:ci(col(['clasific','clasificacion','clasificación']),33),zona:ci(col(['complejo']),34),
    unidad:ci(col(['descripcion unidad','descripción unidad']),3),
  };
  if(CI.dias<0){return{error:'No se encontró columna "Días sin ocupar".'};}
  let totalPosiciones=0,totalOcupadas=0;
  const rows=[];
  const allPos=new Set();
  const posZonaMap=new Map();
  for(let i=hIdx+1;i<mx.length;i++){
    const r=mx[i];
    if(!r||r.every(c=>c===''||c==null))continue;
    totalPosiciones++;
    const posStrRow=String(r[CI.pos]||'').trim();
    if(posStrRow)allPos.add(posStrRow);
    // Zona de esta fila (se guarda para TODAS las filas, ocupadas o vacantes,
    // para poder filtrar el universo de posiciones por Complejo al calcular
    // "Movimiento entre semanas" cuando hay un filtro de Complejo activo).
    if(posStrRow){
      const zonaRow=reclassifyZona(r[CI.local],r[CI.unidad],r[CI.ccosto],r[CI.desc],r[CI.zona])||normZona(String(r[CI.zona]||''));
      posZonaMap.set(posStrRow,zonaRow);
    }
    const estPosVal=norm(r[CI.estPos]||'');
    if(estPosVal==='ocupada'||estPosVal==='ocupado'||estPosVal==='cubierta'||estPosVal==='cubierto'){totalOcupadas++;continue;}
    if(estPosVal!==''&&estPosVal!=='vacante'&&estPosVal!=='vacantes'){totalOcupadas++;continue;}
    if(estPosVal===''&&(r[21]||r[22])){totalOcupadas++;continue;}
    const dias=isNaN(parseFloat(r[CI.dias]))?0:Math.round(parseFloat(r[CI.dias]));
    const posStr=String(r[CI.pos]||'').trim();
    let vigDias=9999;
    const jInfo=justMap[posStr];
    if(jInfo&&jInfo.diasAsig&&jInfo.diasAsig!==9999){vigDias=jInfo.diasAsig;}
    else{const vigRaw=r[CI.vig];const vigNum=parseFloat(String(vigRaw||'').replace(/[^0-9.]/g,''));if(!isNaN(vigNum)&&vigNum>0&&vigNum<9000){vigDias=Math.round(vigNum);}else{vigDias=getVig(r[CI.sub]);}}
    let est;
    if(vigDias>=9999){est=getEst(dias,getVig(r[CI.sub]));}
    else{const pct=dias/vigDias;if(pct<=0.75)est='Verde';else if(pct<=0.87)est='Amarillo';else est='Rojo';}
    if(!est)est='Verde';
    let dv=0;const dvRaw=r[CI.dv];
    if(dvRaw!==null&&dvRaw!==undefined&&String(dvRaw).trim()!==''&&String(dvRaw).trim()!==' '){const dvNum=parseFloat(String(dvRaw).trim());if(!isNaN(dvNum))dv=Math.round(dvNum);}
    if(dv===0&&vigDias<9999)dv=Math.max(0,dias-vigDias);
    const comentario=jInfo?.comentario||'';
    const razonVal=String(r[CI.razon]||'').trim();
    const justVal=String(r[CI.just]||'').trim();
    function esFormula(v){return v&&v.startsWith('=')}
    const justDisplay=(comentario&&!esFormula(comentario)&&comentario!=='0')?comentario:(razonVal&&!esFormula(razonVal)&&razonVal!=='0')?razonVal:(justVal&&!esFormula(justVal)&&justVal!=='0')?justVal:(jInfo?.tipo||'—');
    const vigLabel=vigDias>=9999?'Planta':vigDias+' días';
    const g=r[CI.sub];
    rows.push({
      _idx:rows.length,posicion:posStr,descripcion:String(r[CI.desc]||''),
      nivel:String(r[CI.nivel]||''),subnivel:String(g||''),sind:esSindicalizado(g),
      zona:reclassifyZona(r[CI.local],r[CI.unidad],r[CI.ccosto],r[CI.desc],r[CI.zona])||normZona(String(r[CI.zona]||'')),localidad:String(r[CI.local]||'—'),
      grupo:String(r[CI.grupo]||'—'),dias,vigencia:vigDias,vigLabel,estatus:est,
      estatExcel:String(r[CI.estatExcel]||'—'),diasVencidos:dv,jefeRH:String(r[CI.jrh]||'—'),
      justificacion:justDisplay,tipoJust:jInfo?.tipo||justVal||'—',clasificacion:String(r[CI.clasif]||'—'),
    });
  }
  return{rows,totalPosiciones,totalOcupadas,allPos,posZonaMap};
}

// ── Leer Excel Semáforo ──
function handleFile(file){
  if(!file)return;clearError();
  document.getElementById('fname-text').textContent=file.name;
  document.getElementById('file-badge').style.display='flex';
  document.getElementById('upload-section').style.display='none';
  // Una sola carga alimenta también Detalle de Posiciones y Status de Proyectos
  handleDetFile(file);
  const reader=new FileReader();
  reader.onload=function(e){
    try{
      const wb=XLSX.read(e.target.result,{type:'array',cellFormula:false,cellNF:false,raw:false});
      const result=parseVacantesWorkbook(wb);
      if(result.error){showError(result.error);return;}
      allRows=result.rows;
      window._totalPosiciones=result.totalPosiciones;window._totalOcupadas=result.totalOcupadas;
      if(!allRows.length){showError('No se encontraron vacantes.');return;}
      document.getElementById('dashboard').style.display='block';
      document.getElementById('export-bar').style.display='flex';
      document.getElementById('analytics-bar').style.display='block';
      buildZonaBar();refresh();
    }catch(err){showError('Error al procesar: '+err.message)}
  };
  reader.readAsArrayBuffer(file);
}

function buildZonaBar(){
  const zonaSet=new Set();
  allRows.forEach(r=>{if(r.zona&&r.zona!=='')zonaSet.add(r.zona);});
  const order=['Norte','Sur','Centro','Corporativo','Polanco'];
  const sorted=[...order.filter(z=>zonaSet.has(z)),...[...zonaSet].sort().filter(z=>!order.includes(z))];
  const bar=document.getElementById('zona-bar');
  bar.querySelectorAll('.zbtn-dyn').forEach(b=>b.remove());
  sorted.forEach(z=>{
    const count=allRows.filter(r=>r.zona===z).length;
    const btn=document.createElement('button');
    btn.className='zbtn zbtn-dyn';
    btn.onclick=function(){setZona(z,btn)};
    btn.innerHTML=z+' <span class="zona-count">'+count+'</span>';
    bar.appendChild(btn);
  });
  document.getElementById('zc-todas').textContent=allRows.length;
}

function refresh(){
  const totalPos=window._totalPosiciones||allRows.length;
  const totalOcup=window._totalOcupadas||0;
  document.getElementById('r-total-pos').textContent=totalPos;
  document.getElementById('r-ocupadas').textContent=totalOcup;
  document.getElementById('r-vacantes').textContent=allRows.length;
  document.getElementById('zc-todas').textContent=allRows.length;
  document.querySelectorAll('.zbtn-dyn').forEach(btn=>{
    const zonaText=btn.firstChild?btn.firstChild.textContent.trim():'';
    const count=allRows.filter(r=>r.zona===zonaText).length;
    const spanEl=btn.querySelector('.zona-count');
    if(spanEl)spanEl.textContent=count;
  });
  const base=activeZona?allRows.filter(r=>r.zona===activeZona):allRows;
  const nV=base.filter(r=>r.estatus==='Verde').length;
  const nA=base.filter(r=>r.estatus==='Amarillo').length;
  const nR=base.filter(r=>r.estatus==='Rojo').length;
  document.getElementById('s-total').textContent=base.length;
  document.getElementById('s-total-sub').textContent=activeZona?'Vacantes disponibles — '+activeZona:'Total vacantes disponibles';
  document.getElementById('s-rojo').textContent=nR;
  document.getElementById('s-amarillo').textContent=nA;
  document.getElementById('s-verde').textContent=nV;
  if(analyticsOpen) setTimeout(renderCharts,50);
  ['verde','amarillo','rojo'].forEach(c=>{
    const est=c.charAt(0).toUpperCase()+c.slice(1);
    const rows=base.filter(r=>r.estatus===est);
    applyQ(c,rows,state[c].q);
    if(state[c].open)renderAcord(c);
  });
}

function setZona(val,el){
  activeZona=val;
  document.querySelectorAll('#zona-bar .zbtn').forEach(b=>b.classList.remove('z-active'));
  el.classList.add('z-active');
  refresh();
}

function setCompZona(val,el){
  compActiveZona=val;
  document.querySelectorAll('#comp-zona-bar .zbtn').forEach(b=>b.classList.remove('z-active'));
  el.classList.add('z-active');
  renderComparativo();
}

function toggleAcord(color){
  const s=state[color];s.open=!s.open;
  const wrap=document.getElementById('acord-'+color);
  const kpi=document.getElementById('kpi-'+color);
  if(s.open){wrap.classList.add('open');kpi.classList.add('open');s.page=0;renderAcord(color)}
  else{wrap.classList.remove('open');kpi.classList.remove('open')}
}

function filterAcord(color,q){
  state[color].q=q;
  const base=activeZona?allRows.filter(r=>r.zona===activeZona):allRows;
  const est=color.charAt(0).toUpperCase()+color.slice(1);
  const rows=base.filter(r=>r.estatus===est);
  applyQ(color,rows,q);state[color].page=0;renderAcord(color);
}
function applyQ(color,rows,q){
  const nq=norm(q);
  state[color].filtered=!nq?rows:rows.filter(r=>norm(r.posicion).includes(nq)||norm(r.descripcion).includes(nq)||norm(r.zona).includes(nq)||norm(r.localidad).includes(nq)||norm(r.jefeRH).includes(nq)||norm(r.justificacion).includes(nq));
  document.getElementById('acord-'+color+'-count').textContent=state[color].filtered.length+' vacantes';
}
function pageAcord(color,dir){
  const s=state[color];const pages=Math.ceil(s.filtered.length/PAGE);
  s.page=Math.max(0,Math.min(pages-1,s.page+dir));renderAcord(color);
}

const COLS=[
  {k:'posicion',l:'No. Posicion'},
  {k:'descripcion',l:'Descripcion',render:r=>`<span style="font-size:12px;max-width:160px;display:block;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${r.descripcion}">${r.descripcion||'—'}</span>`},
  {k:'nivel',l:'Nivel',render:r=>{const v=r.nivel||r.subnivel;return r.sind?`<span class="nivel-sind">SIND</span>`:`<span class="nivel-badge">${v||'—'}</span>`;}},
  {k:'subnivel',l:'Subnivel',render:r=>r.sind?`<span style="font-size:11px">${r.subnivel}</span>`:`<span style="font-size:12px">${r.subnivel||'—'}</span>`},
  {k:'grupo',l:'Desc. Grupo',render:r=>`<span style="font-size:12px;max-width:140px;display:block;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${r.grupo}">${r.grupo||'—'}</span>`},
  {k:'localidad',l:'Desc. Localidad',render:r=>`<span style="font-size:12px;max-width:130px;display:block;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${r.localidad}">${r.localidad}</span>`},
  {k:'estatExcel',l:'Estatus',render:r=>{const v=(r.estatExcel||'—').toUpperCase();const color=v==='VENCIDA'?'#C8102E':v==='VIGENTE'?'#16a34a':'#94a3b8';return`<span style="font-size:11px;font-weight:700;color:${color}">${r.estatExcel||'—'}</span>`;}},
  {k:'dias',l:'Dias sin ocupar',render:r=>`<strong>${r.dias}</strong>`},
  {k:'jefeRH',l:'Jefe RH',render:r=>`<span style="font-size:12px;max-width:120px;display:block;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${r.jefeRH}">${r.jefeRH}</span>`},
  {k:'vigLabel',l:'Vigencia',render:r=>{if(r.vigencia>=9999)return`<span style="font-size:11px;color:var(--p-gray)">Planta</span>`;const pct=r.dias/r.vigencia;const color=pct>0.87?'var(--rojo-text)':pct>0.75?'var(--amarillo-text)':'var(--verde-text)';return`<span style="font-size:12px;font-weight:700;color:${color}">${r.vigencia} días</span>`;}},
  {k:'justificacion',l:'Tipo Just.',render:r=>`<span class="justif-tag" title="${r.justificacion}">${r.justificacion}</span>`},
  {k:'clasificacion',l:'Clasificacion',render:r=>`<span style="font-size:11px;color:var(--p-gray)">${r.clasificacion||'—'}</span>`},
  {k:'_plan',l:'Plan de Accion',noSort:true,render:r=>{const key=rowKey(r);const val=(comentarios[key]||{}).plan||'';return`<div class="comment-cell"><textarea class="comment-input plan" rows="3" placeholder="Escribe el plan de accion..." onchange="saveComentario('${key}','plan',this.value)">${escHtml(val)}</textarea></div>`;}},
  {k:'_responsable',l:'Responsable',noSort:true,render:r=>{const key=rowKey(r);const val=(comentarios[key]||{}).responsable||'';return`<div class="comment-cell"><input class="resp-input" type="text" placeholder="Responsable..." value="${escHtml(val)}" onchange="saveComentario('${key}','responsable',this.value)"></div>`;}},
  {k:'_fecha',l:'Fecha',noSort:true,render:r=>{const key=rowKey(r);const val=(comentarios[key]||{}).fecha||'';return`<div class="comment-cell"><input class="date-input" type="date" value="${escHtml(val)}" onchange="saveComentario('${key}','fecha',this.value)"></div>`;}},
];

function rowKey(r){return 'r_'+(r.posicion||'')+r._idx}
function escHtml(s){return(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;')}
function saveComentario(key,field,val){if(!comentarios[key])comentarios[key]={plan:'',responsable:'',fecha:''};comentarios[key][field]=val;}

const sortState={};
function renderAcord(color){
  const s=state[color];const start=s.page*PAGE;const rows=s.filtered.slice(start,start+PAGE);
  const total=s.filtered.length;const pages=Math.ceil(total/PAGE);
  const thead=document.getElementById('thead-'+color);
  thead.innerHTML=COLS.map(c=>c.noSort?`<th style="min-width:160px">${c.l}</th>`:`<th onclick="sortAcord('${color}','${c.k}',this)">${c.l} <span class="sa">↕</span></th>`).join('');
  const tbody=document.getElementById('tbody-'+color);
  if(!rows.length){tbody.innerHTML=`<tr><td colspan="${COLS.length}" class="no-rows">Sin resultados</td></tr>`;document.getElementById('pag-'+color).style.display='none';return}
  document.getElementById('pag-'+color).style.display='flex';
  const cls=color==='rojo'?'row-rojo':color==='amarillo'?'row-amarillo':'';
  tbody.innerHTML=rows.map(r=>`<tr class="${cls}">${COLS.map(c=>`<td>${c.render?c.render(r):(r[c.k]||'—')}</td>`).join('')}</tr>`).join('');
  document.getElementById('pag-'+color+'-info').textContent=`${start+1}–${Math.min(start+PAGE,total)} de ${total}`;
  document.getElementById('prev-'+color).disabled=s.page===0;
  document.getElementById('next-'+color).disabled=s.page>=pages-1;
}

const EORD={'Rojo':0,'Amarillo':1,'Verde':2,'—':3};
function sortAcord(color,key,thEl){
  if(!key||key.startsWith('_'))return;
  const sc=sortState[color]||{key:null,asc:true};
  if(sc.key===key)sc.asc=!sc.asc;else{sc.key=key;sc.asc=true}
  sortState[color]=sc;
  state[color].filtered.sort((a,b)=>{
    let va=a[key],vb=b[key];
    if(key==='estatus'){va=EORD[va]??99;vb=EORD[vb]??99}
    if(typeof va==='number'&&typeof vb==='number')return sc.asc?va-vb:vb-va;
    return sc.asc?String(va).localeCompare(String(vb),'es'):String(vb).localeCompare(String(va),'es');
  });
  state[color].page=0;renderAcord(color);
  document.querySelectorAll('#thead-'+color+' th').forEach(t=>t.classList.remove('sorted'));
  if(thEl)thEl.classList.add('sorted');
}

function showError(msg){const el=document.getElementById('error-box');el.textContent='Aviso: '+msg;el.style.display='block'}
function clearError(){document.getElementById('error-box').style.display='none'}

// =============================================================
// ── HELPER COMPARTIDO: GENERAR PDF POR IMPRESIÓN NATIVA ──
// Se probaron dos versiones basadas en html2pdf.js/html2canvas
// (capturar el reporte como imagen) pero salían en blanco de forma
// consistente y no hay forma de depurar esa librería sin acceso a
// un navegador real. Ahora se usa la impresión nativa del
// navegador: se abre el reporte ya armado (mismo HTML, mismos
// colores y formato que en pantalla) en una pestaña nueva y se
// dispara el diálogo de impresión automáticamente — ahí solo hay
// que elegir "Guardar como PDF" como destino/impresora. Esto no
// depende de rasterizar nada, así que no puede salir en blanco.
// =============================================================
function abrirVentanaImpresion(titulo, htmlContent, pageCss){
  const fullHtml=`<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><title>${escHtml(titulo)}</title><style>*{box-sizing:border-box;margin:0;padding:0}body{font-family:'Abadi','Century Gothic','Trebuchet MS',Arial,sans-serif;font-size:12px;background:#fff;padding:20px;-webkit-print-color-adjust:exact;print-color-adjust:exact;color-adjust:exact}${pageCss||'@page{size:letter landscape;margin:12mm}'}@media print{body{padding:0}.no-print{display:none!important}tr{page-break-inside:avoid}table{page-break-inside:auto}}.print-btn{display:block;margin:0 auto 20px auto;padding:10px 30px;background:#003087;color:#fff;border:none;border-radius:6px;font-size:14px;font-weight:700;cursor:pointer}.print-hint{text-align:center;font-size:12px;color:#666;margin-bottom:16px}</style></head><body>
    <button class="no-print print-btn" onclick="window.print()">🖨️ Imprimir / Guardar como PDF</button>
    <div class="no-print print-hint">Se abrió el diálogo de impresión automáticamente — elige <strong>"Guardar como PDF"</strong> como destino/impresora. Si no se abrió, usa el botón de arriba.</div>
    ${htmlContent}
    <script>window.onload=function(){setTimeout(function(){window.print()},400)}<\/script>
  </body></html>`;
  try{
    const blob=new Blob([fullHtml],{type:'text/html'});
    const url=URL.createObjectURL(blob);
    const win=window.open(url,'_blank');
    if(!win){
      URL.revokeObjectURL(url);
      alert('El navegador bloqueó la ventana del reporte. Habilita las ventanas emergentes para este sitio e intenta de nuevo.');
      return false;
    }
    setTimeout(()=>URL.revokeObjectURL(url),60000);
    return true;
  }catch(err){
    console.error('Error generando el reporte:',err);
    alert('Ocurrió un error al generar el reporte: '+err.message);
    return false;
  }
}

// =============================================================
// ── EXPORTAR PDF SEMÁFORO ──
// =============================================================
async function exportarPDF(){
  const base=activeZona?allRows.filter(r=>r.zona===activeZona):allRows;
  const rojos=base.filter(r=>r.estatus==='Rojo');
  const amarillos=base.filter(r=>r.estatus==='Amarillo');
  const verdes=base.filter(r=>r.estatus==='Verde');
  const vacantesCriticas=[...rojos,...amarillos];
  const vacantes=[...rojos,...amarillos,...verdes];
  if(!vacantes.length){alert('No hay vacantes para exportar.');return;}
  const btn=document.getElementById('btn-export-pdf');
  btn.disabled=true;btn.textContent='⏳ Generando PDF...';
  try{
  // Gráficas dedicadas para el PDF: se dibujan en canvases fuera de pantalla,
  // con tamaño fijo (no dependen de que el acordeón esté abierto), porcentaje
  // fijo visible (no solo al pasar el mouse) y Top Localidades limitado a 10
  // para que no se amontone ni se vea chico.
  const analisisHtml=await buildAnalisisVacantesPDF(base);
  const hoy=new Date().toLocaleDateString('es-MX',{day:'2-digit',month:'long',year:'numeric'});
  const zonaLabel=activeZona||'Nacional';
  const avgDiasRojo=rojos.length?Math.round(rojos.reduce((a,r)=>a+r.dias,0)/rojos.length):0;
  const avgDiasAmarillo=amarillos.length?Math.round(amarillos.reduce((a,r)=>a+r.dias,0)/amarillos.length):0;
  const avgDiasVerde=verdes.length?Math.round(verdes.reduce((a,r)=>a+r.dias,0)/verdes.length):0;
  const maxDias=vacantesCriticas.length?Math.max(...vacantesCriticas.map(r=>r.dias)):0;
  const sinPlan=vacantes.filter(r=>!(comentarios[rowKey(r)]||{}).plan).length;
  const conPlan=vacantes.length-sinPlan;
  const pctCriticas=base.length?Math.round((rojos.length/base.length)*100):0;
  const tdRojo='padding:6px 8px;font-size:10px;border-bottom:1px solid #FECDD3;vertical-align:top;';
  const tdAmar='padding:6px 8px;font-size:10px;border-bottom:1px solid #FDE68A;vertical-align:top;';
  const tdVerde='padding:6px 8px;font-size:10px;border-bottom:1px solid #A7F3D0;vertical-align:top;';
  const hdrRojo='padding:7px 8px;font-size:10px;font-weight:800;text-align:left;text-transform:uppercase;letter-spacing:.04em;color:#9F1239;border-bottom:2px solid #FECDD3;background:#FFF1F2;white-space:nowrap;';
  const hdrAmar='padding:7px 8px;font-size:10px;font-weight:800;text-align:left;text-transform:uppercase;letter-spacing:.04em;color:#92400E;border-bottom:2px solid #FDE68A;background:#FFFBEB;white-space:nowrap;';
  const hdrVerde='padding:7px 8px;font-size:10px;font-weight:800;text-align:left;text-transform:uppercase;letter-spacing:.04em;color:#166534;border-bottom:2px solid #A7F3D0;background:#ECFDF5;white-space:nowrap;';
  function buildRows(arr,tdStyle,color){
    return arr.map(r=>{
      const key=rowKey(r);const com=comentarios[key]||{};
      return`<tr><td style="${tdStyle}font-weight:800;color:${color}">${escHtml(String(r.posicion||'—'))}</td><td style="${tdStyle}">${escHtml(String(r.descripcion||'—'))}</td><td style="${tdStyle}"><strong>${escHtml(String(r.localidad||'—'))}</strong></td><td style="${tdStyle}">${escHtml(String(r.justificacion||'—'))}</td><td style="${tdStyle}text-align:center"><span style="background:${color};color:#fff;border-radius:4px;padding:2px 5px;font-size:10px;font-weight:800">${r.dias} días</span></td><td style="${tdStyle}">${com.plan?escHtml(com.plan):'<em style="color:#ccc">Pendiente</em>'}</td><td style="${tdStyle}">${com.responsable?escHtml(com.responsable):'<span style="color:#ccc">—</span>'}</td><td style="${tdStyle}">${com.fecha?escHtml(com.fecha):'<span style="color:#ccc">—</span>'}</td></tr>`;
    }).join('');
  }
  const htmlContent=`<div style="font-family:'Abadi','Century Gothic','Trebuchet MS',Arial,sans-serif;color:#111;font-size:12px;padding:20px;background:#fff">
    <div style="background:linear-gradient(135deg,#003087,#C8102E);padding:18px 28px;border-radius:10px 10px 0 0"><div style="color:#fff;font-size:22px;font-weight:900">PILGRIMS</div><div style="color:rgba(255,255,255,.8);font-size:11px;text-transform:uppercase">Atraccion de Talento</div></div>
    <div style="border:2px solid #003087;border-top:none;border-radius:0 0 10px 10px;padding:16px 22px;margin-bottom:18px"><div style="font-size:17px;font-weight:900;color:#003087;text-transform:uppercase;margin-bottom:4px">Reporte de Seguimiento Semanal de Vacantes</div><div style="font-size:11px;color:#777"><strong>Fecha:</strong> ${hoy} &nbsp; <strong>Complejo:</strong> ${zonaLabel} &nbsp; <strong>Total:</strong> ${vacantes.length} vacantes</div></div>
    <div style="display:flex;gap:10px;margin-bottom:20px">
      <div style="background:#FFF1F2;border:1.5px solid #FECDD3;border-radius:8px;padding:12px;text-align:center;flex:1"><div style="font-size:28px;font-weight:900;color:#9F1239">${rojos.length}</div><div style="font-size:10px;font-weight:700;color:#9F1239">CRITICAS ROJO</div><div style="font-size:10px;color:#aaa">Prom. ${avgDiasRojo}d</div></div>
      <div style="background:#FFFBEB;border:1.5px solid #FDE68A;border-radius:8px;padding:12px;text-align:center;flex:1"><div style="font-size:28px;font-weight:900;color:#92400E">${amarillos.length}</div><div style="font-size:10px;font-weight:700;color:#92400E">POR VENCER AMARILLO</div><div style="font-size:10px;color:#aaa">Prom. ${avgDiasAmarillo}d</div></div>
      <div style="background:#ECFDF5;border:1.5px solid #A7F3D0;border-radius:8px;padding:12px;text-align:center;flex:1"><div style="font-size:28px;font-weight:900;color:#166534">${verdes.length}</div><div style="font-size:10px;font-weight:700;color:#166534">VIGENTES VERDE</div><div style="font-size:10px;color:#aaa">Prom. ${avgDiasVerde}d</div></div>
      <div style="background:#EEF2FF;border:1.5px solid #C7D2FE;border-radius:8px;padding:12px;text-align:center;flex:1"><div style="font-size:28px;font-weight:900;color:#3730A3">${conPlan}</div><div style="font-size:10px;font-weight:700;color:#3730A3">CON PLAN DE ACCION</div><div style="font-size:10px;color:#aaa">${sinPlan} pendientes</div></div>
    </div>
    ${analisisHtml}
    ${rojos.length?`<div style="margin-bottom:24px"><div style="background:#9F1239;color:#fff;padding:9px 14px;border-radius:8px 8px 0 0;font-size:12px;font-weight:800;text-transform:uppercase">Vacantes Criticas — Rojo <span style="background:rgba(255,255,255,.2);padding:2px 10px;border-radius:20px;font-size:11px">${rojos.length}</span></div><table style="width:100%;border-collapse:collapse;border:1.5px solid #FECDD3;border-top:none"><thead><tr><th style="${hdrRojo}">No. Posicion</th><th style="${hdrRojo}">Descripcion</th><th style="${hdrRojo}">Localidad</th><th style="${hdrRojo}">Motivo</th><th style="${hdrRojo}">Dias</th><th style="${hdrRojo}">Plan de Accion</th><th style="${hdrRojo}">Responsable</th><th style="${hdrRojo}">Fecha</th></tr></thead><tbody>${buildRows(rojos,tdRojo,'#9F1239')}</tbody></table></div>`:''}
    ${amarillos.length?`<div style="margin-bottom:24px"><div style="background:#D97706;color:#fff;padding:9px 14px;border-radius:8px 8px 0 0;font-size:12px;font-weight:800;text-transform:uppercase">Vacantes Por Vencer — Amarillo <span style="background:rgba(255,255,255,.2);padding:2px 10px;border-radius:20px;font-size:11px">${amarillos.length}</span></div><table style="width:100%;border-collapse:collapse;border:1.5px solid #FDE68A;border-top:none"><thead><tr><th style="${hdrAmar}">No. Posicion</th><th style="${hdrAmar}">Descripcion</th><th style="${hdrAmar}">Localidad</th><th style="${hdrAmar}">Motivo</th><th style="${hdrAmar}">Dias</th><th style="${hdrAmar}">Plan de Accion</th><th style="${hdrAmar}">Responsable</th><th style="${hdrAmar}">Fecha</th></tr></thead><tbody>${buildRows(amarillos,tdAmar,'#D97706')}</tbody></table></div>`:''}
    ${verdes.length?`<div style="margin-bottom:24px"><div style="background:#166534;color:#fff;padding:9px 14px;border-radius:8px 8px 0 0;font-size:12px;font-weight:800;text-transform:uppercase">Vacantes Vigentes — Verde <span style="background:rgba(255,255,255,.2);padding:2px 10px;border-radius:20px;font-size:11px">${verdes.length}</span></div><table style="width:100%;border-collapse:collapse;border:1.5px solid #A7F3D0;border-top:none"><thead><tr><th style="${hdrVerde}">No. Posicion</th><th style="${hdrVerde}">Descripcion</th><th style="${hdrVerde}">Localidad</th><th style="${hdrVerde}">Motivo</th><th style="${hdrVerde}">Dias</th><th style="${hdrVerde}">Plan de Accion</th><th style="${hdrVerde}">Responsable</th><th style="${hdrVerde}">Fecha</th></tr></thead><tbody>${buildRows(verdes,tdVerde,'#166534')}</tbody></table></div>`:''}
    <div style="margin-top:20px;text-align:center;font-size:10px;color:#bbb;border-top:1px solid #e5e7eb;padding-top:10px">Pilgrims · Atraccion de Talento · Documento confidencial · ${hoy} · ${zonaLabel}</div>
  </div>`;
  await abrirVentanaImpresion('Vacantes_'+zonaLabel, htmlContent);
  }catch(err){
    console.error('Error generando el reporte de vacantes:',err);
    alert('Ocurrió un error al generar el reporte: '+err.message);
  }finally{
    btn.disabled=false;btn.innerHTML='Exportar PDF';
  }
}

// =============================================================
// ── ANALYTICS SEMÁFORO ──
// =============================================================
let analyticsOpen=false;
const chartInstances={};
const P_COLORS=['#003087','#C8102E','#5A6070','#1B4F9C','#9F1239','#A8B5C8','#E8EEF8','#DDE1EA','#F5F6F8','#1A1A2E','#6B7280','#374151','#DC2626','#1D4ED8','#047857'];
const CHART_FONT="'Abadi','Century Gothic','Trebuchet MS',Arial,sans-serif";
Chart.defaults.font.family=CHART_FONT;Chart.defaults.font.size=11;Chart.defaults.color='#5A6070';
function makeColors(n,alpha){alpha=alpha||1;return Array.from({length:n},(_,i)=>{const c=P_COLORS[i%P_COLORS.length];if(alpha===1)return c;const r=parseInt(c.slice(1,3),16),g=parseInt(c.slice(3,5),16),b=parseInt(c.slice(5,7),16);return`rgba(${r},${g},${b},${alpha})`;});}

// Pinta la lista de "punto de color · etiqueta · cantidad (porcentaje)" que
// va A UN LADO de una gráfica, en vez de encima de las barras/rebanadas
// (donde se traslapaban con muchas categorías). containerId debe apuntar a
// un <div class="chart-side-stats"> vacío ya presente en el HTML.
function buildSideStats(containerId,labels,data,colors,total,pctOnly){
  const el=document.getElementById(containerId);
  if(!el)return;
  if(!labels.length){el.innerHTML='<div style="font-size:11px;color:#aaa">Sin datos</div>';return;}
  el.innerHTML=labels.map((l,i)=>{
    const v=data[i];
    const pct=total?Math.round(v/total*100):0;
    const valorHtml=pctOnly?`<span style="font-weight:800;color:#111;white-space:nowrap;font-size:10.5px">${pct}%</span>`:`<span style="font-weight:800;color:#111;white-space:nowrap;font-size:10.5px">${v} <span style="font-weight:700;color:#5A6070">(${pct}%)</span></span>`;
    return`<div style="display:flex;align-items:center;gap:6px;font-size:11px;line-height:1.3"><span style="width:9px;height:9px;border-radius:50%;flex-shrink:0;display:inline-block;background:${colors[i]||'#94a3b8'}"></span><span style="flex:1;color:#3A3F4B;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;min-width:0" title="${escHtml(String(l))}">${escHtml(String(l))}</span>${valorHtml}</div>`;
  }).join('');
}
function toggleAnalytics(){
  analyticsOpen=!analyticsOpen;
  const content=document.getElementById('analytics-content');const arrow=document.getElementById('analytics-arrow');
  if(analyticsOpen){content.style.maxHeight='2200px';content.style.opacity='1';arrow.style.transform='rotate(180deg)';setTimeout(renderCharts,60);}
  else{content.style.maxHeight='0';content.style.opacity='0';arrow.style.transform='rotate(0deg)';}
}
function countBy(rows,key,topN){
  const map={};rows.forEach(r=>{let v=String(r[key]||'—').trim();if(!v||v==='—'||v==='')v='Sin dato';map[v]=(map[v]||0)+1;});
  let entries=Object.entries(map).sort((a,b)=>b[1]-a[1]);
  if(topN&&entries.length>topN){const rest=entries.slice(topN).reduce((a,e)=>a+e[1],0);entries=entries.slice(0,topN);if(rest>0)entries.push(['Otros',rest]);}
  return{labels:entries.map(e=>e[0]),data:entries.map(e=>e[1])};
}
function destroyChart(id){if(chartInstances[id]){chartInstances[id].destroy();delete chartInstances[id];}}
// =============================================================
// ── GRÁFICAS DEDICADAS PARA EL PDF DE SEMÁFORO (Análisis de Vacantes) ──
// Se dibujan en canvases fuera de pantalla con tamaño fijo, para que
// el PDF no dependa del acordeón en pantalla, muestren el % de forma
// permanente (no solo al pasar el mouse) y no se amontonen al imprimir.
// =============================================================
async function buildAnalisisVacantesPDF(base){
  if(!base||!base.length)return '';
  function nextFrame(){return new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r)));}
  function crearCanvas(w,h){
    const wrap=document.createElement('div');
    wrap.style.cssText='position:fixed;left:-99999px;top:0;width:'+w+'px;height:'+h+'px;';
    const canvas=document.createElement('canvas');
    canvas.width=w;canvas.height=h;
    wrap.appendChild(canvas);
    document.body.appendChild(wrap);
    return{canvas,wrap};
  }
  const prevAnim=Chart.defaults.animation;
  Chart.defaults.animation=false;
  const temp=[];
  // Lista lateral "color · etiqueta · cantidad (porcentaje)" para el PDF —
  // versión en string (no toca el DOM) del mismo patrón que ya se usa en
  // pantalla, para que el porcentaje quede junto a su color en vez de
  // separado en la leyenda y encima de la rebanada.
  function sideStatsHtml(labels,data,colors,total){
    return labels.map((l,i)=>{
      const v=data[i];
      const pct=total?Math.round(v/total*100):0;
      return `<div style="display:flex;align-items:center;gap:6px;font-size:11px;margin-bottom:7px"><span style="width:9px;height:9px;border-radius:50%;background:${colors[i]||'#94a3b8'};flex-shrink:0;display:inline-block"></span><span style="flex:1;color:#3A3F4B;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:110px">${escHtml(String(l))}</span><span style="font-weight:800;color:#111;white-space:nowrap">${v} <span style="font-weight:700;color:#5A6070">(${pct}%)</span></span></div>`;
    }).join('');
  }
  function makePie(type,labels,data,colors,w,h,cutout){
    const{canvas,wrap}=crearCanvas(w,h);
    const total=data.reduce((a,b)=>a+b,0);
    const chart=new Chart(canvas,{type,data:{labels,datasets:[{data,backgroundColor:colors,borderWidth:2,borderColor:'#fff'}]},options:{responsive:false,animation:false,cutout,layout:{padding:20},plugins:{legend:{display:false},datalabels:{display:false}}}});
    temp.push({canvas,wrap,chart,labels,data,colors,total});
  }
  // Tipo de justificación
  const tj=countBy(base,'tipoJust',8);
  makePie('doughnut',tj.labels,tj.data,makeColors(tj.labels.length),900,440,'58%');
  // Estatus
  const estatusMap={};base.forEach(r=>{const v=String(r.estatExcel||'—').trim();estatusMap[v]=(estatusMap[v]||0)+1;});
  const eL=Object.keys(estatusMap).filter(k=>estatusMap[k]>0);const eD=eL.map(k=>estatusMap[k]);
  const eC=eL.map(k=>{const n=k.toUpperCase();if(n==='VIGENTE')return'#16a34a';if(n==='VENCIDA')return'#C8102E';return'#94a3b8';});
  makePie('pie',eL,eD,eC,900,440,undefined);
  // Top 10 localidades (sin bucket "Otros", tamaño fijo para que se lea bien)
  {
    const locFull=countBy(base,'localidad',999);
    const labels=locFull.labels.slice(0,10),data=locFull.data.slice(0,10);
    const{canvas,wrap}=crearCanvas(900,540);
    const total=base.length;
    const chart=new Chart(canvas,{type:'bar',data:{labels,datasets:[{label:'Vacantes',data,backgroundColor:makeColors(labels.length,0.85),borderColor:makeColors(labels.length),borderWidth:1.5,borderRadius:4}]},options:{indexAxis:'y',responsive:false,animation:false,layout:{padding:{right:70,left:6,top:10,bottom:10}},plugins:{legend:{display:false},datalabels:{display:true,anchor:'end',align:'end',color:'#003087',font:{size:12,weight:'700',family:CHART_FONT},formatter:v=>v+' ('+(total?Math.round(v/total*100):0)+'%)'}},scales:{x:{grid:{color:'#EEF0F5'},ticks:{font:{size:11}}},y:{grid:{display:false},ticks:{font:{size:13}}}}}});
    temp.push({canvas,wrap,chart});
  }
  // Grupos
  const grp=countBy(base,'grupo',10);
  makePie('doughnut',grp.labels,grp.data,makeColors(grp.labels.length),900,440,'55%');
  // Áreas
  const area=countBy(base,'clasificacion',10);
  makePie('pie',area.labels,area.data,makeColors(area.labels.length),900,440,undefined);

  await nextFrame();
  await new Promise(r=>setTimeout(r,120));

  const titles=['Tipo de Justificación','Estatus (Vigente / Vencida)','Top 10 Localidades','Descripción de Grupos','Descripción de Áreas (Clasificación)'];
  const imgs=temp.map((t,i)=>{
    let dataUrl='';
    try{dataUrl=t.canvas.toDataURL('image/png',1.0);}catch(e){}
    return{titulo:titles[i],dataUrl,labels:t.labels,data:t.data,colors:t.colors,total:t.total};
  });
  temp.forEach(t=>{try{t.chart.destroy();}catch(e){}try{t.wrap.remove();}catch(e){}});
  Chart.defaults.animation=prevAnim;

  function box(item,tall){
    if(!item.dataUrl)return `<div style="background:#F8FAFC;border:1px solid #DDE1EA;border-radius:10px;padding:14px;font-size:10px;color:#aaa;text-align:center">${escHtml(item.titulo)}<br>(gráfica no disponible)</div>`;
    const sideHtml=item.labels?`<div style="width:165px;flex-shrink:0;display:flex;flex-direction:column;justify-content:center">${sideStatsHtml(item.labels,item.data,item.colors,item.total)}</div>`:'';
    return `<div style="background:#F8FAFC;border:1px solid #DDE1EA;border-radius:10px;padding:12px;page-break-inside:avoid">
      <div style="font-size:11px;font-weight:800;text-transform:uppercase;letter-spacing:.05em;color:#003087;margin-bottom:8px">${escHtml(item.titulo)}</div>
      <div style="display:flex;gap:10px;align-items:center">
        <img src="${item.dataUrl}" style="flex:1;min-width:0;${tall?'max-height:300px':'max-height:230px'};object-fit:contain">
        ${sideHtml}
      </div>
    </div>`;
  }

  return `<div style="margin-bottom:22px">
    <div style="background:#003087;color:#fff;padding:9px 14px;border-radius:8px 8px 0 0;font-size:12px;font-weight:800;text-transform:uppercase">Análisis de Vacantes</div>
    <div style="border:1.5px solid #DDE1EA;border-top:none;border-radius:0 0 8px 8px;padding:16px">
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:16px">
        ${box(imgs[0])}
        ${box(imgs[1])}
      </div>
      <div style="margin-bottom:16px">${box(imgs[2],true)}</div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px">
        ${box(imgs[3])}
        ${box(imgs[4])}
      </div>
    </div>
  </div>`;
}

function renderCharts(){
  const base=activeZona?allRows.filter(r=>r.zona===activeZona):allRows;
  if(!base.length)return;
  destroyChart('tipojust');
  const tj=countBy(base,'tipoJust',8);
  const tjColors=makeColors(tj.labels.length);
  chartInstances['tipojust']=new Chart(document.getElementById('chart-tipojust'),{type:'doughnut',data:{labels:tj.labels,datasets:[{data:tj.data,backgroundColor:tjColors,borderWidth:2,borderColor:'#fff',hoverOffset:6}]},options:{responsive:true,maintainAspectRatio:false,cutout:'58%',plugins:{legend:{display:false},datalabels:{display:false},tooltip:{callbacks:{label:ctx=>` ${ctx.label}: ${ctx.parsed} (${Math.round(ctx.parsed/base.length*100)}%)`}}}}});
  buildSideStats('chart-tipojust-stats',tj.labels,tj.data,tjColors,base.length);

  destroyChart('estatus');
  const estatusMap={};base.forEach(r=>{const v=String(r.estatExcel||'—').trim();estatusMap[v]=(estatusMap[v]||0)+1;});
  const eL=Object.keys(estatusMap).filter(k=>estatusMap[k]>0);const eD=eL.map(k=>estatusMap[k]);
  const eC=eL.map(k=>{const n=k.toUpperCase();if(n==='VIGENTE')return'#16a34a';if(n==='VENCIDA')return'#C8102E';return'#94a3b8';});
  chartInstances['estatus']=new Chart(document.getElementById('chart-estatus'),{type:'pie',data:{labels:eL,datasets:[{data:eD,backgroundColor:eC,borderWidth:3,borderColor:'#fff',hoverOffset:6}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},datalabels:{display:false},tooltip:{callbacks:{label:ctx=>` ${ctx.label}: ${ctx.parsed} (${Math.round(ctx.parsed/base.length*100)}%)`}}}}});
  buildSideStats('chart-estatus-stats',eL,eD,eC,base.length);

  destroyChart('localidad');
  const loc=countBy(base,'localidad',999);
  const locColors=makeColors(loc.labels.length,0.85);
  const locBorders=makeColors(loc.labels.length);
  const locH=Math.max(200,loc.labels.length*26);
  document.getElementById('chart-localidad').parentElement.style.height=locH+'px';
  chartInstances['localidad']=new Chart(document.getElementById('chart-localidad'),{type:'bar',data:{labels:loc.labels,datasets:[{label:'Vacantes',data:loc.data,backgroundColor:locColors,borderColor:locBorders,borderWidth:1.5,borderRadius:4}]},options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,layout:{padding:{right:56}},plugins:{legend:{display:false},datalabels:{display:true,anchor:'end',align:'end',clamp:true,color:'#003087',font:{size:10,weight:'700'},formatter:v=>v+' ('+(base.length?Math.round(v/base.length*100):0)+'%)'}},scales:{x:{grid:{color:'#EEF0F5'},ticks:{font:{size:10}}},y:{grid:{display:false},ticks:{font:{size:10}}}}}});

  destroyChart('grupos');
  const grp=countBy(base,'grupo',10);
  const grpColors=makeColors(grp.labels.length);
  chartInstances['grupos']=new Chart(document.getElementById('chart-grupos'),{type:'doughnut',data:{labels:grp.labels,datasets:[{data:grp.data,backgroundColor:grpColors,borderWidth:2,borderColor:'#fff',hoverOffset:6}]},options:{responsive:true,maintainAspectRatio:false,cutout:'55%',plugins:{legend:{display:false},datalabels:{display:false},tooltip:{callbacks:{label:ctx=>` ${ctx.label}: ${ctx.parsed} (${Math.round(ctx.parsed/base.length*100)}%)`}}}}});
  buildSideStats('chart-grupos-stats',grp.labels,grp.data,grpColors,base.length);

  destroyChart('areas');
  const area=countBy(base,'clasificacion',10);
  const areaColors=makeColors(area.labels.length);
  chartInstances['areas']=new Chart(document.getElementById('chart-areas'),{type:'pie',data:{labels:area.labels,datasets:[{data:area.data,backgroundColor:areaColors,borderWidth:2,borderColor:'#fff',hoverOffset:6}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false},datalabels:{display:false},tooltip:{callbacks:{label:ctx=>` ${ctx.label}: ${ctx.parsed} (${Math.round(ctx.parsed/base.length*100)}%)`}}}}});
  buildSideStats('chart-areas-stats',area.labels,area.data,areaColors,base.length);
}

// =============================================================
// ── PESTAÑA DETALLE DE POSICIONES ──
// =============================================================
let detAllRows = [];       // todas las filas leídas del Excel
let detFiltered = [];      // tras aplicar filtros
let detComplejosOpen = {}; // { complejo: bool }
let detGerentesOpen = {};  // { complejo+'||'+gerente: bool }

// drag & drop detalle
function detDragOver(e){e.preventDefault();document.getElementById('det-upload-section').classList.add('drag-over')}
function detDragLeave(){document.getElementById('det-upload-section').classList.remove('drag-over')}
function detDrop(e){e.preventDefault();detDragLeave();const f=e.dataTransfer.files[0];if(f)handleDetFile(f)}

function detResetAll(){
  detAllRows=[];detFiltered=[];detComplejosOpen={};detGerentesOpen={};
  projectResetAll(false);
  document.getElementById('det-upload-section').style.display='block';
  document.getElementById('det-file-badge').style.display='none';
  document.getElementById('det-dashboard').style.display='none';
  document.getElementById('det-error-box').style.display='none';
  document.getElementById('detFileInput').value='';
  document.getElementById('det-complejos-wrap').innerHTML='';
}

function handleDetFile(file){
  if(!file)return;
  detShowError('');
  document.getElementById('det-fname-text').textContent=file.name;
  document.getElementById('det-file-badge').style.display='flex';
  document.getElementById('det-upload-section').style.display='none';

  const reader=new FileReader();
  reader.onload=function(e){
    try{
      const wb=XLSX.read(e.target.result,{type:'array',cellFormula:false,cellNF:false,raw:false});
      // Buscar hoja Rep Posiciones
      const repName=wb.SheetNames.find(n=>norm(n).includes('rep pos')||norm(n).includes('rep posicion'));
      if(!repName){detShowError('No se encontró la hoja "Rep Posiciones" en el archivo.');return;}
      const mx=XLSX.utils.sheet_to_json(wb.Sheets[repName],{header:1,defval:''});
      if(!mx||mx.length<2){detShowError('La hoja Rep Posiciones está vacía.');return;}

      // Leer hoja Justificación para vigencias
      const justMap={};
      const justSN=wb.SheetNames.find(n=>norm(n).includes('justif'));
      if(justSN){
        const jr=XLSX.utils.sheet_to_json(wb.Sheets[justSN],{header:1,defval:''});
        for(let i=1;i<jr.length;i++){
          const r=jr[i];if(!r[0]&&r[0]!==0)continue;
          const da=parseFloat(r[8]);
          justMap[String(r[0])]={diasAsig:isNaN(da)?9999:Math.round(da),tipo:String(r[5]||'').trim(),comentario:String(r[4]||'').trim()};
        }
      }

      // Encabezados fila 1 (índice 0)
      const hdr=mx[0].map(c=>norm(c));
      function ci2(kws,fb){for(let i=0;i<hdr.length;i++){if(hdr[i]&&kws.some(k=>hdr[i].includes(k)))return i;}return fb;}

      const C={
        pos:ci2(['posicion','posición'],0), desc:ci2(['descripcion posicion','descripción posición'],1),
        unidad:ci2(['descripcion unidad','descripción unidad'],3),
        ccosto:ci2(['descripcion centro de costos','descripción centro de costos'],5),
        nivel:ci2(['nivel'],6), sub:ci2(['subnivel'],7),
        grupo:ci2(['descripcion grupo','descripción grupo'],13),
        area:ci2(['descripcion area','descripción área','descripción area'],15),
        areaCod:ci2(['area','área'],14),
        localidad:ci2(['descripcion localidad','descripción localidad'],19),
        estado:ci2(['estado'],20),
        empleado:ci2(['nombre empleado'],22),
        dias:ci2(['dias sin ocupar','días sin ocupar'],24),
        jrh:ci2(['jefe rh','jefe de rh'],27),
        vig:ci2(['vigencia'],28),
        estatus:ci2(['estatus'],29),
        diasVenc:ci2(['dias vencidas','días vencidas'],30),
        razon:ci2(['razon','razón'],31),
        tipoJust:ci2(['tipo de just','tipo just'],32),
        clasif:ci2(['clasific','clasificacion','clasificación'],33),
        complejo:ci2(['complejo'],34),
        gerente:ci2(['gerente'],35),
      };

      detAllRows=[];
      for(let i=1;i<mx.length;i++){
        const r=mx[i];
        if(!r[C.pos]&&r[C.pos]!==0)continue;
        const posStr=String(r[C.pos]||'').trim();
        const estadoRaw=String(r[C.estado]||'').trim().toUpperCase();
        const dias=isNaN(parseFloat(r[C.dias]))?0:Math.round(parseFloat(r[C.dias]));
        const jInfo=justMap[posStr];

        // Vigencia
        let vigDias=9999;
        if(jInfo&&jInfo.diasAsig&&jInfo.diasAsig!==9999){vigDias=jInfo.diasAsig;}
        else{const vr=r[C.vig];const vn=parseFloat(String(vr||'').replace(/[^0-9.]/g,''));if(!isNaN(vn)&&vn>0&&vn<9000)vigDias=Math.round(vn);else vigDias=getVig(r[C.sub]);}

        // Semáforo solo para vacantes
        let semaforo='';
        if(estadoRaw==='VACANTE'||estadoRaw==='VACANTE '){
          if(vigDias>=9999){semaforo=getEst(dias,getVig(r[C.sub]))||'Verde';}
          else{const pct=dias/vigDias;semaforo=pct<=0.75?'Verde':pct<=0.87?'Amarillo':'Rojo';}
        }

        const complejoRaw=String(r[C.complejo]||'').trim();
        const complejo=reclassifyZona(r[C.localidad],r[C.unidad],r[C.ccosto],r[C.desc],complejoRaw)||complejoRaw||'Sin Complejo';
        const gerente=String(r[C.gerente]||'').trim()||'Sin Gerente';
        const tipoJust=String(r[C.tipoJust]||'').trim();
        const razon=String(r[C.razon]||'').trim();
        const justDisplay=jInfo?.comentario&&jInfo.comentario!=='0'?jInfo.comentario:razon&&razon!=='0'?razon:tipoJust||'—';

        detAllRows.push({
          posicion:posStr, descripcion:String(r[C.desc]||'').trim(),
          unidad:String(r[C.unidad]||'').trim(),
          nivel:String(r[C.nivel]||'').trim(), subnivel:String(r[C.sub]||'').trim(),
          grupo:String(r[C.grupo]||'').trim(), area:String(r[C.area]||'').trim(),
          localidad:String(r[C.localidad]||'').trim(),
          areaCod:String(r[C.areaCod]||'').trim().toUpperCase(),
          estado:estadoRaw, empleado:String(r[C.empleado]||'').trim(),
          dias, vigDias, estatus:String(r[C.estatus]||'').trim(),
          diasVenc:isNaN(parseFloat(r[C.diasVenc]))?0:Math.round(parseFloat(r[C.diasVenc])),
          tipoJust, justificacion:justDisplay,
          clasif:String(r[C.clasif]||'').trim().toUpperCase(),
          complejo, gerente, semaforo,
        });
      }

      if(!detAllRows.length){detShowError('No se encontraron filas con datos.');return;}
      document.getElementById('det-dashboard').style.display='block';
      detApplyFilters();
      projectLoadFromDetalle(file.name);
    }catch(err){detShowError('Error al procesar: '+err.message);console.error(err);}
  };
  reader.readAsArrayBuffer(file);
}

function detShowError(msg){
  const el=document.getElementById('det-error-box');
  if(!msg){el.style.display='none';return;}
  el.textContent='Aviso: '+msg;el.style.display='block';
}

function detClearFilters(){
  document.getElementById('det-sel-complejo').value='';
  document.getElementById('det-sel-estado').value='';
  document.getElementById('det-sel-clasif').value='';
  document.getElementById('det-search').value='';
  detApplyFilters();
}

// % de Cobertura por Complejo + Nacional — SIEMPRE calculado sobre TODAS
// las posiciones del archivo (detAllRows), sin importar los filtros de
// pantalla activos, igual que el recuadro de Circunstancias en Análisis de
// Proyectos. "Nacional" es el total de los 5 complejos juntos.
function detRenderCoberturaChart(){
  const canvas=document.getElementById('det-chart-cobertura');
  if(!canvas||!detAllRows.length)return;
  const ordCompl=['Norte','Sur','Centro','Corporativo','Polanco'];
  const porComplejo=ordCompl.map(z=>{
    const rows=detAllRows.filter(r=>r.complejo===z);
    const total=rows.length;
    const ocu=rows.filter(r=>r.estado==='OCUPADA'||r.estado==='OCUPADA ').length;
    const pct=total?(ocu/total*100):0;
    return{label:z,total,ocu,vac:total-ocu,pct};
  }).filter(z=>z.total>0);
  const totalNal=detAllRows.length;
  const ocuNal=detAllRows.filter(r=>r.estado==='OCUPADA'||r.estado==='OCUPADA ').length;
  const nacional={label:'Nacional',total:totalNal,ocu:ocuNal,vac:totalNal-ocuNal,pct:totalNal?(ocuNal/totalNal*100):0};
  const items=[...porComplejo,nacional];
  const pctColor=p=>p>=97.5?'#166534':p>=95?'#92400E':'#9F1239';
  const colors=items.map(it=>it.label==='Nacional'?'#003087':pctColor(it.pct));

  destroyChart('det-cobertura');
  chartInstances['det-cobertura']=new Chart(canvas,{
    type:'bar',
    data:{labels:items.map(it=>it.label),datasets:[{data:items.map(it=>+it.pct.toFixed(1)),backgroundColor:colors,borderRadius:6,maxBarThickness:56}]},
    options:{responsive:true,maintainAspectRatio:false,layout:{padding:{top:10}},scales:{y:{min:0,max:100,ticks:{callback:v=>v+'%'},grid:{color:'#EEF0F5'}},x:{grid:{display:false},ticks:{font:{size:11,weight:'700'}}}},plugins:{legend:{display:false},datalabels:{display:false},tooltip:{callbacks:{label:ctx=>{const it=items[ctx.dataIndex];return ` ${it.pct.toFixed(1)}% cubierto (${it.ocu} de ${it.total})`;}}}}}
  });

  const statsEl=document.getElementById('det-chart-cobertura-stats');
  if(statsEl){
    statsEl.innerHTML=items.map(it=>{
      const c=it.label==='Nacional'?'#003087':pctColor(it.pct);
      return`<div style="display:flex;align-items:center;gap:6px;font-size:11px;line-height:1.3"><span style="width:9px;height:9px;border-radius:50%;flex-shrink:0;display:inline-block;background:${c}"></span><span style="flex:1;color:#3A3F4B;white-space:nowrap;font-weight:${it.label==='Nacional'?'800':'600'}">${escHtml(it.label)}</span><span style="font-weight:800;color:#111;white-space:nowrap;font-size:10.5px">${it.pct.toFixed(1)}% <span style="font-weight:700;color:#5A6070">(${it.ocu}/${it.total})</span></span></div>`;
    }).join('');
  }
}

function detApplyFilters(){
  const cComplejo=document.getElementById('det-sel-complejo').value;
  const cEstado=document.getElementById('det-sel-estado').value;
  const cClasif=document.getElementById('det-sel-clasif').value;
  const q=norm(document.getElementById('det-search').value);

  detFiltered=detAllRows.filter(r=>{
    if(cComplejo&&r.complejo!==cComplejo)return false;
    if(cEstado&&r.estado!==cEstado)return false;
    if(cClasif&&r.clasif!==cClasif)return false;
    if(q&&!(norm(r.posicion).includes(q)||norm(r.descripcion).includes(q)||norm(r.localidad).includes(q)||norm(r.gerente).includes(q)||norm(r.area).includes(q)||norm(r.grupo).includes(q)||norm(r.unidad).includes(q)))return false;
    return true;
  });

  detUpdateKPIs();
  detRenderComplejos();
  detRenderCoberturaChart();
}

function detUpdateKPIs(){
  const total=detFiltered.length;
  const vac=detFiltered.filter(r=>r.estado==='VACANTE'||r.estado==='VACANTE ').length;
  const ocu=detFiltered.filter(r=>r.estado==='OCUPADA'||r.estado==='OCUPADA ').length;
  const planta=detFiltered.filter(r=>r.clasif==='PLANTA').length;
  const eventual=detFiltered.filter(r=>r.clasif==='EVENTUAL').length;
  document.getElementById('det-total').textContent=total.toLocaleString();
  document.getElementById('det-vacantes').textContent=vac.toLocaleString();
  document.getElementById('det-ocupadas').textContent=ocu.toLocaleString();
  document.getElementById('det-planta').textContent=planta.toLocaleString();
  document.getElementById('det-eventual').textContent=eventual.toLocaleString();
  document.getElementById('det-export-sub').textContent=`Genera reporte PDF con ${total.toLocaleString()} posiciones filtradas`;
}

function detRenderComplejos(){
  const wrap=document.getElementById('det-complejos-wrap');
  wrap.innerHTML='';

  // Agrupar por complejo
  const complMap={};
  detFiltered.forEach(r=>{
    if(!complMap[r.complejo])complMap[r.complejo]={rows:[],vac:0,ocu:0};
    complMap[r.complejo].rows.push(r);
    if(r.estado==='VACANTE'||r.estado==='VACANTE ')complMap[r.complejo].vac++;
    else if(r.estado==='OCUPADA'||r.estado==='OCUPADA ')complMap[r.complejo].ocu++;
  });

  const ordCompl=['Norte','Centro','Sur','Corporativo','Polanco'];
  const complejos=[...ordCompl.filter(c=>complMap[c]),...Object.keys(complMap).filter(c=>!ordCompl.includes(c)).sort()];

  complejos.forEach(complejo=>{
    const cData=complMap[complejo];
    const isOpen=detComplejosOpen[complejo]!==false; // abierto por defecto
    const clsName=complejo.toLowerCase();
    const badgeClass=`complejo-badge-${clsName}`;

    // Agrupar por gerente
    const gerMap={};
    cData.rows.forEach(r=>{
      if(!gerMap[r.gerente])gerMap[r.gerente]={rows:[],vac:0,ocu:0};
      gerMap[r.gerente].rows.push(r);
      if(r.estado==='VACANTE'||r.estado==='VACANTE ')gerMap[r.gerente].vac++;
      else if(r.estado==='OCUPADA'||r.estado==='OCUPADA ')gerMap[r.gerente].ocu++;
    });
    const gerentes=Object.keys(gerMap).sort();

    // Construir acordeón gerentes
    const gerentesHtml=gerentes.map(gerente=>{
      const gData=gerMap[gerente];
      const gKey=complejo+'||'+gerente;
      const gOpen=detGerentesOpen[gKey]===true;
      const gPct=gData.rows.length?((gData.ocu/gData.rows.length)*100).toFixed(1):'0.0';
      const pctClass=gPct>=97.5?'gs-ocu':gPct>=95?'gs-pct-mid':'gs-vac';
      // vacantes primero (es lo primordial a visualizar)
      const sortedRows=[...gData.rows].sort((a,b)=>{
        const av=(a.estado==='VACANTE'||a.estado==='VACANTE ')?0:1;
        const bv=(b.estado==='VACANTE'||b.estado==='VACANTE ')?0:1;
        return av-bv;
      });

      const rowsHtml=sortedRows.map(r=>{
        const trClass=r.estado==='VACANTE'||r.estado==='VACANTE '?'tr-vacante':'tr-ocupada';
        let semHtml='';
        if(r.semaforo){const sc={Rojo:'det-sem-rojo',Amarillo:'det-sem-amarillo',Verde:'det-sem-verde'}[r.semaforo]||'';semHtml=`<span class="det-sem ${sc}">${r.semaforo}</span>`;}
        else{semHtml=`<span class="det-sem det-sem-ocu">Ocupada</span>`;}
        const estHtml=r.estado==='VACANTE'||r.estado==='VACANTE '?`<span class="det-estado-v">● VACANTE</span>`:`<span class="det-estado-o">● OCUPADA</span>`;
        return`<tr class="${trClass}"><td style="white-space:nowrap;font-weight:700;color:#003087">${escHtml(r.posicion)}</td><td style="max-width:180px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${escHtml(r.descripcion)}">${escHtml(r.descripcion)}</td><td>${estHtml}</td><td>${semHtml}</td><td style="white-space:nowrap;font-weight:600;color:#003087">${r.dias>0?r.dias+'d':'—'}</td><td style="max-width:150px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${escHtml(r.localidad)}">${escHtml(r.localidad)}</td><td style="max-width:130px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${escHtml(r.area)}">${escHtml(r.area)}</td><td style="max-width:120px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${escHtml(r.grupo)}">${escHtml(r.grupo)}</td><td>${escHtml(r.clasif)}</td><td style="max-width:140px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${escHtml(r.justificacion)}">${escHtml(r.justificacion)}</td><td style="font-size:10px;color:#999">${escHtml(r.estatus)}</td></tr>`;
      }).join('');

      return`<div class="gerente-section">
        <div class="gerente-header" onclick="detToggleGerente('${escHtml(complejo)}','${escHtml(gerente)}',this)">
          <div class="gerente-name">${escHtml(gerente)}</div>
          <div style="display:flex;align-items:center;gap:10px">
            <div class="gerente-stats">
              <span class="gerente-stat gs-tot">${gData.rows.length} pos.</span>
              <span class="gerente-stat gs-vac">${gData.vac} vac.</span>
              <span class="gerente-stat gs-ocu">${gData.ocu} ocu.</span>
              <span class="gerente-stat ${pctClass}">${gPct}% cubierto</span>
            </div>
            <span class="gerente-chev ${gOpen?'open':''}">▼</span>
          </div>
        </div>
        <div class="gerente-body ${gOpen?'open':''}">
          <div class="det-tbl-wrap">
            <table class="det-tbl">
              <thead><tr>
                <th>No. Posicion</th><th>Descripcion</th><th>Estado</th><th>Semáforo</th>
                <th>Días</th><th>Localidad</th><th>Área</th><th>Grupo</th>
                <th>Clasif.</th><th>Justificación</th><th>Estatus</th>
              </tr></thead>
              <tbody>${rowsHtml}</tbody>
            </table>
          </div>
        </div>
      </div>`;
    }).join('');

    const section=document.createElement('div');
    section.className='complejo-section';
    section.innerHTML=`
      <div class="complejo-header ${clsName}" onclick="detToggleComplejo('${escHtml(complejo)}',this)">
        <div class="complejo-title">
          <span>${escHtml(complejo)}</span>
          <span class="${badgeClass}">${gerentes.length} gerentes</span>
        </div>
        <div style="display:flex;align-items:center;gap:14px">
          <div class="complejo-stats">
            <span class="complejo-stat" style="color:var(--p-blue);font-weight:700">${cData.rows.length.toLocaleString()} posiciones</span>
            <span class="complejo-stat cs-vac">${cData.vac} vacantes</span>
            <span class="complejo-stat cs-ocu">${cData.ocu} ocupadas</span>
          </div>
          <span class="complejo-chevron ${isOpen?'open':''}">▼</span>
        </div>
      </div>
      <div class="complejo-body ${isOpen?'open':''}">
        <div class="complejo-body-inner">${gerentesHtml}</div>
      </div>`;
    wrap.appendChild(section);
  });

  if(complejos.length===0){
    wrap.innerHTML='<div style="text-align:center;padding:3rem;color:var(--p-gray);font-size:14px">No hay datos que coincidan con los filtros aplicados.</div>';
  }
}

function detToggleComplejo(complejo, headerEl){
  const section=headerEl.closest('.complejo-section');
  const body=section.querySelector('.complejo-body');
  const chev=section.querySelector('.complejo-chevron');
  const isOpen=body.classList.contains('open');
  body.classList.toggle('open');
  chev.classList.toggle('open');
  detComplejosOpen[complejo]=!isOpen;
}

function detToggleGerente(complejo, gerente, headerEl){
  const section=headerEl.closest('.gerente-section');
  const body=section.querySelector('.gerente-body');
  const chev=section.querySelector('.gerente-chev');
  const isOpen=body.classList.contains('open');
  body.classList.toggle('open');
  chev.classList.toggle('open');
  detGerentesOpen[complejo+'||'+gerente]=!isOpen;
}

// =============================================================
// ── EXPORTAR PDF DETALLE ──
// =============================================================

// =============================================================
// ── PESTAÑA ANÁLISIS DE PROYECTOS ──
// =============================================================
let projectRows = [];
let projectFiltered = [];
let projectActions = {};
let projComplejosOpen = {};
let projGerentesOpen = {};
let projActiveComplejo = '';
let projActiveCircunstancia = '';

function handleProjectFile(file){
  if(!file)return;
  handleDetFile(file);
}

function projectResetAll(clearFileInput=true){
  projectRows=[];projectFiltered=[];projComplejosOpen={};projGerentesOpen={};projActiveComplejo='';projActiveCircunstancia='';
  const empty=document.getElementById('project-empty');
  const dash=document.getElementById('project-dashboard');
  if(empty)empty.style.display='block';
  if(dash)dash.style.display='none';
  const wrap=document.getElementById('project-complejos-wrap');
  if(wrap)wrap.innerHTML='';
  const sel=document.getElementById('project-sel-gerente');
  if(sel)sel.innerHTML='<option value="">Todos</option>';
  const locSel=document.getElementById('project-sel-localidad');
  if(locSel)locSel.innerHTML='<option value="">Todas</option>';
  const locCount=document.getElementById('project-localidad-count');
  if(locCount)locCount.textContent='';
  const zbar=document.getElementById('project-zona-bar');
  if(zbar)zbar.innerHTML='';
  if(clearFileInput){const inp=document.getElementById('projectFileInput');if(inp)inp.value='';}
  projectUpdateOtrasCircunstancias();
}

function projectLoadFromDetalle(fileName){
  if(!detAllRows.length){projectResetAll(false);return;}
  projectRows=projectBuildRows(detAllRows);
  projectFillGerentes();
  projectFillLocalidades();
  const empty=document.getElementById('project-empty');
  const dash=document.getElementById('project-dashboard');
  if(empty)empty.style.display='none';
  if(dash)dash.style.display='block';
  projectApplyFilters();
}

// Cada posición (Área 2M/4M) es su propia fila — ya NO se agrupan varias
// posiciones bajo un mismo "proyecto" combinado. La circunstancia se
// clasifica leyendo la columna "Tipo de Just" (categoría limpia del Excel:
// Proyecto / Incapacidad / Global Talent / Capacitación), no la Razón
// (texto libre), porque es la columna que realmente trae la categoría.
const CIRC_LABELS={proyecto:'Proyecto',incapacidad:'Incapacidad',globaltalent:'Global Talent',capacitacion:'Capacitación',otro:'Otro'};
function projectCircunstancia(tipoJust){
  const t=norm(tipoJust);
  if(t.includes('global talent')||t.includes('globaltalent'))return 'globaltalent';
  if(t.includes('incapacidad'))return 'incapacidad';
  if(t.includes('capacitacion'))return 'capacitacion';
  if(t.includes('proyecto'))return 'proyecto';
  return 'otro';
}
function projectBuildRows(rows){
  return rows.filter(r=>['2M','4M'].includes(r.areaCod)).map((r,i)=>{
    const est=projectStatus(r);
    const complejo=r.complejo||'Sin Complejo';
    const gerente=r.gerente||'Sin Gerente';
    const razon=(r.justificacion&&r.justificacion!=='—')?r.justificacion:(r.descripcion||'Sin razón / proyecto');
    const circunstancia=projectCircunstancia(r.tipoJust);
    const estado=(r.estado||'').trim().toUpperCase();
    const key=(r.posicion&&String(r.posicion).trim())?String(r.posicion).trim():('sinpos-'+i);
    return{
      key,posicion:String(r.posicion||'—').trim(),complejo,estatus:est,gerente,razon,
      tipoJust:r.tipoJust||'',circunstancia,estado,
      descripcion:r.descripcion||'',localidad:r.localidad||'',
      diasVenc:Number(r.diasVenc||0),dias:Number(r.dias||0),
      accion:projectActions[key]||'',
    };
  }).sort((a,b)=>{
    const co=a.complejo.localeCompare(b.complejo,'es');
    if(co)return co;
    const eo=(a.estatus==='VENCIDA'?0:1)-(b.estatus==='VENCIDA'?0:1);
    if(eo)return eo;
    return a.gerente.localeCompare(b.gerente,'es')||a.razon.localeCompare(b.razon,'es');
  });
}

function projectStatus(r){
  const est=norm(r.estatus||'');
  if(est.includes('venc'))return 'VENCIDA';
  if(Number(r.diasVenc||0)>0)return 'VENCIDA';
  return 'VIGENTE';
}

function projectSuggestedAction(item){
  if(item.estatus==='VENCIDA')return 'Revisar si son requeridas posiciones para solicitar planta; en caso contrario cancelar';
  return 'Solicitar ampliación o justificación del proyecto según aplique';
}

function renderProjZonaBar(){
  const zbar=document.getElementById('project-zona-bar');
  if(!zbar)return;
  const orderZ=['Norte','Sur','Centro','Corporativo','Polanco'];
  const presentes=[...new Set(projectRows.map(r=>r.complejo))];
  const complejos=[...orderZ.filter(z=>presentes.includes(z)),...presentes.filter(z=>!orderZ.includes(z)).sort((a,b)=>a.localeCompare(b,'es'))];
  if(projActiveComplejo&&!complejos.includes(projActiveComplejo))projActiveComplejo='';
  zbar.innerHTML='<span class="zona-label">Complejo</span>'+
    `<button class="zbtn${projActiveComplejo===''?' z-active':''}" onclick="setProjZona('',this)">Todos <span class="zona-count">${projectRows.length}</span></button>`+
    complejos.map(z=>`<button class="zbtn${projActiveComplejo===z?' z-active':''}" onclick="setProjZona('${escHtml(z)}',this)">${escHtml(z)} <span class="zona-count">${projectRows.filter(r=>r.complejo===z).length}</span></button>`).join('');
}

function setProjZona(val,el){
  projActiveComplejo=val;
  document.querySelectorAll('#project-zona-bar .zbtn').forEach(b=>b.classList.remove('z-active'));
  el.classList.add('z-active');
  projectApplyFilters();
}

function projectFillGerentes(){
  const sel=document.getElementById('project-sel-gerente');
  if(!sel)return;
  const prev=sel.value;
  const gerentes=[...new Set(projectRows.map(r=>r.gerente))].sort((a,b)=>a.localeCompare(b,'es'));
  sel.innerHTML='<option value="">Todos</option>'+gerentes.map(g=>`<option value="${escHtml(g)}">${escHtml(g)}</option>`).join('');
  if(gerentes.includes(prev))sel.value=prev;
}

function projectFillLocalidades(){
  const sel=document.getElementById('project-sel-localidad');
  if(!sel)return;
  const prev=sel.value;
  const localidades=[...new Set(projectRows.map(r=>r.localidad||'—'))].sort((a,b)=>a.localeCompare(b,'es'));
  sel.innerHTML='<option value="">Todas</option>'+localidades.map(l=>`<option value="${escHtml(l)}">${escHtml(l)}</option>`).join('');
  if(localidades.includes(prev))sel.value=prev;
}

function projectApplyFilters(){
  renderProjZonaBar();
  const comp=projActiveComplejo;
  const est=document.getElementById('project-sel-estatus')?.value||'';
  const ger=document.getElementById('project-sel-gerente')?.value||'';
  const loc=document.getElementById('project-sel-localidad')?.value||'';
  const q=norm(document.getElementById('project-search')?.value||'');
  projectFiltered=projectRows.filter(r=>{
    if(comp&&r.complejo!==comp)return false;
    if(est&&r.estatus!==est)return false;
    if(ger&&r.gerente!==ger)return false;
    if(loc&&(r.localidad||'—')!==loc)return false;
    if(projActiveCircunstancia&&r.circunstancia!==projActiveCircunstancia)return false;
    if(q&&!(norm(r.gerente).includes(q)||norm(r.razon).includes(q)||norm(r.accion).includes(q)||norm(r.posicion).includes(q)||norm(r.descripcion).includes(q)||norm(r.localidad).includes(q)))return false;
    return true;
  });
  document.getElementById('project-localidad-count').textContent=loc?`${projectFiltered.filter(r=>r.estado==='VACANTE').length} vacantes en ${loc} · ${projectFiltered.length} posiciones`:`${projectFiltered.filter(r=>r.estado==='VACANTE').length} vacantes en todas las localidades`;
  renderProjCircChip();
  projectUpdateKPIs();
  projectRenderTable();
}

function projectClearFilters(){
  projActiveComplejo='';
  projActiveCircunstancia='';
  document.getElementById('project-sel-estatus').value='';
  document.getElementById('project-sel-gerente').value='';
  document.getElementById('project-sel-localidad').value='';
  document.getElementById('project-search').value='';
  projectApplyFilters();
}

// Clic en uno de los 4 números del recuadro "Otras Posiciones por
// Circunstancia" filtra la tabla de abajo para mostrar solo esos proyectos.
// Un segundo clic sobre la misma categoría (o clic en "Total") lo quita.
function setProjCircunstancia(val){
  projActiveCircunstancia=(projActiveCircunstancia===val)?'':val;
  projectApplyFilters();
  const wrap=document.getElementById('project-complejos-wrap');
  if(wrap)wrap.scrollIntoView({behavior:'smooth',block:'start'});
}

function renderProjCircChip(){
  const chip=document.getElementById('project-circ-filter-chip');
  if(!chip)return;
  if(!projActiveCircunstancia){chip.style.display='none';chip.innerHTML='';return;}
  const label=CIRC_LABELS[projActiveCircunstancia]||projActiveCircunstancia;
  chip.style.display='block';
  chip.innerHTML=`<span style="display:inline-flex;align-items:center;gap:8px;background:#6D28D9;color:#fff;padding:6px 14px;border-radius:20px;font-size:12px;font-weight:700">Filtrando proyectos por circunstancia: ${escHtml(label)} <span style="cursor:pointer;font-weight:900" onclick="setProjCircunstancia('${projActiveCircunstancia}')" title="Quitar filtro">✕</span></span>`;
}

function projectUpdateKPIs(){
  const totalProjects=projectFiltered.length;
  const vencidas=projectFiltered.filter(r=>r.estatus==='VENCIDA').length;
  const vigentes=projectFiltered.filter(r=>r.estatus==='VIGENTE').length;
  const vacantes=projectFiltered.filter(r=>r.estado==='VACANTE').length;
  document.getElementById('project-total-projects').textContent=totalProjects.toLocaleString();
  document.getElementById('project-vencidas').textContent=vencidas.toLocaleString();
  document.getElementById('project-vigentes').textContent=vigentes.toLocaleString();
  document.getElementById('project-vacantes').textContent=vacantes.toLocaleString();
  projectUpdateOtrasCircunstancias();
}

// Recuadro "Otras Posiciones por Circunstancia": ahora cada posición
// individual ya trae su categoría precalculada (r.circunstancia, leída de
// la columna "Tipo de Just"), así que solo se cuenta directo — el total
// de las 4 categorías debe coincidir con "Proyectos" de arriba (todas las
// posiciones Área 2M/4M caen en alguna de estas 4 categorías). Respeta los
// mismos filtros activos (Estatus/Gerente/Complejo/búsqueda/circunstancia).
function projectUpdateOtrasCircunstancias(){
  const box=document.getElementById('project-otras-circunstancias');
  if(!box)return;
  const counts={proyecto:0,incapacidad:0,globaltalent:0,capacitacion:0};
  projectFiltered.forEach(r=>{if(counts[r.circunstancia]!==undefined)counts[r.circunstancia]++;});
  const total=counts.proyecto+counts.incapacidad+counts.globaltalent+counts.capacitacion;
  document.getElementById('poc-proyecto').textContent=counts.proyecto.toLocaleString();
  document.getElementById('poc-incapacidad').textContent=counts.incapacidad.toLocaleString();
  document.getElementById('poc-globaltalent').textContent=counts.globaltalent.toLocaleString();
  document.getElementById('poc-capacitacion').textContent=counts.capacitacion.toLocaleString();
  document.getElementById('poc-total').textContent=total.toLocaleString();
  ['proyecto','incapacidad','globaltalent','capacitacion'].forEach(k=>{
    const el=document.getElementById('poc-item-'+k);
    if(el)el.classList.toggle('poc-active',projActiveCircunstancia===k);
  });
}

function projectSaveAction(key,val){
  projectActions[key]=val;
  const row=projectRows.find(r=>r.key===key);
  if(row)row.accion=val;
}

function projectRenderTable(){
  const wrap=document.getElementById('project-complejos-wrap');
  if(!wrap)return;
  if(!projectFiltered.length){wrap.innerHTML='<div style="text-align:center;padding:3rem;color:var(--p-gray);font-size:14px">No hay proyectos que coincidan con los filtros o el archivo no trae posiciones Eventuales con Área 2M o 4M.</div>';return;}

  // Agrupar solo visualmente por Complejo -> Gerencia (cada fila interna
  // sigue siendo UNA posición individual, no se combinan varias en una).
  const complMap={};
  projectFiltered.forEach(r=>{
    if(!complMap[r.complejo])complMap[r.complejo]={gerentes:{}};
    if(!complMap[r.complejo].gerentes[r.gerente])complMap[r.complejo].gerentes[r.gerente]=[];
    complMap[r.complejo].gerentes[r.gerente].push(r);
  });

  const ordCompl=['Norte','Centro','Sur','Corporativo','Polanco'];
  const complejos=[...ordCompl.filter(c=>complMap[c]),...Object.keys(complMap).filter(c=>!ordCompl.includes(c)).sort()];

  let grand={ocupada:0,vacante:0,total:0};
  let html='';

  complejos.forEach(complejo=>{
    const cData=complMap[complejo];
    const cIsOpen=projComplejosOpen[complejo]!==false;
    const clsName=complejo.toLowerCase();
    const badgeClass=`complejo-badge-${clsName}`;
    let cSum={ocupada:0,vacante:0,total:0};
    const gerentes=Object.keys(cData.gerentes).sort();

    const gerentesHtml=gerentes.map(gerente=>{
      const gRows=[...cData.gerentes[gerente]].sort((a,b)=>(a.estado==='VACANTE'?0:1)-(b.estado==='VACANTE'?0:1)||(a.estatus==='VENCIDA'?0:1)-(b.estatus==='VENCIDA'?0:1)||a.razon.localeCompare(b.razon,'es'));
      const gKey=complejo+'||'+gerente;
      const gOpen=projGerentesOpen[gKey]===true;
      let gSum={ocupada:0,vacante:0,total:0};
      const rowsHtml=gRows.map(r=>{
        gSum.total++;
        if(r.estado==='VACANTE')gSum.vacante++;else gSum.ocupada++;
        const rowClass=r.estatus==='VENCIDA'?'project-vencida':'project-vigente';
        const estadoHtml=r.estado==='VACANTE'?'<span style="color:#9F1239;font-weight:800">Vacante</span>':'<span style="color:#166534;font-weight:700">Ocupada</span>';
        const circLabel=CIRC_LABELS[r.circunstancia]||r.tipoJust||'—';
        return`<tr class="${rowClass}"><td class="project-status-cell">${r.estatus}</td><td style="white-space:nowrap;font-weight:700;color:#003087">${escHtml(r.posicion)} <span style="font-weight:500;color:#5A6070;white-space:normal">${escHtml(r.descripcion||'')}</span></td><td class="project-razon">${escHtml(r.razon)}</td><td>${escHtml(r.localidad||'—')}</td><td><span class="justif-tag">${escHtml(circLabel)}</span></td><td class="project-num">${estadoHtml}</td><td class="project-num">${r.diasVenc||''}</td><td><textarea class="project-action" placeholder="Escribe el plan de acción" data-key="${escHtml(r.key)}" onchange="projectSaveAction(this.dataset.key,this.value)">${escHtml(r.accion)}</textarea></td></tr>`;
      }).join('');
      cSum.ocupada+=gSum.ocupada;cSum.vacante+=gSum.vacante;cSum.total+=gSum.total;
      const gPct=gSum.total?((gSum.ocupada/gSum.total)*100).toFixed(1):'0.0';
      const pctClass=gPct>=97.5?'gs-ocu':gPct>=95?'gs-pct-mid':'gs-vac';
      return`<div class="gerente-section">
        <div class="gerente-header" onclick="projToggleGerente('${escHtml(complejo)}','${escHtml(gerente)}',this)">
          <div class="gerente-name">${escHtml(gerente)}</div>
          <div style="display:flex;align-items:center;gap:10px">
            <div class="gerente-stats">
              <span class="gerente-stat gs-tot">${gSum.total} pos.</span>
              <span class="gerente-stat gs-vac">${gSum.vacante} vac.</span>
              <span class="gerente-stat gs-ocu">${gSum.ocupada} ocu.</span>
              <span class="gerente-stat ${pctClass}">${gPct}% cubierto</span>
            </div>
            <span class="gerente-chev ${gOpen?'open':''}">▼</span>
          </div>
        </div>
        <div class="gerente-body ${gOpen?'open':''}">
          <div class="project-tbl-wrap" style="max-height:none">
            <table class="project-tbl" style="min-width:1000px">
              <thead><tr><th>Estatus</th><th>Posición</th><th>Razón / Proyecto</th><th>Localidad</th><th>Tipo</th><th>Estado</th><th>Días vencida</th><th>Acción</th></tr></thead>
              <tbody>${rowsHtml}</tbody>
            </table>
          </div>
        </div>
      </div>`;
    }).join('');

    grand.ocupada+=cSum.ocupada;grand.vacante+=cSum.vacante;grand.total+=cSum.total;

    html+=`<div class="complejo-section">
      <div class="complejo-header ${clsName}" onclick="projToggleComplejo('${escHtml(complejo)}',this)">
        <div class="complejo-title">
          <span>${escHtml(complejo)}</span>
          <span class="${badgeClass}">${gerentes.length} gerentes</span>
        </div>
        <div style="display:flex;align-items:center;gap:14px">
          <div class="complejo-stats">
            <span class="complejo-stat" style="color:var(--p-blue);font-weight:700">${cSum.total} posiciones</span>
            <span class="complejo-stat cs-vac">${cSum.vacante} vacantes</span>
            <span class="complejo-stat cs-ocu">${cSum.ocupada} ocupadas</span>
          </div>
          <span class="complejo-chevron ${cIsOpen?'open':''}">▼</span>
        </div>
      </div>
      <div class="complejo-body ${cIsOpen?'open':''}">
        <div class="complejo-body-inner">${gerentesHtml}</div>
      </div>
    </div>`;
  });

  html+=`<div style="background:#061018;color:#fff;border-radius:12px;padding:14px 20px;display:flex;justify-content:space-between;align-items:center;font-weight:900;margin-top:.5rem">
    <span style="text-transform:uppercase;letter-spacing:.05em;font-size:12px">Total general</span>
    <span style="display:flex;gap:18px;font-size:13px">
      <span>${grand.ocupada} ocupadas</span><span>${grand.vacante} vacantes</span><span>${grand.total} posiciones</span>
    </span>
  </div>`;

  wrap.innerHTML=html;
}

function projToggleComplejo(complejo, headerEl){
  const section=headerEl.closest('.complejo-section');
  const body=section.querySelector('.complejo-body');
  const chev=section.querySelector('.complejo-chevron');
  const isOpen=body.classList.contains('open');
  body.classList.toggle('open');
  chev.classList.toggle('open');
  projComplejosOpen[complejo]=!isOpen;
}

function projToggleGerente(complejo, gerente, headerEl){
  const section=headerEl.closest('.gerente-section');
  const body=section.querySelector('.gerente-body');
  const chev=section.querySelector('.gerente-chev');
  const isOpen=body.classList.contains('open');
  body.classList.toggle('open');
  chev.classList.toggle('open');
  projGerentesOpen[complejo+'||'+gerente]=!isOpen;
}
// =============================================================
// ── EXPORTAR PDF ANÁLISIS DE PROYECTOS ──
// =============================================================
async function exportarProyectosPDF(){
  if(!projectFiltered.length){alert('No hay proyectos para exportar.');return;}
  const btn=document.getElementById('btn-project-export-pdf');
  btn.disabled=true;btn.textContent='⏳ Generando...';
  try{
    const hoy=new Date().toLocaleDateString('es-MX',{day:'2-digit',month:'long',year:'numeric'});
    const cComplejo=projActiveComplejo;
    const cEstatus=document.getElementById('project-sel-estatus')?.value||'Todos';
    const cGerente=document.getElementById('project-sel-gerente')?.value||'Todos';
    const cCircunstancia=CIRC_LABELS[projActiveCircunstancia]||'Todas';
    const totalProjects=projectFiltered.length;
    const vencidas=projectFiltered.filter(r=>r.estatus==='VENCIDA').length;
    const vigentes=projectFiltered.filter(r=>r.estatus==='VIGENTE').length;
    const vacantes=projectFiltered.filter(r=>r.estado==='VACANTE').length;
    const circCounts={proyecto:0,incapacidad:0,globaltalent:0,capacitacion:0};
    projectFiltered.forEach(r=>{if(circCounts[r.circunstancia]!==undefined)circCounts[r.circunstancia]++;});
    const circTotal=circCounts.proyecto+circCounts.incapacidad+circCounts.globaltalent+circCounts.capacitacion;

    const complMap={};
    projectFiltered.forEach(r=>{
      if(!complMap[r.complejo])complMap[r.complejo]={gerentes:{}};
      if(!complMap[r.complejo].gerentes[r.gerente])complMap[r.complejo].gerentes[r.gerente]=[];
      complMap[r.complejo].gerentes[r.gerente].push(r);
    });
    const ordCompl=['Norte','Centro','Sur','Corporativo','Polanco'];
    const complejos=[...ordCompl.filter(c=>complMap[c]),...Object.keys(complMap).filter(c=>!ordCompl.includes(c)).sort()];

    const thS='padding:6px 7px;font-size:9px;font-weight:800;text-transform:uppercase;letter-spacing:.04em;color:#fff;background:#6D28D9;text-align:left;white-space:nowrap;';
    const tdS='padding:5px 7px;font-size:9.5px;border-bottom:1px solid #EEF0F5;vertical-align:top;';

    let tablesHtml='';
    complejos.forEach(complejo=>{
      const cData=complMap[complejo];
      let cSum={ocupada:0,vacante:0,total:0};
      const colorC=complejo==='Norte'?'#1D4ED8':complejo==='Centro'?'#92400E':complejo==='Sur'?'#065F46':complejo==='Corporativo'?'#6D28D9':complejo==='Polanco'?'#BE185D':'#555';
      let gerentesHtml='';
      const gerentes=Object.keys(cData.gerentes).sort();
      gerentes.forEach(gerente=>{
        const gRows=[...cData.gerentes[gerente]].sort((a,b)=>(a.estado==='VACANTE'?0:1)-(b.estado==='VACANTE'?0:1)||(a.estatus==='VENCIDA'?0:1)-(b.estatus==='VENCIDA'?0:1)||a.razon.localeCompare(b.razon,'es'));
        let gSum={ocupada:0,vacante:0,total:0};
        let rowsHtml='';
        gRows.forEach(r=>{
          gSum.total++;
          if(r.estado==='VACANTE')gSum.vacante++;else gSum.ocupada++;
          const estColor=r.estatus==='VENCIDA'?'#C2410C':'#166534';
          const circLabel=CIRC_LABELS[r.circunstancia]||r.tipoJust||'—';
          const estadoTxt=r.estado==='VACANTE'?'Vacante':'Ocupada';
          rowsHtml+=`<tr><td style="${tdS}font-weight:800;color:${estColor}">${r.estatus}</td><td style="${tdS}font-weight:700;color:#003087">${escHtml(r.posicion)} ${escHtml(r.descripcion||'')}</td><td style="${tdS}">${escHtml(r.razon)}</td><td style="${tdS}">${escHtml(r.localidad||'—')}</td><td style="${tdS}">${escHtml(circLabel)}</td><td style="${tdS}text-align:center">${estadoTxt}</td><td style="${tdS}text-align:center">${r.diasVenc||''}</td><td style="${tdS}">${r.accion?escHtml(r.accion):'<em style="color:#ccc">Escribe el plan de acción</em>'}</td></tr>`;
        });
        cSum.ocupada+=gSum.ocupada;cSum.vacante+=gSum.vacante;cSum.total+=gSum.total;
        const gPct=gSum.total?((gSum.ocupada/gSum.total)*100).toFixed(1):'0.0';
        gerentesHtml+=`<div style="border:1px solid #DDE1EA;border-top:none;margin-bottom:2px;page-break-inside:avoid">
          <div style="background:#F3EEFB;padding:7px 14px;font-size:11px;font-weight:800;color:#6D28D9;border-bottom:1px solid #DDE1EA;display:flex;justify-content:space-between;align-items:center">
            <span>${escHtml(gerente)} &nbsp;<span style="font-weight:500;color:#5A6070;font-size:10px">${gSum.total} pos. · ${gSum.vacante} vac. · ${gSum.ocupada} ocu.</span></span>
            <span style="background:${gPct>=97.5?'#166534':gPct>=95?'#92400E':'#9F1239'};color:#fff;border-radius:20px;padding:3px 10px;font-size:10px;font-weight:800">${gPct}% cubierto</span>
          </div>
          <table style="width:100%;border-collapse:collapse">
            <thead><tr><th style="${thS}">Estatus</th><th style="${thS}">Posición</th><th style="${thS}">Razón / Proyecto</th><th style="${thS}">Localidad</th><th style="${thS}">Tipo</th><th style="${thS}">Estado</th><th style="${thS}">Días venc.</th><th style="${thS}">Plan de Acción</th></tr></thead>
            <tbody>${rowsHtml}</tbody>
          </table>
        </div>`;
      });
      tablesHtml+=`<div style="margin-bottom:24px">
        <div style="background:${colorC};color:#fff;padding:10px 16px;border-radius:8px 8px 0 0;font-size:13px;font-weight:900;letter-spacing:.05em;page-break-after:avoid">
          ${escHtml(complejo)} &nbsp;<span style="font-weight:400;opacity:.8;font-size:11px">${cSum.total} posiciones · ${cSum.vacante} vacantes · ${cSum.ocupada} ocupadas</span>
        </div>
        ${gerentesHtml}
      </div>`;
    });

    const htmlContent=`<div style="font-family:'Abadi','Century Gothic','Trebuchet MS',Arial,sans-serif;color:#111;font-size:11px;padding:20px;background:#fff">
      <div style="background:linear-gradient(135deg,#6D28D9,#003087);padding:16px 24px;border-radius:10px 10px 0 0">
        <div style="color:#fff;font-size:20px;font-weight:900">PILGRIMS · ATRACCIÓN DE TALENTO</div>
        <div style="color:rgba(255,255,255,.75);font-size:11px;text-transform:uppercase;letter-spacing:.08em;margin-top:3px">Reporte de Análisis de Proyectos — Por Complejo y Gerencia</div>
      </div>
      <div style="border:2px solid #6D28D9;border-top:none;border-radius:0 0 10px 10px;padding:14px 20px;margin-bottom:16px">
        <div style="font-size:12px;color:#555"><strong>Fecha:</strong> ${hoy} &nbsp;&nbsp; <strong>Complejo:</strong> ${escHtml(cComplejo||'Todos (general)')} &nbsp;&nbsp; <strong>Estatus:</strong> ${escHtml(cEstatus)} &nbsp;&nbsp; <strong>Gerente:</strong> ${escHtml(cGerente)} &nbsp;&nbsp; <strong>Circunstancia:</strong> ${escHtml(cCircunstancia)}</div>
      </div>
      <div style="display:flex;gap:10px;margin-bottom:14px">
        <div style="background:#EEF2FF;border:1.5px solid #C7D2FE;border-radius:8px;padding:10px;text-align:center;flex:1"><div style="font-size:24px;font-weight:900;color:#3730A3">${totalProjects}</div><div style="font-size:10px;font-weight:700;color:#5A6070;text-transform:uppercase">Proyectos</div></div>
        <div style="background:#FFF7ED;border:1.5px solid #FDBA74;border-radius:8px;padding:10px;text-align:center;flex:1"><div style="font-size:24px;font-weight:900;color:#C2410C">${vencidas}</div><div style="font-size:10px;font-weight:700;color:#C2410C;text-transform:uppercase">Vencidas</div></div>
        <div style="background:#ECFDF5;border:1.5px solid #A7F3D0;border-radius:8px;padding:10px;text-align:center;flex:1"><div style="font-size:24px;font-weight:900;color:#166534">${vigentes}</div><div style="font-size:10px;font-weight:700;color:#166534;text-transform:uppercase">Vigentes</div></div>
        <div style="background:#FFF1F2;border:1.5px solid #FECDD3;border-radius:8px;padding:10px;text-align:center;flex:1"><div style="font-size:24px;font-weight:900;color:#9F1239">${vacantes}</div><div style="font-size:10px;font-weight:700;color:#9F1239;text-transform:uppercase">Vacantes</div></div>
      </div>
      <div style="border:2px solid #6D28D9;border-radius:8px;padding:12px 16px;margin-bottom:18px;page-break-inside:avoid">
        <div style="font-size:10px;font-weight:800;text-transform:uppercase;letter-spacing:.05em;color:#6D28D9;margin-bottom:10px">Otras Posiciones por Circunstancia &middot; Global</div>
        <div style="display:flex;gap:10px">
          <div style="background:#FAF5FF;border-radius:8px;padding:8px;text-align:center;flex:1"><div style="font-size:20px;font-weight:900;color:#6D28D9">${circCounts.proyecto}</div><div style="font-size:9px;font-weight:700;color:#5A6070;text-transform:uppercase">Proyecto</div></div>
          <div style="background:#FAF5FF;border-radius:8px;padding:8px;text-align:center;flex:1"><div style="font-size:20px;font-weight:900;color:#6D28D9">${circCounts.incapacidad}</div><div style="font-size:9px;font-weight:700;color:#5A6070;text-transform:uppercase">Incapacidad</div></div>
          <div style="background:#FAF5FF;border-radius:8px;padding:8px;text-align:center;flex:1"><div style="font-size:20px;font-weight:900;color:#6D28D9">${circCounts.globaltalent}</div><div style="font-size:9px;font-weight:700;color:#5A6070;text-transform:uppercase">Global Talent</div></div>
          <div style="background:#FAF5FF;border-radius:8px;padding:8px;text-align:center;flex:1"><div style="font-size:20px;font-weight:900;color:#6D28D9">${circCounts.capacitacion}</div><div style="font-size:9px;font-weight:700;color:#5A6070;text-transform:uppercase">Capacitación</div></div>
          <div style="background:#6D28D9;border-radius:8px;padding:8px;text-align:center;flex:1"><div style="font-size:20px;font-weight:900;color:#fff">${circTotal}</div><div style="font-size:9px;font-weight:700;color:#fff;text-transform:uppercase">Total</div></div>
        </div>
      </div>
      ${tablesHtml}
      <div style="margin-top:18px;text-align:center;font-size:10px;color:#bbb;border-top:1px solid #e5e7eb;padding-top:8px">Pilgrims · Atracción de Talento · Documento confidencial de uso interno · ${hoy}</div>
    </div>`;

    await abrirVentanaImpresion('Analisis_Proyectos_'+(cComplejo||'Todos')+'_'+hoy, htmlContent, '@page{size:legal landscape;margin:10mm}');
  }catch(err){
    console.error('Error generando el análisis de proyectos:',err);
    alert('Ocurrió un error al generar el reporte: '+err.message);
  }finally{
    btn.disabled=false;btn.innerHTML='Exportar PDF';
  }
}

async function exportarDetallePDF(){
  if(!detFiltered.length){alert('No hay datos para exportar.');return;}
  const btn=document.getElementById('btn-det-export-pdf');
  btn.disabled=true;btn.textContent='⏳ Generando...';
  try{
  const hoy=new Date().toLocaleDateString('es-MX',{day:'2-digit',month:'long',year:'numeric'});
  const cComplejo=document.getElementById('det-sel-complejo').value||'Todos los Complejos';
  const cEstado=document.getElementById('det-sel-estado').value||'Todos los Estados';
  const cClasif=document.getElementById('det-sel-clasif').value||'Todas las Clasificaciones';
  const total=detFiltered.length;
  const vac=detFiltered.filter(r=>r.estado==='VACANTE'||r.estado==='VACANTE ').length;
  const ocu=detFiltered.filter(r=>r.estado==='OCUPADA'||r.estado==='OCUPADA ').length;

  // Agrupar
  const complMap={};
  detFiltered.forEach(r=>{
    if(!complMap[r.complejo])complMap[r.complejo]={rows:[],gerentes:{}};
    if(!complMap[r.complejo].gerentes[r.gerente])complMap[r.complejo].gerentes[r.gerente]=[];
    complMap[r.complejo].gerentes[r.gerente].push(r);
    complMap[r.complejo].rows.push(r);
  });

  const tdS='padding:5px 7px;font-size:9.5px;border-bottom:1px solid #EEF0F5;vertical-align:top;';
  const thS='padding:6px 7px;font-size:9px;font-weight:800;text-transform:uppercase;letter-spacing:.04em;color:#fff;background:#003087;text-align:left;white-space:nowrap;';

  let tablesHtml='';
  const ordCompl=['Norte','Centro','Sur','Corporativo','Polanco'];
  const complejos=[...ordCompl.filter(c=>complMap[c]),...Object.keys(complMap).filter(c=>!ordCompl.includes(c)).sort()];

  complejos.forEach(complejo=>{
    const cData=complMap[complejo];
    const cVac=cData.rows.filter(r=>r.estado==='VACANTE'||r.estado==='VACANTE ').length;
    const cOcu=cData.rows.filter(r=>r.estado==='OCUPADA'||r.estado==='OCUPADA ').length;
    const colorC=complejo==='Norte'?'#1D4ED8':complejo==='Centro'?'#92400E':complejo==='Sur'?'#065F46':complejo==='Corporativo'?'#6D28D9':complejo==='Polanco'?'#BE185D':'#555';
    tablesHtml+=`<div style="margin-bottom:24px">
      <div style="background:${colorC};color:#fff;padding:10px 16px;border-radius:8px 8px 0 0;font-size:13px;font-weight:900;letter-spacing:.05em;page-break-after:avoid">
        ${escHtml(complejo)} &nbsp;<span style="font-weight:400;opacity:.8;font-size:11px">${cData.rows.length} posiciones · ${cVac} vacantes · ${cOcu} ocupadas</span>
      </div>`;

    const gerentes=Object.keys(cData.gerentes).sort();
    gerentes.forEach(gerente=>{
      const gRowsRaw=cData.gerentes[gerente];
      const gVac=gRowsRaw.filter(r=>r.estado==='VACANTE'||r.estado==='VACANTE ').length;
      const gOcu=gRowsRaw.filter(r=>r.estado==='OCUPADA'||r.estado==='OCUPADA ').length;
      const gPct=gRowsRaw.length?((gOcu/gRowsRaw.length)*100).toFixed(1):'0.0';
      const pctColor=gPct>=97.5?'#166534':gPct>=95?'#92400E':'#9F1239';
      // vacantes primero (es lo primordial a visualizar)
      const gRows=[...gRowsRaw].sort((a,b)=>{
        const av=(a.estado==='VACANTE'||a.estado==='VACANTE ')?0:1;
        const bv=(b.estado==='VACANTE'||b.estado==='VACANTE ')?0:1;
        return av-bv;
      });
      tablesHtml+=`<div style="border:1px solid #DDE1EA;border-top:none;margin-bottom:2px;page-break-inside:avoid">
        <div style="background:#F0F4FA;padding:7px 14px;font-size:11px;font-weight:800;color:#003087;border-bottom:1px solid #DDE1EA;display:flex;justify-content:space-between;align-items:center">
          <span>${escHtml(gerente)} &nbsp;<span style="font-weight:500;color:#5A6070;font-size:10px">${gRowsRaw.length} pos. · ${gVac} vac. · ${gOcu} ocu.</span></span>
          <span style="background:${pctColor};color:#fff;border-radius:20px;padding:3px 10px;font-size:10px;font-weight:800">${gPct}% cubierto</span>
        </div>
        <table style="width:100%;border-collapse:collapse">
          <thead><tr>
            <th style="${thS}">Posición</th><th style="${thS}">Descripción</th>
            <th style="${thS}">Estado</th><th style="${thS}">Semáforo</th>
            <th style="${thS}">Días</th><th style="${thS}">Localidad</th>
            <th style="${thS}">Área</th><th style="${thS}">Clasif.</th><th style="${thS}">Justificación</th>
          </tr></thead>
          <tbody>`;
      gRows.forEach(r=>{
        const esVac=r.estado==='VACANTE'||r.estado==='VACANTE ';
        const bgRow=esVac?'background:#FFF8F8;':'';
        const semColor={Rojo:'#9F1239',Amarillo:'#92400E',Verde:'#166534'}[r.semaforo]||'#64748B';
        const semBg={Rojo:'#FFF1F2',Amarillo:'#FFFBEB',Verde:'#ECFDF5'}[r.semaforo]||'#F1F5F9';
        const semTxt=r.semaforo||'Ocupada';
        tablesHtml+=`<tr style="${bgRow}">
          <td style="${tdS}font-weight:700;color:#003087">${escHtml(r.posicion)}</td>
          <td style="${tdS}max-width:160px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${escHtml(r.descripcion)}</td>
          <td style="${tdS}font-weight:700;color:${esVac?'#9F1239':'#166534'}">${r.estado}</td>
          <td style="${tdS}"><span style="background:${semBg};color:${semColor};border-radius:4px;padding:2px 6px;font-size:9px;font-weight:800">${semTxt}</span></td>
          <td style="${tdS}text-align:center;font-weight:700">${r.dias>0?r.dias:''}</td>
          <td style="${tdS}max-width:140px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${escHtml(r.localidad)}</td>
          <td style="${tdS}max-width:120px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${escHtml(r.area)}</td>
          <td style="${tdS}">${escHtml(r.clasif)}</td>
          <td style="${tdS}max-width:130px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${escHtml(r.justificacion)}</td>
        </tr>`;
      });
      tablesHtml+='</tbody></table></div>';
    });
    tablesHtml+='</div>';
  });

  // Captura la gráfica de % Cobertura por Complejo (ya dibujada en pantalla)
  // como imagen, igual que se hace en el PDF de Comparativo Semanal — junto
  // con su lista lateral de porcentajes, que ya trae estilos en línea.
  let coberturaChartHtml='';
  const coberturaCanvas=document.getElementById('det-chart-cobertura');
  if(coberturaCanvas&&coberturaCanvas.width&&coberturaCanvas.height){
    let coberturaImg='';
    try{coberturaImg=coberturaCanvas.toDataURL('image/png',1.0);}catch(e){}
    const coberturaStats=document.getElementById('det-chart-cobertura-stats');
    const coberturaStatsHtml=coberturaStats&&coberturaStats.innerHTML.trim()?`<div style="width:190px;flex-shrink:0;display:flex;flex-direction:column;justify-content:center">${coberturaStats.innerHTML}</div>`:'';
    if(coberturaImg){
      coberturaChartHtml=`<div style="background:#F8FAFC;border:1px solid #DDE1EA;border-radius:10px;padding:12px;margin-bottom:18px;page-break-inside:avoid">
        <div style="font-size:11px;font-weight:800;text-transform:uppercase;letter-spacing:.05em;color:#003087;margin-bottom:8px">% de Cobertura por Complejo · Nacional</div>
        <div style="display:flex;gap:12px;align-items:center">
          <img src="${coberturaImg}" style="flex:1;min-width:0;max-height:220px;object-fit:contain">
          ${coberturaStatsHtml}
        </div>
      </div>`;
    }
  }

  const htmlContent=`<div style="font-family:'Abadi','Century Gothic','Trebuchet MS',Arial,sans-serif;color:#111;font-size:11px;padding:20px;background:#fff">
    <div style="background:linear-gradient(135deg,#003087,#C8102E);padding:16px 24px;border-radius:10px 10px 0 0">
      <div style="color:#fff;font-size:20px;font-weight:900">PILGRIMS · ATRACCIÓN DE TALENTO</div>
      <div style="color:rgba(255,255,255,.75);font-size:11px;text-transform:uppercase;letter-spacing:.08em;margin-top:3px">Reporte de Detalle de Posiciones — Todos los Complejos</div>
    </div>
    <div style="border:2px solid #003087;border-top:none;border-radius:0 0 10px 10px;padding:14px 20px;margin-bottom:16px">
      <div style="font-size:12px;color:#555"><strong>Fecha:</strong> ${hoy} &nbsp;&nbsp; <strong>Complejo:</strong> ${escHtml(cComplejo)} &nbsp;&nbsp; <strong>Estado:</strong> ${escHtml(cEstado)} &nbsp;&nbsp; <strong>Clasif.:</strong> ${escHtml(cClasif)}</div>
    </div>
    <div style="display:flex;gap:10px;margin-bottom:18px">
      <div style="background:#EEF2FF;border:1.5px solid #C7D2FE;border-radius:8px;padding:10px;text-align:center;flex:1"><div style="font-size:26px;font-weight:900;color:#003087">${total.toLocaleString()}</div><div style="font-size:10px;font-weight:700;color:#5A6070;text-transform:uppercase">Total Posiciones</div></div>
      <div style="background:#FFF1F2;border:1.5px solid #FECDD3;border-radius:8px;padding:10px;text-align:center;flex:1"><div style="font-size:26px;font-weight:900;color:#9F1239">${vac.toLocaleString()}</div><div style="font-size:10px;font-weight:700;color:#9F1239;text-transform:uppercase">Vacantes</div></div>
      <div style="background:#ECFDF5;border:1.5px solid #A7F3D0;border-radius:8px;padding:10px;text-align:center;flex:1"><div style="font-size:26px;font-weight:900;color:#166534">${ocu.toLocaleString()}</div><div style="font-size:10px;font-weight:700;color:#166534;text-transform:uppercase">Ocupadas</div></div>
    </div>
    ${coberturaChartHtml}
    ${tablesHtml}
    <div style="margin-top:18px;text-align:center;font-size:10px;color:#bbb;border-top:1px solid #e5e7eb;padding-top:8px">Pilgrims · Atracción de Talento · Documento confidencial de uso interno · ${hoy}</div>
  </div>`;

  await abrirVentanaImpresion('Detalle_Posiciones_'+hoy, htmlContent, '@page{size:legal landscape;margin:10mm}');
  }catch(err){
    console.error('Error generando el detalle de posiciones:',err);
    alert('Ocurrió un error al generar el reporte: '+err.message);
  }finally{
    btn.disabled=false;btn.innerHTML='Exportar PDF';
  }
}
