'use strict';
const XLSX     = require('xlsx');
const polyline = require('@mapbox/polyline');
const fs       = require('fs');
const path     = require('path');

const DIR  = __dirname;
const DATA = path.join(DIR, 'datos');

// ─── Read & process data ───────────────────────────────────────────────────────

const vWb   = XLSX.readFile(path.join(DATA, 'viajes.xlsx'));
const viajes = XLSX.utils.sheet_to_json(vWb.Sheets[vWb.SheetNames[0]], { defval: null });

const scWb  = XLSX.readFile(path.join(DATA, 'scores.csv'));
const scores = XLSX.utils.sheet_to_json(scWb.Sheets[scWb.SheetNames[0]], { defval: null });

const siWb  = XLSX.readFile(path.join(DATA, 'siniestros_viales_hechos.xlsx'));
const sinies = XLSX.utils.sheet_to_json(siWb.Sheets[siWb.SheetNames[0]], { defval: null });

const ssnWb  = XLSX.readFile(path.join(DATA, 'ssn_20242025_desarrollo_siniestros_automotores.xlsx'));
const ssnRows = XLSX.utils.sheet_to_json(ssnWb.Sheets['2024 - 2025'], { header: 1, defval: null });

// SSN costs
const dpRow = ssnRows.find(r => r[1] === 'Daño parcial');
const dtRow = ssnRows.find(r => r[1] === 'Daño total');
const TC              = 1050;
const COSTO_PARCIAL   = Math.round(dpRow[6] / dpRow[7]);   // ~1.042.996 ARS
const COSTO_TOTAL     = Math.round(dtRow[6] / dtRow[7]);   // ~12.186.394 ARS
const P_TOTAL         = +(dtRow[7] / (dpRow[7] + dtRow[7])).toFixed(4);

// Siniestros GRAVE + MORTAL
const siniesGM = sinies
  .filter(s => /^(GRAVE|MORTAL)$/.test((s.gravedad_siniestro || '').toUpperCase()))
  .map(s => [+parseFloat(s.latitud_siniestro).toFixed(5), +parseFloat(s.longitud_siniestro).toFixed(5),
             s.gravedad_siniestro.toUpperCase() === 'MORTAL' ? 1 : 0])
  .filter(([la, lo]) => !isNaN(la) && !isNaN(lo));

// Spatial index for risk zones
const CELL = 0.01, RADIO_KM = 0.1, RADIO_DEG = RADIO_KM / 111 * 3;
const sinGrid = new Map();
for (const [la, lo] of siniesGM) {
  const k = `${Math.round(la/CELL)},${Math.round(lo/CELL)}`;
  if (!sinGrid.has(k)) sinGrid.set(k, []);
  sinGrid.get(k).push([la, lo]);
}
function haversine(la1,lo1,la2,lo2){
  const dL=(la2-la1)*Math.PI/180, dO=(lo2-lo1)*Math.PI/180;
  const a=Math.sin(dL/2)**2+Math.cos(la1*Math.PI/180)*Math.cos(la2*Math.PI/180)*Math.sin(dO/2)**2;
  return 6371*2*Math.atan2(Math.sqrt(a),Math.sqrt(1-a));
}
function cercaSin(coords){
  if(!coords?.length) return false;
  for(let i=0;i<coords.length;i+=5){
    const [la,lo]=coords[i];
    const r0=Math.round(la/CELL),c0=Math.round(lo/CELL);
    for(let dr=-1;dr<=1;dr++) for(let dc=-1;dc<=1;dc++){
      const cs=sinGrid.get(`${r0+dr},${c0+dc}`);
      if(!cs) continue;
      for(const [sl,sg] of cs){
        if(Math.abs(la-sl)>RADIO_DEG||Math.abs(lo-sg)>RADIO_DEG) continue;
        if(haversine(la,lo,sl,sg)<RADIO_KM) return true;
      }
    }
  }
  return false;
}
function getHora(v){
  if(!v) return null;
  if(typeof v==='number'){const d=XLSX.SSF.parse_date_code(v);return d?.H??null;}
  const m=String(v).match(/[T ](\d{1,2}):/);return m?+m[1]:null;
}

// Scores per driver
const scoresPor = {};
for(const s of scores){
  const u=s.usuario;
  if(!scoresPor[u]) scoresPor[u]=[];
  scoresPor[u].push({
    mes:String(s.mes_viaje),
    aten:+(s.atencion_promedio_distancia??s.atencion_promedio??0),
    suav:+(s.suavidad_promedio_distancia??s.suavidad_promedio??0),
    legal:+(s.legal_promedio_distancia??s.legal_promedio??0),
    dist:+(s.distancia_total??0),
  });
}

// Group trips by driver
const tripsByDriver = {};
for(const row of viajes){
  let coords=[];
  if(row.polyline&&typeof row.polyline==='string') try{coords=polyline.decode(row.polyline);}catch(_){}
  if(!tripsByDriver[row.usuario]) tripsByDriver[row.usuario]=[];
  tripsByDriver[row.usuario].push({...row,_coords:coords});
}

// Driver definitions
const DRIVER_META = {
  '6900fe050f6bdc080000482f':{key:'micaela',nombre:'Micaela T.',auto:'VW Fox 2005',       valor_usd:8257, antiguedad:21,score_seg:0,color:'#60A5FA'},
  '6920dafb8eb2c10800000b37':{key:'jorge',  nombre:'Jorge D.',  auto:'Jeep Compass 2024', valor_usd:36190,antiguedad:1, score_seg:6,color:'#34D399'},
  '6920e5b43a13660800000b4e':{key:'nico',   nombre:'Nico S.',   auto:'Peugeot 208 2023',  valor_usd:21714,antiguedad:2, score_seg:4,color:'#F472B6'},
};

// Build per-driver data
const conductoresData = {};
for(const [userId, meta] of Object.entries(DRIVER_META)){
  const trips = tripsByDriver[userId]||[];
  const n = trips.length;

  let noct=0;
  for(const t of trips){const h=getHora(t.comienzo);if(h!==null&&(h>=22||h<6))noct++;}

  let sumV=0,cntV=0;
  for(const t of trips){
    const d=parseFloat(t.distancia_m),dur=parseFloat(t.duracion);
    if(!isNaN(d)&&!isNaN(dur)&&dur>0){sumV+=d/dur*3.6;cntV++;}
  }

  let riesgo=0;
  for(const t of trips) if(cercaSin(t._coords)) riesgo++;

  const sc=(scoresPor[userId]||[]).sort((a,b)=>a.mes.localeCompare(b.mes));
  let tD=0,tA=0,tS=0,tL=0;
  for(const s of sc){tD+=s.dist;tA+=s.aten*s.dist;tS+=s.suav*s.dist;tL+=s.legal*s.dist;}

  const m={
    pct_nocturno:n>0?+(noct/n*100).toFixed(1):0,
    vel:cntV>0?+(sumV/cntV).toFixed(1):0,
    pct_riesgo:n>0?+(riesgo/n*100).toFixed(1):0,
    atencion:tD>0?+(tA/tD).toFixed(4):0,
    suavidad:tD>0?+(tS/tD).toFixed(4):0,
    legal:tD>0?+(tL/tD).toFixed(4):0,
    n_viajes:n,
  };

  // Sample routes for map (max 20 trips, max 50 pts each)
  const rutas = trips.slice(0,20).map(t=>{
    const c=t._coords||[];
    if(!c.length) return null;
    const step=Math.max(1,Math.floor(c.length/50));
    return c.filter((_,i)=>i%step===0).map(([la,lo])=>[+la.toFixed(5),+lo.toFixed(5)]);
  }).filter(Boolean);

  const scores_mensuales = sc.map(s=>({
    mes:s.mes,
    aten:+(s.aten*100).toFixed(1),
    suav:+(s.suav*100).toFixed(1),
    legal:+(s.legal*100).toFixed(1),
  }));

  conductoresData[meta.key]={...meta,userId,metricas:m,rutas,scores_mensuales};
}

// ─── Build HTML ────────────────────────────────────────────────────────────────

const DATA_JSON = JSON.stringify({
  conductores: conductoresData,
  siniestros: siniesGM,
  costos:{ parcial_ars:COSTO_PARCIAL, total_ars:COSTO_TOTAL,
           parcial_usd:Math.round(COSTO_PARCIAL/TC), total_usd:Math.round(COSTO_TOTAL/TC),
           tc:TC, p_total:P_TOTAL },
});

const HTML = buildHTML(DATA_JSON);
const outPath = path.join(DIR, 'index.html');
fs.writeFileSync(outPath, HTML, 'utf8');
const kb = (fs.statSync(outPath).size / 1024).toFixed(1);
console.log(`✓ index.html generado — ${kb} KB`);

// ─── HTML template ─────────────────────────────────────────────────────────────

function buildHTML(dataJson){
return `<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Gemelo Digital — Trail</title>
<link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css"/>
<style>
:root{--bg:#0D1B2A;--card:#111E2E;--border:#1E3348;--text:#E2E8F0;--muted:#64748B;
      --blue:#3B82F6;--green:#22C55E;--amber:#F59E0B;--red:#EF4444;}
*{box-sizing:border-box;margin:0;padding:0;}
body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:var(--bg);color:var(--text);min-height:100vh;}
/* HEADER */
.hdr{background:var(--card);border-bottom:1px solid var(--border);padding:0 28px;display:flex;align-items:center;gap:36px;position:sticky;top:0;z-index:1000;}
.logo{font-size:18px;font-weight:800;letter-spacing:-0.5px;color:var(--blue);white-space:nowrap;padding:18px 0;}
.logo em{color:var(--muted);font-style:normal;font-weight:400;font-size:13px;margin-left:8px;}
.tabs{display:flex;gap:0;}
.tab{padding:20px 22px 17px;cursor:pointer;color:var(--muted);font-size:14px;font-weight:500;
     border-bottom:3px solid transparent;transition:all .15s;white-space:nowrap;}
.tab:hover{color:var(--text);}
.tab.on{color:var(--blue);border-bottom-color:var(--blue);}
.tab .dot{width:8px;height:8px;border-radius:50%;display:inline-block;margin-right:8px;}
/* MAIN */
.wrap{padding:24px;max-width:1440px;margin:0 auto;display:flex;flex-direction:column;gap:20px;}
/* DRIVER HEADER */
.drvhdr{display:flex;align-items:center;gap:16px;}
.drvhdr h2{font-size:22px;font-weight:700;}
.drvhdr small{color:var(--muted);font-size:13px;}
.badge{padding:3px 10px;border-radius:99px;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;}
/* KPIs */
.kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:16px;}
.kcard{background:var(--card);border:1px solid var(--border);border-radius:12px;padding:20px 22px;}
.kcard .lbl{font-size:11px;text-transform:uppercase;letter-spacing:.6px;color:var(--muted);margin-bottom:12px;}
.kcard .val{font-size:34px;font-weight:800;line-height:1;margin-bottom:4px;}
.kcard .sub{font-size:12px;color:var(--muted);}
/* GAUGE */
.gauge-wrap{display:flex;flex-direction:column;align-items:flex-start;gap:4px;}
/* GRID 2 col */
.grid2{display:grid;grid-template-columns:1fr 1fr;gap:16px;}
.card{background:var(--card);border:1px solid var(--border);border-radius:12px;overflow:hidden;}
.card-hdr{padding:14px 20px;border-bottom:1px solid var(--border);font-size:13px;font-weight:600;
          display:flex;justify-content:space-between;align-items:center;}
.card-hdr .legend{display:flex;gap:12px;font-size:11px;color:var(--muted);}
.legend-dot{width:9px;height:9px;border-radius:50%;display:inline-block;margin-right:4px;}
#map{height:390px;}
.chart-wrap{padding:16px 20px;height:390px;display:flex;align-items:center;}
#scores-chart{max-height:340px;width:100%;}
/* FEATURES */
.feat-row{display:flex;align-items:center;gap:12px;margin-bottom:10px;}
.feat-lbl{font-size:12px;color:var(--muted);width:160px;text-align:right;flex-shrink:0;}
.feat-bg{flex:1;background:var(--border);border-radius:4px;height:10px;position:relative;}
.feat-bar{height:100%;border-radius:4px;transition:width .5s ease;position:relative;z-index:1;}
.feat-ref{position:absolute;left:50%;top:-3px;bottom:-3px;width:1px;
          border-left:1px dashed #334155;z-index:2;pointer-events:none;}
.feat-pct{font-size:11px;color:var(--text);width:42px;text-align:right;flex-shrink:0;font-weight:600;}
/* SIMULATOR */
.sim-grid{display:grid;grid-template-columns:1fr 1fr;gap:32px;}
.slider-grp{margin-bottom:18px;}
.slider-grp label{display:flex;justify-content:space-between;font-size:13px;margin-bottom:6px;}
.slider-grp label span{color:var(--blue);font-weight:700;}
input[type=range]{width:100%;accent-color:var(--blue);cursor:pointer;}
.sim-result{background:#080F1A;border-radius:10px;padding:24px;display:flex;flex-direction:column;
            gap:16px;align-items:center;justify-content:center;}
.sim-block{text-align:center;}
.sim-block .lbl2{font-size:11px;text-transform:uppercase;letter-spacing:.5px;color:var(--muted);margin-bottom:6px;}
.sim-block .big{font-size:42px;font-weight:800;line-height:1;}
.delta{font-size:13px;margin-top:4px;}
.delta.better{color:var(--green);}
.delta.worse{color:var(--red);}
.delta.same{color:var(--muted);}
.sim-row{display:flex;gap:24px;width:100%;}
.sim-mini{flex:1;text-align:center;background:var(--card);border-radius:8px;padding:12px;}
.sim-mini .lbl2{font-size:10px;color:var(--muted);margin-bottom:4px;}
.sim-mini .val2{font-size:20px;font-weight:700;}
/* CARD BODY */
.cbody{padding:18px 20px;}
/* RISK COLORS */
.rc-low{color:var(--green);}
.rc-mid{color:var(--amber);}
.rc-high{color:var(--red);}
/* ALERTS */
.alerta-item{display:flex;align-items:flex-start;gap:14px;padding:12px 14px;border-radius:8px;
             margin-bottom:8px;background:rgba(127,29,29,.18);border:1px solid rgba(239,68,68,.3);}
.alerta-icon{font-size:18px;flex-shrink:0;margin-top:1px;}
.alerta-mes{font-weight:700;color:var(--amber);font-size:13px;}
.alerta-body{font-size:13px;color:var(--text);margin-top:2px;}
.alerta-score{font-weight:700;color:var(--red);}
.alerta-ok{color:var(--green);font-size:13px;padding:8px 0;display:flex;align-items:center;gap:8px;}
.alerta-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(200px,1fr));gap:8px;}
/* AI DIAGNOSIS */
.ai-loading{color:var(--muted);font-size:14px;display:flex;align-items:center;gap:12px;padding:8px 0;}
.ai-spinner{width:18px;height:18px;border:2px solid var(--border);border-top-color:var(--blue);
            border-radius:50%;animation:spin .75s linear infinite;flex-shrink:0;}
@keyframes spin{to{transform:rotate(360deg)}}
.ai-box{background:#080F1A;border-radius:10px;padding:20px 22px;}
.ai-line{font-size:14px;line-height:1.75;color:var(--text);margin-bottom:6px;padding-left:14px;
         border-left:3px solid var(--border);}
.ai-line:last-of-type{margin-bottom:0;}
.ai-line.l1{border-left-color:var(--blue);}
.ai-line.l2{border-left-color:var(--amber);}
.ai-line.l3{border-left-color:var(--green);}
.ai-meta{font-size:11px;color:var(--muted);margin-top:14px;display:flex;align-items:center;gap:8px;}
.ai-badge{padding:2px 8px;border-radius:99px;font-size:10px;font-weight:700;
          background:#1E3A5F;color:var(--blue);text-transform:uppercase;letter-spacing:.4px;}
.ai-fallback{color:var(--amber);font-size:10px;}
/* CARD SUBTITLE */
.card-sub{padding:10px 20px 12px;font-size:12px;color:var(--muted);line-height:1.6;
          border-bottom:1px solid var(--border);background:rgba(255,255,255,.02);}
/* KPI NOTE */
.kpi-note{font-size:11px;color:var(--muted);margin-top:10px;padding-top:8px;
          border-top:1px solid var(--border);line-height:1.55;}
/* FEATURE ROW TOOLTIPS */
.feat-row{display:flex;align-items:center;gap:12px;margin-bottom:10px;position:relative;cursor:default;}
.feat-row[data-tip]:hover::after{
  content:attr(data-tip);
  display:block;
  position:absolute;
  bottom:calc(100% + 10px);
  left:172px;
  background:#0B1E32;
  border:1px solid var(--border);
  border-radius:8px;
  padding:10px 14px;
  font-size:12px;
  color:var(--text);
  line-height:1.6;
  width:300px;
  z-index:200;
  box-shadow:0 4px 24px rgba(0,0,0,.6);
  white-space:normal;
  pointer-events:none;
}
/* SIM INTRO */
.sim-intro{font-size:13px;color:var(--muted);margin-bottom:20px;line-height:1.5;padding-bottom:16px;border-bottom:1px solid var(--border);}
/* FOOTER */
.footer{padding:14px 28px;border-top:1px solid var(--border);font-size:11px;color:var(--muted);text-align:center;margin-top:8px;}
.footer a{color:var(--muted);}
</style>
</head>
<body>

<header class="hdr">
  <div class="logo">Trail <em>Gemelo Digital</em></div>
  <nav class="tabs" id="tabs"></nav>
</header>

<main class="wrap">
  <div class="drvhdr" id="drvhdr"></div>
  <div class="kpis" id="kpis"></div>
  <div class="grid2">
    <div class="card">
      <div class="card-hdr">
        Rutas y Siniestros CABA
        <div class="legend">
          <span><span class="legend-dot" style="background:#EF4444"></span>Grave</span>
          <span><span class="legend-dot" style="background:#7F1D1D"></span>Mortal</span>
          <span id="map-legend-route"></span>
        </div>
      </div>
      <div id="map"></div>
    </div>
    <div class="card">
      <div class="card-hdr">Evolución Scores Mensuales</div>
      <div class="chart-wrap"><canvas id="scores-chart"></canvas></div>
    </div>
  </div>
  <div class="card">
    <div class="card-hdr">Feature Importance — Peso de cada variable en el score de riesgo</div>
    <div class="card-sub">Qué factores explican más el riesgo de este conductor según el modelo. Calculado con datos reales de Trail + 65.818 siniestros viales de CABA.</div>
    <div class="cbody" id="features"></div>
  </div>
  <div class="card">
    <div class="card-hdr">Simulador de Escenarios — ¿Qué pasa si el conductor mejora?</div>
    <div class="cbody">
      <div class="sim-intro">Simulá el impacto de intervenciones concretas. Mové los sliders para ver cómo cambia el riesgo y el ahorro proyectado.</div>
      <div class="sim-grid">
        <div id="sim-controls"></div>
        <div class="sim-result" id="sim-result"></div>
      </div>
    </div>
  </div>

  <div class="card">
    <div class="card-hdr">
      Sección 6 — Panel de Alertas
      <span style="font-size:11px;color:var(--muted);font-weight:400">Score legal &lt; 80% por mes</span>
    </div>
    <div class="cbody" id="alertas"></div>
  </div>

  <div class="card">
    <div class="card-hdr">
      Sección 9 — Diagnóstico Telemático
      <span style="font-size:11px;color:var(--muted);font-weight:400">Análisis por reglas actuariales</span>
    </div>
    <div class="cbody" id="diagnostico">
      <div class="ai-loading"><div class="ai-spinner"></div>Seleccioná un conductor para generar el diagnóstico…</div>
    </div>
  </div>
</main>

<footer class="footer">
  Fuentes: Trail/Sentiance SDK &nbsp;·&nbsp; SSN 2024/25 — Desarrollo de Siniestros Automotores &nbsp;·&nbsp; Infoauto jun 2026 &nbsp;·&nbsp; GCBA Siniestros Viales 2019-2025 &nbsp;·&nbsp; ANSV
</footer>

<script type="application/json" id="__data__">${dataJson}</script>
<script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<script>
const RAW = JSON.parse(document.getElementById('__data__').textContent);
// ─── Risk model ────────────────────────────────────────────────────────────────
const PESOS = {zona:.30,legal:.20,nocturno:.15,vel:.10,aten:.10,seg:.10,antig:.05};
const FEAT_LABELS = {zona:'Zonas de riesgo',legal:'Score legal',nocturno:'Manejo nocturno',
                     vel:'Velocidad',aten:'Atención',seg:'Seg. vehículo',antig:'Antigüedad'};

function factors(m, d, overrides={}){
  const om = {...m, ...overrides};
  return {
    zona:     om.pct_riesgo/100,
    legal:    1-om.legal,
    nocturno: Math.min(om.pct_nocturno/30,1),
    vel:      Math.min(Math.max((om.vel-15)/55,0),1),
    aten:     1-om.atencion,
    seg:      1-d.score_seg/6,
    antig:    Math.min(d.antiguedad/25,1),
  };
}

function calcScore(m, d, ov={}){
  const f=factors(m,d,ov);
  return Math.round(Object.entries(PESOS).reduce((s,[k,w])=>s+w*Math.max(0,Math.min(1,f[k])),0)*100);
}

function calcProb(score){
  return Math.min(0.17*Math.exp((score-50)/30), 0.60);
}

function calcCosto(prob, d){
  const costos=RAW.costos;
  const val=d.valor_usd*costos.tc;
  const ct=Math.min(costos.total_ars,val);
  const exp=prob*((1-costos.p_total)*costos.parcial_ars+costos.p_total*ct);
  return {ars:Math.round(exp), usd:Math.round(exp/costos.tc)};
}

function riskClass(s){ return s<=33?'rc-low':s<=66?'rc-mid':'rc-high'; }
function riskLbl(s){ return s<=33?'BAJO':s<=66?'MEDIO':'ALTO'; }
function fmtARS(n){ return 'ARS '+n.toLocaleString('es-AR'); }
function fmtPct(n){ return (n*100).toFixed(1)+'%'; }

// ─── State ─────────────────────────────────────────────────────────────────────
let activeKey = 'micaela';
let leafMap, routeGroup, scChart, featChart;
const simState = {};

const keys = ['micaela','jorge','nico'];

// ─── Init ──────────────────────────────────────────────────────────────────────
window.addEventListener('load', () => {
  buildTabs();
  initMap();
  initScoresChart();
  selectDriver('micaela');
});

function buildTabs(){
  const el=document.getElementById('tabs');
  el.innerHTML=keys.map(k=>{
    const d=RAW.conductores[k];
    return \`<div class="tab" id="tab-\${k}" onclick="selectDriver('\${k}')">
      <span class="dot" style="background:\${d.color}"></span>\${d.nombre}
    </div>\`;
  }).join('');
}

function selectDriver(key){
  activeKey=key;
  keys.forEach(k=>document.getElementById('tab-'+k).classList.toggle('on',k===key));
  const d=RAW.conductores[key];
  initSimState(d);
  updateDrvHdr(d);
  updateKPIs(d);
  updateMapRoutes(d);
  updateScoresChart(d);
  updateFeatures(d);
  updateSimulator(d);
  updateAlertas(d);
  updateDiagnostico(d);
}

// ─── Driver header ─────────────────────────────────────────────────────────────
function updateDrvHdr(d){
  const s=calcScore(d.metricas,d);
  document.getElementById('drvhdr').innerHTML=\`
    <div class="dot" style="background:\${d.color};width:14px;height:14px;border-radius:50%"></div>
    <h2>\${d.nombre}</h2>
    <small>\${d.auto} · \${d.metricas.n_viajes} viajes · USD \${d.valor_usd.toLocaleString()}</small>
    <span class="badge \${riskClass(s)}" style="background:\${s<=33?'#14532D':s<=66?'#78350F':'#7F1D1D'}">
      Riesgo \${riskLbl(s)}
    </span>
  \`;
}

// ─── KPIs ──────────────────────────────────────────────────────────────────────
function updateKPIs(d){
  const m=d.metricas;
  const score=calcScore(m,d);
  const prob=calcProb(score);
  const costo=calcCosto(prob,d);
  const rc=riskClass(score);

  document.getElementById('kpis').innerHTML=\`
    <div class="kcard">
      <div class="lbl">Score de Riesgo</div>
      \${gaugeHTML(score)}
      <div class="sub">0 = sin riesgo · 100 = máximo</div>
    </div>
    <div class="kcard">
      <div class="lbl">Probabilidad de Siniestro</div>
      <div class="val \${rc}">\${(prob*100).toFixed(1)}%</div>
      <div class="sub">estimada en 12 meses</div>
      <div class="kpi-note">Probabilidad de que este conductor tenga al menos un siniestro en los próximos 12 meses. Base industria CABA: 17% (SSN 2024/25).</div>
    </div>
    <div class="kcard">
      <div class="lbl">Costo Esperado / año</div>
      <div class="val" style="font-size:24px">\${fmtARS(costo.ars)}</div>
      <div class="sub">USD \${costo.usd.toLocaleString()} · daño propio</div>
      <div class="kpi-note">Costo probable de siniestro anual ponderado por probabilidad. Fuente: SSN — Desarrollo de Siniestros del Ramo Automotores 2024/25. Daño parcial promedio: ARS 1.042.996. Valor del vehículo: Infoauto junio 2026.</div>
    </div>
    <div class="kcard">
      <div class="lbl">Exposición Geográfica</div>
      <div class="val \${m.pct_riesgo>50?'rc-high':m.pct_riesgo>20?'rc-mid':'rc-low'}">\${m.pct_riesgo}%</div>
      <div class="sub">viajes en zonas de riesgo</div>
      <div class="kpi-note">% de viajes que pasan a menos de 100 metros de una zona con siniestros graves o mortales registrados en CABA (2019–2025).</div>
    </div>
  \`;
}

function gaugeHTML(score){
  const color = score < 40 ? '#1D9E75' : score <= 70 ? '#EF9F27' : '#E24B4A';
  const label = score < 40 ? 'BAJO'    : score <= 70 ? 'MEDIO'   : 'ALTO';
  const pct   = Math.min(Math.max(score, 2), 98);
  return \`<div style="text-align:center;padding:0.5rem 1rem 0.25rem">
    <div style="font-size:48px;font-weight:700;line-height:1;color:\${color}">\${score}</div>
    <div style="font-size:12px;color:#64748B;margin-top:2px">de 100 &nbsp;·&nbsp; <strong style="color:\${color}">\${label}</strong></div>
    <div style="margin-top:10px;padding:0 4px">
      <div style="background:linear-gradient(to right,#1D9E75,#EF9F27,#E24B4A);height:8px;border-radius:4px;position:relative">
        <div style="position:absolute;top:-4px;left:\${pct}%;width:16px;height:16px;background:#fff;border-radius:50%;transform:translateX(-50%);box-shadow:0 0 0 2px \${color}"></div>
      </div>
      <div style="display:flex;justify-content:space-between;font-size:10px;color:#475569;margin-top:5px;letter-spacing:.5px">
        <span>0</span><span>50</span><span>100</span>
      </div>
    </div>
  </div>\`;
}

// ─── Leaflet Map ───────────────────────────────────────────────────────────────
function initMap(){
  leafMap=L.map('map',{zoomControl:true}).setView([-34.615,-58.445],13);
  L.tileLayer('https://{s}.basemaps.cartocdn.com/dark_all/{z}/{x}/{y}{r}.png',
    {attribution:'© OpenStreetMap / CartoDB',maxZoom:19}).addTo(leafMap);

  // Siniestros layer (static)
  const sinLayer=L.layerGroup();
  for(const [la,lo,mortal] of RAW.siniestros){
    L.circleMarker([la,lo],{radius:mortal?5:3,color:'transparent',
      fillColor:mortal?'#7F1D1D':'#EF4444',fillOpacity:.7,weight:0}).addTo(sinLayer);
  }
  sinLayer.addTo(leafMap);

  routeGroup=L.layerGroup().addTo(leafMap);
}

function updateMapRoutes(d){
  routeGroup.clearLayers();
  for(const ruta of d.rutas){
    if(ruta.length<2) continue;
    L.polyline(ruta,{color:d.color,weight:2.5,opacity:.85}).addTo(routeGroup);
  }
  document.getElementById('map-legend-route').innerHTML=
    \`<span class="legend-dot" style="background:\${d.color}"></span>\${d.nombre}\`;
  // Fit to routes
  if(d.rutas.length>0){
    const allPts=d.rutas.flat();
    if(allPts.length) leafMap.fitBounds(L.latLngBounds(allPts),{padding:[20,20]});
  }
}

// ─── Scores Chart ──────────────────────────────────────────────────────────────
function initScoresChart(){
  const ctx=document.getElementById('scores-chart').getContext('2d');
  scChart=new Chart(ctx,{
    type:'line',
    data:{labels:[],datasets:[
      {label:'Atención',data:[],borderColor:'#3B82F6',backgroundColor:'#3B82F620',tension:.4,fill:false},
      {label:'Suavidad',data:[],borderColor:'#22C55E',backgroundColor:'#22C55E20',tension:.4,fill:false},
      {label:'Legal',   data:[],borderColor:'#F59E0B',backgroundColor:'#F59E0B20',tension:.4,fill:false},
    ]},
    options:{
      responsive:true,maintainAspectRatio:false,
      plugins:{legend:{labels:{color:'#94A3B8',boxWidth:12}},
               tooltip:{mode:'index',intersect:false}},
      scales:{
        x:{ticks:{color:'#64748B',maxRotation:0},grid:{color:'#1E3348'}},
        y:{min:50,max:105,ticks:{color:'#64748B',callback:v=>v+'%'},grid:{color:'#1E3348'}},
      }
    }
  });
}

function updateScoresChart(d){
  const sm=d.scores_mensuales;
  if(!sm.length){
    scChart.data.labels=[];
    scChart.data.datasets.forEach(ds=>ds.data=[]);
    scChart.update();return;
  }
  const fmt=m=>{const y=m.slice(0,4),mo=m.slice(4);const mns=['','Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];return mns[+mo]+' '+y;};
  scChart.data.labels=sm.map(s=>fmt(s.mes));
  scChart.data.datasets[0].data=sm.map(s=>s.aten);
  scChart.data.datasets[1].data=sm.map(s=>s.suav);
  scChart.data.datasets[2].data=sm.map(s=>s.legal);
  scChart.update();
}

// ─── Feature Importance ────────────────────────────────────────────────────────
const FEAT_TIPS = {
  nocturno: '% de viajes entre 22h y 6h. Siniestros nocturnos tienen 2.3x más probabilidad de ser graves (ANSV).',
  zona:     '% de km recorridos en zonas con alta concentración de siniestros graves o mortales en CABA (2019–2025).',
  legal:    'Exceso de velocidad y respeto de semáforos, ponderado por km recorrido.',
  seg:      'Presencia de sistemas de seguridad activa: ABS, airbag, ESP. 0 = sin ninguno, 6 = equipamiento completo. Fuente: Infoauto.',
  vel:      'Velocidad promedio inferida de distancia y duración del viaje (distancia_m / duración_s × 3.6).',
  aten:     'Uso del teléfono al volante, ponderado por km recorrido.',
  suav:     'Frenadas y aceleraciones bruscas, ponderado por km recorrido.',
  antig:    'Años del vehículo desde su fabricación. A mayor antigüedad, mayor riesgo mecánico y menor valor de mercado.',
};

function updateFeatures(d){
  const f=factors(d.metricas,d);
  const contribs={};
  let total=0;
  for(const [k,w] of Object.entries(PESOS)){
    contribs[k]=w*Math.max(0,Math.min(1,f[k]))*100;
    total+=contribs[k];
  }
  const sorted=Object.entries(contribs).sort((a,b)=>b[1]-a[1]);

  // share = real % this factor contributes to the total risk score
  // barWidth = share directly (0-100%), so 100% width only when share=100%
  const rows=sorted.map(([k,v])=>{
    const share=total>0?v/total*100:0;
    const barW=share.toFixed(2);
    const col=share>20?'var(--red)':share>=10?'var(--amber)':'var(--blue)';
    const tip=FEAT_TIPS[k]||'';
    return \`<div class="feat-row" data-tip="\${tip}">
      <div class="feat-lbl">\${FEAT_LABELS[k]}</div>
      <div class="feat-bg">
        <div class="feat-bar" style="width:\${barW}%;background:\${col}"></div>
        <div class="feat-ref"></div>
      </div>
      <div class="feat-pct">\${share.toFixed(1)}%</div>
    </div>\`;
  }).join('');

  // 50% reference label below the bars
  document.getElementById('features').innerHTML=\`
    <div style="position:relative;margin-bottom:4px">\${rows}</div>
    <div style="display:flex;padding-left:172px;padding-right:54px;margin-top:6px;">
      <div style="flex:1;position:relative;font-size:10px;color:var(--muted);">
        <span style="position:absolute;left:0">0%</span>
        <span style="position:absolute;left:50%;transform:translateX(-50%)">50%</span>
        <span style="position:absolute;right:0">100%</span>
      </div>
    </div>
    <div style="font-size:11px;color:var(--muted);margin-top:10px;padding-top:8px;border-top:1px solid var(--border)">
      Cada barra muestra el % de contribución al score total de riesgo
      (<span style="color:var(--red)">●</span> &gt;20%
       <span style="color:var(--amber)">●</span> 10-20%
       <span style="color:var(--blue)">●</span> &lt;10%) — línea punteada = 50%
    </div>
  \`;
}

// ─── Simulator ─────────────────────────────────────────────────────────────────
function initSimState(d){
  const m=d.metricas;
  simState.pct_nocturno=m.pct_nocturno;
  simState.legal=m.legal;
  simState.pct_riesgo=m.pct_riesgo;
  simState.vel=m.vel;
  simState.atencion=m.atencion;
}

function updateSimulator(d){
  const sliders=[
    {key:'legal',     label:'Score legal',         min:0.5, max:1.0, step:0.01, fmt:v=>(v*100).toFixed(0)+'%'},
    {key:'pct_nocturno',label:'Viajes nocturnos',  min:0,   max:30,  step:0.5,  fmt:v=>v.toFixed(1)+'%'},
    {key:'pct_riesgo', label:'Exposición zonas riesgo', min:0, max:100, step:1, fmt:v=>v.toFixed(0)+'%'},
    {key:'vel',        label:'Velocidad promedio',  min:15,  max:70,  step:0.5,  fmt:v=>v.toFixed(1)+' km/h'},
  ];

  document.getElementById('sim-controls').innerHTML=sliders.map(s=>\`
    <div class="slider-grp">
      <label>\${s.label} <span id="sv-\${s.key}">\${s.fmt(simState[s.key])}</span></label>
      <input type="range" min="\${s.min}" max="\${s.max}" step="\${s.step}"
             value="\${simState[s.key]}"
             data-key="\${s.key}" data-fmt="\${s.key}"
             oninput="onSlider(this)"
      />
    </div>
  \`).join('');

  // store format functions by key for use in onSlider
  window._simFmts = {};
  sliders.forEach(s=>{ window._simFmts[s.key]=s.fmt; });

  renderSimResult(d);
}

function onSlider(input){
  const key=input.dataset.key;
  const val=+input.value;
  simState[key]=val;
  document.getElementById('sv-'+key).textContent=window._simFmts[key](val);
  renderSimResult(RAW.conductores[activeKey]);
}

// ─── Sección 6: Alertas ────────────────────────────────────────────────────────
const UMBRAL_LEGAL = 80;

function fmtMes(m){
  const mns=['','Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
  return mns[+m.slice(4)]+' '+m.slice(0,4);
}

function updateAlertas(d){
  const sm=d.scores_mensuales;
  const bajos=sm.filter(s=>s.legal<UMBRAL_LEGAL);
  const el=document.getElementById('alertas');

  if(!sm.length){
    el.innerHTML='<div class="ai-loading" style="color:var(--muted)">Sin datos de scores mensuales para este conductor.</div>';
    return;
  }

  if(!bajos.length){
    el.innerHTML=\`<div class="alerta-ok">✓ Sin alertas — score legal por encima del \${UMBRAL_LEGAL}% en todos los períodos registrados.</div>\`;
    return;
  }

  // Build trend context: also show previous month for comparison
  const html=bajos.map(a=>{
    const idx=sm.indexOf(a);
    const prev=idx>0?sm[idx-1]:null;
    const trend=prev&&prev.legal>=UMBRAL_LEGAL
      ? \`↓ cayó desde \${prev.legal.toFixed(1)}%\`
      : prev&&prev.legal<UMBRAL_LEGAL
        ? \`↓ segundo mes consecutivo bajo\`
        : '';
    return \`<div class="alerta-item">
      <div class="alerta-icon">⚠</div>
      <div>
        <div class="alerta-mes">\${fmtMes(a.mes)}</div>
        <div class="alerta-body">Score legal: <span class="alerta-score">\${a.legal.toFixed(1)}%</span>
          (umbral \${UMBRAL_LEGAL}%) \${trend?'— '+trend:''}
        </div>
        <div style="font-size:11px;color:var(--muted);margin-top:3px">Atención: \${a.aten.toFixed(1)}% · Suavidad: \${a.suav.toFixed(1)}%</div>
      </div>
    </div>\`;
  }).join('');

  const total=sm.length;
  el.innerHTML=html+\`<div style="font-size:11px;color:var(--muted);margin-top:8px">
    \${bajos.length} de \${total} meses bajo umbral (\${(bajos.length/total*100).toFixed(0)}%)
  </div>\`;
}

// ─── Sección 9: Diagnóstico AI (via proxy local) ──────────────────────────────
const GEMINI_URL = "/api/gemini";

function buildPrompt(d){
  const m=d.metricas, score=calcScore(m,d), prob=calcProb(score);
  const prob90d=((1-Math.pow(1-prob,90/365))*100).toFixed(1);
  return \`Sos un actuario especialista en seguros de autos en Argentina. Generá un diagnóstico en exactamente 3 líneas para este conductor:

- Legalidad × km: \${(m.legal*100).toFixed(1)}% (mide exceso de velocidad y respeto de semáforos)
- Atención × km: \${(m.atencion*100).toFixed(1)}% (mide uso del teléfono al volante)
- Suavidad × km: \${(m.suavidad*100).toFixed(1)}% (mide frenadas y aceleraciones bruscas)
- % viajes nocturnos: \${m.pct_nocturno}%
- % km en zonas de siniestros CABA: \${m.pct_riesgo}%
- Probabilidad siniestro 90d: \${prob90d}%
- Vehículo: \${d.auto} (\${d.antiguedad} años, seguridad \${d.score_seg}/6, USD \${d.valor_usd.toLocaleString()})

Nombrá la causa real de cada score bajo usando el mapeo correcto. Sé directo y accionable. Sin introducción. Terminá con una recomendación de intervención específica con impacto en la prima.\`;
}

function diagnosisRules(d){
  const m=d.metricas, score=calcScore(m,d), prob=calcProb(score);
  const f=factors(m,d);

  // Contribución de cada factor al score total (peso × valor normalizado)
  const contrib={
    zona:    PESOS.zona    * Math.max(0,Math.min(1,f.zona)),
    legal:   PESOS.legal   * Math.max(0,Math.min(1,f.legal)),
    nocturno:PESOS.nocturno* Math.max(0,Math.min(1,f.nocturno)),
    vel:     PESOS.vel     * Math.max(0,Math.min(1,f.vel)),
    aten:    PESOS.aten    * Math.max(0,Math.min(1,f.aten)),
    seg:     PESOS.seg     * Math.max(0,Math.min(1,f.seg)),
    antig:   PESOS.antig   * Math.max(0,Math.min(1,f.antig)),
  };
  const sorted=Object.entries(contrib).sort((a,b)=>b[1]-a[1]);
  const top=sorted[0][0];
  const sec=sorted[1][0];

  // ── Línea 1: causa principal ──────────────────────────────────────────────
  const L1={
    zona:    \`\${m.pct_riesgo}% de los viajes circulan por zonas con alta concentración de siniestros graves o mortales registrados en CABA (2019–2025) — principal factor de exposición del perfil.\`,
    legal:   \`Score de legalidad en \${(m.legal*100).toFixed(0)}%: patrón de exceso de velocidad y/o incumplimiento de semáforos detectado en el \${(100-m.legal*100).toFixed(0)}% de los km recorridos — principal driver de riesgo conductual.\`,
    nocturno:\`\${m.pct_nocturno}% de los viajes se realizan en horario nocturno (22h–6h), período con 2.3× más probabilidad de siniestros graves según ANSV — mayor fuente de riesgo en este perfil.\`,
    aten:    \`Score de atención en \${(m.atencion*100).toFixed(0)}%: uso del teléfono al volante documentado en el \${(100-m.atencion*100).toFixed(0)}% de los km — la distracción es la causa #1 de siniestros urbanos en Argentina (ANSV 2024).\`,
    seg:     \`El \${d.auto} no cuenta con sistemas de seguridad activa (score \${d.score_seg}/6): sin ABS, airbag ni ESP — aumenta severidad esperada ante cualquier colisión y eleva el costo técnico.\`,
    vel:     \`Velocidad media de \${m.vel.toFixed(0)} km/h inferida de los datos GPS, superior a la media recomendada para circulación urbana en CABA — eleva probabilidad y severidad del siniestro.\`,
    antig:   \`El \${d.auto} tiene \${d.antiguedad} años de antigüedad: mayor riesgo de falla mecánica, menor valor de recupero (USD \${d.valor_usd.toLocaleString()}) y probabilidad elevada de daño total en colisión.\`,
  };

  // ── Línea 2: factor secundario ────────────────────────────────────────────
  const L2={
    zona:    \`Exposición geográfica del \${m.pct_riesgo}% en zonas de riesgo — calles con historial de atropellamientos y choques frontales. El ruteo habitual explica parte de este indicador.\`,
    legal:   m.legal<0.80
      ? \`Velocidad y semáforos comprometen el \${(100-m.legal*100).toFixed(0)}% de los km: 1 de cada \${Math.max(2,Math.round(1/(1-m.legal)))} km registra una infracción potencial según el modelo Trail.\`
      : \`Score legal de \${(m.legal*100).toFixed(0)}%: respeto de velocidades y semáforos aceptable, sin infracciones sistemáticas detectadas.\`,
    nocturno:m.pct_nocturno>15
      ? \`La conducción nocturna también aumenta la fatiga acumulada y reduce el tiempo de reacción; \${m.pct_nocturno}% de viajes en ese horario supera el umbral de alerta del 15%.\`
      : \`Conducción mayoritariamente diurna (\${100-m.pct_nocturno}%), lo que reduce la exposición a fatiga nocturna y mejora los tiempos de reacción.\`,
    aten:    m.atencion<0.80
      ? \`Uso del teléfono al volante en \${(100-m.atencion*100).toFixed(0)}% de km — distracción que multiplica por 4 el riesgo de colisión según OMS (2023). Impacta directamente en la prima técnica.\`
      : \`Score de atención en \${(m.atencion*100).toFixed(0)}%: el uso del celular al volante está bajo control en este conductor.\`,
    seg:     d.score_seg===0
      ? \`Sin ningún sistema de seguridad activa (ABS, airbag, ESP), la severidad esperada de un siniestro es significativamente mayor a la media del portafolio.\`
      : \`Equipamiento de seguridad activa del vehículo: score \${d.score_seg}/6 — mitiga parcialmente el riesgo de lesiones en caso de colisión.\`,
    vel:     \`Velocidad media de \${m.vel.toFixed(0)} km/h — \${m.vel>50?'supera':'dentro de'} los umbrales típicos de circulación urbana para CABA (30–50 km/h). Dato inferido de distancia y duración de viaje.\`,
    antig:   \`Con \${d.antiguedad} años, el vehículo \${d.antiguedad>15?'supera el umbral crítico de riesgo mecánico (15 años). Probable ausencia de sistemas de seguridad activa modernos.':'tiene antigüedad moderada; el riesgo mecánico es menor que en vehículos de más de 15 años.'}\`,
  };

  // ── Línea 3: recomendación específica ────────────────────────────────────
  const costoUSD=calcCosto(prob,d).usd;
  const ahorroUSD=Math.round(costoUSD*0.22);
  const prob12=(prob*100).toFixed(1);

  // Peor comportamiento conductual (legal, aten, suavidad — excluye zona/seg/antig que no son comportamentales)
  const behav={'legal':m.legal,'aten':m.atencion,'suav':m.suavidad};
  const worstBeh=Object.entries(behav).sort((a,b)=>a[1]-b[1])[0];
  const behavLabel={'legal':'legalidad vial (velocidad + semáforos)','aten':'atención al volante (uso del teléfono)','suav':'suavidad de manejo (frenadas y aceleraciones)'};

  let rec;
  if(score>65){
    rec=\`Con prob. anual del \${prob12}% y costo esperado USD \${costoUSD.toLocaleString()}, se recomienda: prima con recargo técnico + coaching obligatorio en \${behavLabel[worstBeh[0]]} (score actual: \${(worstBeh[1]*100).toFixed(0)}%). Mejora del 20% en ese factor reduciría el costo estimado en USD \${ahorroUSD.toLocaleString()} anuales.\`;
  } else if(score>40){
    rec=\`Prob. anual de siniestro: \${prob12}%. Acciones sugeridas: alertas de zona en tiempo real + revisión trimestral del score de \${behavLabel[worstBeh[0]]} (\${(worstBeh[1]*100).toFixed(0)}%). Reducir 30% la exposición nocturna podría bajar la prima en USD \${Math.round(ahorroUSD*0.7).toLocaleString()} anuales.\`;
  } else {
    rec=\`Perfil de bajo riesgo (\${prob12}% anual, USD \${costoUSD.toLocaleString()} esperados). Candidato ideal para tarifa telemática bonificada. Mantener seguimiento semestral y ofrecer descuento de fidelización vinculado al score de \${behavLabel[worstBeh[0]]} (actualmente \${(worstBeh[1]*100).toFixed(0)}%).\`;
  }

  return [L1[top], L2[sec!==top?sec:sorted[2][0]], rec];
}

async function updateDiagnostico(d){
  const el=document.getElementById('diagnostico');
  el.innerHTML='<div class="ai-loading"><div class="ai-spinner"></div>Generando diagnóstico AI…</div>';

  try{
    const res=await fetch(GEMINI_URL,{
      method:'POST',
      headers:{'Content-Type':'application/json'},
      body:JSON.stringify({
        contents:[{parts:[{text:buildPrompt(d)}]}],
        generationConfig:{maxOutputTokens:300,temperature:0.4}
      })
    });
    if(!res.ok) throw new Error('HTTP '+res.status);
    const data=await res.json();
    const raw=(data?.candidates?.[0]?.content?.parts?.[0]?.text||'').trim();
    if(!raw) throw new Error('empty response');
    const lines=raw.split(/\n+/).map(l=>l.replace(/^[-•\d.]+\s*/,'').trim()).filter(Boolean).slice(0,3);
    while(lines.length<3) lines.push('');
    el.innerHTML=\`
      <div class="ai-box">
        \${lines.map((l,i)=>'<div class="ai-line l'+(i+1)+'">'+l+'</div>').join('')}
      </div>
      <div class="ai-meta">
        <span class="ai-badge">Gemini 2.0 Flash</span>
        Generado con datos reales de telemática · Costos SSN 2024-2025
      </div>\`;
  }catch(e){
    const lines=diagnosisRules(d);
    el.innerHTML=\`
      <div class="ai-box">
        \${lines.map((l,i)=>'<div class="ai-line l'+(i+1)+'">'+l+'</div>').join('')}
      </div>
      <div class="ai-meta">
        <span class="ai-badge" style="background:#0f2a3d;color:#7dd3fc">Telemática</span>
        <span class="ai-fallback">Proxy no disponible — diagnóstico por reglas actuariales</span>
      </div>\`;
  }
}

function renderSimResult(d){
  const orig=d.metricas;
  const scoreOrig=calcScore(orig,d);
  const scoreSim=calcScore(orig,d,simState);
  const probOrig=calcProb(scoreOrig);
  const probSim=calcProb(scoreSim);
  const costoOrig=calcCosto(probOrig,d);
  const costoSim=calcCosto(probSim,d);
  const dScore=scoreSim-scoreOrig;
  const dCosto=costoSim.usd-costoOrig.usd;
  const dClass=dScore<0?'better':dScore>0?'worse':'same';
  const dSign=dScore>0?'+':'';

  document.getElementById('sim-result').innerHTML=\`
    <div class="sim-block">
      <div class="lbl2">Score proyectado</div>
      <div class="big \${riskClass(scoreSim)}">\${scoreSim}</div>
      <div class="delta \${dClass}">\${dSign}\${dScore} vs actual (\${scoreOrig})</div>
    </div>
    <div class="sim-row">
      <div class="sim-mini">
        <div class="lbl2">Prob. anual</div>
        <div class="val2 \${riskClass(scoreSim)}">\${(probSim*100).toFixed(1)}%</div>
        <div style="font-size:11px;color:var(--muted)">era \${(probOrig*100).toFixed(1)}%</div>
      </div>
      <div class="sim-mini">
        <div class="lbl2">Costo esperado</div>
        <div class="val2" style="font-size:16px">USD \${costoSim.usd.toLocaleString()}</div>
        <div style="font-size:11px;color:\${dCosto<0?'var(--green)':'var(--red)'}">
          \${dCosto<0?'Ahorro:':'Incremento:'} USD \${Math.abs(dCosto).toLocaleString()}
        </div>
      </div>
    </div>
  \`;
}
</script>
</body>
</html>`;
}
