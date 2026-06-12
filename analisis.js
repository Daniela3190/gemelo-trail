const XLSX     = require('xlsx');
const polyline = require('@mapbox/polyline');
const turf     = require('@turf/turf');
const path     = require('path');

const DIR = __dirname;

// ─── Helpers ──────────────────────────────────────────────────────────────────

function readExcel(file) {
  const wb = XLSX.readFile(path.join(DIR, file));
  const ws = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, { defval: null });
}

function showFile(name, rows) {
  console.log('\n' + '='.repeat(65));
  console.log('ARCHIVO: ' + name);
  console.log('='.repeat(65));
  console.log(`Filas: ${rows.length}`);
  const cols = rows.length > 0 ? Object.keys(rows[0]) : [];
  console.log(`Columnas (${cols.length}): ${JSON.stringify(cols)}`);
  console.log('\n--- 3 filas de muestra ---');
  rows.slice(0, 3).forEach((r, i) => {
    const preview = {};
    for (const [k, v] of Object.entries(r)) {
      preview[k] = typeof v === 'string' && v.length > 80 ? v.slice(0, 80) + '…' : v;
    }
    console.log(`[${i}] ${JSON.stringify(preview)}`);
  });
}

// ─── 1. Leer archivos ─────────────────────────────────────────────────────────

const viajes = readExcel('viajes.xlsx');
const scores = readExcel('scores.csv');
const sinies = readExcel('siniestros_viales_hechos.xlsx');
const ssn    = readExcel('ssn_20242025_desarrollo_siniestros_automotores.xlsx');

showFile('viajes.xlsx',   viajes);
showFile('scores.csv',    scores);
showFile('siniestros_viales_hechos.xlsx', sinies);
showFile('ssn_20242025_desarrollo_siniestros_automotores.xlsx', ssn);

// ─── 2. Decodificar polylines ─────────────────────────────────────────────────

console.log('\n' + '='.repeat(65));
console.log('DECODIFICANDO POLYLINES');
console.log('='.repeat(65));

const viajesGeo = viajes.map(row => {
  let coords = [];
  const enc = row['polyline'];
  if (enc && typeof enc === 'string' && enc.trim().length > 0) {
    try { coords = polyline.decode(enc); } catch (_) {}
  }
  return { ...row, _coords: coords };
});

const nDecoded = viajesGeo.filter(v => v._coords.length > 0).length;
console.log(`Polylines decodificadas: ${nDecoded} / ${viajes.length}`);
console.log(`Primer punto viaje[0]: ${JSON.stringify(viajesGeo[0]?._coords?.[0])}`);

// ─── Índice espacial: siniestros GRAVE + MORTAL, radio 100 m ─────────────────

const siniesGM = sinies.filter(s => {
  const g = (s['gravedad_siniestro'] || '').toUpperCase();
  return g === 'GRAVE' || g === 'MORTAL';
});
console.log(`\nSiniestros filtrados (GRAVE+MORTAL): ${siniesGM.length} de ${sinies.length} totales`);

const CELL     = 0.01;   // ~1 km de resolución
const RADIO_KM = 0.1;    // 100 m
const BBOX_DEG = (RADIO_KM / 111) * 3; // margen bounding-box para descarte rápido

const sinGrid = new Map();
for (const s of siniesGM) {
  const lat = parseFloat(s['latitud_siniestro']);
  const lng = parseFloat(s['longitud_siniestro']);
  if (isNaN(lat) || isNaN(lng)) continue;
  const key = `${Math.round(lat / CELL)},${Math.round(lng / CELL)}`;
  if (!sinGrid.has(key)) sinGrid.set(key, []);
  sinGrid.get(key).push([lat, lng]);
}
console.log(`Celdas en grilla espacial: ${sinGrid.size}`);

// Devuelve true si algún punto del viaje está a <100 m de un siniestro GRAVE/MORTAL
function cercaSiniestro(coords) {
  if (!coords || coords.length === 0) return false;
  for (let i = 0; i < coords.length; i += 5) {   // muestreo cada 5 puntos
    const [lat, lng] = coords[i];
    const pt = turf.point([lng, lat]);             // turf usa [lng, lat]
    const r0 = Math.round(lat / CELL);
    const c0 = Math.round(lng / CELL);
    for (let dr = -1; dr <= 1; dr++) {
      for (let dc = -1; dc <= 1; dc++) {
        const cands = sinGrid.get(`${r0 + dr},${c0 + dc}`);
        if (!cands) continue;
        for (const [slat, slng] of cands) {
          if (Math.abs(lat - slat) > BBOX_DEG || Math.abs(lng - slng) > BBOX_DEG) continue;
          const sinPt = turf.point([slng, slat]);
          if (turf.distance(pt, sinPt, { units: 'kilometers' }) < RADIO_KM) return true;
        }
      }
    }
  }
  return false;
}

// ─── Helpers de hora y scores ─────────────────────────────────────────────────

function getHora(val) {
  if (!val) return null;
  if (typeof val === 'number') {
    const d = XLSX.SSF.parse_date_code(val);
    return d ? d.H : null;
  }
  const m = String(val).match(/[T ](\d{1,2}):/);
  return m ? parseInt(m[1]) : null;
}

// Promedio ponderado por distancia de cada score
const scoresPor = {};
for (const s of scores) {
  const u = s['usuario'];
  if (!scoresPor[u]) scoresPor[u] = { aten: 0, suav: 0, legal: 0, dist: 0 };
  const d   = s['distancia_total'] || 0;
  scoresPor[u].aten  += (s['atencion_promedio_distancia']  ?? s['atencion_promedio']  ?? 0) * d;
  scoresPor[u].suav  += (s['suavidad_promedio_distancia']  ?? s['suavidad_promedio']  ?? 0) * d;
  scoresPor[u].legal += (s['legal_promedio_distancia']     ?? s['legal_promedio']     ?? 0) * d;
  scoresPor[u].dist  += d;
}

function sc(id, campo) {
  const d = scoresPor[id];
  if (!d || d.dist === 0) return '—';
  return (d[campo] / d.dist * 100).toFixed(1) + '%';
}

// ─── 3. Calcular métricas por conductor ───────────────────────────────────────

const CONDUCTORES = {
  '6900fe050f6bdc080000482f': { nombre: 'Micaela T.', auto: 'Fox 2005' },
  '6920dafb8eb2c10800000b37': { nombre: 'Jorge D.',   auto: 'Jeep 2024' },
  '6920e5b43a13660800000b4e': { nombre: 'Nico S.',    auto: 'Peugeot 208 2023' },
};

const porConductor = {};
for (const v of viajesGeo) {
  const c = v['usuario'];
  if (!porConductor[c]) porConductor[c] = [];
  porConductor[c].push(v);
}

console.log('\nConductores encontrados:', Object.keys(porConductor));
console.log('Calculando pct_zonas_riesgo (GRAVE+MORTAL, radio 100 m)...\n');

const tabla = Object.entries(porConductor).map(([id, trips]) => {
  const info = CONDUCTORES[id] || { nombre: id.slice(-6), auto: '?' };
  const n    = trips.length;

  // pct_nocturno: viajes que empezaron entre 22h y 6h
  let nocturnos = 0;
  for (const t of trips) {
    const h = getHora(t['comienzo']);
    if (h !== null && (h >= 22 || h < 6)) nocturnos++;
  }

  // velocidad_promedio: distancia_m / duracion_s * 3.6
  let sumVel = 0, cntVel = 0;
  for (const t of trips) {
    const dist = parseFloat(t['distancia_m']);
    const dur  = parseFloat(t['duracion']);
    if (!isNaN(dist) && !isNaN(dur) && dur > 0) { sumVel += (dist / dur) * 3.6; cntVel++; }
  }

  // variabilidad_ruta: cantidad de polylines únicas
  const rutasUnicas = new Set(trips.map(t => t['polyline'])).size;

  // pct_zonas_riesgo
  let riesgo = 0;
  for (const t of trips) {
    if (cercaSiniestro(t._coords)) riesgo++;
  }

  return {
    conductor:     info.nombre,
    auto:          info.auto,
    n_viajes:      n,
    pct_nocturno:  (nocturnos / n * 100).toFixed(1) + '%',
    vel_prom:      cntVel > 0 ? (sumVel / cntVel).toFixed(1) + ' km/h' : 'N/A',
    var_ruta:      rutasUnicas,
    pct_riesgo:    (riesgo / n * 100).toFixed(1) + '%',
    score_aten:    sc(id, 'aten'),
    score_suav:    sc(id, 'suav'),
    score_legal:   sc(id, 'legal'),
  };
});

// ─── 4. Tabla resumen ─────────────────────────────────────────────────────────

console.log('='.repeat(80));
console.log('TABLA RESUMEN — GEMELO DIGITAL TRAIL');
console.log('='.repeat(80));
console.table(tabla);
