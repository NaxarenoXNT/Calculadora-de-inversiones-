// build_indices.js
// Genera indices.json e indices.js a partir de las fuentes oficiales:
//   - CER diario (2010-2016): BCRA, archivos cerAAAA.xls
//   - IPC mensual (dic-2016 en adelante): INDEC, serie_ipc_divisiones.csv
//   - Tasa pasiva digital BIP (30 días, TNAV): Banco Provincia, PDF "tasas_frecuentes_historico"
//
// Uso: node build_indices.js
// Requiere: npm install xlsx pdf-parse

const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const { PDFParse } = require('pdf-parse');

const DATA_DIR = path.join(__dirname, 'data');
const CORTE = '2016-12-31';
const CER_DESDE = '2010-01-01';
const CER_ANIOS = [2010, 2011, 2012, 2013, 2014, 2015, 2016];
const IPC_URL = 'https://www.indec.gob.ar/ftp/cuadros/economia/serie_ipc_divisiones.csv';
const BIP_URL = 'https://www.bancoprovincia.com.ar/CDN/Get/tasas_frecuentes_historico';
const UA = { headers: { 'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)' } };

// ---- utilidades ----
function parseNumero(s) {
  if (s === null || s === undefined) return NaN;
  s = String(s).trim();
  if (!s) return NaN;
  const tieneComa = s.includes(',');
  const tienePunto = s.includes('.');
  if (tieneComa && tienePunto) {
    // coma = decimal, punto = separador de miles
    return parseFloat(s.replace(/\./g, '').replace(',', '.'));
  }
  if (tieneComa) return parseFloat(s.replace(',', '.'));
  return parseFloat(s);
}

function hoyLocalISO() {
  const d = new Date();
  const p = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}`;
}

function fmtFecha(ymd) {
  // ymd = YYYYMMDD
  return `${ymd.slice(0, 4)}-${ymd.slice(4, 6)}-${ymd.slice(6, 8)}`;
}

// ---- descarga de archivos (cache en data/) ----
async function descargar(url, dest) {
  if (fs.existsSync(dest)) {
    process.stdout.write(`  (cache) ${path.basename(dest)}\n`);
    return fs.readFileSync(dest);
  }
  process.stdout.write(`  (descarga) ${url}\n`);
  const res = await fetch(url, UA);
  if (!res.ok) throw new Error(`HTTP ${res.status} para ${url}`);
  const buf = Buffer.from(await res.arrayBuffer());
  fs.mkdirSync(path.dirname(dest), { recursive: true });
  fs.writeFileSync(dest, buf);
  return buf;
}

// ---- extraccion CER ----
function extraerCer(buf) {
  const wb = XLSX.read(buf, { type: 'buffer' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '' });

  let fechaCol = -1, coefCol = -1;
  for (const row of rows) {
    let f = -1, c = -1;
    for (let i = 0; i < row.length; i++) {
      const cell = String(row[i]).trim().toLowerCase();
      if (cell === 'fecha') f = i;
      if (cell === 'coef001') c = i;
    }
    if (f >= 0 && c >= 0) { fechaCol = f; coefCol = c; break; }
  }
  if (fechaCol < 0 || coefCol < 0) throw new Error('No se encontraron las columnas fecha/coef001');

  const out = [];
  for (const row of rows) {
    const f = String(row[fechaCol] ?? '').trim();
    const v = String(row[coefCol] ?? '').trim();
    if (/^\d{8}$/.test(f) && v !== '') {
      const valor = parseNumero(v);
      if (!isNaN(valor)) out.push({ fecha: f, valor });
    }
  }
  return out;
}

// ---- extraccion IPC ----
function extraerIpc(buf) {
  const csv = new TextDecoder('latin1').decode(buf);
  const lines = csv.split(/\r?\n/);
  const mapa = {};
  for (const line of lines) {
    const cols = line.split(';');
    if (cols.length < 8) continue;
    const desc = cols[1].trim().toUpperCase();
    const region = cols[7].trim();
    if (desc === 'NIVEL GENERAL' && region === 'Nacional') {
      const periodo = cols[3].trim();
      const valor = parseNumero(cols[4]);
      if (/^\d{6}$/.test(periodo) && !isNaN(valor)) {
        const fecha = `${periodo.slice(0, 4)}-${periodo.slice(4, 6)}-01`;
        mapa[fecha] = valor;
      }
    }
  }
  return Object.keys(mapa).sort().map(f => ({ fecha: f, valor: mapa[f] }));
}

// ---- extraccion tasa pasiva BIP (Banco Provincia, PDF) ----
async function extraerBipSerie(buf) {
  const pdf = new PDFParse({ data: new Uint8Array(buf) });
  const t = await pdf.getText();
  const paginas = t.pages.filter(p => (p.text || '').includes('Cuadro V-'));
  if (!paginas.length) throw new Error('No se encontró el Cuadro V en el PDF');
  const texto = paginas.map(p => p.text).join('\n');

  const fmt = (d, m, y) => `${y}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
  const toNum = s => parseFloat(s.replace(/\./g, '').replace(',', '.'));

  const mapa = {};
  for (const line of texto.split('\n')) {
    const dm = line.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (!dm) continue;
    const fecha = fmt(+dm[1], +dm[2], +dm[3]);
    const tasas = [...line.matchAll(/(\d{1,3}(?:[.,]\d{1,3})?)%/g)].map(m => toNum(m[1]));
    if (!tasas.length) continue;
    mapa[fecha] = Math.max(mapa[fecha] || 0, Math.max(...tasas));
  }

  // Corrección de artefacto de extracción del PDF: "1/8/2024 23,50%" es en realidad 1/8/2016.
  if (mapa['2024-08-01'] !== undefined && mapa['2016-08-01'] === undefined) {
    mapa['2016-08-01'] = mapa['2024-08-01'];
    delete mapa['2024-08-01'];
  }

  const fechas = Object.keys(mapa).sort();
  const bip = fechas.map(f => ({ fecha: f, tasaTNAV: Number(mapa[f].toFixed(2)) }));

  const act = texto.match(/Actualizado al (\d{1,2})\/(\d{1,2})\/(\d{4})/);
  await pdf.destroy();

  return {
    bip,
    bipMeta: {
      tipo: 'TNAV (Tasa Nominal Anual Vencida)',
      periodo: '30 días',
      modalidad: 'Plazo Fijo Digital (BIP)',
      fuente: 'Banco Provincia - "Tasas de uso judicial e históricas de consulta frecuente" - B - Cuadro V',
      url: BIP_URL,
      criterio: 'tasa pasiva más alta (SCBA)',
      columna: 'máximo entre todas las columnas publicadas (Personas Humanas y Personas Jurídicas, todos los tramos de monto)',
      desde: bip.length ? bip[0].fecha : null,
      actualizado: act ? fmt(+act[1], +act[2], +act[3]) : hoyLocalISO()
    }
  };
}

// ---- generar archivo único autocontenido ----
function generarAutocontenido(indices) {
  const calcPath = path.join(__dirname, 'Calc.html');
  const xlsxPath = path.join(__dirname, 'xlsx.full.min.js');
  if (!fs.existsSync(calcPath) || !fs.existsSync(xlsxPath)) {
    console.log('[aviso] Falta Calc.html o xlsx.full.min.js; se saltea el archivo único.');
    return;
  }

  let html = fs.readFileSync(calcPath, 'utf8');
  const xlsx = fs.readFileSync(xlsxPath, 'utf8');

  const indicesInline = 'window.INDICES = ' + JSON.stringify(indices) + ';';
  // evita que un '</script>' dentro del JS cierre el <script> antes de tiempo
  const xlsxInline = xlsx.replace(/<\/script/gi, '<\\/script');

  // usar función como reemplazo evita que '$' dentro del contenido se interprete mal
  html = html.replace('<script src="indices.js"></script>', () => '<script>' + indicesInline + '</script>');
  html = html.replace('<script src="xlsx.full.min.js"></script>', () => '<script>' + xlsxInline + '</script>');

  const out = path.join(__dirname, 'CalculadoraLiquidaciones.html');
  fs.writeFileSync(out, html, 'utf8');
  console.log('[CalculadoraLiquidaciones.html] generado (' + (html.length / 1024).toFixed(0) + ' KB).');
}

// ---- main ----
(async () => {
  console.log('=== Generando series CER + IPC ===\n');
  fs.mkdirSync(DATA_DIR, { recursive: true });

  // 1) CER diario
  console.log('Descargando/leyendo CER (BCRA)...');
  const cerBruto = [];
  for (const anio of CER_ANIOS) {
    const url = `https://www.bcra.gob.ar/pdfs/PublicacionesEstadisticas/cer${anio}.xls`;
    const buf = await descargar(url, path.join(DATA_DIR, `cer${anio}.xls`));
    const datos = extraerCer(buf);
    console.log(`  cer${anio}.xls -> ${datos.length} filas`);
    cerBruto.push(...datos);
  }

  // merge + dedup por fecha
  const cerMapa = {};
  for (const d of cerBruto) cerMapa[d.fecha] = d.valor;
  const cer = Object.keys(cerMapa).sort()
    .filter(f => f >= '20100101' && f <= '20161231')
    .map(f => ({ fecha: fmtFecha(f), valor: cerMapa[f] }));

  console.log(`\nCER final: ${cer.length} valores (${cer[0].fecha} .. ${cer[cer.length - 1].fecha})`);

  // 2) IPC mensual
  console.log('\nDescargando/leyendo IPC (INDEC)...');
  const ipcBuf = await descargar(IPC_URL, path.join(DATA_DIR, 'serie_ipc_divisiones.csv'));
  const ipc = extraerIpc(ipcBuf);
  console.log(`IPC final: ${ipc.length} valores (${ipc[0].fecha} .. ${ipc[ipc.length - 1].fecha})`);
  console.log(`  ultimo IPC: ${ipc[ipc.length - 1].fecha} = ${ipc[ipc.length - 1].valor}`);

  // 3) tasa pasiva digital BIP (Banco Provincia)
  console.log('\nDescargando/leyendo tasa pasiva BIP (Banco Provincia)...');
  let bip = [], bipMeta = null;
  try {
    const bipBuf = await descargar(BIP_URL, path.join(DATA_DIR, 'tasas_frecuentes_historico.pdf'));
    const resBip = await extraerBipSerie(bipBuf);
    bip = resBip.bip; bipMeta = resBip.bipMeta;
    console.log(`BIP final: ${bip.length} valores (${bip[0].fecha} .. ${bip[bip.length - 1].fecha})`);
  } catch (err) {
    console.log(`  [aviso] No se pudo extraer la serie BIP: ${err.message}`);
  }

  // 4) armar objeto
  const indices = {
    meta: {
      fuenteCer: 'BCRA - Coeficiente de Estabilización de Referencia (CER), base 2.2.2002=1',
      fuenteIpc: 'INDEC - IPC Nivel general, Total Nacional, base dic-2016=100',
      cerDesde: CER_DESDE,
      corte: CORTE,
      actualizado: hoyLocalISO(),
    },
    cer,
    ipc,
  };
  if (bip.length) {
    indices.bip = bip;
    indices.bipMeta = bipMeta;
  }

  // 5) escribir indices.json
  fs.writeFileSync(path.join(__dirname, 'indices.json'), JSON.stringify(indices, null, 2), 'utf8');
  console.log('\n[indices.json] generado.');

  // 6) escribir indices.js (para cargar sin CORS en file://)
  const js = '/* Generado por build_indices.js - no editar a mano */\nwindow.INDICES = ' +
    JSON.stringify(indices) + ';\n';
  fs.writeFileSync(path.join(__dirname, 'indices.js'), js, 'utf8');
  console.log('[indices.js] generado.');

  // 7) generar archivo único autocontenido (para distribuir/descargar)
  generarAutocontenido(indices);

  // 8) validaciones rapidas
  console.log('\n=== Validaciones ===');
  const cerPuntos = { '2010-01-01': 2.3691, '2013-01-01': 3.1857, '2016-12-31': 6.8378 };
  for (const [f, esperado] of Object.entries(cerPuntos)) {
    const real = cer.find(x => x.fecha === f);
    console.log(`  CER ${f}: ${real ? real.valor : 'NO ENCONTRADO'} (esperado ~${esperado})`);
  }
  const ipcBase = ipc.find(x => x.fecha === '2016-12-01');
  console.log(`  IPC 2016-12-01: ${ipcBase ? ipcBase.valor : 'NO ENCONTRADO'} (esperado 100)`);
  const bipPuntos = { '2008-08-19': 12, '2016-08-01': 23.5, '2023-10-17': 133, '2025-12-11': 25 };
  for (const [f, esperado] of Object.entries(bipPuntos)) {
    const real = bip.find(x => x.fecha === f);
    console.log(`  BIP ${f}: ${real ? real.tasaTNAV : 'NO ENCONTRADO'} (esperado ${esperado})`);
  }
})().catch(err => {
  console.error('\nERROR:', err.message);
  process.exit(1);
});

