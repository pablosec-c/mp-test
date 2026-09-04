/**********************************************************************
 *  SYNC MERCADO PAGO (Take 15) → SUPABASE (tabla LEDGER, doble entrada)
 *  Corre automático cada día desde GitHub Actions.
 *
 *  Flujo:
 *   1. Lista los reportes de MP y agarra el más nuevo procesado
 *   2. Lo descarga (xlsx)
 *   3. Lee cada movimiento y aplica las 5 reglas + memoria
 *   4. Chequea duplicados contra LEDGER (mp_payment_id)
 *   5. Inserta los nuevos en LEDGER como 2 filas por movimiento
 *      (line_type='cuenta' + line_type='categoria', mismo entry_id,
 *       montos opuestos → suma cero, como exige el constraint)
 **********************************************************************/

const XLSX = require('xlsx');
const crypto = require('crypto');

// ===== CONSTANTES =====
const SUPABASE_URL = 'https://kqngnjbtkddkhiahcsdo.supabase.co';
const ACCOUNT_ID_TAKE15 = '45af468c-d87e-44d6-a86a-b809a77456ba';
const MI_CUIT = '20406633587';

// Columnas del Excel de MP (0-based)
const COL = { FECHA: 0, ID: 4, TIPO_MEDIO: 6, TIPO_OP: 9, VALOR: 10, COMISION: 12, LIQUIDADO: 39, CUIT: 55, PAGADOR: 56 };

// User-Agent de navegador (obligatorio para pasar el PolicyAgent de MP)
const UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36';

const MP_TOKEN = process.env.MP_TOKEN;
const SUPABASE_KEY = process.env.SUPABASE_KEY;

const norm = (s) => String(s || '').trim().toLowerCase().replace(/\s+/g, ' ');

// ===== MP: listar reportes =====
async function listarReportes() {
  const r = await fetch('https://api.mercadopago.com/v1/account/settlement_report/list', {
    headers: { Authorization: `Bearer ${MP_TOKEN}`, 'User-Agent': UA, Accept: 'application/json' },
  });
  if (!r.ok) throw new Error(`Error listando reportes: ${r.status} ${await r.text()}`);
  return await r.json();
}

// ===== MP: descargar reporte =====
async function descargarReporte(fileName) {
  const r = await fetch(`https://api.mercadopago.com/v1/account/settlement_report/${fileName}`, {
    headers: { Authorization: `Bearer ${MP_TOKEN}`, 'User-Agent': UA },
  });
  if (!r.ok) throw new Error(`Error descargando ${fileName}: ${r.status}`);
  return Buffer.from(await r.arrayBuffer());
}

// ===== SUPABASE: mp_payment_id ya importados (en LEDGER) =====
async function getExistentes() {
  // Ahora consulta LEDGER (fuente única de verdad), no la tabla vieja transactions.
  // Trae solo las filas de tipo 'cuenta' porque ahí guardamos el mp_payment_id.
  // Paginación por si hay más de 1000 filas.
  const existentes = new Set();
  let from = 0;
  const CHUNK = 1000;
  while (true) {
    const url = `${SUPABASE_URL}/rest/v1/ledger?select=mp_payment_id&line_type=eq.cuenta&origen_mp=eq.true&account_id=eq.${ACCOUNT_ID_TAKE15}&mp_payment_id=not.is.null`;
    const r = await fetch(url, {
      headers: {
        apikey: SUPABASE_KEY,
        Authorization: `Bearer ${SUPABASE_KEY}`,
        Range: `${from}-${from + CHUNK - 1}`,
        'Range-Unit': 'items',
      },
    });
    if (!r.ok) throw new Error(`Error consultando existentes: ${r.status} ${await r.text()}`);
    const arr = await r.json();
    arr.forEach((x) => existentes.add(String(x.mp_payment_id)));
    if (arr.length < CHUNK) break;
    from += CHUNK;
  }
  return existentes;
}

// ===== SUPABASE: memoria de destinatarios =====
async function getMemoria() {
  const opts = { headers: { apikey: SUPABASE_KEY, Authorization: `Bearer ${SUPABASE_KEY}` } };
  const [dests, idents] = await Promise.all([
    fetch(`${SUPABASE_URL}/rest/v1/mp_destinatarios?select=id,categoria,subcategoria,unidad_negocio,provider_id`, opts).then((r) => r.json()),
    fetch(`${SUPABASE_URL}/rest/v1/mp_destinatario_identificadores?select=destinatario_id,tipo,valor`, opts).then((r) => r.json()),
  ]);
  const byId = {};
  dests.forEach((d) => (byId[d.id] = d));
  const mapa = {};
  idents.forEach((i) => {
    const d = byId[i.destinatario_id];
    if (d) mapa[norm(i.valor)] = d;
  });
  return mapa;
}

// ===== CLASIFICADOR (las 5 reglas) =====
function clasificar(r, memoria) {
  const valor = parseFloat(r[COL.VALOR]) || 0;
  const tipoOp = String(r[COL.TIPO_OP] || '');
  const pagador = String(r[COL.PAGADOR] || '').trim();
  const cuit = String(r[COL.CUIT] || '').trim();
  const tipoMedio = String(r[COL.TIPO_MEDIO] || '').trim();
  const liquidado = String(r[COL.LIQUIDADO] || '').trim();

  // R1: PAYOUTS
  if (tipoOp.toUpperCase().includes('PAYOUT')) return { categoria: 'Traspaso', subcategoria: 'Traspaso a Banco Francés', provider_id: null };
  // R2: mi propio CUIT saliendo → retiro personal
  if (cuit === MI_CUIT && valor < 0) return { categoria: 'Personales', subcategoria: 'Pablo Seco', provider_id: null };
  // Memoria (busca por nombre/CUIT/CBU normalizado)
  for (const k of [norm(cuit), norm(pagador)]) {
    if (k && memoria[k]) {
      const m = memoria[k];
      return { categoria: m.categoria, subcategoria: m.subcategoria, provider_id: m.provider_id };
    }
  }
  // R3: rendimientos
  if (valor > 0 && tipoMedio === '' && liquidado === 'false') return { categoria: 'Financiero', subcategoria: 'Rendimientos', provider_id: null };
  // R4: entra plata sin match → cobranza
  if (valor > 0) return { categoria: 'Cobranzas clientes', subcategoria: null, provider_id: null };
  // R5: sale sin match → queda sin categoría, admin decide
  return { categoria: null, subcategoria: null, provider_id: null };
}

function armarDesc(r) {
  const pagador = String(r[COL.PAGADOR] || '').trim();
  const tipoOp = String(r[COL.TIPO_OP] || '').trim();
  const comision = parseFloat(r[COL.COMISION]) || 0;
  let d = pagador || tipoOp || 'Movimiento MP';
  if (comision !== 0) d += ` · Comisión MP: $${Math.abs(comision).toFixed(2)}`;
  return d;
}

// ===== SUPABASE: insertar lote de LÍNEAS en LEDGER =====
async function insertarLineas(lineas) {
  const r = await fetch(`${SUPABASE_URL}/rest/v1/ledger`, {
    method: 'POST',
    headers: {
      apikey: SUPABASE_KEY,
      Authorization: `Bearer ${SUPABASE_KEY}`,
      'Content-Type': 'application/json',
      Prefer: 'return=minimal',
    },
    body: JSON.stringify(lineas),
  });
  if (!r.ok) throw new Error(`Error insertando: ${r.status} ${await r.text()}`);
}

// Arma las 2 líneas de ledger para un movimiento MP
function armarLineas(row, clasif) {
  const entryId = crypto.randomUUID();
  const fecha = String(row[COL.FECHA]).slice(0, 10);
  const valor = parseFloat(row[COL.VALOR]) || 0;   // signo lo pone MP, no lo tocamos
  const desc = armarDesc(row);
  const mpId = String(row[COL.ID]).trim();
  const pagador = String(row[COL.PAGADOR] || '').trim() || null;

  // Subcategoría que va en la línea de CUENTA (por consistencia con lo que hay en la BD):
  // - si hay subcategoria clasificada → esa
  // - si no, si es ingreso con pagador → nombre del pagador
  // - si no → null
  const subCuenta = clasif.subcategoria || (valor > 0 ? pagador : null);

  return [
    // Línea de CUENTA: lleva el account_id, el signo real, el mp_payment_id y el estado_revision
    {
      entry_id: entryId,
      line_type: 'cuenta',
      account_id: ACCOUNT_ID_TAKE15,
      categoria: null,
      subcategoria: subCuenta,
      descripcion: desc,
      monto: valor,                            // signed (+ ingreso, − egreso)
      fecha,
      mp_payment_id: mpId,
      origen_mp: true,
      estado_revision: 'pendiente_categorizar',
      cliente: valor > 0 ? pagador : null,
      provider_id: clasif.provider_id,
    },
    // Línea de CATEGORÍA: monto opuesto para que sume cero (doble entrada)
    {
      entry_id: entryId,
      line_type: 'categoria',
      account_id: null,
      categoria: clasif.categoria,             // puede ser null → queda pendiente
      subcategoria: clasif.subcategoria,
      descripcion: desc,
      monto: -valor,                           // signo opuesto
      fecha,
      mp_payment_id: null,
      origen_mp: false,
      estado_revision: null,
      cliente: valor > 0 ? pagador : null,
      provider_id: clasif.provider_id,
    },
  ];
}

// ===== MAIN =====
(async () => {
  console.log('=== INICIO sync MP → Supabase (LEDGER) ===');

  // 1. Buscar el reporte más nuevo procesado
  const reportes = await listarReportes();
  const procesados = reportes.filter((r) => r.status === 'processed' && r.file_name);
  if (procesados.length === 0) {
    console.log('❌ No hay reportes procesados disponibles');
    return;
  }
  procesados.sort((a, b) => new Date(b.date_created) - new Date(a.date_created));
  const ultimo = procesados[0];
  console.log(`✅ Reporte más nuevo: ${ultimo.file_name} (${ultimo.date_created})`);

  // 2. Descargar
  const buffer = await descargarReporte(ultimo.file_name);
  console.log(`✅ Descargado: ${buffer.length} bytes`);

  // 3. Leer Excel
  const wb = XLSX.read(buffer, { type: 'buffer' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1 }).slice(1); // saltar headers
  console.log(`✅ Excel leído: ${rows.length} filas`);

  // 4. Cargar existentes (de LEDGER) + memoria
  const [existentes, memoria] = await Promise.all([getExistentes(), getMemoria()]);
  console.log(`   Ya en ledger: ${existentes.size} | Destinatarios en memoria: ${Object.keys(memoria).length}`);

  // 5. Armar líneas (2 por movimiento nuevo)
  let lineasParaInsertar = [];
  let contNuevos = 0;
  let contDupIntraArchivo = 0;
  const vistosEnEsteArchivo = new Set();

  for (const r of rows) {
    const mpId = String(r[COL.ID] || '').trim();
    if (!mpId) continue;
    if (existentes.has(mpId)) continue;
    // Dedupe también dentro del mismo archivo (por si MP repite filas)
    if (vistosEnEsteArchivo.has(mpId)) { contDupIntraArchivo++; continue; }
    vistosEnEsteArchivo.add(mpId);

    const valor = parseFloat(r[COL.VALOR]) || 0;
    if (valor === 0) continue;

    const c = clasificar(r, memoria);
    const [linCuenta, linCat] = armarLineas(r, c);
    lineasParaInsertar.push(linCuenta, linCat);
    contNuevos++;
  }
  console.log(`✅ Movimientos NUEVOS a importar: ${contNuevos} (= ${lineasParaInsertar.length} líneas en ledger)`);
  if (contDupIntraArchivo > 0) console.log(`   (Ignorados por repetirse dentro del mismo archivo: ${contDupIntraArchivo})`);

  if (lineasParaInsertar.length === 0) {
    console.log('=== FIN: nada nuevo ===');
    return;
  }

  // 6. Insertar en lotes de 200 LÍNEAS (= 100 movimientos por lote)
  //    Importante: los pares (cuenta+categoria) del mismo entry_id
  //    van siempre juntos, porque los generamos consecutivos y usamos
  //    tamaño de lote par. Así el constraint AFTER de "suma cero por
  //    entry_id" nunca falla por lote cortado a la mitad.
  const BATCH = 200;
  for (let i = 0; i < lineasParaInsertar.length; i += BATCH) {
    await insertarLineas(lineasParaInsertar.slice(i, i + BATCH));
  }
  console.log(`✅ IMPORTADOS: ${contNuevos} movimientos como PENDIENTES (badge REVISAR)`);
  console.log('=== FIN sync MP ===');
})().catch((e) => {
  console.error('❌ ERROR:', e.message);
  process.exit(1);
});
