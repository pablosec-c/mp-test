/**********************************************************************
 *  SYNC MERCADO PAGO (Take 15) → SUPABASE
 *  Corre automático cada día desde GitHub Actions.
 *
 *  Flujo:
 *   1. Lista los reportes de MP y agarra el más nuevo procesado
 *   2. Lo descarga (xlsx)
 *   3. Lee cada movimiento y aplica las 5 reglas + memoria
 *   4. Chequea duplicados (mp_payment_id)
 *   5. Inserta los nuevos en Supabase como "pendiente_categorizar"
 **********************************************************************/

const XLSX = require('xlsx');

// ===== CONSTANTES =====
const SUPABASE_URL = 'https://kqngnjbtkddkhiahcsdo.supabase.co';
const ACCOUNT_ID_TAKE15 = '45af468c-d87e-44d6-a86a-b809a77456ba';
const TYPE_INGRESO = '7f38e08e-5d25-45b7-a58b-07e70fe8474c';   // Cobranza Cliente (provisorio)
const TYPE_EGRESO = '6670f68b-41dc-4de1-9a41-ffa1e4d178d4';    // Transferencia (provisorio)
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

// ===== SUPABASE: mp_payment_id ya importados =====
async function getExistentes() {
  const url = `${SUPABASE_URL}/rest/v1/transactions?select=mp_payment_id&origen_mp=eq.true&account_id=eq.${ACCOUNT_ID_TAKE15}&mp_payment_id=not.is.null`;
  const r = await fetch(url, { headers: { apikey: SUPABASE_KEY, Authorization: `Bearer ${SUPABASE_KEY}` } });
  if (!r.ok) throw new Error(`Error consultando existentes: ${r.status}`);
  const arr = await r.json();
  return new Set(arr.map((x) => String(x.mp_payment_id)));
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
  // Memoria
  const k = norm(pagador);
  if (k && memoria[k]) {
    const m = memoria[k];
    return { categoria: m.categoria, subcategoria: m.subcategoria, provider_id: m.provider_id };
  }
  // R3: rendimientos
  if (valor > 0 && tipoMedio === '' && liquidado === 'false') return { categoria: 'Financiero', subcategoria: 'Rendimientos', provider_id: null };
  // R4: entra plata sin match
  if (valor > 0) return { categoria: 'Cobranzas clientes', subcategoria: null, provider_id: null };
  // R5: sale sin match
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

// ===== SUPABASE: insertar lote =====
async function insertar(lote) {
  const r = await fetch(`${SUPABASE_URL}/rest/v1/transactions`, {
    method: 'POST',
    headers: {
      apikey: SUPABASE_KEY,
      Authorization: `Bearer ${SUPABASE_KEY}`,
      'Content-Type': 'application/json',
      Prefer: 'return=minimal',
    },
    body: JSON.stringify(lote),
  });
  if (!r.ok) throw new Error(`Error insertando: ${r.status} ${await r.text()}`);
}

// ===== MAIN =====
(async () => {
  console.log('=== INICIO sync MP → Supabase ===');

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

  // 4. Cargar existentes + memoria
  const [existentes, memoria] = await Promise.all([getExistentes(), getMemoria()]);
  console.log(`   Ya importados: ${existentes.size} | Destinatarios en memoria: ${Object.keys(memoria).length}`);

  // 5. Armar movimientos nuevos
  const nuevos = [];
  for (const r of rows) {
    const mpId = String(r[COL.ID] || '').trim();
    if (!mpId) continue;
    if (existentes.has(mpId)) continue;
    const valor = parseFloat(r[COL.VALOR]) || 0;
    if (valor === 0) continue;

    const c = clasificar(r, memoria);
    nuevos.push({
      transaction_date: String(r[COL.FECHA]).slice(0, 10),
      amount: Math.abs(valor),
      description: armarDesc(r),
      account_id: ACCOUNT_ID_TAKE15,
      transaction_type_id: valor >= 0 ? TYPE_INGRESO : TYPE_EGRESO,
      categoria: null,
      categoria_sugerida: c.categoria,
      subcategoria: c.subcategoria,
      provider_id: c.provider_id,
      cliente: valor >= 0 ? String(r[COL.PAGADOR] || '') || null : null,
      mp_payment_id: mpId,
      origen_mp: true,
      estado_revision: 'pendiente_categorizar',
    });
  }
  console.log(`✅ Movimientos NUEVOS a importar: ${nuevos.length}`);

  if (nuevos.length === 0) {
    console.log('=== FIN: nada nuevo ===');
    return;
  }

  // 6. Insertar en lotes de 100
  for (let i = 0; i < nuevos.length; i += 100) {
    await insertar(nuevos.slice(i, i + 100));
  }
  console.log(`✅ IMPORTADOS: ${nuevos.length} movimientos como PENDIENTES (badge NUEVO azul)`);
  console.log('=== FIN sync MP ===');
})().catch((e) => {
  console.error('❌ ERROR:', e.message);
  process.exit(1);
});
