const express = require('express');
const cors = require('cors');
const http = require('http');
const https = require('https');
const { execFile } = require('child_process');
const XLSX = require('xlsx');
const os = require('os');
const net = require('net');
const path = require('path');
const fs = require('fs');
const dgram = require('dgram');
const multer = require('multer');
const { Server } = require('socket.io');
const { io: createSocketClient } = require('socket.io-client');
const { v4: uuidv4 } = require('uuid');
const crypto = require('crypto');
const {
  decryptDatabaseSecrets,
  encryptDatabaseSecrets,
  isSecretMigrationPending,
  clearSecretMigrationPending
} = require('./secret-store');

const DATA_DIR = process.env.DATA_DIR || path.join(__dirname, 'data');
const PORT = Number(process.env.PORT || 4100);
const DISCOVERY_PORT = Number(process.env.DISCOVERY_PORT || 41234);
const ADMIN_NAME = process.env.ADMIN_NAME || 'Cremería El Güero';
const RENDER_MIRROR_ONLY = String(process.env.RENDER_MIRROR_ONLY || '').trim().toLowerCase() === 'true';
const RENDER_REMOTE_API_KEY = String(process.env.REMOTE_API_KEY || '').trim();
if (RENDER_MIRROR_ONLY && !RENDER_REMOTE_API_KEY) {
  throw new Error('El puente requiere configurar la variable REMOTE_API_KEY en Render.');
}
const SAT_PORTAL_DEFAULT = 'https://portalcfdi.facturaelectronica.sat.gob.mx/';
const SAT_INFO_PORTAL_DEFAULT = 'https://wwwmat.sat.gob.mx/aplicacion/75169/servicio-de-facturacion-cfdi-version-4.0-%28vigente-a-partir-del-1-de-enero-de-2022%29';
const SAT_SIMULATOR_DEFAULT = 'https://wwwmat.sat.gob.mx/aplicacion/80484/simulador-genera-tu-factura-version-4.0';
const SAT_CANCELLATION_DEFAULT = 'https://www.sat.gob.mx/minisitio/Factura/cancela_procesocancelacion.htm';

const UPLOADS_DIR = path.join(DATA_DIR, 'uploads');
const DB_FILE = path.join(DATA_DIR, 'db.json');
const BACKUPS_DIR = path.join(DATA_DIR, 'backups');
const RECOVERY_DIR = path.join(DATA_DIR, 'recovery');
const JOURNALS_DIR = path.join(DATA_DIR, 'journals');
const SAFE_RELEASE_TAG = '13.6.44';
let startupBackupMade = false;
let DB_SCHEMA_SIGNATURE_13650 = '';

function dbFileSignature13650() {
  try {
    const stat = fs.statSync(DB_FILE);
    return `${stat.size}:${stat.mtimeMs}:${stat.ctimeMs}`;
  } catch (_) {
    return 'missing';
  }
}

if (!RENDER_MIRROR_ONLY) {
  fs.mkdirSync(DATA_DIR, { recursive: true });
  fs.mkdirSync(UPLOADS_DIR, { recursive: true });
  fs.mkdirSync(BACKUPS_DIR, { recursive: true });
  fs.mkdirSync(RECOVERY_DIR, { recursive: true });
  fs.mkdirSync(JOURNALS_DIR, { recursive: true });
}

let RENDER_RELAY_AUTH_DB = null;
let branchRelayClient = null;
let branchRelayClientKey = '';
let branchRelayHeartbeatTimer = null;
let branchRelayClientStatus = { connected: false, status: 'IDLE', lastError: '' };

function createStartupSafetyBackup() {
  if (startupBackupMade || !fs.existsSync(DB_FILE)) return;
  startupBackupMade = true;
  try {
    // Validar primero: si el archivo está dañado NO lo sustituimos por seed().
    JSON.parse(fs.readFileSync(DB_FILE, 'utf8'));
    const backupsDir = path.join(DATA_DIR, 'backups');
    fs.mkdirSync(backupsDir, { recursive: true });
    const stamp = new Date().toISOString().replace(/[:.]/g, '-');
    const target = path.join(backupsDir, `db-before-${SAFE_RELEASE_TAG}-${stamp}.json`);
    fs.copyFileSync(DB_FILE, target);
  } catch (err) {
    // La lectura/migración normal mostrará el error. Nunca reinicializamos silenciosamente.
    console.error('[SAFE BACKUP] No se pudo respaldar/validar db.json:', err.message);
  }
}

function round2(n) {
  return Math.round((Number(n || 0) + Number.EPSILON) * 100) / 100;
}

function roundQuantity(n) {
  return Math.round((Number(n || 0) + Number.EPSILON) * 1000000) / 1000000;
}

function calcRetailPrice(costPrice, retailMarginPct) {
  return round2(Number(costPrice || 0) * (1 + (Number(retailMarginPct || 0) / 100)));
}

function safeNumber(n, fallback = 0) {
  const parsed = Number(n);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function slugify(value = '') {
  return String(value || '')
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/[^a-zA-Z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .toUpperCase();
}


const INVENTORY_IMPORT_HEADERS = [
  'Nombre', 'Categoria', 'SKU', 'CodigoBarras', 'Unidad', 'PermiteFraccion', 'Costo', 'MargenMenudeo',
  'PrecioMenudeo', 'PrecioMayoreo', 'Stock', 'Minimo', 'ClaveSATProducto', 'ClaveSATUnidad',
  'UnidadFiscal', 'DescripcionFiscal', 'ObjetoImpuesto', 'TasaIVA'
];

function execFileAsync(file, args = [], options = {}) {
  return new Promise((resolve, reject) => {
    execFile(file, args, { windowsHide: true, maxBuffer: 8 * 1024 * 1024, ...options }, (error, stdout, stderr) => {
      if (error) {
        error.stdout = stdout;
        error.stderr = stderr;
        reject(error);
        return;
      }
      resolve({ stdout, stderr });
    });
  });
}

function isFacturamaTlsError(err) {
  const text = String(err?.message || err || '').toLowerCase();
  return [
    'unable_to_get_issuer_cert_locally',
    'unable to get local issuer certificate',
    'unable to get issuer certificate',
    'self signed certificate',
    'certificate has expired',
    'unable to verify the first certificate'
  ].some((token) => text.includes(token));
}

function parseCurlResponse(stdout = '') {
  const text = String(stdout || '');
  const marker = '__STATUS__:';
  const idx = text.lastIndexOf(marker);
  if (idx === -1) {
    return { status: 0, headers: {}, raw: text, error: '' };
  }
  const raw = text.slice(0, idx).trim();
  const statusText = text.slice(idx + marker.length).trim();
  return { status: Number(statusText || 0), headers: {}, raw, error: '' };
}

async function factoramaWindowsRequest(urlString, { method = 'GET', headers = {}, body = undefined, timeoutMs = 20000 } = {}) {
  const args = [
    '--silent', '--show-error', '--location', '--ipv4', '--tlsv1.2',
    '--connect-timeout', String(Math.max(5, Math.ceil(timeoutMs / 1000))),
    '--max-time', String(Math.max(10, Math.ceil(timeoutMs / 1000))),
    '--request', String(method || 'GET').toUpperCase(),
    '--header', `Accept: ${String(headers.Accept || 'application/json')}`,
    '--header', `Authorization: ${String(headers.Authorization || '')}`,
    '--header', 'User-Agent: CremeriaElGueroPOS/1.1.2',
    '--write-out', '__STATUS__:%{http_code}'
  ];
  if (body !== undefined) {
    const payload = typeof body === 'string' ? body : JSON.stringify(body);
    args.push('--header', 'Content-Type: application/json', '--data-binary', payload);
  }
  args.push(String(urlString));
  try {
    const result = await execFileAsync('curl.exe', args, { timeout: timeoutMs + 5000 });
    return parseCurlResponse(result.stdout || '');
  } catch (err) {
    const stderr = String(err.stderr || err.stdout || err.message || 'curl failed').trim();
    throw new Error(stderr || 'curl failed');
  }
}

function factoramaBuildInventoryTemplateBuffer() {
  const sample = {
    Nombre: 'Queso Oaxaca',
    Categoria: 'Lácteos',
    SKU: 'QOA-001',
    CodigoBarras: '7501000000012',
    Unidad: 'KG',
    PermiteFraccion: 'SI',
    Costo: 90,
    MargenMenudeo: 20,
    PrecioMenudeo: 108,
    PrecioMayoreo: 110,
    Stock: 25,
    Minimo: 10,
    ClaveSATProducto: '50131802',
    ClaveSATUnidad: 'KGM',
    UnidadFiscal: 'Kilogramo',
    DescripcionFiscal: 'Queso Oaxaca a granel',
    ObjetoImpuesto: '02',
    TasaIVA: 0.16
  };
  const instructions = [
    ['Campo', 'Obligatorio', 'Descripción'],
    ['Nombre', 'Sí', 'Nombre del producto'],
    ['Categoria', 'Sí', 'Categoría visible en catálogo e inventario'],
    ['SKU', 'Recomendado', 'Clave interna del producto'],
    ['CodigoBarras', 'Recomendado', 'Código de barras usado por el escáner'],
    ['Unidad', 'Sí', 'Ejemplo: PZA, KG, LT'],
    ['PermiteFraccion', 'No', 'SI/NO'],
    ['Costo', 'Sí', 'Costo base'],
    ['MargenMenudeo', 'No', 'Porcentaje de margen usado para precio menudeo'],
    ['PrecioMenudeo', 'Sí', 'Precio de venta menudeo'],
    ['PrecioMayoreo', 'Sí', 'Precio mayoreo'],
    ['Stock', 'Sí', 'Existencia inicial o actual'],
    ['Minimo', 'Sí', 'Stock mínimo para alertas'],
    ['ClaveSATProducto', 'No', 'Para facturación CFDI'],
    ['ClaveSATUnidad', 'No', 'Para facturación CFDI'],
    ['UnidadFiscal', 'No', 'Nombre fiscal de la unidad'],
    ['DescripcionFiscal', 'No', 'Descripción CFDI'],
    ['ObjetoImpuesto', 'No', 'Normalmente 02'],
    ['TasaIVA', 'No', 'Ejemplo 0.16']
  ];
  const catalogs = [
    ['Categorias sugeridas', 'Unidades sugeridas'],
    ['Lácteos', 'PZA'],
    ['Embutidos', 'KG'],
    ['Bebidas', 'LT'],
    ['Panadería', 'KG']
  ];
  const wb = XLSX.utils.book_new();
  const wsData = XLSX.utils.json_to_sheet([sample], { header: INVENTORY_IMPORT_HEADERS });
  const wsInst = XLSX.utils.aoa_to_sheet(instructions);
  const wsCatalogs = XLSX.utils.aoa_to_sheet(catalogs);
  XLSX.utils.book_append_sheet(wb, wsData, 'Inventario');
  XLSX.utils.book_append_sheet(wb, wsInst, 'Instrucciones');
  XLSX.utils.book_append_sheet(wb, wsCatalogs, 'Catalogos');
  return XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' });
}

function normalizeInventoryImportRow(row = {}) {
  const get = (...keys) => {
    for (const key of keys) {
      if (row[key] !== undefined && row[key] !== null && row[key] !== '') return row[key];
    }
    return '';
  };
  return {
    name: String(get('Nombre', 'name')).trim(),
    category: String(get('Categoria', 'Categoría', 'category')).trim(),
    sku: String(get('SKU', 'sku')).trim(),
    barcode: String(get('CodigoBarras', 'CódigoBarras', 'Código de barras', 'barcode', 'codigo_barras', 'codigoBarras')).trim(),
    unit: String(get('Unidad', 'unit') || 'PZA').trim().toUpperCase(),
    allowsFraction: /^(1|true|si|sí|yes|y)$/i.test(String(get('PermiteFraccion', 'Permite Fraccion', 'allowsFraction')).trim()),
    costPrice: safeNumber(get('Costo', 'costPrice'), 0),
    retailMarginPct: safeNumber(get('MargenMenudeo', 'Margen Menudeo', 'retailMarginPct'), 0),
    retailPrice: safeNumber(get('PrecioMenudeo', 'Precio Menudeo', 'retailPrice', 'price'), 0),
    wholesalePrice: safeNumber(get('PrecioMayoreo', 'Precio Mayoreo', 'wholesalePrice'), 0),
    stock: safeNumber(get('Stock', 'stock'), 0),
    lowStockThreshold: safeNumber(get('Minimo', 'Mínimo', 'lowStockThreshold'), 0),
    fiscalProductCode: String(get('ClaveSATProducto', 'fiscalProductCode')).trim(),
    fiscalUnitCode: String(get('ClaveSATUnidad', 'fiscalUnitCode')).trim(),
    fiscalUnitName: String(get('UnidadFiscal', 'fiscalUnitName')).trim(),
    fiscalDescription: String(get('DescripcionFiscal', 'DescripciónFiscal', 'fiscalDescription')).trim(),
    fiscalTaxObject: String(get('ObjetoImpuesto', 'fiscalTaxObject') || '02').trim(),
    fiscalVatRate: safeNumber(get('TasaIVA', 'fiscalVatRate'), 0.16)
  };
}

function buildImportedProduct(db, incoming = {}, existing = null) {
  const body = {
    ...(existing || {}),
    ...incoming,
    id: existing?.id,
    createdAt: existing?.createdAt,
    imageUrl: existing?.imageUrl || ''
  };
  const product = normalizeProductPayload(body, existing?.imageUrl || '');
  product.fiscalProductCode = String(incoming.fiscalProductCode || existing?.fiscalProductCode || (String(product.unit).toUpperCase() === 'KG' ? '50131802' : '50202306')).trim();
  product.fiscalUnitCode = String(incoming.fiscalUnitCode || existing?.fiscalUnitCode || (String(product.unit).toUpperCase() === 'KG' ? 'KGM' : 'H87')).trim();
  product.fiscalUnitName = String(incoming.fiscalUnitName || existing?.fiscalUnitName || (String(product.unit).toUpperCase() === 'KG' ? 'Kilogramo' : 'Pieza')).trim();
  product.fiscalDescription = String(incoming.fiscalDescription || existing?.fiscalDescription || `${product.name} de prueba CFDI 4.0`).trim();
  product.fiscalTaxObject = String(incoming.fiscalTaxObject || existing?.fiscalTaxObject || '02').trim();
  product.fiscalVatRate = safeNumber(incoming.fiscalVatRate ?? existing?.fiscalVatRate, 0.16);
  return product;
}

function defaultSellerSession(employee = {}, data = {}) {
  const now = new Date().toISOString();
  return {
    id: data.id || `seller-session-${employee.id || uuidv4()}`,
    operatorId: employee.id || data.operatorId || '',
    operatorName: employee.name || data.operatorName || 'Operador',
    employeeNumber: employee.employeeNumber || data.employeeNumber || '',
    terminalCode: data.terminalCode || '',
    terminalName: data.terminalName || '',
    branchCode: data.branchCode || 'LOCAL',
    customerName: String(data.customerName || '').trim(),
    customerReference: String(data.customerReference || '').trim(),
    ticketId: String(data.ticketId || '').trim(),
    items: normalizeTicketItems(data.items || []),
    subtotal: round2(data.subtotal ?? sumTicket(data.items || [])),
    status: String(data.status || 'IDLE').trim() || 'IDLE',
    createdAt: data.createdAt || now,
    updatedAt: now,
    lastActivityAt: data.lastActivityAt || now,
    takenBy: String(data.takenBy || '').trim(),
    takenByUsername: String(data.takenByUsername || '').trim()
  };
}

function ensureSellerSessions(db) {
  db.sellerSessions = Array.isArray(db.sellerSessions) ? db.sellerSessions : [];
  const employees = getActiveEmployees(db);
  let changed = false;

  db.sellerSessions = db.sellerSessions.map((session) => {
    const employee = employees.find((entry) => entry.id === session.operatorId) || {};
    const normalized = defaultSellerSession(employee, session);
    normalized.updatedAt = session.updatedAt || normalized.updatedAt;
    normalized.createdAt = session.createdAt || normalized.createdAt;
    normalized.lastActivityAt = session.lastActivityAt || normalized.lastActivityAt;
    return normalized;
  });

  for (const employee of employees) {
    const existing = db.sellerSessions.find((session) => session.operatorId === employee.id);
    if (!existing) {
      db.sellerSessions.push(defaultSellerSession(employee));
      changed = true;
    }
  }

  return changed;
}

function getSellerSessionByOperator(db, operatorId = '') {
  ensureSellerSessions(db);
  return (db.sellerSessions || []).find((session) => session.operatorId === operatorId) || null;
}

function upsertSellerSession(db, body = {}, actor = null) {
  ensureSellerSessions(db);
  const operatorId = String(body.operatorId || '').trim();
  if (!operatorId) throw new Error('Selecciona un operador para guardar la sesión.');
  const employee = getActiveEmployees(db).find((entry) => entry.id === operatorId);
  if (!employee) throw new Error('Operador no encontrado.');
  const now = new Date().toISOString();
  const current = getSellerSessionByOperator(db, operatorId) || defaultSellerSession(employee);
  const items = normalizeTicketItems(body.items || current.items || []);
  const updated = {
    ...current,
    operatorId,
    operatorName: employee.name,
    employeeNumber: employee.employeeNumber || current.employeeNumber || '',
    terminalCode: String(body.terminalCode || current.terminalCode || '').trim(),
    terminalName: String(body.terminalName || current.terminalName || '').trim(),
    branchCode: String(body.branchCode || current.branchCode || 'LOCAL').trim() || 'LOCAL',
    customerName: String(body.customerName || '').trim(),
    customerReference: String(body.customerReference || '').trim(),
    ticketId: String(body.ticketId || '').trim(),
    items,
    subtotal: round2(body.subtotal ?? sumTicket(items)),
    status: String(body.status || (items.length ? 'IN_PROGRESS' : 'IDLE')).trim() || 'IDLE',
    takenBy: String(body.terminalCode || current.terminalCode || '').trim(),
    takenByUsername: String(actor?.username || body.takenByUsername || current.takenByUsername || '').trim(),
    updatedAt: now,
    lastActivityAt: now
  };
  const index = (db.sellerSessions || []).findIndex((session) => session.operatorId === operatorId);
  if (index >= 0) db.sellerSessions[index] = updated;
  else db.sellerSessions.push(updated);
  return updated;
}

function clearSellerSession(db, operatorId = '', status = 'CLEARED') {
  ensureSellerSessions(db);
  const session = getSellerSessionByOperator(db, operatorId);
  if (!session) return null;
  Object.assign(session, {
    items: [],
    subtotal: 0,
    customerName: '',
    customerReference: '',
    ticketId: '',
    status,
    updatedAt: new Date().toISOString(),
    lastActivityAt: new Date().toISOString()
  });
  return session;
}

function normalizeFactoramaEnvironment(value = '') {

  return String(value || '').trim().toLowerCase() === 'production' ? 'production' : 'sandbox';
}

function normalizePaymentFormValue(value = '') {
  const raw = String(value || '').trim().toUpperCase();
  if (!raw) return '01';
  if (/^(01|02|03|04|05|06|08|12|13|14|15|17|23|24|25|26|27|28|29|30|31|99)$/.test(raw)) return raw;
  if (raw.includes('EFECTIVO')) return '01';
  if (raw.includes('TRANSFER')) return '03';
  if (raw.includes('TARJETA DE CREDITO') || raw.includes('CRÉDITO')) return '04';
  if (raw.includes('TARJETA DE DEBITO') || raw.includes('DÉBITO')) return '28';
  if (raw.includes('POR DEFINIR')) return '99';
  return '01';
}

function normalizePaymentMethodValue(value = '') {
  const raw = String(value || '').trim().toUpperCase();
  if (!raw) return 'PUE';
  if (raw === 'PUE' || raw === 'PPD') return raw;
  if (raw.includes('PARCIAL') || raw.includes('DIFERID') || raw.includes('CREDITO') || raw.includes('CRÉDITO')) return 'PPD';
  return 'PUE';
}

function normalizeCfdiUseValue(value = '') {
  const raw = String(value || '').trim().toUpperCase();
  if (/^[A-Z][0-9]{2}$/.test(raw)) return raw;
  if (raw === 'XML' || !raw) return 'S01';
  return 'S01';
}

function normalizeFiscalRegimeValue(value = '', fallback = '605') {
  const raw = String(value || '').trim();
  if (!raw) return fallback;
  if (/^\d{3}$/.test(raw)) return raw;
  const normalized = raw.normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase();
  if (normalized.includes('ACTIVIDADES EMPRESARIALES') || normalized.includes('ACTIVIDAD EMPRESARIAL') || normalized.includes('RESICO')) return '612';
  if (normalized.includes('SIN OBLIGACIONES')) return '616';
  if (normalized.includes('SUELDOS')) return '605';
  if (normalized.includes('GENERAL')) return '626';
  if (normalized.includes('PERSONA FISICA')) return fallback;
  return fallback;
}

function defaultFiscalTestCustomer(now = new Date().toISOString()) {
  return {
    id: 'fiscal-publico-general',
    name: 'PUBLICO EN GENERAL',
    rfc: 'XAXX010101000',
    regimenFiscal: '616',
    zip: '55117',
    email: '',
    useCfdi: 'S01',
    createdAt: now,
    updatedAt: now,
    isDefault: true
  };
}

function getFacturamaBaseUrl(config = {}) {
  return normalizeFactoramaEnvironment(config.factoramaEnvironment || config.pacMode) === 'production'
    ? 'https://api.facturama.mx'
    : 'https://apisandbox.facturama.mx';
}

function getFacturamaCredentials(config = {}) {
  return {
    username: String(config.factoramaApiUsername || '').trim(),
    password: String(config.factoramaApiPassword || '').trim()
  };
}

function hasFacturamaCredentials(config = {}) {
  const creds = getFacturamaCredentials(config);
  return !!(creds.username && creds.password);
}

function buildFacturamaAuthHeader(config = {}) {
  const creds = getFacturamaCredentials(config);
  if (!creds.username || !creds.password) {
    throw new Error('Faltan las credenciales de Facturama. Configura usuario y contraseña API.');
  }
  return `Basic ${Buffer.from(`${creds.username}:${creds.password}`).toString('base64')}`;
}


function factoramaHttpsRequest(urlString, { method = 'GET', headers = {}, body = undefined, timeoutMs = 15000 } = {}) {
  return new Promise((resolve, reject) => {
    const url = new URL(urlString);
    const payload = body === undefined ? undefined : (typeof body === 'string' ? body : JSON.stringify(body));
    const req = https.request({
      protocol: url.protocol,
      hostname: url.hostname,
      port: url.port || 443,
      path: `${url.pathname}${url.search}`,
      method,
      headers: {
        ...headers,
        'User-Agent': 'CremeriaElGueroPOS/1.1.2',
        Connection: 'close',
        ...(payload !== undefined ? { 'Content-Length': Buffer.byteLength(payload) } : {})
      },
      timeout: timeoutMs,
      family: 4,
      servername: url.hostname,
      minVersion: 'TLSv1.2',
      rejectUnauthorized: true,
      agent: false
    }, (res) => {
      const chunks = [];
      res.on('data', (chunk) => chunks.push(chunk));
      res.on('end', () => resolve({ status: res.statusCode || 0, headers: res.headers || {}, raw: Buffer.concat(chunks).toString('utf8'), error: '' }));
    });
    req.on('timeout', () => req.destroy(new Error('ETIMEDOUT')));
    req.on('error', reject);
    if (payload !== undefined) req.write(payload);
    req.end();
  });
}

async function factoramaRawRequest(urlString, options = {}) {
  const transportErrors = [];

  if (process.platform === 'win32') {
    try {
      return await factoramaWindowsRequest(urlString, options);
    } catch (err) {
      transportErrors.push(err);
    }
  }

  try {
    return await factoramaHttpsRequest(urlString, options);
  } catch (err) {
    transportErrors.push(err);
    if (process.platform === 'win32' && isFacturamaTlsError(err)) {
      return await factoramaWindowsRequest(urlString, options);
    }
    throw transportErrors[0];
  }
}

async function factoramaFetch(pathOrUrl, { method = 'GET', body = undefined, config = {}, expectJson = true } = {}) {
  const url = /^https?:/i.test(pathOrUrl) ? pathOrUrl : `${getFacturamaBaseUrl(config)}${pathOrUrl}`;
  const headers = {
    Authorization: buildFacturamaAuthHeader(config),
    Accept: 'application/json'
  };
  const options = { method, headers, timeoutMs: 20000 };
  if (body !== undefined) {
    headers['Content-Type'] = 'application/json';
    options.body = typeof body === 'string' ? body : JSON.stringify(body);
  }
  let response;
  try {
    response = await factoramaRawRequest(url, options);
  } catch (networkErr) {
    const err = new Error(`No se pudo conectar con Facturama (${networkErr?.message || 'error de red'}). Revisa internet, firewall, antivirus/proxy y el entorno seleccionado.`);
    err.status = 502;
    err.payload = { cause: String(networkErr?.message || networkErr || 'network failed'), url };
    throw err;
  }
  const raw = String(response.raw || '');
  let data = raw;
  try { data = raw ? JSON.parse(raw) : {}; } catch (_) {}
  const status = Number(response.status || 0);
  if (!(status >= 200 && status < 300)) {
    const message = typeof data === 'object' && data
      ? (data.Message || data.message || data.ModelState?.Message || response.error || JSON.stringify(data))
      : String(raw || response.error || `HTTP ${status || 0}`);
    const err = new Error(`Facturama: ${message}`);
    err.status = status || 502;
    err.payload = data;
    throw err;
  }
  return expectJson ? data : raw;
}

async function factoramaFetchFirst(paths = [], options = {}) {
  let lastError = null;
  for (const path of paths) {
    try {
      return await factoramaFetch(path, options);
    } catch (err) {
      lastError = err;
      if (![404, 405].includes(Number(err.status || 0))) throw err;
    }
  }
  throw lastError || new Error('No se pudo resolver la ruta de Facturama.');
}


function normalizeFacturamaInvoiceRecord(record = {}, sale = null) {
  const total = safeNumber(record.Total ?? sale?.total, 0);
  return {
    id: record.Id || record.id || uuidv4(),
    uuid: record.Uuid || record.uuid || '',
    folio: record.Folio || record.folio || sale?.folio || '',
    serie: record.Serie || record.serie || '',
    status: String(record.Status || record.status || 'issued').toUpperCase(),
    createdAt: record.Date || record.createdAt || new Date().toISOString(),
    receiverRfc: record.Receiver?.Rfc || record.receiverRfc || '',
    receiverName: record.Receiver?.Name || record.receiverName || '',
    total: round2(total),
    cfdiType: record.CfdiType || record.cfdiType || 'I',
    paymentForm: record.PaymentForm || record.paymentForm || '',
    paymentMethod: record.PaymentMethod || record.paymentMethod || '',
    pdfBase64: record.PdfBase64 || '',
    xmlBase64: record.XmlBase64 || '',
    saleId: record.saleId || sale?.id || '',
    customerId: record.customerId || '',
    observations: record.Observations || record.observations || '',
    cancellation: record.cancellation || null,
    apiMode: normalizeFactoramaEnvironment(record.apiMode || ''),
    provider: 'FACTORAMA'
  };
}

function extractBase64Content(payload) {
  if (!payload) return '';
  if (typeof payload === 'string') {
    const trimmed = payload.trim();
    if (!trimmed) return '';
    if ((trimmed.startsWith('{') && trimmed.endsWith('}')) || (trimmed.startsWith('[') && trimmed.endsWith(']'))) {
      try { return extractBase64Content(JSON.parse(trimmed)); } catch (_) { return trimmed; }
    }
    return trimmed.replace(/^"|"$/g, '');
  }
  if (typeof payload === 'object') {
    return payload.Content || payload.content || payload.Data || payload.data || payload.Base64 || payload.base64 || payload.Xml || payload.Pdf || '';
  }
  return '';
}

function buildInvoiceLineItemsSnapshot(sale = {}) {
  return (Array.isArray(sale.items) ? sale.items : []).map((item) => ({
    name: item.name || '',
    quantity: safeNumber(item.quantity),
    unit: item.unit || '',
    unitPrice: round2(item.unitPrice),
    total: round2(safeNumber(item.quantity) * safeNumber(item.unitPrice))
  }));
}

function getDefaultItemFiscalData(item = {}, product = null) {
  const unit = String(item.unit || product?.unit || 'PZA').toUpperCase();
  const unitCode = unit === 'KG' ? 'KGM' : 'H87';
  return {
    ProductCode: product?.fiscalProductCode || item.fiscalProductCode || '10111302',
    UnitCode: product?.fiscalUnitCode || item.fiscalUnitCode || unitCode,
    Unit: product?.fiscalUnitName || item.fiscalUnitName || (unit === 'KG' ? 'Kilogramo' : 'Pieza'),
    Description: product?.fiscalDescription || item.fiscalDescription || item.name || product?.name || 'Producto de prueba CFDI 4.0',
    IdentificationNumber: item.productId || product?.sku || product?.id || 'PRUEBA',
    TaxObject: product?.fiscalTaxObject || item.fiscalTaxObject || '02',
    VatRate: safeNumber(product?.fiscalVatRate ?? item.fiscalVatRate, 0.16)
  };
}

function sanitizeFacturamaIdentifier(value = '', fallback = 'ITEM') {
  const cleaned = String(value || '')
    .trim()
    .replace(/\s+/g, '-')
    .replace(/[^A-Za-z0-9_\-\.]/g, '')
    .slice(0, 100);
  return cleaned || fallback;
}

function buildFacturamaInvoicePayload({ db, config, sale, customer, overrides = {} }) {
  let paymentForm = normalizePaymentFormValue(overrides.paymentForm || config.defaultPaymentForm || '01');
  const paymentMethod = normalizePaymentMethodValue(overrides.paymentMethod || config.defaultPaymentMethod || 'PUE');
  const cfdiUse = normalizeCfdiUseValue(overrides.useCfdi || customer?.useCfdi || config.defaultUseCfdi || 'S01');
  const series = String(config.defaultSeries || config.series || 'CREG').trim() || 'CREG';
  if (paymentMethod === 'PPD' && paymentForm !== '99') paymentForm = '99';
  const receiverName = String(customer.name || '').trim().replace(/\s+/g, ' ').toUpperCase();

  const items = (Array.isArray(sale.items) ? sale.items : []).map((item, idx) => {
    const product = (db.products || []).find((p) => p.id === item.productId) || null;
    const fiscal = getDefaultItemFiscalData(item, product);
    const quantity = Math.max(0.001, safeNumber(item.quantity, 1));
    const unitPrice = round2(safeNumber(item.unitPrice, 0));
    const grossSubtotal = round2(quantity * unitPrice);
    const discountAmount = Math.max(0, Math.min(grossSubtotal, round2(safeNumber(item.discountAmount ?? item.discount ?? 0, 0))));
    const subtotal = round2(grossSubtotal - discountAmount);
    const vatRate = fiscal.TaxObject === '02' ? safeNumber(fiscal.VatRate, 0.16) : 0;
    const taxTotal = fiscal.TaxObject === '02' ? round2(subtotal * vatRate) : 0;
    const total = round2(subtotal + taxTotal);
    const identification = sanitizeFacturamaIdentifier(
      item.barcode || product?.barcode || product?.sku || item.productId || product?.id || `ITEM${idx + 1}`,
      `ITEM${idx + 1}`
    );

    const node = {
      Quantity: String(quantity),
      ProductCode: fiscal.ProductCode,
      UnitCode: fiscal.UnitCode,
      Unit: fiscal.Unit,
      Description: fiscal.Description,
      IdentificationNumber: identification,
      UnitPrice: unitPrice.toFixed(2),
      Subtotal: subtotal.toFixed(2),
      TaxObject: fiscal.TaxObject,
      Total: total.toFixed(2)
    };

    if (discountAmount > 0) node.Discount = discountAmount.toFixed(2);

    if (fiscal.TaxObject === '02') {
      node.Taxes = [{
        Name: 'IVA',
        Rate: vatRate.toFixed(2),
        Total: taxTotal.toFixed(2),
        Base: subtotal.toFixed(2),
        IsRetention: false,
        IsFederalTax: true
      }];
    }
    return node;
  });

  return {
    CfdiType: 'I',
    NameId: '1',
    ExpeditionPlace: String(config.expeditionZip || '55750').trim(),
    Serie: series || null,
    Folio: sanitizeFacturamaIdentifier(String(sale.folio || '').replace(/[^A-Za-z0-9_-]/g, '').slice(-40) || `VENTA-${Date.now()}`, `VENTA-${Date.now()}`),
    PaymentForm: paymentForm,
    PaymentMethod: paymentMethod,
    Currency: 'MXN',
    Decimals: 2,
    Exportation: '01',
    Receiver: {
      Name: receiverName,
      CfdiUse: cfdiUse,
      Rfc: String(customer.rfc || '').trim().toUpperCase(),
      FiscalRegime: normalizeFiscalRegimeValue(customer.regimenFiscal || '', '616'),
      TaxZipCode: String(customer.zip || config.expeditionZip || '55750').trim()
    },
    Items: items,
    Observations: `Venta ${sale.folio || ''} · ${config.commercialName || config.emitterName || ADMIN_NAME}`.trim()
  };
}


function defaultBranchProfile(now = new Date().toISOString()) {
  return {
    name: ADMIN_NAME,
    code: 'ADMIN-01',
    publicUrl: String(process.env.RENDER_EXTERNAL_URL || '').trim(),
    mode: 'LOCAL',
    description: '',
    updatedAt: now
  };
}

function defaultPublicationConfig(now = new Date().toISOString()) {
  return {
    enabled: false,
    mirrorUrl: '',
    authType: 'token',
    apiKey: '',
    remoteAdminUsername: 'admin',
    remoteAdminPassword: 'admin123',
    autoSyncEnabled: false,
    syncIntervalSec: 60,
    lastSyncAt: '',
    lastSyncStatus: 'IDLE',
    lastSyncError: '',
    lastCommandPollAt: '',
    lastCommandApplyAt: '',
    updatedAt: now
  };
}

function defaultMirrorReplica(now = new Date().toISOString()) {
  return {
    branchCode: '',
    sourceInstanceId: '',
    snapshot: null,
    lastSnapshotAt: '',
    lastSnapshotSource: '',
    commands: [],
    commandSequence: 0,
    updatedAt: now
  };
}

function seed() {
  const now = new Date().toISOString();
  return {
    users: [
      { id: 'u-admin', username: 'admin', password: 'admin123', role: 'ADMIN', name: 'Administrador General', active: true, employeeNumber: 'ADMIN-01' },
      { id: 'u-caja1', username: 'caja1', password: '1234', role: 'POS', name: 'Caja 1', active: true, employeeNumber: '1001' },
      { id: 'u-caja2', username: 'caja2', password: '1234', role: 'POS', name: 'Caja 2', active: true, employeeNumber: '1002' },
      { id: 'u-caja3', username: 'caja3', password: '1234', role: 'POS', name: 'Caja 3', active: true, employeeNumber: '1003' }
    ],
    security: {
      requireAdminPasswordForEdits: false,
      updatedAt: now,
      remoteAccess: {
        apiKey: RENDER_REMOTE_API_KEY || `REMOTE-${uuidv4().replace(/-/g, '').slice(0, 24).toUpperCase()}`,
        allowAdminCredentials: !RENDER_MIRROR_ONLY,
        permissions: {
          readSummary: true,
          readInventory: true,
          writeInventory: true,
          syncBridge: true
        }
      },
      notifications: {
        whatsappRestockNumber: ''
      }
    },
    fiscalConfig: {
      commercialName: 'Studio Code',
      emitterRfc: 'AEVE040724ST8',
      emitterName: 'STUDIO CODE INC',
      emitterRegimen: '612',
      expeditionZip: '55117',
      pacProvider: 'FACTORAMA',
      preferredInvoiceProvider: 'SAT_ASSISTED',
      satPortalUrl: SAT_PORTAL_DEFAULT,
      satInfoUrl: SAT_INFO_PORTAL_DEFAULT,
      satSimulatorUrl: SAT_SIMULATOR_DEFAULT,
      satCancellationUrl: SAT_CANCELLATION_DEFAULT,
      satWorkflowMode: 'SIMULATOR',
      pacMode: 'SANDBOX',
      factoramaEnvironment: 'sandbox',
      factoramaApiUsername: 'edgar117',
      factoramaApiPassword: 'Ae24072310@.',
      defaultSeries: 'CREG',
      defaultPaymentForm: '01',
      defaultPaymentMethod: 'PUE',
      defaultUseCfdi: 'S01',
      defaultCancelMotive: '02'
    },
    products: [
      {
        id: 'p-queso',
        name: 'Queso Oaxaca',
        category: 'Lácteos',
        sku: 'QOA-001',
        unit: 'KG',
        allowsFraction: true,
        costPrice: 96,
        retailMarginPct: 25,
        retailPrice: 120,
        wholesalePrice: 110,
        price: 120,
        stock: 200,
        lowStockThreshold: 15,
        imageUrl: '',
        fiscalProductCode: '50131802',
        fiscalUnitCode: 'KGM',
        fiscalUnitName: 'Kilogramo',
        fiscalDescription: 'Queso Oaxaca de prueba CFDI 4.0',
        fiscalTaxObject: '02',
        fiscalVatRate: 0.16,
        active: true,
        createdAt: now,
        updatedAt: now
      },
      {
        id: 'p-coca',
        name: 'Coca Cola 600ml',
        category: 'Bebidas',
        sku: 'COC-600',
        unit: 'PZA',
        allowsFraction: false,
        costPrice: 12,
        retailMarginPct: 50,
        retailPrice: 18,
        wholesalePrice: 15,
        price: 18,
        stock: 150,
        lowStockThreshold: 24,
        imageUrl: '',
        fiscalProductCode: '50202306',
        fiscalUnitCode: 'H87',
        fiscalUnitName: 'Pieza',
        fiscalDescription: 'Bebida de prueba CFDI 4.0',
        fiscalTaxObject: '02',
        fiscalVatRate: 0.16,
        active: true,
        createdAt: now,
        updatedAt: now
      }
    ],
    openTickets: [],
    sales: [],
    saleConflicts: [],
    fiscalCustomers: [defaultFiscalTestCustomer(now)],
    invoices: [],
    branches: [],
    branchProfile: defaultBranchProfile(now),
    publication: defaultPublicationConfig(now),
    mirrorReplica: defaultMirrorReplica(now),
    mirrorReplicas: {},
    syncInstanceId: uuidv4(),
    inventoryAdjustments: [],
    counters: {
      salesSequence: 0
    },
    terminals: [
      {
        id: 't-admin',
        terminalCode: 'ADMIN-01',
        terminalType: 'ADMIN',
        terminalName: ADMIN_NAME,
        lastSeenAt: now
      }
    ],
    promotions: [],
    cashOpenings: [],
    cashEntries: [],
    cashClosings: []
  };
}

function validateSalePrices({ name, costPrice, retailPrice, wholesalePrice }) {
  const safeCost = safeNumber(costPrice);
  const safeRetail = safeNumber(retailPrice);
  const safeWholesale = safeNumber(wholesalePrice);

  if (safeCost < 0) {
    throw new Error(`El costo no puede ser negativo en ${name || 'el producto'}`);
  }
  if (safeRetail < safeCost) {
    throw new Error(`El precio menudeo no puede ser menor al costo en ${name || 'el producto'}`);
  }
  if (safeWholesale < safeCost) {
    throw new Error(`El precio mayoreo no puede ser menor al costo en ${name || 'el producto'}`);
  }
}

function normalizeProductPayload(body, imageUrl = '') {
  const costPrice = round2(body.costPrice);
  const retailMarginPct = round2(body.retailMarginPct);
  const retailPrice = round2(
    body.retailPrice !== undefined && body.retailPrice !== ''
      ? body.retailPrice
      : calcRetailPrice(costPrice, retailMarginPct)
  );
  const wholesalePrice = round2(
    body.wholesalePrice !== undefined && body.wholesalePrice !== ''
      ? body.wholesalePrice
      : retailPrice
  );
  const lowStockThreshold = safeNumber(body.lowStockThreshold);
  const now = new Date().toISOString();

  const product = {
    id: body.id || uuidv4(),
    name: String(body.name || '').trim(),
    category: String(body.category || '').trim(),
    sku: String(body.sku || '').trim(),
    barcode: String(body.barcode || body.codigoBarras || body.codigo_barras || '').trim(),
    unit: body.unit || 'PZA',
    allowsFraction:
      String(body.allowsFraction) === 'true' ||
      body.allowsFraction === true ||
      body.allowsFraction === 'on',
    costPrice,
    retailMarginPct,
    retailPrice,
    wholesalePrice,
    price: retailPrice,
    stock: safeNumber(body.stock),
    lowStockThreshold,
    imageUrl: imageUrl || body.imageUrl || '',
    active: body.active !== false,
    createdAt: body.createdAt || now,
    updatedAt: now
  };

  if (!product.name) {
    throw new Error('El nombre del producto es obligatorio');
  }
  if (!product.category) {
    throw new Error('La categoría del producto es obligatoria');
  }

  validateSalePrices(product);

  return product;
}


function normalizeEmployeePayload(body = {}) {
  const name = String(body.name || '').trim();
  const employeeNumber = String(body.employeeNumber || '').trim();
  if (!name) throw new Error('El nombre completo del empleado es obligatorio');
  if (!employeeNumber) throw new Error('El número de empleado es obligatorio');

  const usernameBase = String(body.username || `emp-${employeeNumber}`).trim().toLowerCase();
  return {
    id: body.id || uuidv4(),
    role: 'POS',
    name,
    employeeNumber,
    username: usernameBase,
    password: String(body.password || employeeNumber || '1234').trim(),
    active: body.active !== false,
    createdAt: body.createdAt || new Date().toISOString(),
    updatedAt: new Date().toISOString()
  };
}

function getActiveEmployees(db) {
  return (db.users || [])
    .filter((u) => u.role === 'POS' && u.active !== false)
    .sort((a, b) => String(a.name || '').localeCompare(String(b.name || ''), 'es'));
}

function parseDateInput(value, fallback = new Date()) {
  if (!value) return fallback;
  const parsed = new Date(value);
  return Number.isNaN(parsed.getTime()) ? fallback : parsed;
}

function getRangeBounds(range = 'day', anchorDate) {
  const base = parseDateInput(anchorDate, new Date());
  const start = new Date(base);
  const end = new Date(base);

  if (range === 'month') {
    start.setDate(1);
    start.setHours(0, 0, 0, 0);
    end.setMonth(start.getMonth() + 1, 1);
    end.setHours(0, 0, 0, 0);
  } else if (range === 'week') {
    const day = start.getDay();
    const diff = day === 0 ? -6 : 1 - day;
    start.setDate(start.getDate() + diff);
    start.setHours(0, 0, 0, 0);
    end.setTime(start.getTime());
    end.setDate(start.getDate() + 7);
  } else {
    start.setHours(0, 0, 0, 0);
    end.setTime(start.getTime());
    end.setDate(start.getDate() + 1);
  }

  return { start, end };
}

function getBranchSummarySkeleton(branch) {
  return {
    id: branch.id,
    branchId: branch.id,
    name: branch.name,
    code: branch.code || '',
    color: /^#[0-9a-f]{6}$/i.test(String(branch.color || '')) ? branch.color : '#1f6feb',
    remoteUrl: branch.remoteUrl,
    authType: branch.authType || 'credentials',
    apiKey: '',
    hasApiKey: branch.authType === 'token' && !!String(branch.apiKey || '').trim(),
    remoteAdminUsername: branch.remoteAdminUsername || '',
    mode: branch.mode || 'render',
    status: branch.lastStatus || 'PENDING',
    lastCheckedAt: branch.lastCheckedAt || '',
    lastError: branch.lastError || '',
    lastRemoteSummary: branch.lastRemoteSummary || null
  };
}

function getBranchProfile(db) {
  return {
    ...defaultBranchProfile(),
    ...(db.branchProfile || {})
  };
}

function getPublicationConfig(db) {
  return {
    ...defaultPublicationConfig(),
    ...(db.publication || {})
  };
}

function persistBranchRelayStatus(status, message = '') {
  branchRelayClientStatus = {
    connected: status === 'ONLINE',
    status,
    lastError: message
  };
  if (RENDER_MIRROR_ONLY) return;
  try {
    const db = readDb();
    db.publication ||= defaultPublicationConfig();
    db.publication.lastSyncAt = new Date().toISOString();
    db.publication.lastSyncStatus = status;
    db.publication.lastSyncError = message;
    db.publication.updatedAt = new Date().toISOString();
    writeDb(db);
  } catch (err) {
    console.warn('[BRANCH RELAY] No se pudo guardar el estado local:', err.message);
  }
}

function buildBranchRelayReadPayload(db, request = {}) {
  const branchCode = normalizeMirrorBranchCode(request.branchCode || '');
  const profile = getBranchProfile(db);
  if (!branchCode || branchCode !== normalizeMirrorBranchCode(profile.code || '')) {
    const err = new Error('El código solicitado no coincide con el código de esta sucursal.');
    err.status = 409;
    throw err;
  }

  const endpoint = String(request.endpoint || '');
  const query = request.query && typeof request.query === 'object' ? request.query : {};
  const permissions = getRemotePermissions(db);
  let payload;
  switch (endpoint) {
    case '/api/remote/publication-info':
      payload = {
        ok: true,
        branchProfile: profile,
        remote: {
          adminName: ADMIN_NAME,
          hostname: os.hostname(),
          port: PORT,
          service: 'embedded-admin-backend',
          serviceMode: 'LIVE_RELAY',
          requestedBranchCode: branchCode,
          permissions
        },
        lastSnapshotAt: '',
        authMode: 'token'
      };
      break;
    case '/api/remote/summary':
      payload = buildRemoteSummary(db, { profile, serviceMode: 'LIVE_RELAY' });
      break;
    case '/api/remote/cash-overview':
      payload = { ok: true, branchProfile: profile, overview: buildCashOverview(db, normalizeDateKey(query.date)) };
      break;
    case '/api/remote/users':
      payload = {
        ok: true,
        users: (db.users || []).map((user) => ({
          id: user.id,
          name: user.name || user.username || 'Usuario',
          employeeNumber: user.employeeNumber || '',
          role: user.role || 'POS',
          active: user.active !== false
        }))
      };
      break;
    case '/api/remote/inventory':
      payload = buildRemoteInventory(db, query);
      payload.permissions = permissions;
      break;
    case '/api/remote/sales':
      payload = buildRemoteSales(db, query);
      break;
    case '/api/remote/employee-report':
      payload = buildRemoteEmployeeReport(db, query);
      break;
    case '/api/remote/stock-alerts':
    case '/api/remote/alerts':
      payload = buildRemoteStockAlerts(db);
      break;
    default: {
      const err = new Error('Esta consulta no está disponible a través del puente.');
      err.status = 404;
      throw err;
    }
  }
  payload.authMode ||= 'token';
  return payload;
}

function handleBranchRelayQuery(request, acknowledge) {
  if (typeof acknowledge !== 'function') return;
  try {
    const db = readDb();
    const payload = buildBranchRelayReadPayload(db, request || {});
    acknowledge({ ok: true, payload });
  } catch (err) {
    acknowledge({ ok: false, status: err.status || 500, message: err.message || 'No se pudo atender la consulta.' });
  }
}

function configureBranchRelayClient(db = null) {
  if (RENDER_MIRROR_ONLY) return { ok: false, message: 'El conector de sucursal no se inicia en el servicio Render.' };
  const sourceDb = db || readDb();
  const publication = getPublicationConfig(sourceDb);
  const profile = getBranchProfile(sourceDb);
  const branchCode = normalizeMirrorBranchCode(profile.code || '');
  const base = String(publication.mirrorUrl || '').trim().replace(/\/$/, '');
  const apiKey = String(publication.apiKey || '').trim();
  const shouldConnect = publication.enabled === true && publication.autoSyncEnabled === true;

  if (!shouldConnect || !base || !branchCode || !apiKey) {
    if (branchRelayClient) {
      branchRelayClient.removeAllListeners();
      branchRelayClient.disconnect();
      branchRelayClient = null;
    }
    clearInterval(branchRelayHeartbeatTimer);
    branchRelayHeartbeatTimer = null;
    branchRelayClientKey = '';
    if (!shouldConnect) persistBranchRelayStatus('IDLE', 'Conexión del puente deshabilitada.');
    else if (!apiKey) persistBranchRelayStatus('ERROR', 'Falta la API key del servicio Render.');
    return { ok: false, message: !shouldConnect ? 'La conexión al puente está deshabilitada.' : 'Completa la URL, código de sucursal y API key.' };
  }

  let origin;
  try {
    const parsed = new URL(base);
    if (!['http:', 'https:'].includes(parsed.protocol)) throw new Error('protocolo');
    origin = parsed.origin;
  } catch (_) {
    persistBranchRelayStatus('ERROR', 'La URL del puente no es válida.');
    return { ok: false, message: 'La URL del puente debe ser HTTP o HTTPS.' };
  }

  const connectionKey = `${origin}|${branchCode}|${apiKey}`;
  if (branchRelayClient && branchRelayClientKey === connectionKey) {
    return { ok: true, connected: branchRelayClient.connected === true, client: branchRelayClient };
  }

  if (branchRelayClient) {
    branchRelayClient.removeAllListeners();
    branchRelayClient.disconnect();
  }
  clearInterval(branchRelayHeartbeatTimer);
  branchRelayHeartbeatTimer = null;
  branchRelayClientKey = connectionKey;
  persistBranchRelayStatus('CONNECTING', 'Conectando con el puente Render…');
  branchRelayClient = createSocketClient(origin, {
    auth: { apiKey, branchCode },
    transports: ['websocket', 'polling'],
    reconnection: true,
    reconnectionAttempts: Infinity,
    reconnectionDelay: 1000,
    reconnectionDelayMax: 10000,
    timeout: 20000,
    maxHttpBufferSize: 16 * 1024 * 1024
  });
  branchRelayClient.on('connect', () => {
    persistBranchRelayStatus('ONLINE', '');
    clearInterval(branchRelayHeartbeatTimer);
    branchRelayHeartbeatTimer = setInterval(() => {
      if (branchRelayClient?.connected) branchRelayClient.emit('branch:heartbeat', { branchCode, at: Date.now() });
    }, 60_000);
    branchRelayHeartbeatTimer.unref?.();
  });
  branchRelayClient.on('disconnect', (reason) => {
    clearInterval(branchRelayHeartbeatTimer);
    branchRelayHeartbeatTimer = null;
    persistBranchRelayStatus('OFFLINE', `Conexión interrumpida: ${reason}`);
  });
  branchRelayClient.on('connect_error', (err) => {
    branchRelayClientStatus = { connected: false, status: 'CONNECTING', lastError: err.message || 'Sin conexión' };
  });
  branchRelayClient.on('branch:query', handleBranchRelayQuery);
  return { ok: true, connected: false, client: branchRelayClient };
}

async function waitForBranchRelayConnection(db = null, timeoutMs = 120000) {
  const configured = configureBranchRelayClient(db);
  if (!configured.ok) throw new Error(configured.message);
  if (configured.connected || configured.client?.connected) return { ok: true, connected: true };
  const client = configured.client;
  return new Promise((resolve, reject) => {
    let settled = false;
    const finish = (err) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      client.off('connect', onConnect);
      client.off('connect_error', onConnectError);
      if (err) reject(err);
      else resolve({ ok: true, connected: true });
    };
    const onConnect = () => finish();
    const onConnectError = (err) => {
      if (/no autorizada|unauthorized|forbidden/i.test(String(err?.message || ''))) {
        finish(new Error('Render rechazó la conexión. Revisa que la API key sea exactamente la misma en Render y en esta sucursal.'));
      }
    };
    const timer = setTimeout(() => finish(new Error(branchRelayClientStatus.lastError || 'Render no respondió a tiempo. Verifica la URL, la API key y que la publicación esté habilitada.')), timeoutMs);
    client.once('connect', onConnect);
    client.on('connect_error', onConnectError);
  });
}

function normalizeMirrorBranchCode(value = '') {
  return String(value || '')
    .trim()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toUpperCase()
    .replace(/[^A-Z0-9_-]+/g, '-')
    .replace(/[-_]{2,}/g, '-')
    .replace(/^[-_]+|[-_]+$/g, '')
    .slice(0, 48);
}

function getMirrorReplica(db, branchCode = '') {
  const code = normalizeMirrorBranchCode(branchCode);
  if (code) {
    const replica = db.mirrorReplicas?.[code];
    if (replica) return { ...defaultMirrorReplica(), ...replica };
    const legacyCode = normalizeMirrorBranchCode(db.mirrorReplica?.snapshot?.branchProfile?.code || '');
    if (legacyCode === code && db.mirrorReplica) return { ...defaultMirrorReplica(), ...db.mirrorReplica };
    return defaultMirrorReplica();
  }

  const legacy = { ...defaultMirrorReplica(), ...(db.mirrorReplica || {}) };
  if (legacy.snapshot) return legacy;
  const replicas = Object.values(db.mirrorReplicas || {})
    .filter((replica) => replica?.snapshot)
    .sort((a, b) => String(b.lastSnapshotAt || '').localeCompare(String(a.lastSnapshotAt || '')));
  return replicas.length ? { ...defaultMirrorReplica(), ...replicas[0] } : legacy;
}

function getMirrorReplicaCodeList(db) {
  return Object.entries(db.mirrorReplicas || {})
    .filter(([, replica]) => replica?.snapshot)
    .map(([code]) => normalizeMirrorBranchCode(code))
    .filter(Boolean);
}

function resolveRemoteMirrorContext(req, res, db, { allowEmpty = false } = {}) {
  const requestedCode = String(req.headers['x-remote-branch-code'] || req.query?.branchCode || '').trim();
  let branchCode = normalizeMirrorBranchCode(requestedCode);
  if (requestedCode && !branchCode) {
    res.status(400).json({ ok: false, message: 'El código de sucursal enviado no es válido.' });
    return null;
  }

  if (RENDER_MIRROR_ONLY && !branchCode) {
    const codes = getMirrorReplicaCodeList(db);
    const legacyCode = normalizeMirrorBranchCode(db.mirrorReplica?.snapshot?.branchProfile?.code || '');
    if (codes.length === 1) branchCode = codes[0];
    else if (!codes.length && legacyCode) branchCode = legacyCode;
    else if (codes.length > 1) {
      res.status(400).json({ ok: false, message: 'Indica el código de la sucursal que deseas consultar.' });
      return null;
    }
  }

  const replica = getMirrorReplica(db, branchCode);
  if (RENDER_MIRROR_ONLY && !allowEmpty && !replica.snapshot) {
    res.status(404).json({ ok: false, message: `La sucursal ${branchCode || ''} aún no ha publicado su primera sincronización.`.trim() });
    return null;
  }
  return { branchCode, replica };
}

function getWritableMirrorReplica(db, branchCode = '') {
  const code = normalizeMirrorBranchCode(branchCode);
  if (code) {
    db.mirrorReplicas ||= {};
    const existing = db.mirrorReplicas[code]
      || (normalizeMirrorBranchCode(db.mirrorReplica?.snapshot?.branchProfile?.code || '') === code ? db.mirrorReplica : null);
    return { branchCode: code, replica: { ...defaultMirrorReplica(), ...(existing || {}) } };
  }
  db.mirrorReplica ||= defaultMirrorReplica();
  return { branchCode: '', replica: { ...defaultMirrorReplica(), ...db.mirrorReplica } };
}

function storeMirrorReplica(db, branchCode, replica) {
  const code = normalizeMirrorBranchCode(branchCode);
  if (code) {
    db.mirrorReplicas ||= {};
    db.mirrorReplicas[code] = replica;
  } else {
    db.mirrorReplica = replica;
  }
}

function getNearStockThreshold(product) {
  const low = Math.max(0, safeNumber(product?.lowStockThreshold, 0));
  if (low <= 0) return 0;
  return round2(Math.max(low + (product?.allowsFraction ? 0.5 : 1), low * 1.25));
}

function getInventoryAlertStatus(product) {
  const stock = safeNumber(product?.stock, 0);
  const low = Math.max(0, safeNumber(product?.lowStockThreshold, 0));
  if (low > 0 && stock <= low) return 'missing';
  const nearThreshold = getNearStockThreshold(product);
  if (low > 0 && stock > low && stock <= nearThreshold) return 'near';
  return 'normal';
}

function enrichInventoryRow(product = {}) {
  const alertStatus = getInventoryAlertStatus(product);
  return {
    ...product,
    alertStatus,
    nearStockThreshold: getNearStockThreshold(product)
  };
}

function filterInventoryRows(rows = [], filters = {}) {
  const query = String(filters.query || '').trim().toLowerCase();
  const category = String(filters.category || '').trim().toLowerCase();
  const alertMode = String(filters.alertMode || (String(filters.lowOnly || '').trim() === '1' ? 'attention' : 'all')).trim().toLowerCase();
  let result = (Array.isArray(rows) ? rows : []).map(enrichInventoryRow);

  if (query) {
    result = result.filter((product) => [product.name, product.sku, product.category].some((value) => String(value || '').toLowerCase().includes(query)));
  }
  if (category) {
    result = result.filter((product) => String(product.category || '').trim().toLowerCase() === category);
  }
  if (alertMode === 'missing') {
    result = result.filter((product) => product.alertStatus === 'missing');
  } else if (alertMode === 'near') {
    result = result.filter((product) => product.alertStatus === 'near');
  } else if (alertMode === 'attention') {
    result = result.filter((product) => product.alertStatus === 'missing' || product.alertStatus === 'near');
  }

  result.sort((a, b) => String(a.name || '').localeCompare(String(b.name || ''), 'es'));
  return result;
}

function buildInventoryPayloadFromRows(rows = [], filters = {}) {
  const enrichedRows = (Array.isArray(rows) ? rows : []).map(enrichInventoryRow);
  const allCategories = Array.from(new Set(enrichedRows.map((product) => String(product.category || '').trim()).filter(Boolean))).sort((a, b) => a.localeCompare(b, 'es'));
  const filtered = filterInventoryRows(enrichedRows, filters);
  const missingCount = filtered.filter((product) => product.alertStatus === 'missing').length;
  const nearCount = filtered.filter((product) => product.alertStatus === 'near').length;
  return {
    ok: true,
    products: filtered,
    categories: allCategories,
    totals: {
      products: filtered.length,
      stockUnits: round2(filtered.reduce((acc, product) => acc + safeNumber(product.stock), 0)),
      lowStockCount: missingCount,
      missingCount,
      nearStockCount: nearCount,
      attentionCount: missingCount + nearCount
    },
    generatedAt: new Date().toISOString()
  };
}

function getSalesRangeBounds(filters = {}) {
  const range = String(filters.range || filters.period || filters.preset || '').trim().toLowerCase();
  const fromRaw = String(filters.from || '').trim();
  const toRaw = String(filters.to || '').trim();

  if (fromRaw || toRaw) {
    const start = fromRaw ? new Date(`${fromRaw}T00:00:00`) : null;
    const end = toRaw ? new Date(`${toRaw}T23:59:59.999`) : null;
    return {
      range: range || 'custom',
      from: fromRaw,
      to: toRaw,
      start,
      end
    };
  }

  if (['today', 'day', 'week', 'month'].includes(range)) {
    const now = new Date();
    const start = new Date(now);
    const end = new Date(now);
    start.setHours(0, 0, 0, 0);
    end.setHours(23, 59, 59, 999);

    if (range === 'week') {
      const day = start.getDay();
      const diff = day === 0 ? -6 : 1 - day;
      start.setDate(start.getDate() + diff);
      start.setHours(0, 0, 0, 0);
    } else if (range === 'month') {
      start.setDate(1);
      start.setHours(0, 0, 0, 0);
    }

    return {
      range: range === 'day' ? 'today' : range,
      from: start.toISOString().slice(0, 10),
      to: end.toISOString().slice(0, 10),
      start,
      end
    };
  }

  return { range: 'all', from: '', to: '', start: null, end: null };
}

function buildRemoteSalesPayloadFromRows(sourceRows = [], filters = {}, generatedAt = new Date().toISOString()) {
  const limit = Math.max(1, Math.min(5000, safeNumber(filters.limit, 5000)));
  const q = String(filters.query || '').trim().toLowerCase();
  const employee = String(filters.employee || '').trim().toLowerCase();
  const customer = String(filters.customer || '').trim().toLowerCase();
  const rangeInfo = getSalesRangeBounds(filters);

  let rows = (sourceRows || []).slice().sort((a, b) => new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime());

  if (rangeInfo.start || rangeInfo.end) {
    rows = rows.filter((sale) => {
      const saleDate = new Date(sale.createdAt);
      if (Number.isNaN(saleDate.getTime())) return false;
      if (rangeInfo.start && saleDate < rangeInfo.start) return false;
      if (rangeInfo.end && saleDate > rangeInfo.end) return false;
      return true;
    });
  }

  if (q) {
    rows = rows.filter((sale) => {
      const haystack = [
        sale.folio,
        sale.customerName,
        sale.customerReference,
        sale.operatorName,
        sale.employeeNumber,
        sale.terminalCode,
        ...(Array.isArray(sale.items) ? sale.items.map((item) => item.name) : [])
      ].map((value) => String(value || '').toLowerCase()).join(' ');
      return haystack.includes(q);
    });
  }

  if (employee) {
    rows = rows.filter((sale) => [sale.operatorName, sale.employeeNumber]
      .some((value) => String(value || '').toLowerCase().includes(employee)));
  }

  if (customer) {
    rows = rows.filter((sale) => [sale.customerName, sale.customerReference]
      .some((value) => String(value || '').toLowerCase().includes(customer)));
  }

  const selected = rows.slice(0, limit);
  return {
    ok: true,
    filters: {
      limit,
      query: q,
      employee,
      customer,
      range: rangeInfo.range,
      preset: rangeInfo.range,
      from: rangeInfo.from,
      to: rangeInfo.to
    },
    totals: {
      count: selected.length,
      amount: round2(selected.reduce((acc, sale) => acc + saleCashTotal(sale), 0)),
      internalOwnerAmount: round2(selected.filter((sale) => isOwnerConsumptionSale(sale)).reduce((acc, sale) => acc + safeNumber(sale.total), 0))
    },
    sales: selected.map((sale) => ({
      id: sale.id,
      folio: sale.folio,
      createdAt: sale.createdAt,
      saleType: sale.saleType || sale.transactionType || '',
      affectsCash: sale.affectsCash !== false,
      paymentMethod: sale.paymentMethod || '',
      total: sale.total,
      operatorName: sale.operatorName || '',
      employeeNumber: sale.employeeNumber || '',
      customerName: sale.customerName || '',
      customerReference: sale.customerReference || '',
      terminalCode: sale.terminalCode || '',
      itemsCount: Array.isArray(sale.items) ? sale.items.length : 0,
      items: Array.isArray(sale.items) ? sale.items.map((item) => ({
        name: item.name,
        quantity: item.quantity,
        unit: item.unit || '',
        unitPrice: item.unitPrice,
        total: round2(safeNumber(item.quantity) * safeNumber(item.unitPrice))
      })) : []
    })),
    generatedAt
  };
}

function buildRemoteSales(db, filters = {}) {
  return buildRemoteSalesPayloadFromRows(db.sales || [], filters, new Date().toISOString());
}

function buildRemoteEmployeeReport(db, filters = {}) {
  const { range = 'day', anchorDate = '', from = '', to = '' } = filters || {};
  let start;
  let end;
  if (from && to) {
    start = parseDateInput(from, new Date());
    end = parseDateInput(to, new Date());
    end.setDate(end.getDate() + 1);
  } else {
    ({ start, end } = getRangeBounds(range, anchorDate));
  }

  const salesInRange = (db.sales || []).filter((sale) => {
    if (isOwnerConsumptionSale(sale)) return false;
    const d = new Date(sale.createdAt);
    return d >= start && d < end;
  });

  const map = new Map();
  for (const sale of salesInRange) {
    const key = sale.operatorId || `${sale.employeeNumber || ''}-${sale.operatorName || 'SIN_OPERADOR'}`;
    if (!map.has(key)) {
      map.set(key, {
        operatorId: sale.operatorId || '',
        name: sale.operatorName || 'Sin asignar',
        employeeNumber: sale.employeeNumber || '',
        salesCount: 0,
        total: 0
      });
    }
    const row = map.get(key);
    row.salesCount += 1;
    row.total = round2(row.total + safeNumber(sale.total));
  }

  const rows = Array.from(map.values()).sort((a, b) => b.total - a.total);
  return {
    ok: true,
    range,
    from: start.toISOString(),
    to: end.toISOString(),
    totals: {
      salesCount: rows.reduce((acc, row) => acc + safeNumber(row.salesCount), 0),
      total: round2(rows.reduce((acc, row) => acc + safeNumber(row.total), 0))
    },
    rows,
    generatedAt: new Date().toISOString()
  };
}

function buildRemoteStockAlerts(db) {
  const rows = (db.products || []).filter((product) => product.active !== false).map(sanitizeInventoryProduct).map(enrichInventoryRow);
  const missing = rows.filter((product) => product.alertStatus === 'missing');
  const near = rows.filter((product) => product.alertStatus === 'near');
  return {
    ok: true,
    totals: {
      missingCount: missing.length,
      nearStockCount: near.length,
      attentionCount: missing.length + near.length
    },
    missing,
    near,
    generatedAt: new Date().toISOString()
  };
}

function buildSyncSnapshot(db) {
  const profile = getBranchProfile(db);
  const summary = buildRemoteSummary(db, { profile, serviceMode: 'LOCAL_ADMIN' });
  const inventory = buildInventoryPayloadFromRows((db.products || []).filter((product) => product.active !== false).map(sanitizeInventoryProduct), {});
  const sales = buildRemoteSales(db, { limit: 5000 });
  const employeeReport = buildRemoteEmployeeReport(db, { range: 'day' });
  const stockAlerts = buildRemoteStockAlerts(db);
  const cashOverview = buildCashOverview(db, normalizeDateKey());
  const users = (db.users || []).map((user) => ({
    id: user.id,
    name: user.name || user.username || 'Usuario',
    employeeNumber: user.employeeNumber || '',
    role: user.role || 'POS',
    active: user.active !== false
  }));
  return {
    branchProfile: profile,
    branchCode: normalizeMirrorBranchCode(profile.code),
    sourceInstanceId: db.syncInstanceId || '',
    summary,
    inventory,
    sales,
    employeeReport,
    stockAlerts,
    cashOverview,
    users,
    sentAt: new Date().toISOString(),
    sourceAdminName: ADMIN_NAME,
    sourceHostname: os.hostname()
  };
}

function getEffectiveRemoteSummary(db, branchCode = '') {
  const replica = getMirrorReplica(db, branchCode);
  if (replica.snapshot?.summary) {
    const payload = JSON.parse(JSON.stringify(replica.snapshot.summary));
    payload.branchProfile = replica.snapshot.branchProfile || payload.branchProfile || getBranchProfile(db);
    payload.remote = {
      ...(payload.remote || {}),
      serviceMode: 'REMOTE_MIRROR',
      mirrorUpdatedAt: replica.lastSnapshotAt || replica.snapshot.sentAt || '',
      permissions: getRemotePermissions(db)
    };
    payload.generatedAt = replica.snapshot.sentAt || payload.generatedAt || new Date().toISOString();
    return payload;
  }
  return buildRemoteSummary(db, { profile: getBranchProfile(db), serviceMode: 'LOCAL_ADMIN' });
}

function getEffectiveRemoteInventory(db, filters = {}, branchCode = '') {
  const replica = getMirrorReplica(db, branchCode);
  if (replica.snapshot?.inventory?.products) {
    return buildInventoryPayloadFromRows(replica.snapshot.inventory.products, filters);
  }
  return buildRemoteInventory(db, filters);
}

function getEffectiveRemoteSales(db, filters = {}, branchCode = '') {
  const replica = getMirrorReplica(db, branchCode);
  if (Array.isArray(replica.snapshot?.sales?.sales)) {
    return buildRemoteSalesPayloadFromRows(
      replica.snapshot.sales.sales,
      { ...(replica.snapshot.sales.filters || {}), ...(filters || {}) },
      replica.lastSnapshotAt || replica.snapshot.sentAt || replica.snapshot.sales.generatedAt || new Date().toISOString()
    );
  }
  return buildRemoteSales(db, filters);
}

function getEffectiveRemoteEmployeeReport(db, filters = {}, branchCode = '') {
  const replica = getMirrorReplica(db, branchCode);
  if (replica.snapshot?.employeeReport) {
    const payload = JSON.parse(JSON.stringify(replica.snapshot.employeeReport));
    payload.generatedAt = replica.lastSnapshotAt || replica.snapshot.sentAt || payload.generatedAt || new Date().toISOString();
    return payload;
  }
  return buildRemoteEmployeeReport(db, filters);
}

function getEffectiveRemoteStockAlerts(db, branchCode = '') {
  const replica = getMirrorReplica(db, branchCode);
  if (replica.snapshot?.stockAlerts) {
    const payload = JSON.parse(JSON.stringify(replica.snapshot.stockAlerts));
    payload.generatedAt = replica.lastSnapshotAt || replica.snapshot.sentAt || payload.generatedAt || new Date().toISOString();
    return payload;
  }
  if (replica.snapshot?.inventory?.products) {
    const rows = (replica.snapshot.inventory.products || []).map(enrichInventoryRow);
    return {
      ok: true,
      totals: {
        missingCount: rows.filter((product) => product.alertStatus === 'missing').length,
        nearStockCount: rows.filter((product) => product.alertStatus === 'near').length,
        attentionCount: rows.filter((product) => product.alertStatus !== 'normal').length
      },
      missing: rows.filter((product) => product.alertStatus === 'missing'),
      near: rows.filter((product) => product.alertStatus === 'near'),
      generatedAt: replica.lastSnapshotAt || replica.snapshot.sentAt || new Date().toISOString()
    };
  }
  return buildRemoteStockAlerts(db);
}

function getEffectiveRemoteCashOverview(db, dateKey = normalizeDateKey(), branchCode = '') {
  const replica = getMirrorReplica(db, branchCode);
  const overview = replica.snapshot?.cashOverview;
  if (overview && normalizeDateKey(overview.dateKey || '') === dateKey) {
    return JSON.parse(JSON.stringify(overview));
  }
  if (RENDER_MIRROR_ONLY && replica.snapshot) {
    return { dateKey, unavailable: true, message: 'La sucursal publica el resumen de caja del día actual.' };
  }
  return buildCashOverview(db, dateKey);
}

function getEffectiveRemoteUsers(db, branchCode = '') {
  const replica = getMirrorReplica(db, branchCode);
  if (Array.isArray(replica.snapshot?.users)) {
    return JSON.parse(JSON.stringify(replica.snapshot.users));
  }
  return (db.users || []).map((user) => ({
    id: user.id,
    name: user.name || user.username || 'Usuario',
    employeeNumber: user.employeeNumber || '',
    role: user.role || 'POS',
    active: user.active !== false
  }));
}

function queueMirrorInventoryCommand(db, payload = {}, branchCode = '') {
  const target = getWritableMirrorReplica(db, branchCode);
  const replica = target.replica;
  replica.commands ||= [];
  replica.commandSequence = safeNumber(replica.commandSequence, 0) + 1;
  const command = {
    id: uuidv4(),
    sequence: replica.commandSequence,
    commandType: 'INVENTORY_ADJUST',
    productId: payload.productId,
    operation: payload.operation === 'set' ? 'set' : (payload.operation === 'subtract' ? 'subtract' : 'add'),
    quantity: round2(safeNumber(payload.quantity)),
    targetStock: payload.targetStock !== undefined ? round2(safeNumber(payload.targetStock)) : null,
    reason: String(payload.reason || '').trim() || 'Ajuste remoto de inventario',
    actorName: String(payload.actorName || '').trim() || 'Administrador maestro',
    actorUsername: String(payload.actorUsername || '').trim() || 'admin',
    actorBranch: String(payload.actorBranch || '').trim(),
    status: 'PENDING',
    createdAt: new Date().toISOString(),
    acknowledgedAt: ''
  };
  replica.commands.push(command);
  replica.commands = replica.commands.slice(-500);
  replica.updatedAt = new Date().toISOString();
  storeMirrorReplica(db, target.branchCode, replica);
  return command;
}

function applyMirrorCommandBatch(db, commands = []) {
  const applied = [];
  const failed = [];
  for (const command of Array.isArray(commands) ? commands : []) {
    try {
      if (command.commandType === 'INVENTORY_ADJUST') {
        const result = applyInventoryAdjustment(db, {
          productId: command.productId,
          operation: command.operation,
          quantity: command.quantity,
          targetStock: command.targetStock,
          reason: command.reason,
          sourceType: 'REMOTE_MIRROR_COMMAND',
          actorName: command.actorName,
          actorUsername: command.actorUsername,
          actorBranch: command.actorBranch
        });
        applied.push({ commandId: command.id, productId: result.product.id, movement: result.movement });
      }
    } catch (err) {
      failed.push({ commandId: command.id, message: err.message });
    }
  }
  return { applied, failed };
}

function normalizeTicketItems(items = []) {
  return (Array.isArray(items) ? items : []).map((item) => {
    const quantity = safeNumber(item.quantity);
    const originalUnitPrice = round2(item.originalUnitPrice ?? item.unitPrice);
    const discountAmount = round2(item.discountAmount ?? 0);
    const effectiveUnitPrice = Math.max(0, round2(item.unitPrice ?? (originalUnitPrice - discountAmount)));
    const lineTotal = round2(item.lineTotal ?? (quantity * effectiveUnitPrice));
    return {
      productId: item.productId,
      name: item.name,
      unit: item.unit || 'PZA',
      quantity,
      originalUnitPrice,
      discountAmount,
      discountPercent: round2(item.discountPercent ?? 0),
      unitPrice: effectiveUnitPrice,
      selectedPromotionId: item.selectedPromotionId || '',
      promotionId: item.promotionId || '',
      promotionName: item.promotionName || '',
      saleMode: item.saleMode === 'WHOLESALE' ? 'WHOLESALE' : 'RETAIL',
      saleByAmount: item.saleByAmount === true,
      saleAmountTarget: round2(item.saleAmountTarget ?? 0),
      saleAmountUnitLabel: String(item.saleAmountUnitLabel || '').trim(),
      lineTotal
    };
  });
}

function getItemEffectiveUnitPrice(item = {}) {
  if (Number.isFinite(Number(item.unitPrice))) return round2(item.unitPrice);
  return Math.max(0, round2(safeNumber(item.originalUnitPrice) - safeNumber(item.discountAmount)));
}

function sumTicket(items = []) {
  return round2((Array.isArray(items) ? items : []).reduce((acc, item) => acc + safeNumber(item.lineTotal, safeNumber(item.quantity) * getItemEffectiveUnitPrice(item)), 0));
}

function normalizePromotionPayload(body = {}) {
  const type = ['percent', 'fixed_price', 'free'].includes(String(body.type || '').trim()) ? String(body.type || '').trim() : 'percent';
  const value = type === 'free' ? 100 : round2(body.value ?? 0);
  const productIds = Array.from(new Set((Array.isArray(body.productIds) ? body.productIds : [body.productId]).map((id) => String(id || '').trim()).filter(Boolean)));
  return {
    id: String(body.id || uuidv4()).trim(),
    name: String(body.name || '').trim() || 'Promoción',
    type,
    value,
    productIds,
    isActive: body.isActive !== false,
    startAt: body.startAt ? new Date(body.startAt).toISOString() : '',
    endAt: body.endAt ? new Date(body.endAt).toISOString() : '',
    createdAt: body.createdAt || new Date().toISOString(),
    updatedAt: new Date().toISOString()
  };
}

function ensurePromotionCollections(db) {
  db.promotions = Array.isArray(db.promotions) ? db.promotions : [];
}

function getApplicablePromotionsForProduct(db, productId, atDate = new Date()) {
  ensurePromotionCollections(db);
  const now = atDate instanceof Date ? atDate : new Date(atDate || Date.now());
  return (db.promotions || []).filter((promotion) => {
    if (promotion.isActive === false) return false;
    const productIds = Array.isArray(promotion.productIds) ? promotion.productIds.map((id) => String(id)) : [];
    if (!productIds.includes(String(productId))) return false;
    if (promotion.startAt) {
      const start = new Date(promotion.startAt);
      if (!Number.isNaN(start.getTime()) && now < start) return false;
    }
    if (promotion.endAt) {
      const end = new Date(promotion.endAt);
      if (!Number.isNaN(end.getTime()) && now > end) return false;
    }
    return true;
  });
}

function getActivePromotionForProduct(db, productId, atDate = new Date(), selectedPromotionId = '') {
  // v13.6.4: las promociones de producto son manuales. Si el cajero no selecciona una promoción, no se aplica ninguna.
  if (!selectedPromotionId || selectedPromotionId === '__none__') return null;
  const promotions = getApplicablePromotionsForProduct(db, productId, atDate);
  return promotions.find((promotion) => promotion.id === selectedPromotionId) || null;
}

function applyPromotionToTicketItem(db, item = {}, product = {}) {
  const normalized = normalizeTicketItems([item])[0] || null;
  if (!normalized) return null;
  const promotion = getActivePromotionForProduct(db, normalized.productId, new Date(), normalized.selectedPromotionId || '');
  const basePrice = safeNumber(normalized.originalUnitPrice || normalized.unitPrice || product.retailPrice || product.price, 0);
  let discountAmount = 0;
  let discountPercent = 0;
  let effectiveUnitPrice = basePrice;
  if (promotion) {
    if (promotion.type === 'free') {
      discountAmount = basePrice;
      discountPercent = 100;
      effectiveUnitPrice = 0;
    } else if (promotion.type === 'fixed_price') {
      effectiveUnitPrice = Math.max(0, round2(promotion.value));
      discountAmount = Math.max(0, round2(basePrice - effectiveUnitPrice));
      discountPercent = basePrice > 0 ? round2((discountAmount / basePrice) * 100) : 0;
    } else {
      discountPercent = Math.max(0, Math.min(100, round2(promotion.value)));
      discountAmount = round2(basePrice * (discountPercent / 100));
      effectiveUnitPrice = Math.max(0, round2(basePrice - discountAmount));
    }
  } else if (String(normalized.promotionId || '').startsWith('manual_') || String(normalized.promotionName || '').toLowerCase().includes('condon')) {
    discountAmount = Math.max(0, Math.min(basePrice, round2(normalized.discountAmount || 0)));
    discountPercent = basePrice > 0 ? round2((discountAmount / basePrice) * 100) : 0;
    effectiveUnitPrice = Math.max(0, round2(basePrice - discountAmount));
  }
  return normalizeTicketItems([{ ...normalized, originalUnitPrice: basePrice, discountAmount, discountPercent, unitPrice: effectiveUnitPrice, promotionId: promotion?.id || normalized.promotionId || '', promotionName: promotion?.name || normalized.promotionName || '', lineTotal: round2(safeNumber(normalized.quantity) * effectiveUnitPrice) }])[0];
}

function normalizeSalePayments(method = 'cash', payments = [], total = 0, amountReceived = 0) {
  const paymentMethod = ['cash', 'card', 'transfer', 'mixed'].includes(String(method || '').trim()) ? String(method || '').trim() : 'cash';
  let normalized = (Array.isArray(payments) ? payments : []).map((line) => ({
    method: ['cash', 'card', 'transfer'].includes(String(line.method || '').trim()) ? String(line.method || '').trim() : 'cash',
    amount: round2(line.amount),
    reference: String(line.reference || '').trim(),
    bank: String(line.bank || '').trim()
  })).filter((line) => safeNumber(line.amount) > 0);
  if (!normalized.length) {
    if (paymentMethod === 'cash') {
      const received = Math.max(round2(amountReceived || total), round2(total));
      normalized = [{ method: 'cash', amount: received, reference: '', bank: '' }];
    } else if (paymentMethod === 'card') {
      normalized = [{ method: 'card', amount: round2(total), reference: '', bank: '' }];
    } else if (paymentMethod === 'transfer') {
      normalized = [{ method: 'transfer', amount: round2(total), reference: '', bank: '' }];
    }
  }
  const totalPaid = round2(normalized.reduce((acc, line) => acc + safeNumber(line.amount), 0));
  const changeDue = normalized.some((line) => line.method === 'cash') ? Math.max(0, round2(totalPaid - round2(total))) : 0;
  if (round2(total) > 0 && totalPaid + 0.009 < round2(total)) throw new Error('La suma de los pagos no cubre el total de la venta.');
  return { paymentMethod, payments: normalized, amountReceived: totalPaid, changeDue };
}

function normalizeDateKey(value) {
  // La caja opera por fecha de America/Mexico_City, aunque el servidor esté en otra zona.
  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(trimmed)) return trimmed;
  }
  const parsed = value ? new Date(value) : new Date();
  const d = Number.isNaN(parsed.getTime()) ? new Date() : parsed;
  const parts = new Intl.DateTimeFormat('en-US', {
    timeZone: 'America/Mexico_City', year: 'numeric', month: '2-digit', day: '2-digit'
  }).formatToParts(d);
  const values = Object.fromEntries(parts.map((part) => [part.type, part.value]));
  const yyyy = values.year;
  const mm = values.month;
  const dd = values.day;
  return `${yyyy}-${mm}-${dd}`;
}

function isOwnerConsumptionSale(sale = {}) {
  const type = String(sale.saleType || sale.transactionType || sale.internalType || '').trim().toLowerCase();
  return type === 'owner_consumption'
    || type === 'internal_owner_consumption'
    || type === 'owner-consumption'
    || sale.affectsCash === false
    || sale.paymentMethod === 'internal';
}

function saleCashTotal(sale = {}) {
  if (isOwnerConsumptionSale(sale)) return 0;
  return round2(safeNumber(sale.total));
}

function ensureCashCollections(db) {
  db.cashOpenings = Array.isArray(db.cashOpenings) ? db.cashOpenings : [];
  db.cashEntries = Array.isArray(db.cashEntries) ? db.cashEntries : [];
  db.cashClosings = Array.isArray(db.cashClosings) ? db.cashClosings : [];
}

function buildCashOverview(db, dateKey = normalizeDateKey()) {
  ensureCashCollections(db);
  db.procurementPurchases = Array.isArray(db.procurementPurchases) ? db.procurementPurchases : [];
  const openings = (db.cashOpenings || []).filter((entry) => normalizeDateKey(entry.dateKey || entry.createdAt) === dateKey);
  const entries = (db.cashEntries || []).filter((entry) => normalizeDateKey(entry.dateKey || entry.createdAt) === dateKey);
  const restockPurchases = entries.filter((entry) => entry.kind === 'restock_purchase');
  const purchases = (db.procurementPurchases || []).filter((entry) => normalizeDateKey(entry.dateKey || entry.createdAt) === dateKey);
  const allSalesForDay = (db.sales || []).filter((sale) => normalizeDateKey(sale.dateKey || sale.createdAt) === dateKey);
  const sales = allSalesForDay.filter((sale) => !isOwnerConsumptionSale(sale));
  const internalOwnerSales = allSalesForDay.filter((sale) => isOwnerConsumptionSale(sale));
  const productMap = new Map((db.products || []).map((product) => [product.id, product]));
  const totalsByMethod = { cash: 0, card: 0, transfer: 0 };
  let acquisitionCost = 0;
  for (const sale of sales) {
    const payments = normalizeSalePaymentBreakdownForCashbox(sale);
    for (const payment of payments) {
      const key = ['cash', 'card', 'transfer'].includes(payment.method) ? payment.method : 'cash';
      totalsByMethod[key] = round2(totalsByMethod[key] + safeNumber(payment.amount));
    }
    for (const item of (sale.items || [])) {
      const product = productMap.get(item.productId) || {};
      const cost = safeNumber(item.costPrice ?? item.unitCost ?? product.costPrice);
      acquisitionCost = round2(acquisitionCost + safeNumber(item.quantity) * cost);
    }
  }
  const openingAmount = round2(openings.reduce((acc, entry) => acc + safeNumber(entry.amount), 0));
  const operationalExpenses = entries.filter((entry) => entry.kind === 'operational_expense');
  const ownerConsumptions = entries.filter((entry) => entry.kind === 'owner_consumption');
  const cashIncomeEntries = entries.filter((entry) => ['cash_income', 'cash_inflow', 'cash_contribution'].includes(String(entry.kind || '').toLowerCase()));
  const cashOutflowEntries = entries.filter((entry) => entry.kind !== 'sale_income' && !['cash_income', 'cash_inflow', 'cash_contribution'].includes(String(entry.kind || '').toLowerCase()));
  const expensesTotal = round2(operationalExpenses.reduce((acc, entry) => acc + safeNumber(entry.accountingAmount ?? entry.amountCashImpact), 0));
  const expensesCashImpactTotal = round2(cashOutflowEntries.reduce((acc, entry) => acc + safeNumber(entry.amountCashImpact), 0));
  const cashInflowTotal = round2(cashIncomeEntries.reduce((acc, entry) => acc + safeNumber(entry.amountCashImpact ?? entry.amount), 0));
  const restockTotal = round2(restockPurchases.reduce((acc, entry) => acc + safeNumber(entry.totalCost ?? entry.accountingAmount), 0));
  const restockCashOutflowTotal = round2(restockPurchases.reduce((acc, entry) => acc + safeNumber(entry.amountCashImpact), 0));
  const ownerValueTotal = round2(ownerConsumptions.reduce((acc, entry) => acc + safeNumber(entry.inventoryValue), 0));
  const salesDiscountTotal = round2(sales.reduce((acc, sale) => acc + safeNumber(sale.discountTotal), 0));
  const procurementTotal = round2(purchases.reduce((acc, entry) => acc + safeNumber(entry.totalCost ?? entry.accountingAmount), 0));
  const accountingOutflowTotal = round2(expensesTotal + restockTotal + procurementTotal + ownerValueTotal);
  const expensesByPayment = operationalExpenses.reduce((bucket, entry) => addAmountToBucket(bucket, entry.paymentMethod || 'other', entry.accountingAmount ?? entry.amountCashImpact), { cash: 0, card: 0, transfer: 0, other: 0, internal: 0 });
  const procurementByPayment = purchases.reduce((bucket, entry) => addAmountToBucket(bucket, entry.paymentMethod || 'other', entry.totalCost), { cash: 0, card: 0, transfer: 0, other: 0, internal: 0 });
  const ownerByPayment = ownerConsumptions.reduce((bucket, entry) => addAmountToBucket(bucket, entry.paymentMethod || 'internal', entry.inventoryValue), { cash: 0, card: 0, transfer: 0, other: 0, internal: 0 });
  const expectedCash = round2(openingAmount + totalsByMethod.cash + cashInflowTotal - expensesCashImpactTotal);
  return {
    dateKey,
    openingAmount,
    salesCount: sales.length,
    salesTotal: round2(sales.reduce((acc, sale) => acc + safeNumber(sale.total), 0)),
    salesDiscountTotal,
    acquisitionCost: round2(acquisitionCost),
    grossProfit: round2(sales.reduce((acc, sale) => acc + safeNumber(sale.total), 0) - acquisitionCost),
    payments: totalsByMethod,
    expensesTotal,
    expensesCashImpactTotal,
    cashInflowTotal,
    expensesByPayment,
    restockTotal,
    restockCashOutflowTotal,
    restockPurchases,
    procurementTotal,
    procurementByPayment,
    ownerValueTotal,
    ownerByPayment,
    accountingOutflowTotal,
    ownerInternalCount: internalOwnerSales.length,
    expectedCash,
    openings,
    entries,
    purchases,
    sales: sales.map((sale) => ({
      id: sale.id,
      folio: sale.folio,
      dateKey: sale.dateKey || normalizeDateKey(sale.createdAt),
      total: round2(safeNumber(sale.total)),
      subtotalBeforeDiscount: round2(safeNumber(sale.subtotalBeforeDiscount)),
      discountTotal: round2(safeNumber(sale.discountTotal)),
      paymentMethod: sale.paymentMethod || 'cash',
      payments: normalizeSalePaymentBreakdownForCashbox(sale),
      amountReceived: round2(safeNumber(sale.amountReceived)),
      changeDue: round2(safeNumber(sale.changeDue)),
      terminalCode: sale.terminalCode || '',
      operatorName: sale.operatorName || '',
      employeeNumber: sale.employeeNumber || '',
      customerName: sale.customerName || '',
      createdAt: sale.createdAt || '',
      cashEntryId: sale.cashEntryId || ''
    })),
    closings: (db.cashClosings || []).filter((entry) => normalizeDateKey(entry.dateKey || entry.createdAt) === dateKey)
  };
}

function isLowStock(product) {
  return getInventoryAlertStatus(product) === 'missing';
}


function normalizeBusinessPaymentMethod(value = '') {
  const raw = String(value || '').trim().toLowerCase();
  if (['cash', 'efectivo'].includes(raw)) return 'cash';
  if (['card', 'tarjeta', 'debit', 'credit'].includes(raw)) return 'card';
  if (['transfer', 'transferencia', 'spei'].includes(raw)) return 'transfer';
  if (['internal', 'none', 'sin_pago', 'sin-pago', 'no_cash'].includes(raw)) return 'internal';
  return 'other';
}

function paymentMethodLabel(value = '') {
  const method = normalizeBusinessPaymentMethod(value);
  if (method === 'cash') return 'Efectivo';
  if (method === 'card') return 'Tarjeta';
  if (method === 'transfer') return 'Transferencia';
  if (method === 'internal') return 'Sin pago / interno';
  return 'Otro / apunte';
}

function addAmountToBucket(bucket, method, amount) {
  const key = normalizeBusinessPaymentMethod(method);
  bucket[key] = round2(safeNumber(bucket[key]) + safeNumber(amount));
  return bucket;
}

function normalizeSalePaymentBreakdownForCashbox(sale = {}) {
  const total = round2(safeNumber(sale.total));
  let payments = Array.isArray(sale.payments) && sale.payments.length
    ? sale.payments.map((line) => ({
        method: normalizeBusinessPaymentMethod(line.method || sale.paymentMethod || 'cash'),
        amount: round2(line.amount),
        reference: String(line.reference || '').trim(),
        bank: String(line.bank || '').trim()
      })).filter((line) => safeNumber(line.amount) > 0)
    : [{ method: normalizeBusinessPaymentMethod(sale.paymentMethod || 'cash'), amount: total, reference: '', bank: '' }];

  // Si se capturó efectivo recibido mayor al total, changeDue no debe sumarse a caja.
  let pendingChange = round2(safeNumber(sale.changeDue));
  payments = payments.map((line) => {
    if (pendingChange > 0 && line.method === 'cash') {
      const discount = Math.min(safeNumber(line.amount), pendingChange);
      pendingChange = round2(pendingChange - discount);
      return { ...line, amount: round2(safeNumber(line.amount) - discount) };
    }
    return line;
  }).filter((line) => safeNumber(line.amount) > 0);

  const paidTotal = round2(payments.reduce((acc, line) => acc + safeNumber(line.amount), 0));
  if (paidTotal > total + 0.009 && total > 0) {
    // Última defensa: si por cualquier causa sigue sobrando, ajusta primero efectivo.
    let excess = round2(paidTotal - total);
    payments = payments.map((line) => {
      if (excess > 0 && line.method === 'cash') {
        const cut = Math.min(excess, safeNumber(line.amount));
        excess = round2(excess - cut);
        return { ...line, amount: round2(safeNumber(line.amount) - cut) };
      }
      return line;
    }).filter((line) => safeNumber(line.amount) > 0);
  }
  return payments;
}

function saleCashAmountForCashbox(sale = {}) {
  if (isOwnerConsumptionSale(sale)) return 0;
  return round2(normalizeSalePaymentBreakdownForCashbox(sale)
    .filter((line) => line.method === 'cash')
    .reduce((acc, line) => acc + safeNumber(line.amount), 0));
}

function buildSaleCashEntry(sale = {}, reqUser = {}) {
  if (!sale || !sale.id || isOwnerConsumptionSale(sale)) return null;
  const payments = normalizeSalePaymentBreakdownForCashbox(sale);
  const dateKey = normalizeDateKey(sale.dateKey || sale.createdAt);
  return {
    id: sale.cashEntryId || uuidv4(),
    kind: 'sale_income',
    saleId: sale.id,
    folio: sale.folio || '',
    dateKey,
    amountCashImpact: saleCashAmountForCashbox(sale),
    accountingAmount: round2(safeNumber(sale.total)),
    inventoryValue: 0,
    paymentMethod: normalizeBusinessPaymentMethod(sale.paymentMethod || payments[0]?.method || 'cash'),
    paymentMethodLabel: paymentMethodLabel(sale.paymentMethod || payments[0]?.method || 'cash'),
    payments,
    description: `Venta cobrada ${sale.folio || ''}`.trim(),
    notes: sale.customerName ? `Cliente: ${sale.customerName}` : '',
    terminalCode: sale.terminalCode || 'POS-01',
    userId: sale.operatorId || reqUser?.id || '',
    userName: sale.operatorName || reqUser?.name || reqUser?.username || '',
    createdAt: sale.createdAt || new Date().toISOString()
  };
}

function ensureSaleCashEntry(db, sale = {}, reqUser = {}) {
  ensureCashCollections(db);
  if (!sale || !sale.id || isOwnerConsumptionSale(sale)) return null;
  const existing = (db.cashEntries || []).find((entry) => entry.saleId === sale.id && entry.kind === 'sale_income');
  if (existing) {
    sale.cashEntryId = existing.id;
    return existing;
  }
  const entry = buildSaleCashEntry(sale, reqUser);
  if (!entry) return null;
  sale.cashEntryId = entry.id;
  db.cashEntries.push(entry);
  return entry;
}


const RBAC_MODULES = [
  'dashboard', 'catalog', 'products', 'inventory', 'sales', 'reports', 'promotions', 'expenses', 'branches', 'fiscal', 'terminals', 'security'
];

// =============================================================
// v13.6.41 — Permisos de VISUALIZACIÓN por terminal/nodo.
// La identidad/rol del usuario se conserva para autenticación y
// operaciones; qué módulos se muestran depende del dispositivo.
// =============================================================
const TERMINAL_VIEW_MODULES_13641 = RBAC_MODULES.slice();

function defaultTerminalViewPermissions13641(terminalType = 'POS') {
  const type = String(terminalType || 'POS').toUpperCase();
  const result = {};
  for (const moduleKey of TERMINAL_VIEW_MODULES_13641) result[moduleKey] = type === 'ADMIN';
  if (type !== 'ADMIN') result.catalog = true; // Venta rápida: valor seguro por defecto para cajas/POS.
  return result;
}

function normalizeTerminalViewPermissions13641(input = {}, terminalType = 'POS') {
  const base = defaultTerminalViewPermissions13641(terminalType);
  const hasInput = input && typeof input === 'object' && Object.keys(input).length > 0;
  if (!hasInput) return base;
  for (const moduleKey of TERMINAL_VIEW_MODULES_13641) {
    base[moduleKey] = input[moduleKey] === true;
  }
  return base;
}

function ensureTerminalViewPermissions13641(terminal) {
  if (!terminal) return defaultTerminalViewPermissions13641('POS');
  const normalized = normalizeTerminalViewPermissions13641(terminal.viewPermissions || {}, terminal.terminalType || 'POS');
  terminal.viewPermissions = normalized;
  return normalized;
}


// v13.6.42 — Persistencia verificable y sincronización robusta de permisos por terminal.
function terminalViewPermissionsEqual13642(a = {}, b = {}) {
  return TERMINAL_VIEW_MODULES_13641.every((key) => (a?.[key] === true) === (b?.[key] === true));
}

function findTerminalForPermissions13642(db, identifier = '') {
  const key = String(identifier || '').trim();
  if (!key) return null;
  return (db.terminals || []).find((terminal) =>
    String(terminal.id || '') === key ||
    String(terminal.terminalCode || '').trim() === key ||
    String(terminal.nodeId || '').trim() === key
  ) || null;
}

function defaultModulePermissions(level = 'none') {
  const template = {};
  for (const moduleKey of RBAC_MODULES) {
    template[moduleKey] = { read: false, write: false, delete: false };
  }
  if (level === 'full') {
    for (const moduleKey of RBAC_MODULES) template[moduleKey] = { read: true, write: true, delete: true };
  }
  if (level === 'pos') {
    template.dashboard.read = true;
    template.catalog = { read: true, write: true, delete: false };
    template.sales = { read: true, write: false, delete: false };
  }
  return template;
}

function normalizePermissions(input = {}, fallbackLevel = 'none') {
  const base = defaultModulePermissions(fallbackLevel);
  for (const moduleKey of RBAC_MODULES) {
    const row = input[moduleKey] || {};
    base[moduleKey] = {
      read: row.read === true,
      write: row.write === true,
      delete: row.delete === true
    };
  }
  return base;
}

function isPrincipalAdmin(db, user = {}) {
  if (!user || user.role !== 'ADMIN') return false;
  const principalId = db.security?.principalUserId || 'u-admin';
  return user.id === principalId || user.username === 'admin' || user.isPrincipalAdmin === true;
}

function ensureRbacCollections(db) {
  let changed = false;
  db.security ||= {};
  if (!db.security.principalProfile) {
    db.security.principalProfile = { name: 'Hanna Jiménez', label: 'Administrador Principal' };
    changed = true;
  }
  if (!db.security.principalUserId) {
    const admin = (db.users || []).find((u) => u.role === 'ADMIN') || (db.users || [])[0];
    db.security.principalUserId = admin?.id || 'u-admin';
    changed = true;
  }
  db.rbacModules = RBAC_MODULES.slice();
  for (const user of (db.users || [])) {
    const shouldFull = user.role === 'ADMIN' && (user.id === db.security.principalUserId || user.username === 'admin');
    const fallback = shouldFull ? 'full' : (user.role === 'ADMIN' ? 'full' : 'pos');
    const next = normalizePermissions(user.permissions || {}, fallback);
    if (JSON.stringify(user.permissions || {}) !== JSON.stringify(next)) {
      user.permissions = next;
      changed = true;
    }
    if (typeof user.profileLabel !== 'string') {
      user.profileLabel = shouldFull ? 'Administrador Principal' : (user.role === 'ADMIN' ? 'Sub-administrador' : 'Cajero');
      changed = true;
    }
  }
  return changed;
}

function getUserPermissions(db, user = {}) {
  if (!user) return defaultModulePermissions('none');
  if (isPrincipalAdmin(db, user)) return defaultModulePermissions('full');
  return normalizePermissions(user.permissions || {}, user.role === 'ADMIN' ? 'full' : 'pos');
}

function userHasPermission(db, user = {}, moduleKey = '', actionKey = 'read') {
  if (isPrincipalAdmin(db, user)) return true;
  const permissions = getUserPermissions(db, user);
  return permissions?.[moduleKey]?.[actionKey] === true;
}

function moduleFromRequestPath(pathName = '') {
  const p = String(pathName || '').split('?')[0];
  if (p.includes('/products')) return 'products';
  if (p.includes('/inventory')) return 'inventory';
  if (p.includes('/sales') || p.includes('/tickets')) return 'sales';
  if (p.includes('/reports')) return 'reports';
  if (p.includes('/promotions')) return 'promotions';
  if (p.includes('/cash') || p.includes('/accounting') || p.includes('/procurement')) return 'expenses';
  if (p.includes('/branches') || p.includes('/branch-profile')) return 'branches';
  if (p.includes('/fiscal') || p.includes('/invoice')) return 'fiscal';
  if (p.includes('/terminals') || p.includes('/conflicts')) return 'terminals';
  if (p.includes('/security') || p.includes('/backup') || p.includes('/system/network')) return 'security';
  return 'dashboard';
}

function actionFromMethod(method = 'GET') {
  const m = String(method || 'GET').toUpperCase();
  if (m === 'GET' || m === 'HEAD') return 'read';
  if (m === 'DELETE') return 'delete';
  return 'write';
}

function normalizeSystemNetworkConfig(config = {}) {
  return {
    preferredHost: String(config.preferredHost || 'auto').trim() || 'auto',
    preferredPort: safeNumber(config.preferredPort, PORT) || PORT,
    lastKnownUrl: String(config.lastKnownUrl || '').trim(),
    autoReconnect: config.autoReconnect !== false,
    lockPort: config.lockPort === true,
    updatedAt: config.updatedAt || new Date().toISOString()
  };
}

function getLocalNetworkInfo(db) {
  const interfaces = os.networkInterfaces();
  const ips = [];
  for (const entries of Object.values(interfaces)) {
    for (const entry of entries || []) {
      if (entry.family === 'IPv4' && !entry.internal) ips.push(entry.address);
    }
  }
  const config = normalizeSystemNetworkConfig(db.systemNetwork || {});
  const selectedHost = config.preferredHost && config.preferredHost !== 'auto' ? config.preferredHost : (ips[0] || '127.0.0.1');
  return { ips, port: PORT, selectedHost, localUrl: `http://${selectedHost}:${PORT}`, config };
}

function getReportDateRange(query = {}) {
  const range = String(query.range || 'day');
  let start;
  let end;
  if (range === 'custom' && query.from && query.to) {
    start = parseDateInput(query.from, new Date());
    start.setHours(0,0,0,0);
    end = parseDateInput(query.to, new Date());
    end.setHours(23,59,59,999);
  } else if (range === 'week') {
    ({ start, end } = getRangeBounds('week', query.anchorDate || query.date || new Date()));
    end = new Date(end.getTime() - 1);
  } else {
    ({ start, end } = getRangeBounds('day', query.anchorDate || query.date || new Date()));
    end = new Date(end.getTime() - 1);
  }
  return { range, start, end, from: start.toISOString().slice(0,10), to: end.toISOString().slice(0,10) };
}

function buildFinancialReport(db, query = {}) {
  const dateRange = getReportDateRange(query);
  const productMap = new Map((db.products || []).map((p) => [p.id, p]));
  const rows = new Map();
  const sales = (db.sales || []).filter((sale) => {
    if (isOwnerConsumptionSale(sale)) return false;
    const d = new Date(sale.createdAt);
    return d >= dateRange.start && d <= dateRange.end;
  });
  for (const sale of sales) {
    for (const item of (sale.items || [])) {
      const key = item.productId || item.name || uuidv4();
      const product = productMap.get(item.productId) || {};
      if (!rows.has(key)) {
        rows.set(key, { productId: item.productId || '', name: item.name || product.name || 'Producto', unit: item.unit || product.unit || 'PZA', quantity: 0, acquisitionCost: 0, revenue: 0, retailRevenue: 0, wholesaleRevenue: 0, profit: 0 });
      }
      const row = rows.get(key);
      const qty = safeNumber(item.quantity);
      const costUnit = safeNumber(product.costPrice ?? item.costPrice, 0);
      const lineRevenue = round2(safeNumber(item.lineTotal, qty * safeNumber(item.unitPrice)));
      const lineCost = round2(qty * costUnit);
      row.quantity = round2(row.quantity + qty);
      row.acquisitionCost = round2(row.acquisitionCost + lineCost);
      row.revenue = round2(row.revenue + lineRevenue);
      if (item.saleMode === 'WHOLESALE') row.wholesaleRevenue = round2(row.wholesaleRevenue + lineRevenue);
      else row.retailRevenue = round2(row.retailRevenue + lineRevenue);
      row.profit = round2(row.profit + (lineRevenue - lineCost));
    }
  }
  const entries = (db.cashEntries || []).filter((entry) => {
    const d = new Date(entry.createdAt || entry.dateKey);
    return d >= dateRange.start && d <= dateRange.end;
  });
  const purchaseEntries = (db.procurementPurchases || []).filter((entry) => {
    const d = new Date(entry.createdAt || entry.dateKey);
    return d >= dateRange.start && d <= dateRange.end;
  });
  const operationalExpenses = round2(entries.filter((e) => e.kind === 'operational_expense').reduce((a,e) => a + safeNumber(e.accountingAmount ?? e.amountCashImpact), 0));
  const internalConsumptions = round2(entries.filter((e) => e.kind === 'owner_consumption').reduce((a,e) => a + safeNumber(e.inventoryValue), 0));
  const merchandisePurchased = round2(purchaseEntries.reduce((a,e) => a + safeNumber(e.totalCost), 0));
  const productRows = Array.from(rows.values()).sort((a,b) => b.revenue - a.revenue);
  const totals = {
    productsCount: productRows.length,
    salesCount: sales.length,
    acquisitionCost: round2(productRows.reduce((a,r) => a + r.acquisitionCost, 0)),
    totalRevenue: round2(productRows.reduce((a,r) => a + r.revenue, 0)),
    retailRevenue: round2(productRows.reduce((a,r) => a + r.retailRevenue, 0)),
    wholesaleRevenue: round2(productRows.reduce((a,r) => a + r.wholesaleRevenue, 0)),
    grossProfit: round2(productRows.reduce((a,r) => a + r.profit, 0)),
    operationalExpenses,
    internalConsumptions,
    merchandisePurchased
  };
  totals.netProfit = round2(totals.grossProfit - operationalExpenses - internalConsumptions - merchandisePurchased);
  return { ok: true, range: dateRange, totals, products: productRows, generatedAt: new Date().toISOString() };
}

function buildExpenseTicketHtml(entry = {}) {
  const amount = safeNumber(entry.accountingAmount ?? entry.amountCashImpact ?? entry.totalCost, 0);
  return `<!doctype html><html><head><meta charset="utf-8"><title>Ticket gasto</title><style>@page{size:58mm auto;margin:2mm}body{font-family:Arial,sans-serif;color:#000}.ticket{width:54mm;margin:0 auto;font-size:10px}.center{text-align:center}.line{display:flex;justify-content:space-between;border-bottom:1px dashed #000;padding:3px 0}.warn{border:1px solid #000;padding:4px;margin:6px 0;text-align:center;font-weight:700}</style></head><body><div class="ticket"><h2 class="center">Cremería El Güero</h2><div class="center">Comprobante de gasto</div><div class="warn">APUNTE CONTABLE · NO DESCUENTA CAJA AUTOMÁTICAMENTE</div><div class="line"><span>Folio</span><strong>${entry.folio || entry.id || ''}</strong></div><div class="line"><span>Fecha</span><span>${entry.createdAt || ''}</span></div><div class="line"><span>Concepto</span><span>${entry.description || entry.productName || ''}</span></div><div class="line"><span>Monto</span><strong>$${amount.toFixed(2)}</strong></div><p>${entry.notes || ''}</p><hr><div class="center">Conserve este ticket para auditoría.</div></div></body></html>`;
}

function buildInventoryExcelHtml(products = []) {
  const rows = products.map((p) => {
    const missing = safeNumber(p.stock) <= 0 || getInventoryAlertStatus(p) === 'missing';
    const near = getInventoryAlertStatus(p) === 'near';
    const style = missing ? ' style="background:#ffcccc;color:#990000;font-weight:bold"' : (near ? ' style="background:#fff3cd"' : '');
    const cost = safeNumber(p.costPrice);
    const retail = safeNumber(p.retailPrice ?? p.price);
    const wholesale = safeNumber(p.wholesalePrice ?? p.retailPrice ?? p.price);
    const retailProfit = round2(retail - cost);
    const wholesaleProfit = round2(wholesale - cost);
    const retailMargin = retail > 0 ? round2((retailProfit / retail) * 100) : 0;
    const wholesaleMargin = wholesale > 0 ? round2((wholesaleProfit / wholesale) * 100) : 0;
    return `<tr${style}><td>${p.name || ''}</td><td>${p.category || ''}</td><td>${p.sku || ''}</td><td>${p.barcode || p.codigoBarras || ''}</td><td>${p.unit || ''}</td><td>${safeNumber(p.stock)}</td><td>${safeNumber(p.lowStockThreshold)}</td><td>${cost.toFixed(2)}</td><td>${retail.toFixed(2)}</td><td>${retailProfit.toFixed(2)}</td><td>${retailMargin.toFixed(2)}%</td><td>${wholesale.toFixed(2)}</td><td>${wholesaleProfit.toFixed(2)}</td><td>${wholesaleMargin.toFixed(2)}%</td><td>${missing ? 'FALTANTE' : near ? 'PRÓXIMO AL MÍNIMO' : 'OK'}</td></tr>`;
  }).join('');
  return `<!doctype html><html><head><meta charset="utf-8"><style>table{border-collapse:collapse;font-family:Arial,sans-serif;font-size:12px}td,th{border:1px solid #333;padding:6px}th{background:#1f3655;color:white}</style></head><body><table><thead><tr><th>Producto</th><th>Categoría</th><th>SKU</th><th>Código de barras</th><th>Unidad</th><th>Stock</th><th>Mínimo</th><th>Costo</th><th>Menudeo</th><th>Utilidad menudeo</th><th>Margen menudeo</th><th>Mayoreo</th><th>Utilidad mayoreo</th><th>Margen mayoreo</th><th>Estatus</th></tr></thead><tbody>${rows}</tbody></table></body></html>`;
}



// v13.6.33 / conserva protección antiapagón: lectura con autorrecuperación y backups diferidos.
let LAST_AUTOHEAL_STATUS = { ok: true, recovered: false, message: 'Sin recuperación pendiente.', at: new Date().toISOString() };
let BACKUP_SCHEDULE_TIMER = null;
let LAST_BACKUP_AT = 0;

function timestampForFile13633() {
  return new Date().toISOString().replace(/[:.]/g, '-');
}

function safeParseJsonFile13633(filePath) {
  try {
    const raw = fs.readFileSync(filePath, 'utf8');
    return { ok: true, data: JSON.parse(raw), raw };
  } catch (err) {
    return { ok: false, message: String(err.message || err) };
  }
}

function looksLikeCremeriaDb13633(db) {
  return !!(db && typeof db === 'object' && Array.isArray(db.products) && Array.isArray(db.sales) && Array.isArray(db.users));
}

function listValidBackups13633() {
  try {
    if (!fs.existsSync(BACKUPS_DIR)) return [];
    return fs.readdirSync(BACKUPS_DIR)
      .filter((name) => name.endsWith('.json'))
      .map((name) => path.join(BACKUPS_DIR, name))
      .sort((a, b) => fs.statSync(b).mtimeMs - fs.statSync(a).mtimeMs)
      .filter((file) => {
        const parsed = safeParseJsonFile13633(file);
        return parsed.ok && looksLikeCremeriaDb13633(parsed.data);
      });
  } catch (_) {
    return [];
  }
}

function createBackupCopy13633(reason = 'auto') {
  try {
    if (!fs.existsSync(DB_FILE)) return null;
    const parsed = safeParseJsonFile13633(DB_FILE);
    if (!parsed.ok || !looksLikeCremeriaDb13633(parsed.data)) return null;
    const name = `db-backup-${timestampForFile13633()}-${reason}.json`;
    const target = path.join(BACKUPS_DIR, name);
    fs.copyFileSync(DB_FILE, target);
    pruneBackups13633(120);
    LAST_BACKUP_AT = Date.now();
    return target;
  } catch (_) {
    return null;
  }
}

function scheduleBackup13633(reason = 'auto') {
  // No bloquea ventas: se difiere y se agrupa para evitar copiar la base en cada operación.
  if (Date.now() - LAST_BACKUP_AT < 60_000) return;
  clearTimeout(BACKUP_SCHEDULE_TIMER);
  BACKUP_SCHEDULE_TIMER = setTimeout(() => createBackupCopy13633(reason), 1500);
}

function pruneBackups13633(maxFiles = 120) {
  try {
    const files = fs.readdirSync(BACKUPS_DIR)
      .filter((name) => name.endsWith('.json'))
      .map((name) => path.join(BACKUPS_DIR, name))
      .sort((a, b) => fs.statSync(b).mtimeMs - fs.statSync(a).mtimeMs);
    for (const file of files.slice(maxFiles)) {
      try { fs.unlinkSync(file); } catch (_) {}
    }
  } catch (_) {}
}

function autoRecoverDbIfNeeded13633() {
  try {
    if (!fs.existsSync(DB_FILE)) return false;
    const parsed = safeParseJsonFile13633(DB_FILE);
    if (parsed.ok && looksLikeCremeriaDb13633(parsed.data)) {
      LAST_AUTOHEAL_STATUS = { ok: true, recovered: false, message: 'Base sana.', at: new Date().toISOString() };
      return false;
    }

    const corruptTarget = path.join(RECOVERY_DIR, `db-corrupt-${timestampForFile13633()}.json`);
    try { fs.copyFileSync(DB_FILE, corruptTarget); } catch (_) {}

    const backups = listValidBackups13633();
    if (!backups.length) {
      LAST_AUTOHEAL_STATUS = { ok: false, recovered: false, message: `Base dañada y sin respaldo válido: ${parsed.message || 'JSON inválido'}`, corruptCopy: corruptTarget, at: new Date().toISOString() };
      return false;
    }

    fs.copyFileSync(backups[0], DB_FILE);
    LAST_AUTOHEAL_STATUS = { ok: true, recovered: true, message: 'Se restauró automáticamente el último respaldo sano después de un cierre inesperado.', backup: backups[0], corruptCopy: corruptTarget, at: new Date().toISOString() };
    return true;
  } catch (err) {
    LAST_AUTOHEAL_STATUS = { ok: false, recovered: false, message: String(err.message || err), at: new Date().toISOString() };
    return false;
  }
}

function readDbFileWithAutoRecovery13633() {
  autoRecoverDbIfNeeded13633();
  const parsed = safeParseJsonFile13633(DB_FILE);
  if (!parsed.ok) throw new Error(`Base de datos dañada: ${parsed.message}`);
  decryptDatabaseSecrets(parsed.data);
  if (!looksLikeCremeriaDb13633(parsed.data)) throw new Error('Base de datos inválida: estructura incompleta.');
  return parsed.data;
}

function appendSaleJournal13633(stage, payload = {}) {
  try {
    const line = JSON.stringify({ stage, at: new Date().toISOString(), ...payload }) + '\n';
    fs.appendFileSync(path.join(JOURNALS_DIR, 'sale-journal.jsonl'), line, 'utf8');
  } catch (_) {}
}

function ensureDb() {
  if (!fs.existsSync(DB_FILE)) {
    writeDb(seed());
    return;
  }

  createStartupSafetyBackup();
  const db = readDbFileWithAutoRecovery13633();
  let changed = false;

  db.users ||= seed().users;
  db.security ||= { requireAdminPasswordForEdits: false, updatedAt: new Date().toISOString(), remoteAccess: {}, notifications: {} };
  db.security.requireAdminPasswordForEdits = false;
  db.security.notifications ||= {};
  if (typeof db.security.notifications.whatsappRestockNumber !== 'string') {
    db.security.notifications.whatsappRestockNumber = '';
    changed = true;
  }
  db.security.remoteAccess ||= {};
  if (!db.security.remoteAccess.apiKey) {
    db.security.remoteAccess.apiKey = RENDER_REMOTE_API_KEY || `REMOTE-${uuidv4().replace(/-/g, '').slice(0, 24).toUpperCase()}`;
    changed = true;
  }
  if (typeof db.security.remoteAccess.allowAdminCredentials !== 'boolean') {
    db.security.remoteAccess.allowAdminCredentials = !RENDER_MIRROR_ONLY;
    changed = true;
  }
  if (RENDER_MIRROR_ONLY && db.security.remoteAccess.allowAdminCredentials !== false) {
    db.security.remoteAccess.allowAdminCredentials = false;
    changed = true;
  }
  db.security.remoteAccess.permissions ||= {};
  if (typeof db.security.remoteAccess.permissions.readSummary !== 'boolean') {
    db.security.remoteAccess.permissions.readSummary = true;
    changed = true;
  }
  if (typeof db.security.remoteAccess.permissions.readInventory !== 'boolean') {
    db.security.remoteAccess.permissions.readInventory = true;
    changed = true;
  }
  if (typeof db.security.remoteAccess.permissions.writeInventory !== 'boolean') {
    db.security.remoteAccess.permissions.writeInventory = true;
    changed = true;
  }
  if (typeof db.security.remoteAccess.permissions.syncBridge !== 'boolean') {
    db.security.remoteAccess.permissions.syncBridge = true;
    changed = true;
  }
  db.fiscalConfig ||= seed().fiscalConfig;
  db.fiscalConfig = {
    ...seed().fiscalConfig,
    ...(db.fiscalConfig || {})
  };
  db.fiscalConfig.commercialName = String(db.fiscalConfig.commercialName || 'Studio Code').trim() || 'Studio Code';
  db.fiscalConfig.emitterRfc = String(db.fiscalConfig.emitterRfc || '').trim().toUpperCase();
  db.fiscalConfig.emitterName = String(db.fiscalConfig.emitterName || '').trim();
  db.fiscalConfig.emitterRegimen = normalizeFiscalRegimeValue(db.fiscalConfig.emitterRegimen || '', '605');
  db.fiscalConfig.expeditionZip = String(db.fiscalConfig.expeditionZip || '55750').trim();
  db.fiscalConfig.pacProvider = 'FACTURAMA';
  db.fiscalConfig.preferredInvoiceProvider = String(db.fiscalConfig.preferredInvoiceProvider || 'SAT_ASSISTED').toUpperCase() === 'FACTURAMA' ? 'FACTURAMA' : 'SAT_ASSISTED';
  db.fiscalConfig.satPortalUrl = String(db.fiscalConfig.satPortalUrl || SAT_PORTAL_DEFAULT).trim() || SAT_PORTAL_DEFAULT;
  db.fiscalConfig.satInfoUrl = String(db.fiscalConfig.satInfoUrl || SAT_INFO_PORTAL_DEFAULT).trim() || SAT_INFO_PORTAL_DEFAULT;
  db.fiscalConfig.satSimulatorUrl = String(db.fiscalConfig.satSimulatorUrl || SAT_SIMULATOR_DEFAULT).trim() || SAT_SIMULATOR_DEFAULT;
  db.fiscalConfig.satCancellationUrl = String(db.fiscalConfig.satCancellationUrl || SAT_CANCELLATION_DEFAULT).trim() || SAT_CANCELLATION_DEFAULT;
  db.fiscalConfig.satWorkflowMode = String(db.fiscalConfig.satWorkflowMode || 'SIMULATOR').toUpperCase() === 'PORTAL' ? 'PORTAL' : 'SIMULATOR';
  db.fiscalConfig.pacMode = normalizeFactoramaEnvironment(db.fiscalConfig.factoramaEnvironment || db.fiscalConfig.pacMode).toUpperCase();
  db.fiscalConfig.factoramaEnvironment = normalizeFactoramaEnvironment(db.fiscalConfig.factoramaEnvironment || db.fiscalConfig.pacMode);
  db.fiscalConfig.factoramaApiUsername = String(db.fiscalConfig.factoramaApiUsername || '').trim();
  db.fiscalConfig.factoramaApiPassword = String(db.fiscalConfig.factoramaApiPassword || '').trim();
  db.fiscalConfig.defaultSeries = String(db.fiscalConfig.defaultSeries || 'CREG').trim() || 'CREG';
  db.fiscalConfig.defaultPaymentForm = normalizePaymentFormValue(db.fiscalConfig.defaultPaymentForm || '01');
  db.fiscalConfig.defaultPaymentMethod = normalizePaymentMethodValue(db.fiscalConfig.defaultPaymentMethod || 'PUE');
  db.fiscalConfig.defaultUseCfdi = normalizeCfdiUseValue(db.fiscalConfig.defaultUseCfdi || 'S01');
  db.fiscalConfig.defaultCancelMotive = String(db.fiscalConfig.defaultCancelMotive || '02').trim() || '02';
  changed = true;
  db.products ||= [];
  db.enterpriseExpenses = Array.isArray(db.enterpriseExpenses) ? db.enterpriseExpenses : [];
  db.procurementPurchases = Array.isArray(db.procurementPurchases) ? db.procurementPurchases : [];
  db.systemNetwork = normalizeSystemNetworkConfig(db.systemNetwork || {});
  changed = ensureRbacCollections(db) || changed;
  db.openTickets ||= [];
  db.sales ||= [];
  db.saleConflicts ||= [];
  if (!Array.isArray(db.fiscalCustomers)) {
    db.fiscalCustomers = [defaultFiscalTestCustomer()];
    changed = true;
  }
  db.invoices ||= [];
  db.fiscalDrafts = Array.isArray(db.fiscalDrafts) ? db.fiscalDrafts : [];
  db.fiscalEvents = Array.isArray(db.fiscalEvents) ? db.fiscalEvents : [];
  db.branches ||= [];
  db.branchProfile ||= defaultBranchProfile();
  db.publication ||= defaultPublicationConfig();
  db.mirrorReplica ||= defaultMirrorReplica();
  if (!db.mirrorReplicas || typeof db.mirrorReplicas !== 'object' || Array.isArray(db.mirrorReplicas)) {
    db.mirrorReplicas = {};
    changed = true;
  }
  if (!/^[0-9a-f-]{36}$/i.test(String(db.syncInstanceId || ''))) {
    db.syncInstanceId = uuidv4();
    changed = true;
  }
  db.inventoryAdjustments ||= [];
  db.counters ||= { salesSequence: 0 };
  db.terminals ||= seed().terminals;
  db.terminals = (db.terminals || []).map((terminal, index) => ({
    id: terminal.id || `terminal-${index + 1}`,
    terminalCode: String(terminal.terminalCode || `POS-${index + 1}`).trim(),
    terminalType: String(terminal.terminalType || 'POS').toUpperCase(),
    terminalName: String(terminal.terminalName || terminal.terminalCode || `POS-${index + 1}`).trim(),
    nodeId: normalizeNodeId13640(terminal.nodeId || ''),
    platform: String(terminal.platform || '').trim(),
    authStatus: String(terminal.authStatus || (terminal.authorized === true ? 'AUTHORIZED' : '')).trim().toUpperCase(),
    authorized: terminal.authorized === true,
    authorizedAt: terminal.authorizedAt || '',
    pairingTokenHash: String(terminal.pairingTokenHash || ''),
    pairingCode: String(terminal.pairingCode || ''),
    pairingCodeHash: String(terminal.pairingCodeHash || ''),
    pairingCodeCreatedAt: terminal.pairingCodeCreatedAt || '',
    pairingCodeExpiresAt: terminal.pairingCodeExpiresAt || '',
    lastSeenAt: terminal.lastSeenAt || '',
    ipAddress: normalizeLanIp(terminal.ipAddress || ''),
    lastIp: normalizeLanIp(terminal.lastIp || ''),
    fixedIp: normalizeLanIp(terminal.fixedIp || ''),
    connectionMode: terminal.fixedIp ? 'FIXED_IP' : (terminal.connectionMode || 'AUTO'),
    enabled: terminal.enabled !== false,
    ipMatches: terminal.fixedIp && terminal.lastIp ? normalizeLanIp(terminal.fixedIp) === normalizeLanIp(terminal.lastIp) : true,
    // v13.6.43 — CRÍTICO: conservar los permisos por terminal durante cada
    // normalización de ensureDb(). En v13.6.42 esta reconstrucción omitía estos
    // campos; como ensureDb() se ejecuta en cada readDb(), los permisos recién
    // escritos desaparecían antes de la verificación posterior.
    viewPermissions: normalizeTerminalViewPermissions13641(
      terminal.viewPermissions || {},
      String(terminal.terminalType || 'POS').toUpperCase()
    ),
    permissionsRevision: Math.max(0, Number(terminal.permissionsRevision || 0)),
    permissionsUpdatedAt: terminal.permissionsUpdatedAt || '',
    createdAt: terminal.createdAt || new Date().toISOString(),
    updatedAt: terminal.updatedAt || new Date().toISOString()
  }));
  ensurePromotionCollections(db);
  ensureCashCollections(db);


  db.users = db.users.map((u) => {
    const migrated = {
      active: u.active !== false,
      employeeNumber: u.employeeNumber || (u.role === 'ADMIN' ? 'ADMIN-01' : ''),
      createdAt: u.createdAt || new Date().toISOString(),
      updatedAt: u.updatedAt || new Date().toISOString(),
      ...u
    };
    if (
      migrated.active !== u.active ||
      migrated.employeeNumber !== u.employeeNumber ||
      migrated.createdAt !== u.createdAt ||
      migrated.updatedAt !== u.updatedAt
    ) {
      changed = true;
    }
    return migrated;
  });

  db.branches = db.branches.map((branch) => {
    const migrated = {
      id: branch.id || uuidv4(),
      name: String(branch.name || '').trim(),
      code: String(branch.code || '').trim(),
      remoteUrl: String(branch.remoteUrl || '').trim(),
      authType: branch.authType || (String(branch.apiKey || '').trim() ? 'token' : 'credentials'),
      apiKey: String(branch.apiKey || '').trim(),
      remoteAdminUsername: String(branch.remoteAdminUsername || '').trim(),
      remoteAdminPassword: String(branch.remoteAdminPassword || '').trim(),
      mode: branch.mode || 'render',
      active: branch.active !== false,
      lastStatus: branch.lastStatus || 'PENDING',
      lastCheckedAt: branch.lastCheckedAt || '',
      lastError: branch.lastError || '',
      lastRemoteSummary: branch.lastRemoteSummary || null,
      createdAt: branch.createdAt || new Date().toISOString(),
      updatedAt: branch.updatedAt || new Date().toISOString()
    };
    if (
      migrated.id !== branch.id ||
      migrated.name !== branch.name ||
      migrated.code !== branch.code ||
      migrated.remoteUrl !== branch.remoteUrl ||
      migrated.authType !== branch.authType ||
      migrated.apiKey !== branch.apiKey ||
      migrated.remoteAdminUsername !== branch.remoteAdminUsername ||
      migrated.remoteAdminPassword !== branch.remoteAdminPassword ||
      migrated.mode !== branch.mode ||
      migrated.active !== branch.active ||
      migrated.lastStatus !== branch.lastStatus ||
      migrated.lastCheckedAt !== branch.lastCheckedAt ||
      migrated.lastError !== branch.lastError ||
      JSON.stringify(migrated.lastRemoteSummary || null) !== JSON.stringify(branch.lastRemoteSummary || null) ||
      migrated.createdAt !== branch.createdAt ||
      migrated.updatedAt !== branch.updatedAt
    ) {
      changed = true;
    }
    return migrated;
  });


  const migratedBranchProfile = {
    ...defaultBranchProfile(),
    ...(db.branchProfile || {})
  };
  migratedBranchProfile.name = String(migratedBranchProfile.name || ADMIN_NAME).trim() || ADMIN_NAME;
  migratedBranchProfile.code = String(migratedBranchProfile.code || 'ADMIN-01').trim() || 'ADMIN-01';
  migratedBranchProfile.publicUrl = String(migratedBranchProfile.publicUrl || '').trim();
  migratedBranchProfile.mode = String(migratedBranchProfile.mode || 'LOCAL').trim() || 'LOCAL';
  migratedBranchProfile.description = String(migratedBranchProfile.description || '').trim();
  migratedBranchProfile.updatedAt = migratedBranchProfile.updatedAt || new Date().toISOString();
  if (JSON.stringify(db.branchProfile || {}) !== JSON.stringify(migratedBranchProfile)) changed = true;
  db.branchProfile = migratedBranchProfile;

  const migratedPublication = {
    ...defaultPublicationConfig(),
    ...(db.publication || {})
  };
  migratedPublication.enabled = migratedPublication.enabled === true;
  migratedPublication.mirrorUrl = String(migratedPublication.mirrorUrl || '').trim().replace(/\/$/, '');
  migratedPublication.authType = migratedPublication.authType === 'token' ? 'token' : 'credentials';
  migratedPublication.apiKey = String(migratedPublication.apiKey || '').trim();
  migratedPublication.remoteAdminUsername = String(migratedPublication.remoteAdminUsername || 'admin').trim() || 'admin';
  migratedPublication.remoteAdminPassword = String(migratedPublication.remoteAdminPassword || '').trim();
  migratedPublication.autoSyncEnabled = migratedPublication.autoSyncEnabled === true;
  migratedPublication.syncIntervalSec = Math.max(15, Math.min(600, safeNumber(migratedPublication.syncIntervalSec, 60)));
  migratedPublication.lastSyncAt = String(migratedPublication.lastSyncAt || '').trim();
  migratedPublication.lastSyncStatus = String(migratedPublication.lastSyncStatus || 'IDLE').trim() || 'IDLE';
  migratedPublication.lastSyncError = String(migratedPublication.lastSyncError || '').trim();
  migratedPublication.lastCommandPollAt = String(migratedPublication.lastCommandPollAt || '').trim();
  migratedPublication.lastCommandApplyAt = String(migratedPublication.lastCommandApplyAt || '').trim();
  migratedPublication.updatedAt = migratedPublication.updatedAt || new Date().toISOString();
  if (JSON.stringify(db.publication || {}) !== JSON.stringify(migratedPublication)) changed = true;
  db.publication = migratedPublication;

  const migratedMirrorReplica = {
    ...defaultMirrorReplica(),
    ...(db.mirrorReplica || {})
  };
  migratedMirrorReplica.snapshot = migratedMirrorReplica.snapshot || null;
  migratedMirrorReplica.lastSnapshotAt = String(migratedMirrorReplica.lastSnapshotAt || '').trim();
  migratedMirrorReplica.lastSnapshotSource = String(migratedMirrorReplica.lastSnapshotSource || '').trim();
  migratedMirrorReplica.commands = Array.isArray(migratedMirrorReplica.commands) ? migratedMirrorReplica.commands : [];
  migratedMirrorReplica.commandSequence = safeNumber(migratedMirrorReplica.commandSequence, 0);
  migratedMirrorReplica.updatedAt = migratedMirrorReplica.updatedAt || new Date().toISOString();
  if (JSON.stringify(db.mirrorReplica || {}) !== JSON.stringify(migratedMirrorReplica)) changed = true;
  db.mirrorReplica = migratedMirrorReplica;


  db.products = db.products.map((p) => {
    const migrated = {
      ...p,
      category: String(p.category || 'General').trim() || 'General',
      costPrice: safeNumber(p.costPrice, 0),
      retailMarginPct: safeNumber(p.retailMarginPct, 0),
      retailPrice: round2(p.retailPrice ?? p.price ?? 0),
      wholesalePrice: round2(p.wholesalePrice ?? p.retailPrice ?? p.price ?? 0),
      price: round2(p.retailPrice ?? p.price ?? 0),
      lowStockThreshold: safeNumber(p.lowStockThreshold, 0),
      fiscalProductCode: String(p.fiscalProductCode || (String(p.unit || '').toUpperCase() === 'KG' ? '50131802' : '50202306')).trim(),
      fiscalUnitCode: String(p.fiscalUnitCode || (String(p.unit || '').toUpperCase() === 'KG' ? 'KGM' : 'H87')).trim(),
      fiscalUnitName: String(p.fiscalUnitName || (String(p.unit || '').toUpperCase() === 'KG' ? 'Kilogramo' : 'Pieza')).trim(),
      fiscalDescription: String(p.fiscalDescription || `${p.name || 'Producto'} de prueba CFDI 4.0`).trim(),
      fiscalTaxObject: String(p.fiscalTaxObject || '02').trim(),
      fiscalVatRate: safeNumber(p.fiscalVatRate, 0.16),
      active: p.active !== false,
      createdAt: p.createdAt || new Date().toISOString(),
      updatedAt: p.updatedAt || new Date().toISOString()
    };

    if (
      migrated.category !== p.category ||
      migrated.costPrice !== p.costPrice ||
      migrated.retailMarginPct !== p.retailMarginPct ||
      migrated.retailPrice !== p.retailPrice ||
      migrated.wholesalePrice !== p.wholesalePrice ||
      migrated.lowStockThreshold !== p.lowStockThreshold ||
      migrated.price !== p.price
    ) {
      changed = true;
    }

    return migrated;
  });

  db.openTickets = db.openTickets.map((t, index) => {
    const migrated = {
      id: t.id || uuidv4(),
      folio: t.folio || `TK-${String(index + 1).padStart(4, '0')}`,
      terminalCode: t.terminalCode || 'POS-01',
      operatorName: t.operatorName || 'Operador',
      customerName: String(t.customerName || '').trim(),
      customerReference: String(t.customerReference || '').trim(),
      takenBy: t.takenBy || '',
      status: t.status || 'OPEN',
      items: normalizeTicketItems(t.items || []),
      subtotal: round2(t.subtotal ?? sumTicket(t.items || [])),
      createdAt: t.createdAt || new Date().toISOString(),
      updatedAt: t.updatedAt || new Date().toISOString()
    };
    if (
      migrated.id !== t.id ||
      migrated.folio !== t.folio ||
      migrated.terminalCode !== t.terminalCode ||
      migrated.operatorName !== t.operatorName ||
      migrated.customerName !== t.customerName ||
      migrated.customerReference !== t.customerReference ||
      migrated.takenBy !== t.takenBy ||
      migrated.status !== t.status ||
      migrated.subtotal !== t.subtotal
    ) {
      changed = true;
    }
    return migrated;
  });

  if (ensureSellerSessions(db)) changed = true;

  db.sales = db.sales.map((sale, index) => {
    const migrated = {
      ...sale,
      id: sale.id || uuidv4(),
      folio: sale.folio || `VTA-${String(index + 1).padStart(6, '0')}`,
      terminalCode: sale.terminalCode || 'POS-01',
      createdAt: sale.createdAt || new Date().toISOString(),
      clientTxnId: sale.clientTxnId || uuidv4(),
      items: normalizeTicketItems(sale.items || []),
      total: round2(sale.total ?? sumTicket(sale.items || [])),
      operatorId: sale.operatorId || '',
      operatorName: sale.operatorName || '',
      employeeNumber: sale.employeeNumber || '',
      customerName: String(sale.customerName || '').trim(),
      customerReference: String(sale.customerReference || '').trim(),
      ticketId: sale.ticketId || '',
      saleType: sale.saleType || sale.transactionType || '',
      transactionType: sale.transactionType || sale.saleType || '',
      paymentMethod: sale.paymentMethod || (sale.affectsCash === false ? 'internal' : 'cash'),
      payments: Array.isArray(sale.payments) ? sale.payments : [],
      amountReceived: round2(sale.amountReceived || 0),
      changeDue: round2(sale.changeDue || 0),
      subtotalBeforeDiscount: round2(sale.subtotalBeforeDiscount ?? sale.total ?? sumTicket(sale.items || [])),
      discountTotal: round2(sale.discountTotal || 0),
      inventoryValue: round2(sale.inventoryValue || (sale.affectsCash === false ? (sale.total || 0) : 0)),
      affectsCash: sale.affectsCash === false ? false : true,
      affectsSalesTotals: sale.affectsSalesTotals === false ? false : true,
      cashEntryId: sale.cashEntryId || '',
      notes: String(sale.notes || '').trim()
    };
    if (
      migrated.id !== sale.id ||
      migrated.folio !== sale.folio ||
      migrated.terminalCode !== sale.terminalCode ||
      migrated.clientTxnId !== sale.clientTxnId ||
      migrated.total !== sale.total ||
      migrated.operatorId !== sale.operatorId ||
      migrated.operatorName !== sale.operatorName ||
      migrated.employeeNumber !== sale.employeeNumber ||
      migrated.customerName !== sale.customerName ||
      migrated.customerReference !== sale.customerReference ||
      migrated.ticketId !== sale.ticketId
    ) {
      changed = true;
    }
    return migrated;
  });

  for (const sale of db.sales) {
    if (!isOwnerConsumptionSale(sale)) {
      const beforeEntryCount = (db.cashEntries || []).length;
      const entry = ensureSaleCashEntry(db, sale);
      if (entry && (db.cashEntries || []).length !== beforeEntryCount) changed = true;
      if (entry && sale.cashEntryId !== entry.id) { sale.cashEntryId = entry.id; changed = true; }
    }
  }


  db.fiscalCustomers = (db.fiscalCustomers || []).map((customer, index) => ({
    id: customer.id || `fiscal-${index + 1}`,
    name: String(customer.name || '').trim().toUpperCase() || 'PUBLICO EN GENERAL',
    rfc: String(customer.rfc || 'XAXX010101000').trim().toUpperCase(),
    regimenFiscal: normalizeFiscalRegimeValue(customer.regimenFiscal || '', customer.rfc === 'XAXX010101000' ? '616' : '605'),
    zip: String(customer.zip || db.fiscalConfig.expeditionZip || '55750').trim(),
    email: String(customer.email || '').trim(),
    useCfdi: normalizeCfdiUseValue(customer.useCfdi || customer.lastUseCfdi || 'S01'),
    lastUseCfdi: normalizeCfdiUseValue(customer.lastUseCfdi || customer.useCfdi || 'S01'),
    lastPaymentMethod: normalizePaymentMethodValue(customer.lastPaymentMethod || db.fiscalConfig.defaultPaymentMethod || 'PUE'),
    lastPaymentForm: normalizePaymentFormValue(customer.lastPaymentForm || db.fiscalConfig.defaultPaymentForm || '01'),
    lastInvoicedAt: customer.lastInvoicedAt || '',
    createdAt: customer.createdAt || new Date().toISOString(),
    updatedAt: customer.updatedAt || new Date().toISOString(),
    isDefault: customer.isDefault === true || customer.rfc === 'XAXX010101000'
  }));

  db.invoices = (db.invoices || []).map((invoice) => ({
    ...invoice,
    provider: invoice.provider || 'FACTURAMA',
    apiMode: normalizeFactoramaEnvironment(invoice.apiMode || db.fiscalConfig.factoramaEnvironment || 'sandbox'),
    status: String(invoice.status || 'DRAFT').toUpperCase(),
    paymentForm: normalizePaymentFormValue(invoice.paymentForm || db.fiscalConfig.defaultPaymentForm || '01'),
    paymentMethod: normalizePaymentMethodValue(invoice.paymentMethod || db.fiscalConfig.defaultPaymentMethod || 'PUE'),
    useCfdi: normalizeCfdiUseValue(invoice.useCfdi || db.fiscalConfig.defaultUseCfdi || 'S01'),
    createdAt: invoice.createdAt || new Date().toISOString(),
    total: round2(safeNumber(invoice.total)),
    cancellation: invoice.cancellation || null
  }));

  db.inventoryAdjustments = (db.inventoryAdjustments || []).map((entry) => ({
    ...entry,
    id: entry.id || uuidv4(),
    type: entry.type || entry.operation || 'inventory_adjustment',
    operation: entry.operation || (safeNumber(entry.delta) < 0 ? 'subtract' : 'add'),
    productId: entry.productId || '',
    productName: entry.productName || '',
    sku: entry.sku || '',
    previousStock: roundQuantity(safeNumber(entry.previousStock)),
    newStock: roundQuantity(safeNumber(entry.newStock)),
    delta: roundQuantity(safeNumber(entry.delta)),
    unit: entry.unit || 'PZA',
    reason: String(entry.reason || '').trim(),
    sourceType: entry.sourceType || 'LOCAL_ADMIN',
    actorName: entry.actorName || '',
    actorUsername: entry.actorUsername || '',
    actorBranch: entry.actorBranch || '',
    deletedAt: entry.deletedAt || '',
    deletedBy: entry.deletedBy || '',
    createdAt: entry.createdAt || new Date().toISOString()
  }));

  const derivedSalesSequence = db.sales.reduce((max, sale, index) => {
    const match = String(sale.folio || '').match(/(\d+)$/);
    const seq = match ? Number(match[1]) : index + 1;
    return Math.max(max, seq);
  }, 0);

  if (safeNumber(db.counters.salesSequence, 0) !== derivedSalesSequence) {
    db.counters.salesSequence = Math.max(safeNumber(db.counters.salesSequence, 0), derivedSalesSequence);
    changed = true;
  }

  if (changed || isSecretMigrationPending()) {
    writeDb(db);
  }
  DB_SCHEMA_SIGNATURE_13650 = dbFileSignature13650();
}

function readDb() {
  if (RENDER_MIRROR_ONLY) {
    if (!RENDER_RELAY_AUTH_DB) {
      RENDER_RELAY_AUTH_DB = {
        security: {
          remoteAccess: {
            apiKey: RENDER_REMOTE_API_KEY,
            allowAdminCredentials: false,
            permissions: { readSummary: true, readInventory: true, writeInventory: false, syncBridge: false }
          }
        },
        mirrorReplicas: {},
        mirrorReplica: null
      };
    }
    return RENDER_RELAY_AUTH_DB;
  }
  if (dbFileSignature13650() !== DB_SCHEMA_SIGNATURE_13650) {
    ensureDb();
    DB_SCHEMA_SIGNATURE_13650 = dbFileSignature13650();
  }
  return readDbFileWithAutoRecovery13633();
}

function migrateLegacyCredentialBackups13647(protectedKey) {
  if (!protectedKey || !fs.existsSync(BACKUPS_DIR)) return;
  let names = [];
  try { names = fs.readdirSync(BACKUPS_DIR).filter((name) => name.toLowerCase().endsWith('.json')); } catch (_) { return; }
  for (const name of names) {
    const file = path.join(BACKUPS_DIR, name);
    try {
      const backup = JSON.parse(fs.readFileSync(file, 'utf8'));
      if (!backup || typeof backup !== 'object' || Array.isArray(backup)) continue;
      backup.security ||= {};
      backup.security.secretStore ||= { version: 1, protectedKey };
      if (!backup.security.secretStore.protectedKey) backup.security.secretStore.protectedKey = protectedKey;
      encryptDatabaseSecrets(backup);
      const temp = `${file}.secure-${process.pid}-${Date.now()}`;
      fs.writeFileSync(temp, JSON.stringify(backup, null, 2), 'utf8');
      fs.renameSync(temp, file);
    } catch (_) {
      // Preserve the backup untouched if it is not a valid supported database snapshot.
    }
  }
}

function writeDb(db) {
  if (RENDER_MIRROR_ONLY) return false;
  // Escritura atómica validada: evita JSON incompleto si Windows, antivirus o apagón interrumpe el guardado.
  const persistedDb = JSON.parse(JSON.stringify(db));
  encryptDatabaseSecrets(persistedDb);
  const payload = JSON.stringify(persistedDb, null, 2);
  JSON.parse(payload);
  const tmpFile = `${DB_FILE}.tmp-${process.pid}-${Date.now()}`;
  fs.writeFileSync(tmpFile, payload, 'utf8');
  fs.renameSync(tmpFile, DB_FILE);
  DB_SCHEMA_SIGNATURE_13650 = dbFileSignature13650();
  if (isSecretMigrationPending()) {
    migrateLegacyCredentialBackups13647(persistedDb.security?.secretStore?.protectedKey || '');
    clearSecretMigrationPending();
  }
  // Respaldo diferido y agrupado: no copia la base en cada venta; evita lentitud en caja.
  scheduleBackup13633('auto');
}

function verifyAdminPassword(db, password) {
  const admin = (db.users || []).find((u) => u.role === 'ADMIN');
  return !!(admin && password && admin.password === password);
}

function getRemotePermissions(db) {
  return {
    readSummary: db.security?.remoteAccess?.permissions?.readSummary !== false,
    readInventory: db.security?.remoteAccess?.permissions?.readInventory !== false,
    writeInventory: db.security?.remoteAccess?.permissions?.writeInventory !== false,
    syncBridge: db.security?.remoteAccess?.permissions?.syncBridge !== false
  };
}

function hasValidRemoteToken(db, token) {
  const supplied = Buffer.from(String(token || '').trim());
  if (!supplied.length) return false;
  const configuredKeys = RENDER_MIRROR_ONLY
    ? [RENDER_REMOTE_API_KEY]
    : [String(db.security?.remoteAccess?.apiKey || '').trim(), RENDER_REMOTE_API_KEY];
  return configuredKeys.some((value) => {
    const configured = Buffer.from(String(value || '').trim());
    return configured.length === supplied.length && configured.length > 0 && crypto.timingSafeEqual(configured, supplied);
  });
}

function hasValidRemoteAdminCredentials(db, username, password) {
  if (RENDER_MIRROR_ONLY) return false;
  if (db.security?.remoteAccess?.allowAdminCredentials === false) return false;
  const admin = (db.users || []).find((u) => u.role === 'ADMIN');
  if (!admin) return false;
  return !!(username && password && String(username).trim() === String(admin.username || '').trim() && String(password).trim() === String(admin.password || '').trim());
}

function authorizeRemoteRequest(req, res, db, permissionKey = 'readSummary') {
  const token = req.headers['x-remote-key'];
  const username = req.headers['x-remote-username'];
  const password = req.headers['x-remote-password'];
  const hasToken = hasValidRemoteToken(db, token);
  const hasCredentials = hasValidRemoteAdminCredentials(db, username, password);

  if (!hasToken && !hasCredentials) {
    res.status(401).json({ ok: false, message: 'Acceso remoto no autorizado' });
    return null;
  }

  const permissions = getRemotePermissions(db);
  if (permissions[permissionKey] === false) {
    res.status(403).json({ ok: false, message: 'La sucursal remota no permite esta operación' });
    return null;
  }

  return {
    authMode: hasToken ? 'token' : 'credentials',
    permissions,
    actorName: hasToken ? 'Integración remota' : String(username || '').trim(),
    actorUsername: hasToken ? 'remote-token' : String(username || '').trim()
  };
}

function sanitizeInventoryProduct(product) {
  return {
    id: product.id,
    name: product.name,
    category: product.category || 'General',
    sku: product.sku || '',
    stock: roundQuantity(safeNumber(product.stock)),
    unit: product.unit || 'PZA',
    allowsFraction: !!product.allowsFraction,
    lowStockThreshold: round2(safeNumber(product.lowStockThreshold)),
    retailPrice: round2(product.retailPrice ?? product.price ?? 0),
    wholesalePrice: round2(product.wholesalePrice ?? product.retailPrice ?? product.price ?? 0),
    costPrice: round2(safeNumber(product.costPrice)),
    imageUrl: product.imageUrl || '',
    active: product.active !== false,
    updatedAt: product.updatedAt || '',
    alertStatus: getInventoryAlertStatus(product),
    nearStockThreshold: getNearStockThreshold(product)
  };
}

function buildRemoteInventory(db, filters = {}) {
  const allProducts = (db.products || []).filter((product) => product.active !== false).map(sanitizeInventoryProduct);
  return buildInventoryPayloadFromRows(allProducts, filters);
}

function applyInventoryAdjustment(db, payload = {}) {
  const product = (db.products || []).find((row) => row.id === payload.productId);
  if (!product) {
    throw new Error('Producto no encontrado');
  }

  const operation = payload.operation === 'set' ? 'set' : (payload.operation === 'subtract' ? 'subtract' : 'add');
  const quantity = roundQuantity(safeNumber(payload.quantity));
  const previousStock = roundQuantity(safeNumber(product.stock));
  let newStock = previousStock;

  if (operation === 'set') {
    newStock = roundQuantity(safeNumber(payload.targetStock, previousStock));
  } else if (operation === 'subtract') {
    if (quantity <= 0) throw new Error('La cantidad a descontar debe ser mayor a 0');
    newStock = roundQuantity(previousStock - quantity);
  } else {
    if (quantity <= 0) throw new Error('La cantidad a agregar debe ser mayor a 0');
    newStock = roundQuantity(previousStock + quantity);
  }

  if (newStock < 0) {
    throw new Error(`El ajuste dejaría stock negativo en ${product.name}`);
  }

  product.stock = newStock;
  product.updatedAt = new Date().toISOString();

  const movement = {
    id: uuidv4(),
    type: payload.type || 'inventory_adjustment',
    operation,
    productId: product.id,
    productName: product.name,
    sku: product.sku || '',
    previousStock,
    newStock,
    delta: roundQuantity(newStock - previousStock),
    unit: product.unit || 'PZA',
    reason: String(payload.reason || '').trim() || 'Ajuste manual de inventario',
    sourceType: payload.sourceType || 'LOCAL_ADMIN',
    actorName: payload.actorName || '',
    actorUsername: payload.actorUsername || '',
    actorBranch: payload.actorBranch || '',
    createdAt: new Date().toISOString()
  };

  db.inventoryAdjustments ||= [];
  db.inventoryAdjustments.unshift(movement);
  db.inventoryAdjustments = db.inventoryAdjustments.slice(0, 5000);

  return { product, movement };
}

function buildRemoteSummary(db, options = {}) {
  const profile = options.profile || getBranchProfile(db);
  const serviceMode = options.serviceMode || 'LOCAL_ADMIN';
  const cashSales = (db.sales || []).filter((sale) => !isOwnerConsumptionSale(sale));
  const totalSales = round2(cashSales.reduce((a, s) => a + safeNumber(s.total), 0));
  const lowStockProducts = (db.products || []).filter((product) => getInventoryAlertStatus(product) === 'missing').slice(0, 8).map((product) => ({
    id: product.id,
    name: product.name,
    sku: product.sku || '',
    stock: safeNumber(product.stock),
    unit: product.unit || 'PZA',
    threshold: safeNumber(product.lowStockThreshold),
    alertStatus: 'missing'
  }));
  const nearStockProducts = (db.products || []).filter((product) => getInventoryAlertStatus(product) === 'near').slice(0, 8).map((product) => ({
    id: product.id,
    name: product.name,
    sku: product.sku || '',
    stock: safeNumber(product.stock),
    unit: product.unit || 'PZA',
    threshold: safeNumber(product.lowStockThreshold),
    nearThreshold: getNearStockThreshold(product),
    alertStatus: 'near'
  }));
  const recentSales = (db.sales || []).slice().sort((a, b) => new Date(b.createdAt || 0).getTime() - new Date(a.createdAt || 0).getTime()).slice(0, 8).map((sale) => ({
    id: sale.id,
    folio: sale.folio,
    total: round2(safeNumber(sale.total)),
    operatorName: sale.operatorName || '-',
    customerName: sale.customerName || '',
    createdAt: sale.createdAt || '',
    itemsCount: Array.isArray(sale.items) ? sale.items.length : 0
  }));
  const recentInventoryAdjustments = (db.inventoryAdjustments || []).slice(0, 8).map((entry) => ({
    id: entry.id,
    productName: entry.productName || '-',
    delta: round2(safeNumber(entry.delta)),
    newStock: round2(safeNumber(entry.newStock)),
    reason: entry.reason || '',
    actorName: entry.actorName || entry.actorUsername || '-',
    createdAt: entry.createdAt || ''
  }));

  return {
    ok: true,
    branchProfile: profile,
    remote: {
      adminName: ADMIN_NAME,
      hostname: os.hostname(),
      port: PORT,
      service: 'embedded-admin-backend',
      serviceMode,
      publicUrl: profile.publicUrl || '',
      permissions: getRemotePermissions(db)
    },
    metrics: {
      products: (db.products || []).length,
      stockUnits: round2((db.products || []).reduce((a, p) => a + safeNumber(p.stock), 0)),
      salesCount: (db.sales || []).length,
      totalSales,
      employees: getActiveEmployees(db).length,
      lowStockCount: (db.products || []).filter((product) => getInventoryAlertStatus(product) === 'missing').length,
      nearStockCount: (db.products || []).filter((product) => getInventoryAlertStatus(product) === 'near').length,
      attentionCount: (db.products || []).filter((product) => getInventoryAlertStatus(product) !== 'normal').length,
      openTickets: (db.openTickets || []).length,
      invoices: (db.invoices || []).length,
      inventoryAdjustments: (db.inventoryAdjustments || []).length
    },
    recentSales,
    lowStockProducts,
    nearStockProducts,
    recentInventoryAdjustments,
    generatedAt: new Date().toISOString()
  };
}

function buildRemoteHeaders(branch) {
  const headers = {};
  if ((branch.authType || 'credentials') === 'token') {
    if (!String(branch.apiKey || '').trim()) {
      throw new Error('Falta el token/API key de la sucursal remota');
    }
    headers['x-remote-key'] = String(branch.apiKey || '').trim();
  } else {
    if (!String(branch.remoteAdminUsername || '').trim() || !String(branch.remoteAdminPassword || '').trim()) {
      throw new Error('Faltan las credenciales del administrador remoto');
    }
    headers['x-remote-username'] = String(branch.remoteAdminUsername || '').trim();
    headers['x-remote-password'] = String(branch.remoteAdminPassword || '').trim();
  }
  const branchCode = normalizeMirrorBranchCode(branch.branchCode || branch.code || '');
  if (branchCode) headers['x-remote-branch-code'] = branchCode;
  return headers;
}

function isPrivateOrLocalRemoteHost(hostname) {
  const host = String(hostname || '').toLowerCase().replace(/^\[|\]$/g, '');
  if (!host || host === 'localhost' || host.endsWith('.localhost') || host.endsWith('.local') || host.endsWith('.lan')) return true;
  const ipVersion = net.isIP(host);
  if (ipVersion === 4) {
    const octets = host.split('.').map(Number);
    return octets[0] === 10
      || octets[0] === 127
      || (octets[0] === 169 && octets[1] === 254)
      || (octets[0] === 172 && octets[1] >= 16 && octets[1] <= 31)
      || (octets[0] === 192 && octets[1] === 168);
  }
  if (ipVersion === 6) {
    return host === '::1'
      || host.startsWith('fc')
      || host.startsWith('fd')
      || /^fe[89ab]/.test(host)
      || (host.startsWith('::ffff:') && isPrivateOrLocalRemoteHost(host.slice(7)));
  }
  return false;
}

function validateRemoteBranchUrl(remoteUrl) {
  let url;
  try {
    url = new URL(String(remoteUrl || '').trim());
  } catch (_) {
    throw new Error('La URL remota de la sucursal no es válida.');
  }
  if (!['http:', 'https:'].includes(url.protocol) || url.username || url.password) {
    throw new Error('La URL remota debe usar HTTP/HTTPS y no incluir credenciales en la dirección.');
  }
  if (url.protocol === 'http:' && !isPrivateOrLocalRemoteHost(url.hostname)) {
    throw new Error('La sucursal remota está fuera de la red local. Configura una URL HTTPS o conéctate mediante una VPN.');
  }
  return url;
}

async function fetchRemoteBranch(branch, endpointPath, options = {}) {
  const base = String(branch.remoteUrl || '').trim().replace(/\/$/, '');
  if (!base) {
    throw new Error('La URL remota de la sucursal está vacía');
  }
  validateRemoteBranchUrl(base);

  const headers = {
    ...buildRemoteHeaders(branch),
    ...(options.headers || {})
  };

  if (options.body && !(options.body instanceof Buffer)) {
    headers['Content-Type'] = 'application/json';
  }

  const controller = new AbortController();
  const isRenderFreeService = new URL(base).hostname.toLowerCase().endsWith('.onrender.com');
  // Render Free may need about a minute to wake after idling. Allow that cold start
  // for remote branch calls while keeping the normal local-network timeout short.
  const timeout = setTimeout(() => controller.abort(), isRenderFreeService ? 120000 : 10000);

  try {
    const url = new URL(`${base}${endpointPath}`);
    if ((options.method || 'GET').toUpperCase() === 'GET') {
      url.searchParams.set('_ts', String(Date.now()));
    }
    const response = await fetch(url.toString(), {
      method: options.method || 'GET',
      headers,
      body: options.body,
      signal: controller.signal,
      cache: 'no-store'
    });
    const data = await response.json().catch(() => ({}));
    if (!response.ok) {
      const err = new Error(data.message || `La sucursal respondió con HTTP ${response.status}`);
      err.status = response.status;
      err.payload = data;
      err.endpointPath = endpointPath;
      throw err;
    }
    return data;
  } finally {
    clearTimeout(timeout);
  }
}

function normalizeRemoteStockAlertsPayload(payload = {}, source = 'stock-alerts') {
  const rawMissing = Array.isArray(payload.missing)
    ? payload.missing
    : (Array.isArray(payload.lowStockProducts)
      ? payload.lowStockProducts
      : (Array.isArray(payload.products) ? payload.products.filter((row) => row.alertStatus === 'missing') : []));
  const rawNear = Array.isArray(payload.near)
    ? payload.near
    : (Array.isArray(payload.nearStockProducts)
      ? payload.nearStockProducts
      : (Array.isArray(payload.products) ? payload.products.filter((row) => row.alertStatus === 'near') : []));

  const mapRow = (row = {}, forcedStatus = '') => ({
    id: row.id,
    name: row.name || row.productName || '-',
    sku: row.sku || '',
    stock: roundQuantity(safeNumber(row.stock)),
    unit: row.unit || 'PZA',
    lowStockThreshold: round2(safeNumber(row.lowStockThreshold, row.threshold)),
    allowsFraction: row.allowsFraction !== false && ['KG', 'LT'].includes(String(row.unit || '').toUpperCase()),
    alertStatus: forcedStatus || row.alertStatus || 'normal'
  });

  const missing = rawMissing.map((row) => mapRow(row, 'missing'));
  const near = rawNear.map((row) => mapRow(row, 'near'));

  return {
    ok: true,
    missing,
    near,
    totals: {
      missingCount: payload.totals?.missingCount ?? missing.length,
      nearStockCount: payload.totals?.nearStockCount ?? near.length
    },
    generatedAt: payload.generatedAt || new Date().toISOString(),
    source
  };
}

function normalizeRemotePublicationInfoPayload(payload = {}, source = 'publication-info') {
  const branchProfile = payload.branchProfile || {};
  const remote = payload.remote || {};
  return {
    ok: true,
    branchProfile,
    remote: {
      adminName: remote.adminName || ADMIN_NAME,
      hostname: remote.hostname || os.hostname(),
      port: remote.port || PORT,
      service: remote.service || 'embedded-admin-backend',
      serviceMode: remote.serviceMode || 'LOCAL_ADMIN',
      publicUrl: remote.publicUrl || branchProfile.publicUrl || '',
      permissions: remote.permissions || getRemotePermissions(readDb())
    },
    lastSnapshotAt: payload.lastSnapshotAt || '',
    authMode: payload.authMode || '',
    source
  };
}

async function fetchRemoteBranchWithFallbacks(branch, attempts = []) {
  let lastError = null;
  for (const attempt of attempts) {
    const config = typeof attempt === 'string' ? { path: attempt } : (attempt || {});
    try {
      const data = await fetchRemoteBranch(branch, config.path, config.options || {});
      return typeof config.transform === 'function'
        ? config.transform(data, config.path)
        : data;
    } catch (err) {
      lastError = err;
      const fallbackOn = config.fallbackOn || [];
      if (!fallbackOn.includes(err.status)) {
        throw err;
      }
    }
  }
  throw lastError || new Error('No se pudo consultar la sucursal remota');
}

async function fetchRemoteBranchSummary(branch) {
  return fetchRemoteBranch(branch, '/api/remote/summary');
}

async function fetchRemoteBranchPublicationInfo(branch) {
  return fetchRemoteBranchWithFallbacks(branch, [
    { path: '/api/remote/publication-info' },
    { path: '/api/remote/summary', fallbackOn: [404, 405], transform: (data) => normalizeRemotePublicationInfoPayload(data, 'summary-fallback') }
  ]);
}

async function fetchRemoteBranchStockAlerts(branch) {
  return fetchRemoteBranchWithFallbacks(branch, [
    { path: '/api/remote/stock-alerts', transform: (data) => normalizeRemoteStockAlertsPayload(data, 'stock-alerts') },
    { path: '/api/remote/summary', fallbackOn: [404, 405], transform: (data) => normalizeRemoteStockAlertsPayload(data, 'summary-fallback') },
    { path: '/api/remote/inventory?alertMode=attention', fallbackOn: [404, 405], transform: (data) => normalizeRemoteStockAlertsPayload(data, 'inventory-fallback') }
  ]);
}

function normalizeRemoteIp(ip = '') {
  return String(ip || '')
    .replace(/^::ffff:/, '')
    .replace(/^::1$/, '127.0.0.1')
    .trim();
}

function normalizeLanIp(value = '') { return normalizeRemoteIp(value); }
function requestLanIp(req) {
  const forwarded = String(req?.headers?.['x-forwarded-for'] || '').split(',')[0].trim();
  return normalizeLanIp(forwarded || req?.socket?.remoteAddress || req?.connection?.remoteAddress || '');
}
function isValidLanIpv4(value = '') {
  const ip = normalizeLanIp(value);
  const parts = ip.split('.');
  return parts.length === 4 && parts.every((part) => /^\d{1,3}$/.test(part) && Number(part) >= 0 && Number(part) <= 255);
}

function touchTerminal(db, payload = {}) {
  db.terminals ||= [];
  const terminalCode = String(payload.terminalCode || 'UNKNOWN').trim() || 'UNKNOWN';
  const existing = (db.terminals || []).find((t) => t.terminalCode === terminalCode);
  const now = new Date().toISOString();
  const ipAddress = normalizeRemoteIp(payload.ipAddress || payload.ip || payload.remoteAddress || payload.lastIp || '');
  const detectedIp = normalizeLanIp(payload.lastIp || payload.ip || payload.ipAddress || payload.remoteAddress || '');

  if (existing) {
    existing.lastSeenAt = now;
    existing.terminalType = payload.terminalType || existing.terminalType;
    existing.terminalName = payload.terminalName || existing.terminalName;
    if (ipAddress) existing.ipAddress = ipAddress;
    if (detectedIp) existing.lastIp = detectedIp;
    existing.fixedIp = normalizeLanIp(existing.fixedIp || '');
    existing.connectionMode = existing.fixedIp ? 'FIXED_IP' : (existing.connectionMode || 'AUTO');
    existing.enabled = existing.enabled !== false;
    existing.ipMatches = !existing.fixedIp || !detectedIp || existing.fixedIp === detectedIp;
    ensureTerminalViewPermissions13641(existing);
  } else {
    const newTerminal = {
      id: uuidv4(),
      terminalCode,
      terminalType: payload.terminalType || 'POS',
      terminalName: payload.terminalName || terminalCode,
      ipAddress,
      lastSeenAt: now,
      lastIp: detectedIp,
      fixedIp: '',
      connectionMode: 'AUTO',
      enabled: true,
      ipMatches: true
    };
    ensureTerminalViewPermissions13641(newTerminal);
    db.terminals.push(newTerminal);
  }
}


// =============================================================
// v13.6.40 — Autorización persistente de nodos mediante código
// de 6 dígitos visible únicamente en el Nodo Administrador.
// =============================================================
const PAIRING_CODE_TTL_MS_13640 = 15 * 60 * 1000;

function hashTerminalSecret13640(value = '') {
  return crypto.createHash('sha256').update(String(value || ''), 'utf8').digest('hex');
}

function normalizeNodeId13640(value = '') {
  return String(value || '').trim().slice(0, 180);
}

function normalizeTerminalCode13640(value = '') {
  return String(value || '').trim().slice(0, 80);
}

function findTerminalForNode13640(db, nodeId = '', terminalCode = '') {
  const cleanNodeId = normalizeNodeId13640(nodeId);
  const cleanCode = normalizeTerminalCode13640(terminalCode);
  if (cleanNodeId) {
    const byNode = (db.terminals || []).find((t) => normalizeNodeId13640(t.nodeId) === cleanNodeId);
    if (byNode) return byNode;
  }
  if (cleanCode) {
    return (db.terminals || []).find((t) => normalizeTerminalCode13640(t.terminalCode) === cleanCode) || null;
  }
  return null;
}

function issuePairingCode13640(terminal, force = false) {
  const now = Date.now();
  const expiresMs = Date.parse(terminal.pairingCodeExpiresAt || '') || 0;
  if (!force && terminal.pairingCode && expiresMs > now + 30_000) {
    return terminal.pairingCode;
  }
  const code = String(crypto.randomInt(0, 1_000_000)).padStart(6, '0');
  terminal.pairingCode = code;
  terminal.pairingCodeHash = hashTerminalSecret13640(code);
  terminal.pairingCodeCreatedAt = new Date().toISOString();
  terminal.pairingCodeExpiresAt = new Date(now + PAIRING_CODE_TTL_MS_13640).toISOString();
  terminal.authStatus = 'PENDING';
  terminal.authorized = false;
  terminal.updatedAt = new Date().toISOString();
  return code;
}

function ensurePendingTerminal13640(db, payload = {}, req = null) {
  db.terminals ||= [];
  const nodeId = normalizeNodeId13640(payload.nodeId);
  const terminalCode = normalizeTerminalCode13640(payload.terminalCode || 'POS-01') || 'POS-01';
  const terminalName = String(payload.terminalName || terminalCode).trim().slice(0, 120) || terminalCode;
  const terminalType = String(payload.terminalType || 'POS').trim().toUpperCase() || 'POS';
  let terminal = findTerminalForNode13640(db, nodeId, terminalCode);

  if (!terminal) {
    terminal = {
      id: uuidv4(),
      terminalCode,
      terminalType,
      terminalName,
      nodeId,
      platform: String(payload.platform || '').trim().slice(0, 80),
      ipAddress: req ? requestLanIp(req) : '',
      lastIp: req ? requestLanIp(req) : '',
      fixedIp: '',
      connectionMode: 'AUTO',
      enabled: true,
      ipMatches: true,
      createdAt: new Date().toISOString(),
      lastSeenAt: new Date().toISOString(),
      authStatus: 'PENDING',
      authorized: false
    };
    db.terminals.push(terminal);
  } else {
    terminal.terminalCode = terminalCode || terminal.terminalCode;
    terminal.terminalType = terminalType || terminal.terminalType;
    terminal.terminalName = terminalName || terminal.terminalName;
    if (nodeId) terminal.nodeId = nodeId;
    if (payload.platform) terminal.platform = String(payload.platform).trim().slice(0, 80);
    if (req) {
      const detected = requestLanIp(req);
      if (detected) {
        terminal.ipAddress = detected;
        terminal.lastIp = detected;
      }
    }
    terminal.lastSeenAt = new Date().toISOString();
    terminal.enabled = terminal.enabled !== false;
  }

  ensureTerminalViewPermissions13641(terminal);
  if (terminal.authStatus !== 'AUTHORIZED' || terminal.authorized !== true) {
    issuePairingCode13640(terminal, false);
  }
  return terminal;
}

function terminalTokenValid13640(terminal, nodeId = '', token = '') {
  if (!terminal || terminal.authorized !== true || terminal.authStatus !== 'AUTHORIZED') return false;
  if (!nodeId || normalizeNodeId13640(terminal.nodeId) !== normalizeNodeId13640(nodeId)) return false;
  if (!token || !terminal.pairingTokenHash) return false;
  const hash = hashTerminalSecret13640(token);
  const stored = String(terminal.pairingTokenHash || '');
  if (stored.length !== hash.length) return false;
  try {
    return crypto.timingSafeEqual(Buffer.from(stored), Buffer.from(hash));
  } catch (_) {
    return false;
  }
}

function requireAuthorizedTerminal13640(req, res, next) {
  const terminalCode = normalizeTerminalCode13640(req.headers['x-terminal-code'] || '');
  // El nodo ADMIN local y las integraciones remotas existentes conservan su flujo.
  if (!terminalCode || terminalCode === 'ADMIN-01') return next();

  const nodeId = normalizeNodeId13640(req.headers['x-terminal-node-id'] || '');
  const token = String(req.headers['x-terminal-token'] || '').trim();
  const db = readDb();
  const terminal = findTerminalForNode13640(db, nodeId, terminalCode);

  if (!terminalTokenValid13640(terminal, nodeId, token)) {
    return res.status(403).json({
      ok: false,
      code: 'NODE_NOT_AUTHORIZED',
      pairingRequired: true,
      message: 'Esta terminal no está autorizada. Solicita al administrador el código de 6 dígitos.'
    });
  }

  terminal.lastSeenAt = new Date().toISOString();
  const detected = requestLanIp(req);
  if (detected) {
    terminal.ipAddress = detected;
    terminal.lastIp = detected;
  }
  req.terminal = terminal;
  writeDb(db);
  next();
}

const app = express();
const server = http.createServer(app);
const io = new Server(server, { cors: { origin: '*' }, maxHttpBufferSize: 16 * 1024 * 1024 });
const activeBranchRelays = new Map();
const branchRelayConnectionWaiters = new Map();
const lanPrintJobs = new Map();

const upload = multer({ dest: UPLOADS_DIR });

app.use(cors());
app.use(express.json({ limit: '12mb' }));
app.use((req, res, next) => {
  if (!RENDER_MIRROR_ONLY || req.method === 'OPTIONS') return next();
  const publicHealthRoute = req.method === 'GET' && (req.path === '/' || req.path === '/health');
  const mirrorApiRoute = /^\/api\/remote(?:\/|$)/.test(req.path || '');
  const socketIoRoute = /^\/socket\.io(?:\/|$)/.test(req.path || '');
  if (publicHealthRoute || mirrorApiRoute || socketIoRoute) return next();
  return res.status(404).json({ ok: false, message: 'Este servicio solo retransmite consultas en vivo de sucursales.' });
});
app.use('/uploads', express.static(UPLOADS_DIR));

io.use((socket, next) => {
  if (RENDER_MIRROR_ONLY) {
    const branchCode = normalizeMirrorBranchCode(socket.handshake.auth?.branchCode || '');
    const token = socket.handshake.auth?.apiKey || '';
    if (!branchCode || !hasValidRemoteToken(readDb(), token)) {
      return next(new Error('Conexión de sucursal no autorizada.'));
    }
    socket.data.branchCode = branchCode;
    return next();
  }
  return next();
});

function relayRemoteRead(req, res, remoteAuth, endpoint) {
  const requestedCode = String(req.headers['x-remote-branch-code'] || req.query?.branchCode || '').trim();
  const branchCode = normalizeMirrorBranchCode(requestedCode);
  if (requestedCode && !branchCode) {
    return res.status(400).json({ ok: false, message: 'El código de sucursal enviado no es válido.' });
  }
  if (!branchCode) {
    return res.status(400).json({ ok: false, message: 'Indica el código de la sucursal que deseas consultar.' });
  }

  const dispatchQuery = (socket) => {
    if (res.headersSent || res.writableEnded) return;
    if (!socket?.connected) {
      return res.status(503).json({ ok: false, message: `La sucursal ${branchCode} no está conectada al puente. Abre la aplicación en esa sucursal e inténtalo de nuevo.` });
    }
    socket.timeout(110000).emit('branch:query', {
      branchCode,
      endpoint,
      query: { ...(req.query || {}) }
    }, (timeoutError, result) => {
    if (timeoutError) {
      return res.status(504).json({ ok: false, message: `La sucursal ${branchCode} no respondió a tiempo.` });
    }
    if (!result?.ok || !result.payload) {
      return res.status(result?.status || 502).json({ ok: false, message: result?.message || 'No se pudo leer la información de la sucursal.' });
    }
    result.payload.authMode = remoteAuth.authMode;
    const readOnlyPermissions = { ...(result.payload.permissions || remoteAuth.permissions), writeInventory: false, syncBridge: false };
    result.payload.permissions = readOnlyPermissions;
    if (result.payload.remote && typeof result.payload.remote === 'object') {
      result.payload.remote.permissions = readOnlyPermissions;
    }
    return res.json(result.payload);
    });
  };

  const socket = activeBranchRelays.get(branchCode);
  if (socket?.connected) return dispatchQuery(socket);

  const waiter = (connectedSocket) => {
    clearTimeout(waitTimer);
    const waiting = branchRelayConnectionWaiters.get(branchCode) || [];
    const remaining = waiting.filter((entry) => entry !== waiter);
    if (remaining.length) branchRelayConnectionWaiters.set(branchCode, remaining);
    else branchRelayConnectionWaiters.delete(branchCode);
    dispatchQuery(connectedSocket);
  };
  const waitTimer = setTimeout(() => {
    const waiting = branchRelayConnectionWaiters.get(branchCode) || [];
    branchRelayConnectionWaiters.set(branchCode, waiting.filter((entry) => entry !== waiter));
    if (!branchRelayConnectionWaiters.get(branchCode)?.length) branchRelayConnectionWaiters.delete(branchCode);
    if (!res.headersSent && !res.writableEnded) {
      res.status(503).json({ ok: false, message: `La sucursal ${branchCode} no se conectó al puente. Comprueba que la aplicación de esa sucursal esté abierta.` });
    }
  }, 105000);
  const waiters = branchRelayConnectionWaiters.get(branchCode) || [];
  waiters.push(waiter);
  branchRelayConnectionWaiters.set(branchCode, waiters);
}

function auth(req, _res, next) {
  const username = req.headers['x-username'];
  const db = readDb();
  req.dataDb13650 = db;
  req.user = (db.users || []).find((u) => u.username === username) || null;
  next();
}

function requireAdmin(req, res, next) {
  if (!req.user || req.user.role !== 'ADMIN') {
    return res.status(403).json({ ok: false, message: 'Solo administrador' });
  }
  const db = req.dataDb13650 || readDb();
  if (isPrincipalAdmin(db, req.user)) return next();
  const moduleKey = moduleFromRequestPath(req.path || req.originalUrl || '');
  const actionKey = actionFromMethod(req.method || 'GET');
  if (!moduleKey || userHasPermission(db, req.user, moduleKey, actionKey)) return next();
  return res.status(403).json({ ok: false, message: 'No tienes permiso para esta acción.' });
}

function broadcastInventory(db, reason = 'SYNC', extra = {}) {
  io.emit('inventory:update', {
    products: db.products,
    lowStock: db.products.filter(isLowStock).map((p) => p.id),
    reason,
    ...extra
  });
}

function emitTickets(db, reason = 'TICKETS_SYNC', extra = {}) {
  io.emit('tickets:update', {
    ok: true,
    reason,
    tickets: (db.openTickets || []).slice().sort((a, b) => {
      return new Date(b.updatedAt).getTime() - new Date(a.updatedAt).getTime();
    }),
    ...extra
  });
}

app.get('/', (_req, res) => {
  res.type('text/plain').send(`Cremería El Güero · ${RENDER_MIRROR_ONLY ? 'Puente en vivo de sucursales' : 'Backend local activo'}\nHealth: /health`);
});

app.get('/health', (_req, res) => {
  const db = readDb();
  res.json({
    ok: true,
    service: RENDER_MIRROR_ONLY ? 'branch-live-relay' : 'embedded-lan-backend',
    port: PORT,
    adminName: ADMIN_NAME,
    hostname: os.hostname(),
    branchProfile: RENDER_MIRROR_ONLY ? null : getBranchProfile(db),
    connectedBranches: RENDER_MIRROR_ONLY ? activeBranchRelays.size : undefined,
    storesBranchData: false
  });
});


app.post('/api/terminal-auth/status', (req, res) => {
  const body = req.body || {};
  const nodeId = normalizeNodeId13640(body.nodeId);
  const terminalCode = normalizeTerminalCode13640(body.terminalCode || 'POS-01');
  const token = String(body.pairingToken || '').trim();
  if (!nodeId) return res.status(400).json({ ok: false, message: 'Identidad de dispositivo no disponible.' });

  const db = readDb();
  let terminal = findTerminalForNode13640(db, nodeId, terminalCode);
  if (terminal && terminalTokenValid13640(terminal, nodeId, token)) {
    terminal.lastSeenAt = new Date().toISOString();
    const detected = requestLanIp(req);
    if (detected) { terminal.ipAddress = detected; terminal.lastIp = detected; }
    writeDb(db);
    return res.json({
      ok: true,
      authorized: true,
      pairingRequired: false,
      terminalCode: terminal.terminalCode,
      terminalName: terminal.terminalName,
      authorizedAt: terminal.authorizedAt || '',
      viewPermissions: ensureTerminalViewPermissions13641(terminal)
    });
  }

  terminal = ensurePendingTerminal13640(db, body, req);
  // Si existía una autorización pero el dispositivo perdió/cambió su token
  // (reinstalación, restauración o terminal nueva con el mismo código),
  // se fuerza un nuevo emparejamiento visible únicamente en ADMIN.
  if (terminal.authorized === true || terminal.authStatus === 'AUTHORIZED') {
    terminal.authorized = false;
    terminal.authStatus = 'PENDING';
    terminal.pairingTokenHash = '';
    issuePairingCode13640(terminal, false);
  }
  writeDb(db);
  io.emit('terminals:update', { ok: true, reason: 'PAIRING_PENDING', terminalCode: terminal.terminalCode });

  return res.json({
    ok: true,
    authorized: false,
    pairingRequired: true,
    terminalCode: terminal.terminalCode,
    terminalName: terminal.terminalName,
    expiresAt: terminal.pairingCodeExpiresAt || '',
    viewPermissions: ensureTerminalViewPermissions13641(terminal),
    message: 'Ingresa el código de 6 dígitos que aparece únicamente en el Nodo Administrador.'
  });
});

app.post('/api/terminal-auth/request', (req, res) => {
  const body = req.body || {};
  const nodeId = normalizeNodeId13640(body.nodeId);
  if (!nodeId) return res.status(400).json({ ok: false, message: 'Identidad de dispositivo no disponible.' });
  const db = readDb();
  const terminal = ensurePendingTerminal13640(db, body, req);
  issuePairingCode13640(terminal, true);
  writeDb(db);
  io.emit('terminals:update', { ok: true, reason: 'PAIRING_CODE_REGENERATED', terminalCode: terminal.terminalCode });

  res.json({
    ok: true,
    authorized: false,
    pairingRequired: true,
    terminalCode: terminal.terminalCode,
    expiresAt: terminal.pairingCodeExpiresAt || '',
    message: 'Se generó un nuevo código. Consúltalo en Terminales del Nodo Administrador.'
  });
});

app.post('/api/terminal-auth/confirm', (req, res) => {
  const body = req.body || {};
  const nodeId = normalizeNodeId13640(body.nodeId);
  const terminalCode = normalizeTerminalCode13640(body.terminalCode || 'POS-01');
  const code = String(body.code || '').replace(/\D/g, '').slice(0, 6);
  if (!nodeId || code.length !== 6) {
    return res.status(400).json({ ok: false, code: 'INVALID_PAIRING_CODE', message: 'Captura exactamente los 6 dígitos.' });
  }

  const db = readDb();
  const terminal = findTerminalForNode13640(db, nodeId, terminalCode);
  if (!terminal) {
    return res.status(404).json({ ok: false, code: 'NODE_NOT_REGISTERED', message: 'La terminal todavía no está pendiente en el administrador.' });
  }
  const expires = Date.parse(terminal.pairingCodeExpiresAt || '') || 0;
  if (!terminal.pairingCodeHash) {
    issuePairingCode13640(terminal, true);
    writeDb(db);
    io.emit('terminals:update', { ok: true, reason: 'PAIRING_CODE_REGENERATED', terminalCode: terminal.terminalCode });
    return res.status(410).json({ ok: false, code: 'PAIRING_CODE_EXPIRED', message: 'El código ya no es válido. En el Nodo Administrador aparece uno nuevo.' });
  }
  if (hashTerminalSecret13640(code) !== terminal.pairingCodeHash) {
    if (expires > 0 && expires <= Date.now()) {
      issuePairingCode13640(terminal, true);
      writeDb(db);
      io.emit('terminals:update', { ok: true, reason: 'PAIRING_CODE_EXPIRED', terminalCode: terminal.terminalCode });
      return res.status(410).json({ ok: false, code: 'PAIRING_CODE_EXPIRED', message: 'El código venció. En el Nodo Administrador ya aparece uno nuevo.' });
    }
    return res.status(401).json({ ok: false, code: 'INVALID_PAIRING_CODE', message: 'Código incorrecto.' });
  }

  const token = crypto.randomBytes(32).toString('hex');
  terminal.nodeId = nodeId;
  terminal.pairingTokenHash = hashTerminalSecret13640(token);
  terminal.authStatus = 'AUTHORIZED';
  terminal.authorized = true;
  terminal.authorizedAt = new Date().toISOString();
  terminal.pairingCode = '';
  terminal.pairingCodeHash = '';
  terminal.pairingCodeCreatedAt = '';
  terminal.pairingCodeExpiresAt = '';
  terminal.lastSeenAt = new Date().toISOString();
  const detected = requestLanIp(req);
  if (detected) { terminal.ipAddress = detected; terminal.lastIp = detected; }
  writeDb(db);
  io.emit('terminals:update', { ok: true, reason: 'PAIRING_AUTHORIZED', terminalCode: terminal.terminalCode });

  res.json({
    ok: true,
    authorized: true,
    pairingRequired: false,
    pairingToken: token,
    terminalCode: terminal.terminalCode,
    terminalName: terminal.terminalName,
    authorizedAt: terminal.authorizedAt,
    viewPermissions: ensureTerminalViewPermissions13641(terminal)
  });
});

app.post('/api/terminal-auth/:terminalCode/regenerate', auth, requireAdmin, (req, res) => {
  const terminalCode = normalizeTerminalCode13640(req.params.terminalCode || '');
  const db = readDb();
  const terminal = (db.terminals || []).find((t) => normalizeTerminalCode13640(t.terminalCode) === terminalCode);
  if (!terminal) return res.status(404).json({ ok: false, message: 'Terminal no encontrada.' });
  terminal.authorized = false;
  terminal.authStatus = 'PENDING';
  terminal.pairingTokenHash = '';
  issuePairingCode13640(terminal, true);
  writeDb(db);
  io.emit('terminals:update', { ok: true, reason: 'PAIRING_CODE_REGENERATED', terminalCode });
  res.json({ ok: true, terminalCode, pairingCode: terminal.pairingCode, expiresAt: terminal.pairingCodeExpiresAt });
});

// A partir de aquí, cualquier petición POS con código de terminal requiere
// identidad persistente + token de emparejamiento vigente.
app.use('/api', requireAuthorizedTerminal13640);


app.get('/api/recovery/status', (_req, res) => {
  const validBackups = listValidBackups13633().slice(0, 5).map((file) => ({ file, mtime: fs.statSync(file).mtime }));
  res.json({ ok: true, status: LAST_AUTOHEAL_STATUS, backups: validBackups });
});

app.post('/api/recovery/backup-now', auth, requireAdmin, (_req, res) => {
  const backup = createBackupCopy13633('manual');
  res.json({ ok: !!backup, backup, message: backup ? 'Respaldo manual creado.' : 'No se pudo crear respaldo manual.' });
});

app.get('/api/remote/publication-info', (req, res) => {
  const db = readDb();
  const remoteAuth = authorizeRemoteRequest(req, res, db, 'readSummary');
  if (!remoteAuth) return;
  if (RENDER_MIRROR_ONLY) return relayRemoteRead(req, res, remoteAuth, '/api/remote/publication-info');
  const context = resolveRemoteMirrorContext(req, res, db, { allowEmpty: true });
  if (!context) return;
  const { branchCode, replica } = context;
  return res.json({
    ok: true,
    branchProfile: replica.snapshot?.branchProfile || getBranchProfile(db),
    remote: {
      adminName: ADMIN_NAME,
      hostname: os.hostname(),
      port: PORT,
      service: 'embedded-admin-backend',
      serviceMode: replica.snapshot ? 'REMOTE_MIRROR' : (RENDER_MIRROR_ONLY ? 'REMOTE_MIRROR_WAITING' : 'LOCAL_ADMIN'),
      requestedBranchCode: branchCode,
      permissions: getRemotePermissions(db)
    },
    lastSnapshotAt: replica.lastSnapshotAt || '',
    authMode: remoteAuth.authMode
  });
});

app.get('/api/remote/summary', (req, res) => {
  const db = readDb();
  const remoteAuth = authorizeRemoteRequest(req, res, db, 'readSummary');
  if (!remoteAuth) return;
  if (RENDER_MIRROR_ONLY) return relayRemoteRead(req, res, remoteAuth, '/api/remote/summary');
  const context = resolveRemoteMirrorContext(req, res, db);
  if (!context) return;

  const payload = getEffectiveRemoteSummary(db, context.branchCode);
  payload.authMode = remoteAuth.authMode;
  return res.json(payload);
});

app.get('/api/remote/cash-overview', (req, res) => {
  const db = readDb();
  const remoteAuth = authorizeRemoteRequest(req, res, db, 'readSummary');
  if (!remoteAuth) return;
  if (RENDER_MIRROR_ONLY) return relayRemoteRead(req, res, remoteAuth, '/api/remote/cash-overview');
  const context = resolveRemoteMirrorContext(req, res, db);
  if (!context) return;
  const dateKey = normalizeDateKey(req.query?.date);
  return res.json({
    ok: true,
    branchProfile: context.replica.snapshot?.branchProfile || getBranchProfile(db),
    overview: getEffectiveRemoteCashOverview(db, dateKey, context.branchCode),
    authMode: remoteAuth.authMode
  });
});

app.get('/api/remote/users', (req, res) => {
  const db = readDb();
  const remoteAuth = authorizeRemoteRequest(req, res, db, 'readSummary');
  if (!remoteAuth) return;
  if (RENDER_MIRROR_ONLY) return relayRemoteRead(req, res, remoteAuth, '/api/remote/users');
  const context = resolveRemoteMirrorContext(req, res, db);
  if (!context) return;
  const users = getEffectiveRemoteUsers(db, context.branchCode);
  return res.json({ ok: true, users, authMode: remoteAuth.authMode });
});

app.get('/api/remote/inventory', (req, res) => {
  const db = readDb();
  const remoteAuth = authorizeRemoteRequest(req, res, db, 'readInventory');
  if (!remoteAuth) return;
  if (RENDER_MIRROR_ONLY) return relayRemoteRead(req, res, remoteAuth, '/api/remote/inventory');
  const context = resolveRemoteMirrorContext(req, res, db);
  if (!context) return;

  const payload = getEffectiveRemoteInventory(db, req.query || {}, context.branchCode);
  payload.authMode = remoteAuth.authMode;
  payload.permissions = remoteAuth.permissions;
  return res.json(payload);
});

app.get('/api/remote/sales', (req, res) => {
  const db = readDb();
  const remoteAuth = authorizeRemoteRequest(req, res, db, 'readSummary');
  if (!remoteAuth) return;
  if (RENDER_MIRROR_ONLY) return relayRemoteRead(req, res, remoteAuth, '/api/remote/sales');
  const context = resolveRemoteMirrorContext(req, res, db);
  if (!context) return;
  const payload = getEffectiveRemoteSales(db, req.query || {}, context.branchCode);
  payload.authMode = remoteAuth.authMode;
  return res.json(payload);
});

app.get('/api/remote/employee-report', (req, res) => {
  const db = readDb();
  const remoteAuth = authorizeRemoteRequest(req, res, db, 'readSummary');
  if (!remoteAuth) return;
  if (RENDER_MIRROR_ONLY) return relayRemoteRead(req, res, remoteAuth, '/api/remote/employee-report');
  const context = resolveRemoteMirrorContext(req, res, db);
  if (!context) return;
  const payload = getEffectiveRemoteEmployeeReport(db, req.query || {}, context.branchCode);
  payload.authMode = remoteAuth.authMode;
  return res.json(payload);
});

app.get('/api/remote/stock-alerts', (req, res) => {
  const db = readDb();
  const remoteAuth = authorizeRemoteRequest(req, res, db, 'readInventory');
  if (!remoteAuth) return;
  if (RENDER_MIRROR_ONLY) return relayRemoteRead(req, res, remoteAuth, '/api/remote/stock-alerts');
  const context = resolveRemoteMirrorContext(req, res, db);
  if (!context) return;
  const payload = getEffectiveRemoteStockAlerts(db, context.branchCode);
  payload.authMode = remoteAuth.authMode;
  return res.json(payload);
});

app.get('/api/remote/alerts', (req, res) => {
  const db = readDb();
  const remoteAuth = authorizeRemoteRequest(req, res, db, 'readInventory');
  if (!remoteAuth) return;
  if (RENDER_MIRROR_ONLY) return relayRemoteRead(req, res, remoteAuth, '/api/remote/alerts');
  const context = resolveRemoteMirrorContext(req, res, db);
  if (!context) return;
  const payload = getEffectiveRemoteStockAlerts(db, context.branchCode);
  payload.authMode = remoteAuth.authMode;
  return res.json(payload);
});

app.post('/api/remote/inventory/adjust', (req, res) => {
  const db = readDb();
  const remoteAuth = authorizeRemoteRequest(req, res, db, 'writeInventory');
  if (!remoteAuth) return;
  const context = resolveRemoteMirrorContext(req, res, db);
  if (!context) return;

  try {
    const replica = context.replica;
    if (replica.snapshot?.inventory?.products) {
      const command = queueMirrorInventoryCommand(db, {
        ...req.body,
        actorName: req.body?.actorName || remoteAuth.actorName,
        actorUsername: req.body?.actorUsername || remoteAuth.actorUsername,
        actorBranch: String(req.body?.actorBranch || '').trim()
      }, context.branchCode);
      writeDb(db);
      return res.json({ ok: true, queued: true, command, message: 'Ajuste remoto enviado a la sucursal para aplicarse en la siguiente sincronización.' });
    }

    const result = applyInventoryAdjustment(db, {
      ...req.body,
      sourceType: 'REMOTE_ADMIN',
      actorName: req.body?.actorName || remoteAuth.actorName,
      actorUsername: req.body?.actorUsername || remoteAuth.actorUsername,
      actorBranch: String(req.body?.actorBranch || '').trim()
    });
    writeDb(db);
    broadcastInventory(db, 'REMOTE_INVENTORY_ADJUSTED', { productId: result.product.id });
    return res.json({ ok: true, product: sanitizeInventoryProduct(result.product), movement: result.movement });
  } catch (err) {
    return res.status(400).json({ ok: false, message: err.message });
  }
});

app.post('/api/remote/snapshot/push', (req, res) => {
  const db = readDb();
  const remoteAuth = authorizeRemoteRequest(req, res, db, 'syncBridge');
  if (!remoteAuth) return;
  const snapshot = req.body?.snapshot || null;
  if (!snapshot?.summary || !snapshot?.inventory) {
    return res.status(400).json({ ok: false, message: 'Snapshot remoto inválido' });
  }
  const branchCode = normalizeMirrorBranchCode(req.headers['x-remote-branch-code'] || req.body?.branchCode || snapshot.branchCode || snapshot.branchProfile?.code || '');
  if (!branchCode) return res.status(400).json({ ok: false, message: 'El snapshot debe indicar el código único de la sucursal.' });
  const profileCode = normalizeMirrorBranchCode(snapshot.branchProfile?.code || '');
  if (profileCode && profileCode !== branchCode) {
    return res.status(400).json({ ok: false, message: 'El código de sucursal del encabezado no coincide con el perfil publicado.' });
  }
  const target = getWritableMirrorReplica(db, branchCode);
  const replica = target.replica;
  const sourceInstanceId = String(snapshot.sourceInstanceId || '').trim();
  if (replica.sourceInstanceId && sourceInstanceId && replica.sourceInstanceId !== sourceInstanceId) {
    return res.status(409).json({ ok: false, message: `El código ${branchCode} ya está vinculado a otra instalación. Asigna un código diferente a cada sucursal.` });
  }
  replica.branchCode = branchCode;
  replica.sourceInstanceId = sourceInstanceId || replica.sourceInstanceId || '';
  replica.snapshot = snapshot;
  replica.lastSnapshotAt = new Date().toISOString();
  replica.lastSnapshotSource = String(snapshot?.branchProfile?.name || remoteAuth.actorName || '').trim();
  replica.updatedAt = new Date().toISOString();
  storeMirrorReplica(db, branchCode, replica);
  writeDb(db);
  return res.json({ ok: true, branchCode, receivedAt: replica.lastSnapshotAt, pendingCommands: (replica.commands || []).filter((row) => row.status === 'PENDING').length });
});

app.get('/api/remote/commands/pull', (req, res) => {
  const db = readDb();
  const remoteAuth = authorizeRemoteRequest(req, res, db, 'syncBridge');
  if (!remoteAuth) return;
  const context = resolveRemoteMirrorContext(req, res, db);
  if (!context) return;
  const pending = (context.replica.commands || []).filter((row) => row.status === 'PENDING');
  return res.json({ ok: true, commands: pending, generatedAt: new Date().toISOString() });
});

app.post('/api/remote/commands/ack', (req, res) => {
  const db = readDb();
  const remoteAuth = authorizeRemoteRequest(req, res, db, 'syncBridge');
  if (!remoteAuth) return;
  const context = resolveRemoteMirrorContext(req, res, db);
  if (!context) return;
  const ids = new Set(Array.isArray(req.body?.commandIds) ? req.body.commandIds.map(String) : []);
  const target = getWritableMirrorReplica(db, context.branchCode);
  const replica = target.replica;
  replica.commands = (replica.commands || []).map((command) => {
    if (ids.has(String(command.id))) {
      return { ...command, status: 'ACKNOWLEDGED', acknowledgedAt: new Date().toISOString() };
    }
    return command;
  }).slice(-500);
  replica.updatedAt = new Date().toISOString();
  storeMirrorReplica(db, context.branchCode, replica);
  writeDb(db);
  return res.json({ ok: true, acknowledged: ids.size });
});

app.post('/auth/login', (req, res) => {
  const { username, password, terminalCode, terminalType, terminalName, nodeId, pairingToken } = req.body || {};
  const db = readDb();
  if (String(terminalType || '').toUpperCase() !== 'ADMIN') {
    const terminal = findTerminalForNode13640(db, nodeId, terminalCode);
    if (!terminalTokenValid13640(terminal, nodeId, pairingToken)) {
      return res.status(403).json({
        ok: false,
        code: 'NODE_NOT_AUTHORIZED',
        pairingRequired: true,
        message: 'Esta terminal requiere autorización mediante el código de 6 dígitos del Nodo Administrador.'
      });
    }
  }
  const user = db.users.find((u) => u.username === username && u.password === password);

  if (!user) {
    return res.status(401).json({ ok: false, message: 'Credenciales inválidas' });
  }

  if (user.role !== 'ADMIN' && user.active === false) {
    return res.status(403).json({ ok: false, message: 'Empleado inactivo' });
  }

  touchTerminal(db, { terminalCode, terminalType, terminalName, ipAddress: req.ip, lastIp: requestLanIp(req) });
  const activeTerminal = String(terminalType || '').toUpperCase() === 'ADMIN'
    ? (db.terminals || []).find((t) => String(t.terminalCode || '') === String(terminalCode || 'ADMIN-01'))
    : findTerminalForNode13640(db, nodeId, terminalCode);
  const terminalViewPermissions = ensureTerminalViewPermissions13641(activeTerminal || { terminalType: terminalType || 'POS' });
  writeDb(db);

  ensureRbacCollections(db);
  const safeUser = { id: user.id, username: user.username, role: user.role, name: user.name, employeeNumber: user.employeeNumber || '' };
  safeUser.isPrincipalAdmin = isPrincipalAdmin(db, user);
  safeUser.permissions = getUserPermissions(db, user);
  res.json({ ok: true, user: safeUser, terminalViewPermissions });
});

app.post('/auth/admin/verify', auth, (req, res) => {
  // v13.6.27: las acciones operativas ya no piden contraseña adicional.
  // La única contraseña obligatoria queda en el inicio de sesión del modo administrador.
  const db = readDb();
  if (req.user?.role !== 'ADMIN') {
    return res.status(403).json({ ok: false, message: 'Solo el administrador puede ejecutar esta acción.' });
  }
  res.json({ ok: true, bypassed: true, authorizedBy: { id: req.user.id, username: req.user.username, name: req.user.name } });
});

app.get('/api/dashboard', auth, (_req, res) => {
  const db = readDb();
  const cashSales = (db.sales || []).filter((sale) => !isOwnerConsumptionSale(sale));
  const ownerConsumptionSales = (db.sales || []).filter((sale) => isOwnerConsumptionSale(sale));
  const totalSales = cashSales.reduce((a, s) => a + safeNumber(s.total), 0);
  const today = new Date().toISOString().slice(0, 10);
  const todaySales = cashSales
    .filter((s) => String(s.createdAt || '').startsWith(today))
    .reduce((a, s) => a + safeNumber(s.total), 0);

  res.json({
    ok: true,
    metrics: {
      products: db.products.length,
      stockUnits: round2(db.products.reduce((a, p) => a + safeNumber(p.stock), 0)),
      salesCount: cashSales.length,
      ownerConsumptionCount: ownerConsumptionSales.length,
      totalSales,
      todaySales,
      invoices: db.invoices.length,
      terminals: db.terminals.length,
      lowStockCount: db.products.filter(isLowStock).length,
      openTickets: db.openTickets.length,
      employees: getActiveEmployees(db).length,
      branches: (db.branches || []).filter((b) => b.active !== false).length
    }
  });
});

app.get('/api/products', auth, (_req, res) => {
  const db = readDb();
  res.json({ ok: true, products: db.products });
});

app.get('/api/inventory', auth, (_req, res) => {
  const db = readDb();
  res.json({
    ok: true,
    inventory: db.products.map((p) => ({
      id: p.id,
      name: p.name,
      sku: p.sku,
      barcode: p.barcode || p.codigoBarras || '',
      stock: p.stock,
      unit: p.unit,
      price: p.price,
      retailPrice: p.retailPrice,
      wholesalePrice: p.wholesalePrice,
      costPrice: p.costPrice,
      lowStockThreshold: p.lowStockThreshold,
      imageUrl: p.imageUrl,
      allowsFraction: p.allowsFraction
    }))
  });
});


app.get('/api/inventory/import-template', auth, requireAdmin, (_req, res) => {
  const buffer = factoramaBuildInventoryTemplateBuffer();
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename="plantilla-inventario.xlsx"');
  res.send(buffer);
});

app.post('/api/inventory/import', auth, requireAdmin, upload.single('file'), (req, res) => {
  try {
    if (!req.file?.path) {
      return res.status(400).json({ ok: false, message: 'Selecciona un archivo Excel para importar.' });
    }
    const workbook = XLSX.readFile(req.file.path, { cellDates: false, raw: false });
    const firstSheet = workbook.SheetNames[0];
    if (!firstSheet) {
      return res.status(400).json({ ok: false, message: 'El archivo no contiene hojas para importar.' });
    }
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheet], { defval: '' });
    if (!rows.length) {
      return res.status(400).json({ ok: false, message: 'La hoja principal está vacía.' });
    }

    const db = readDb();
    const summary = { created: 0, updated: 0, failed: 0, errors: [] };

    rows.forEach((rawRow, index) => {
      try {
        const row = normalizeInventoryImportRow(rawRow);
        if (!row.name || !row.category) {
          throw new Error('Faltan Nombre y/o Categoria.');
        }
        const existing = (db.products || []).find((product) => {
          const productSku = String(product.sku || '').trim();
          const productBarcode = String(product.barcode || product.codigoBarras || '').trim();
          if (row.sku && productSku && productSku === row.sku) return true;
          if (row.barcode && productBarcode && productBarcode === row.barcode) return true;
          return String(product.name || '').trim().toLowerCase() === row.name.toLowerCase();
        }) || null;
        const product = buildImportedProduct(db, row, existing);
        if (existing) {
          Object.assign(existing, product);
          summary.updated += 1;
        } else {
          db.products.push(product);
          summary.created += 1;
        }
      } catch (err) {
        summary.failed += 1;
        summary.errors.push({ row: index + 2, message: err.message });
      }
    });

    writeDb(db);
    broadcastInventory(db, 'INVENTORY_IMPORTED', { created: summary.created, updated: summary.updated, failed: summary.failed });
    res.json({ ok: true, summary });
  } catch (err) {
    res.status(400).json({ ok: false, message: err.message || 'No se pudo importar el inventario.' });
  }
});

app.post('/api/products', auth, requireAdmin, upload.single('image'), (req, res) => {
  try {
    const db = readDb();
    const imageUrl = req.file ? `/uploads/${req.file.filename}` : '';
    const product = normalizeProductPayload(req.body || {}, imageUrl);

    db.products.push(product);
    writeDb(db);
    broadcastInventory(db, 'NEW_PRODUCT', { productId: product.id });

    res.json({ ok: true, product });
  } catch (err) {
    res.status(400).json({ ok: false, message: err.message });
  }
});

app.put('/api/products/:id', auth, requireAdmin, upload.single('image'), (req, res) => {
  try {
    const db = readDb();
    const current = db.products.find((p) => p.id === req.params.id);

    if (!current) {
      return res.status(404).json({ ok: false, message: 'Producto no encontrado' });
    }

    const imageUrl = req.file ? `/uploads/${req.file.filename}` : current.imageUrl;
    const updated = normalizeProductPayload({ ...current, ...(req.body || {}), id: current.id, createdAt: current.createdAt }, imageUrl);

    Object.assign(current, updated);
    writeDb(db);
    broadcastInventory(db, 'PRODUCT_UPDATED', { productId: current.id });

    res.json({ ok: true, product: current });
  } catch (err) {
    res.status(400).json({ ok: false, message: err.message });
  }
});


function removeProductsFromDb(db, productIds = []) {
  const ids = new Set((productIds || []).map((row) => String(row || '').trim()).filter(Boolean));
  if (!ids.size) return { deleted: 0, removedIds: [] };
  const removedIds = (db.products || []).filter((product) => ids.has(String(product.id))).map((product) => product.id);
  db.products = (db.products || []).filter((product) => !ids.has(String(product.id)));
  db.openTickets = (db.openTickets || []).map((ticket) => {
    const filteredItems = normalizeTicketItems(ticket.items).filter((item) => !ids.has(String(item.productId)));
    return {
      ...ticket,
      items: filteredItems,
      subtotal: sumTicket(filteredItems),
      updatedAt: new Date().toISOString()
    };
  });
  return { deleted: removedIds.length, removedIds };
}

app.delete('/api/products/:id', auth, requireAdmin, (req, res) => {
  const db = readDb();
  const result = removeProductsFromDb(db, [req.params.id]);

  if (!result.deleted) {
    return res.status(404).json({ ok: false, message: 'Producto no encontrado' });
  }

  writeDb(db);
  broadcastInventory(db, 'PRODUCT_DELETED', { productId: req.params.id, deleted: result.deleted });
  emitTickets(db, 'PRODUCT_DELETED', { productId: req.params.id, deleted: result.deleted });

  res.json({ ok: true, deleted: result.deleted, productIds: result.removedIds });
});

app.post('/api/products/bulk-delete', auth, requireAdmin, (req, res) => {
  const { productIds } = req.body || {};
  const ids = Array.isArray(productIds) ? productIds : [];
  if (!ids.length) {
    return res.status(400).json({ ok: false, message: 'Selecciona al menos un producto para eliminar.' });
  }

  const db = readDb();
  const result = removeProductsFromDb(db, ids);
  if (!result.deleted) {
    return res.status(404).json({ ok: false, message: 'No se encontraron productos para eliminar.' });
  }

  writeDb(db);
  broadcastInventory(db, 'PRODUCTS_BULK_DELETED', { productIds: result.removedIds, deleted: result.deleted });
  emitTickets(db, 'PRODUCTS_BULK_DELETED', { productIds: result.removedIds, deleted: result.deleted });
  res.json({ ok: true, deleted: result.deleted, productIds: result.removedIds });
});

app.post('/api/products/clear', auth, requireAdmin, (_req, res) => {
  const db = readDb();
  const allIds = (db.products || []).map((product) => product.id);
  const result = removeProductsFromDb(db, allIds);
  writeDb(db);
  broadcastInventory(db, 'PRODUCTS_CLEARED', { deleted: result.deleted, productIds: result.removedIds });
  emitTickets(db, 'PRODUCTS_CLEARED', { deleted: result.deleted, productIds: result.removedIds });
  res.json({ ok: true, deleted: result.deleted, productIds: result.removedIds });
});

app.post('/api/products/restock', auth, requireAdmin, (req, res) => {
  const db = req.dataDb13650 || readDb();
  try {
    ensureCashCollections(db);
    db.inventoryAdjustments = Array.isArray(db.inventoryAdjustments) ? db.inventoryAdjustments : [];
    db.auditLog = Array.isArray(db.auditLog) ? db.auditLog : [];
    db.counters = { ...(db.counters || {}) };
    const product = (db.products || []).find((row) => row.id === req.body?.productId);
    if (!product) return res.status(404).json({ ok: false, message: 'Producto existente no encontrado.' });

    const idempotencyKey = String(req.body?.idempotencyKey || '').trim();
    if (!idempotencyKey || idempotencyKey.length > 160) {
      return res.status(400).json({ ok: false, message: 'No se recibió un identificador único para el Restock. Actualiza la aplicación y vuelve a intentarlo.' });
    }
    const priorMovement = db.inventoryAdjustments.find((entry) => entry.idempotencyKey === idempotencyKey);
    const priorCashEntry = db.cashEntries.find((entry) => entry.idempotencyKey === idempotencyKey);
    if (priorMovement || priorCashEntry) {
      if (!priorMovement || !priorCashEntry) {
        return res.status(409).json({ ok: false, code: 'RESTOCK_INCOMPLETE', message: 'El Restock tiene un registro incompleto y requiere revisión administrativa. No se volvió a aplicar.' });
      }
      return res.json({
        ok: true,
        alreadyRegistered: true,
        product,
        movement: priorMovement,
        cashEntry: priorCashEntry,
        overview: buildCashOverview(db, normalizeDateKey(priorCashEntry.dateKey || priorCashEntry.createdAt))
      });
    }

    const rawQuantity = Number(req.body?.quantity);
    if (!Number.isFinite(rawQuantity) || rawQuantity <= 0) {
      return res.status(400).json({ ok: false, message: 'La cantidad para Restock debe ser mayor a cero.' });
    }
    const quantity = roundQuantity(rawQuantity);
    if (quantity <= 0 || (!product.allowsFraction && !Number.isInteger(quantity))) {
      return res.status(400).json({ ok: false, message: product.allowsFraction
        ? 'La cantidad debe ser mayor a cero y admite hasta tres decimales.'
        : 'Este producto se maneja por pieza; captura una cantidad entera.' });
    }
    const rawUnitCost = Number(product.costPrice);
    if (!Number.isFinite(rawUnitCost) || rawUnitCost <= 0) {
      return res.status(409).json({ ok: false, code: 'ACQUISITION_COST_REQUIRED', message: 'Este producto no tiene un costo de adquisición válido. Captura o corrige su costo en Productos antes de hacer el Restock.' });
    }
    const unitCost = round2(rawUnitCost);
    if (unitCost <= 0) {
      return res.status(409).json({ ok: false, code: 'ACQUISITION_COST_REQUIRED', message: 'El costo de adquisición debe ser mayor a $0.00. Corrígelo antes de hacer el Restock.' });
    }
    const totalCost = round2(quantity * unitCost);
    if (totalCost <= 0) return res.status(400).json({ ok: false, message: 'El importe del Restock debe ser mayor a cero.' });
    const dateKey = normalizeDateKey();
    const overviewBefore = buildCashOverview(db, dateKey);
    if (overviewBefore.expectedCash + 0.0001 < totalCost) {
      const shortage = round2(totalCost - overviewBefore.expectedCash);
      return res.status(409).json({
        ok: false,
        code: 'INSUFFICIENT_CASH',
        message: `Efectivo insuficiente para el Restock. Disponible: $${overviewBefore.expectedCash.toFixed(2)}; requerido: $${totalCost.toFixed(2)}; faltan: $${shortage.toFixed(2)}. Registra primero el fondo o ingreso autorizado de caja.`
      });
    }

    const result = applyInventoryAdjustment(db, {
      productId: product.id,
      operation: 'add',
      quantity,
      reason: String(req.body?.reason || '').trim() || 'Restock de producto existente con egreso de efectivo',
      type: 'restock',
      sourceType: 'LOCAL_ADMIN_RESTOCK',
      actorName: req.user?.name || 'Administrador local',
      actorUsername: req.user?.username || 'admin',
      actorBranch: ADMIN_NAME
    });
    const now = new Date().toISOString();
    const restockSequence = safeNumber(db.counters.restockSequence, 0) + 1;
    db.counters.restockSequence = restockSequence;
    const restockId = uuidv4();
    const cashEntry = {
      id: uuidv4(),
      restockId,
      inventoryMovementId: result.movement.id,
      idempotencyKey,
      folio: `RST-${String(restockSequence).padStart(6, '0')}`,
      kind: 'restock_purchase',
      sourceType: 'LOCAL_ADMIN_RESTOCK',
      dateKey,
      productId: product.id,
      productName: product.name,
      quantity,
      unit: product.unit || 'PZA',
      unitCostSnapshot: unitCost,
      totalCost,
      accountingAmount: totalCost,
      amountCashImpact: totalCost,
      paymentMethod: 'cash',
      paymentMethodLabel: 'Efectivo',
      description: `Restock: ${product.name}`,
      notes: String(req.body?.notes || '').trim(),
      terminalCode: String(req.body?.terminalCode || 'ADMIN-01').trim(),
      branchId: getBranchProfile(db).code || 'LOCAL',
      userId: req.user?.id || '',
      userName: req.user?.name || req.user?.username || '',
      createdAt: now,
      status: 'COMMITTED'
    };
    Object.assign(result.movement, {
      restockId,
      idempotencyKey,
      unitCostSnapshot: unitCost,
      totalCost,
      cashEntryId: cashEntry.id,
      dateKey,
      status: 'COMMITTED'
    });
    db.cashEntries.push(cashEntry);
    db.auditLog.unshift({
      id: uuidv4(), action: 'RESTOCK_CASH_COMMITTED', restockId, idempotencyKey,
      productId: product.id, productName: product.name, quantity, unit: product.unit || 'PZA',
      unitCostSnapshot: unitCost, totalCost, inventoryMovementId: result.movement.id,
      cashEntryId: cashEntry.id, branchId: cashEntry.branchId,
      userId: req.user?.id || '', userName: req.user?.name || req.user?.username || '',
      createdAt: now
    });
    db.auditLog = db.auditLog.slice(0, 10000);
    writeDb(db);
    const overview = {
      ...overviewBefore,
      expensesCashImpactTotal: round2(overviewBefore.expensesCashImpactTotal + cashEntry.amountCashImpact),
      restockTotal: round2(overviewBefore.restockTotal + totalCost),
      restockCashOutflowTotal: round2(overviewBefore.restockCashOutflowTotal + cashEntry.amountCashImpact),
      restockPurchases: [...overviewBefore.restockPurchases, cashEntry],
      accountingOutflowTotal: round2(overviewBefore.accountingOutflowTotal + totalCost),
      expectedCash: round2(overviewBefore.expectedCash - cashEntry.amountCashImpact),
      entries: [...overviewBefore.entries, cashEntry]
    };
    res.json({ ok: true, product: result.product, movement: result.movement, cashEntry, overview });
    // La respuesta al operador sale primero. La sincronización LAN no debe mantener
    // el botón esperando ni obligar a redibujar el formulario que conserva el foco.
    setTimeout(() => {
      try {
        broadcastInventory(db, 'RESTOCK', {
          productId: result.product.id,
          product: result.product,
          movement: result.movement,
          idempotencyKey,
          terminalCode: cashEntry.terminalCode
        });
      } catch (broadcastError) {
        console.error('[RESTOCK] No se pudo sincronizar el inventario:', broadcastError.message || broadcastError);
      }
    }, 500);
  } catch (err) {
    res.status(400).json({ ok: false, message: err.message || 'No se pudo reabastecer el producto.' });
  }
});


app.get('/api/inventory/adjustments', auth, requireAdmin, (req, res) => {
  const db = readDb();
  const limit = Math.max(1, Math.min(1000, safeNumber(req.query?.limit, 250)));
  const productId = String(req.query?.productId || '').trim();
  const rows = (db.inventoryAdjustments || [])
    .filter((entry) => !productId || String(entry.productId || '') === productId)
    .sort((a, b) => new Date(b.createdAt || 0).getTime() - new Date(a.createdAt || 0).getTime())
    .slice(0, limit)
    .map((entry) => ({
      ...entry,
      previousStock: roundQuantity(safeNumber(entry.previousStock)),
      newStock: roundQuantity(safeNumber(entry.newStock)),
      delta: roundQuantity(safeNumber(entry.delta))
    }));
  res.json({ ok: true, adjustments: rows });
});

app.delete('/api/inventory/adjustments/:id', auth, requireAdmin, (req, res) => {
  const db = readDb();
  const before = (db.inventoryAdjustments || []).length;
  db.inventoryAdjustments = (db.inventoryAdjustments || []).filter((entry) => entry.id !== req.params.id);
  const deleted = before - db.inventoryAdjustments.length;
  if (!deleted) return res.status(404).json({ ok: false, message: 'Movimiento de inventario no encontrado.' });
  writeDb(db);
  res.json({ ok: true, deleted });
});


app.post('/api/inventory/adjust', auth, requireAdmin, (req, res) => {
  const db = readDb();

  try {
    const result = applyInventoryAdjustment(db, {
      ...req.body,
      sourceType: 'LOCAL_ADMIN',
      actorName: req.user?.name || 'Administrador local',
      actorUsername: req.user?.username || 'admin',
      actorBranch: ADMIN_NAME
    });
    writeDb(db);
    broadcastInventory(db, 'LOCAL_INVENTORY_ADJUSTED', { productId: result.product.id });
    res.json({ ok: true, product: sanitizeInventoryProduct(result.product), movement: result.movement });
  } catch (err) {
    res.status(400).json({ ok: false, message: err.message });
  }
});

app.get('/api/tickets/open', auth, (_req, res) => {
  const db = readDb();
  const tickets = (db.openTickets || []).filter((ticket) => String(ticket.status || 'OPEN').toUpperCase() === 'OPEN');
  res.json({ ok: true, tickets });
});

app.post('/api/tickets/open', auth, (req, res) => {
  const db = readDb();
  const body = req.body || {};
  const items = normalizeTicketItems(body.items || []);
  const subtotal = sumTicket(items);
  let ticket = (db.openTickets || []).find((t) => t.id === body.id);

  if (ticket) {
    Object.assign(ticket, {
      items,
      subtotal,
      operatorName: body.operatorName || ticket.operatorName,
      terminalCode: body.terminalCode || ticket.terminalCode,
      customerName: String(body.customerName || ticket.customerName || '').trim(),
      customerReference: String(body.customerReference || ticket.customerReference || '').trim(),
      updatedAt: new Date().toISOString(),
      status: 'OPEN'
    });
  } else {
    const seq = (db.openTickets || []).length + 1;
    ticket = {
      id: uuidv4(),
      folio: `TK-${String(seq).padStart(4, '0')}`,
      terminalCode: body.terminalCode || 'POS-01',
      operatorName: body.operatorName || 'Operador',
      customerName: String(body.customerName || '').trim(),
      customerReference: String(body.customerReference || '').trim(),
      takenBy: '',
      status: 'OPEN',
      items,
      subtotal,
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    };
    db.openTickets.push(ticket);
  }

  touchTerminal(db, {
    terminalCode: body.terminalCode || 'POS-01',
    terminalType: 'POS',
    terminalName: body.terminalName || body.terminalCode || 'Punto de Venta'
  });

  writeDb(db);
  emitTickets(db, 'TICKET_UPSERT', { ticketId: ticket.id });

  res.json({ ok: true, ticket });
});

app.post('/api/tickets/take', auth, (req, res) => {
  const db = readDb();
  const { ticketId, terminalCode, operatorName } = req.body || {};
  const ticket = (db.openTickets || []).find((t) => t.id === ticketId);

  if (!ticket) {
    return res.status(404).json({ ok: false, message: 'Ticket no encontrado' });
  }

  ticket.terminalCode = terminalCode || ticket.terminalCode;
  ticket.takenBy = operatorName || ticket.takenBy;
  ticket.status = 'IN_PROGRESS';
  ticket.updatedAt = new Date().toISOString();

  touchTerminal(db, {
    terminalCode: terminalCode || 'POS-01',
    terminalType: 'POS',
    terminalName: terminalCode || 'Punto de Venta'
  });

  writeDb(db);
  emitTickets(db, 'TICKET_TAKEN', { ticketId: ticket.id });

  res.json({ ok: true, ticket });
});

app.post('/api/tickets/close', auth, (req, res) => {
  const db = readDb();
  const { ticketId } = req.body || {};

  db.openTickets = (db.openTickets || []).filter((t) => t.id !== ticketId);
  writeDb(db);
  emitTickets(db, 'TICKET_CLOSED', { ticketId });

  res.json({ ok: true });
});

app.delete('/api/tickets/:id', auth, requireAdmin, (req, res) => {
  const db = readDb();
  const ticket = (db.openTickets || []).find((entry) => entry.id === req.params.id);
  db.openTickets = (db.openTickets || []).filter((entry) => entry.id !== req.params.id);
  writeDb(db);
  emitTickets(db, 'TICKET_DELETED', { ticketId: req.params.id });
  res.json({ ok: true, deleted: !!ticket });
});



// v13.6.31: utilidades de venta segura, idempotencia y verificación de carrito.
let saleCommitQueue13630 = Promise.resolve();
function runSaleCommit13630(task) {
  const run = saleCommitQueue13630.then(task, task);
  saleCommitQueue13630 = run.catch(() => {});
  return run;
}
function normalizeIntegrityNumber13630(value, decimals = 6) {
  const number = Number(value || 0);
  if (!Number.isFinite(number)) return 0;
  return Number(number.toFixed(decimals));
}
function canonicalSaleItemForIntegrity13630(item = {}) {
  return {
    productId: String(item.productId || '').trim(),
    name: String(item.name || '').trim(),
    quantity: normalizeIntegrityNumber13630(item.quantity, 6),
    unit: String(item.unit || 'PZA').trim().toUpperCase(),
    saleMode: item.saleMode === 'WHOLESALE' ? 'WHOLESALE' : 'RETAIL',
    saleByAmount: item.saleByAmount === true,
    saleAmountTarget: normalizeIntegrityNumber13630(item.saleAmountTarget, 2),
    originalUnitPrice: normalizeIntegrityNumber13630(item.originalUnitPrice ?? item.unitPrice ?? 0, 2),
    unitPrice: normalizeIntegrityNumber13630(item.unitPrice ?? item.originalUnitPrice ?? 0, 2),
    discountAmount: normalizeIntegrityNumber13630(item.discountAmount || 0, 2),
    discountPercent: normalizeIntegrityNumber13630(item.discountPercent || 0, 2),
    promotionId: String(item.promotionId || item.selectedPromotionId || '').trim(),
    lineTotal: normalizeIntegrityNumber13630(item.lineTotal ?? (Number(item.quantity || 0) * Number(item.unitPrice ?? item.originalUnitPrice ?? 0)), 2)
  };
}
function buildCartFingerprint13630(items = []) {
  return JSON.stringify((Array.isArray(items) ? items : []).map(canonicalSaleItemForIntegrity13630));
}
function hashText13630(text = '') {
  let hash = 2166136261;
  const input = String(text || '');
  for (let i = 0; i < input.length; i += 1) {
    hash ^= input.charCodeAt(i);
    hash = Math.imul(hash, 16777619) >>> 0;
  }
  return `FNV1A-${hash.toString(16).padStart(8, '0')}`;
}

app.post('/api/sales', auth, async (req, res) => {
  try {
    const result = await runSaleCommit13630(() => {
      const {
        items = [],
        terminalCode = 'UNKNOWN',
        clientTxnId,
        operatorId = '',
        operatorName = '',
        employeeNumber = '',
        ticketId = '',
        customerName = '',
        customerReference = '',
        dateKey = '',
        paymentMethod = 'cash',
        payments = [],
        amountReceived = 0,
        printTicket = false,
        itemCount,
        cartFingerprint = '',
        clientCartHash = '',
        saleSnapshotId = ''
      } = req.body || {};

      const db = readDb();
      const existing = db.sales.find((s) => s.clientTxnId === clientTxnId);
      if (existing) {
        const existingEntry = ensureSaleCashEntry(db, existing, req.user);
        writeDb(db);
        return { status: 200, body: { ok: true, sale: existing, cashEntry: existingEntry, overview: buildCashOverview(db, existing.dateKey || normalizeDateKey(existing.createdAt)), deduped: true } };
      }

      const normalizedItems = normalizeTicketItems(items);
      if (!normalizedItems.length) return { status: 400, body: { ok: false, message: 'No hay productos para vender' } };
      if (itemCount !== undefined && Number(itemCount) !== normalizedItems.length) {
        return { status: 409, body: { ok: false, message: 'La cantidad de productos enviada no coincide con el cobro congelado.' } };
      }
      if (!String(operatorName || '').trim()) return { status: 400, body: { ok: false, message: 'Selecciona un operador válido antes de cobrar' } };

      const productMap = new Map(db.products.map((p) => [p.id, p]));
      const pricedItems = [];
      for (const item of normalizedItems) {
        const product = productMap.get(item.productId);
        if (!product) return { status: 400, body: { ok: false, message: `Producto inexistente: ${item.productId}` } };
        if (safeNumber(product.stock) < safeNumber(item.quantity)) {
          const conflict = { id: uuidv4(), terminalCode, clientTxnId, productId: item.productId, requested: item.quantity, available: safeNumber(product.stock), createdAt: new Date().toISOString() };
          db.saleConflicts.push(conflict);
          writeDb(db);
          return { status: 409, body: { ok: false, message: `Stock insuficiente para ${product.name}`, conflict } };
        }
        pricedItems.push(applyPromotionToTicketItem(db, { ...item, originalUnitPrice: item.originalUnitPrice || item.unitPrice }, product));
      }

      if (cartFingerprint) {
        const serverFingerprint = buildCartFingerprint13630(pricedItems);
        if (serverFingerprint !== String(cartFingerprint)) {
          return { status: 409, body: { ok: false, message: 'La venta enviada no coincide con la venta recalculada por el servidor. Revisa el carrito y vuelve a cobrar.', serverCartHash: hashText13630(serverFingerprint), clientCartHash } };
        }
      }

      const subtotalBeforeDiscount = round2(pricedItems.reduce((acc, item) => acc + safeNumber(item.quantity) * safeNumber(item.originalUnitPrice || item.unitPrice), 0));
      const discountTotal = round2(pricedItems.reduce((acc, item) => acc + safeNumber(item.discountAmount) * safeNumber(item.quantity), 0));
      const total = sumTicket(pricedItems);
      let paymentInfo;
      try { paymentInfo = normalizeSalePayments(paymentMethod, payments, total, amountReceived); }
      catch (err) { return { status: 400, body: { ok: false, message: err.message } }; }

      for (const item of pricedItems) {
        const product = productMap.get(item.productId);
        product.stock = roundQuantity(safeNumber(product.stock) - safeNumber(item.quantity));
        product.updatedAt = new Date().toISOString();
      }

      const nextSalesSequence = safeNumber(db.counters?.salesSequence, 0) + 1;
      db.counters = { ...(db.counters || {}), salesSequence: nextSalesSequence };
      const serverFingerprint = buildCartFingerprint13630(pricedItems);
      const branchProfile = getBranchProfile(db);
      const sale = {
        id: uuidv4(),
        folio: `VTA-${String(nextSalesSequence).padStart(6, '0')}`,
        branchName: branchProfile.name || ADMIN_NAME,
        branchCode: branchProfile.code || 'ADMIN-01',
        terminalCode,
        operatorId: operatorId || '',
        operatorName,
        employeeNumber: employeeNumber || '',
        customerName: String(customerName || '').trim(),
        customerReference: String(customerReference || '').trim(),
        createdAt: new Date().toISOString(),
        dateKey: normalizeDateKey(dateKey || new Date()),
        clientTxnId: clientTxnId || uuidv4(),
        saleSnapshotId: saleSnapshotId || '',
        ticketId: ticketId || '',
        items: pricedItems,
        itemCount: pricedItems.length,
        cartFingerprint: serverFingerprint,
        clientCartHash: clientCartHash || '',
        serverCartHash: hashText13630(serverFingerprint),
        subtotalBeforeDiscount,
        discountTotal,
        total,
        paymentMethod: paymentInfo.paymentMethod,
        payments: paymentInfo.payments,
        amountReceived: paymentInfo.amountReceived,
        changeDue: paymentInfo.changeDue,
        affectsCash: true,
        affectsSalesTotals: true,
        printTicketRequested: printTicket === true
      };
      const cashEntry = ensureSaleCashEntry(db, sale, req.user);
      db.sales.push(sale);
      db.saleCommitLog = Array.isArray(db.saleCommitLog) ? db.saleCommitLog : [];
      db.saleCommitLog.push({ id: uuidv4(), saleId: sale.id, folio: sale.folio, terminalCode, clientTxnId: sale.clientTxnId, clientCartHash: sale.clientCartHash, serverCartHash: sale.serverCartHash, itemCount: sale.itemCount, total: sale.total, createdAt: sale.createdAt });
      if (db.saleCommitLog.length > 5000) db.saleCommitLog = db.saleCommitLog.slice(-5000);
      db.inventoryAdjustments.push({ id: uuidv4(), type: 'sale', saleId: sale.id, createdAt: sale.createdAt, items: pricedItems.map((item) => ({ productId: item.productId, quantity: item.quantity, promotionId: item.promotionId || '' })) });
      if (ticketId) db.openTickets = db.openTickets.filter((t) => t.id !== ticketId);
      if (operatorId) clearSellerSession(db, operatorId, 'COMPLETED');
      touchTerminal(db, { terminalCode, terminalType: 'POS', terminalName: terminalCode });
      writeDb(db);
      return { status: 200, body: { ok: true, sale, cashEntry, overview: buildCashOverview(db, sale.dateKey || normalizeDateKey(sale.createdAt)), inventory: db.products } };
    });

    if (result.status !== 200) return res.status(result.status).json(result.body);
    io.emit('sale:created', result.body.sale);
    broadcastInventory(readDb(), 'SALE_COMMITTED', { saleId: result.body.sale.id, terminalCode: result.body.sale.terminalCode });
    emitTickets(readDb(), 'SALE_COMMITTED', { saleId: result.body.sale.id, ticketId: result.body.sale.ticketId });
    return res.json(result.body);
  } catch (err) {
    console.error('Error en venta segura v13.6.31:', err);
    return res.status(500).json({ ok: false, message: err.message || 'No se pudo registrar la venta.' });
  }
});
app.get('/api/sales', auth, (_req, res) => {
  res.json({ ok: true, sales: readDb().sales.slice().reverse() });
});

app.delete('/api/sales', auth, requireAdmin, (req, res) => {
  const db = readDb();
  const summary = {
    removedSales: (db.sales || []).length,
    removedConflicts: (db.saleConflicts || []).length,
    removedOpenTickets: (db.openTickets || []).length,
    removedCashOpenings: (db.cashOpenings || []).length,
    removedCashEntries: (db.cashEntries || []).length,
    removedCashClosings: (db.cashClosings || []).length,
    removedProcurementPurchases: (db.procurementPurchases || []).length,
    removedEnterpriseExpenses: (db.enterpriseExpenses || []).length,
    removedInventoryMovements: (db.inventoryAdjustments || []).length,
    removedInvoices: (db.invoices || []).length
  };

  db.sales = [];
  db.saleConflicts = [];
  db.openTickets = [];
  db.cashOpenings = [];
  db.cashEntries = [];
  db.cashClosings = [];
  db.procurementPurchases = [];
  db.enterpriseExpenses = [];
  db.inventoryAdjustments = [];
  db.invoices = [];
  db.sellerSessions = [];
  db.counters = {
    ...(db.counters || {}),
    salesSequence: 0,
    ownerConsumptionSequence: 0,
    purchaseSequence: 0
  };

  writeDb(db);
  io.emit('sale:history-cleared', {
    ok: true,
    ...summary,
    removedCount: summary.removedSales,
    clearedConflictsCount: summary.removedConflicts,
    clearedAt: new Date().toISOString()
  });
  emitTickets(db, 'HISTORY_CLEARED', { cleared: true });
  broadcastInventory(db, 'HISTORY_CLEARED', { cleared: true });

  res.json({ ok: true, removedCount: summary.removedSales, clearedConflictsCount: summary.removedConflicts, summary });
});


app.get('/api/employees', auth, (req, res) => {
  const db = readDb();
  const allEmployees = (db.users || [])
    .filter((u) => u.role === 'POS')
    .sort((a, b) => String(a.name || '').localeCompare(String(b.name || ''), 'es'));

  const employees = req.user?.role === 'ADMIN'
    ? allEmployees
    : allEmployees.filter((u) => u.active !== false);

  res.json({
    ok: true,
    employees: employees.map((u) => ({
      id: u.id,
      name: u.name,
      employeeNumber: u.employeeNumber || '',
      username: u.username || '',
      active: u.active !== false,
      createdAt: u.createdAt || '',
      updatedAt: u.updatedAt || ''
    }))
  });
});

app.post('/api/employees', auth, requireAdmin, (req, res) => {
  try {
    const db = readDb();
    const payload = normalizeEmployeePayload(req.body || {});
    const duplicateNumber = (db.users || []).find((u) => u.role === 'POS' && u.employeeNumber === payload.employeeNumber && u.id !== payload.id);
    if (duplicateNumber) {
      return res.status(400).json({ ok: false, message: 'El número de empleado ya existe' });
    }

    const duplicateUsername = (db.users || []).find((u) => u.username === payload.username && u.id !== payload.id);
    if (duplicateUsername) {
      return res.status(400).json({ ok: false, message: 'El usuario generado para este empleado ya existe' });
    }

    const existing = (db.users || []).find((u) => u.id === payload.id && u.role === 'POS');
    if (existing) {
      Object.assign(existing, { ...payload, createdAt: existing.createdAt || payload.createdAt });
    } else {
      db.users.push(payload);
    }

    writeDb(db);
    res.json({ ok: true, employee: payload });
  } catch (err) {
    res.status(400).json({ ok: false, message: err.message });
  }
});

app.delete('/api/employees/:id', auth, requireAdmin, (req, res) => {
  const db = readDb();
  const employee = (db.users || []).find((u) => u.id === req.params.id && u.role === 'POS');
  if (!employee) {
    return res.status(404).json({ ok: false, message: 'Empleado no encontrado' });
  }

  employee.active = false;
  employee.updatedAt = new Date().toISOString();
  writeDb(db);
  res.json({ ok: true });
});

app.get('/api/reports/employees', auth, requireAdmin, (req, res) => {
  const db = readDb();
  const { range = 'day', anchorDate = '', from = '', to = '' } = req.query || {};

  let start;
  let end;
  if (from && to) {
    start = parseDateInput(from, new Date());
    end = parseDateInput(to, new Date());
    end.setDate(end.getDate() + 1);
  } else {
    ({ start, end } = getRangeBounds(range, anchorDate));
  }

  const salesInRange = (db.sales || []).filter((sale) => {
    if (isOwnerConsumptionSale(sale)) return false;
    const d = new Date(sale.createdAt);
    return d >= start && d < end;
  });

  const map = new Map();
  for (const sale of salesInRange) {
    const key = sale.operatorId || `${sale.employeeNumber || ''}-${sale.operatorName || 'SIN_OPERADOR'}`;
    if (!map.has(key)) {
      map.set(key, {
        operatorId: sale.operatorId || '',
        name: sale.operatorName || 'Sin asignar',
        employeeNumber: sale.employeeNumber || '',
        salesCount: 0,
        total: 0
      });
    }
    const row = map.get(key);
    row.salesCount += 1;
    row.total = round2(row.total + safeNumber(sale.total));
  }

  const rows = Array.from(map.values()).sort((a, b) => b.total - a.total);
  res.json({
    ok: true,
    range,
    from: start.toISOString(),
    to: end.toISOString(),
    rows,
    totals: {
      salesCount: salesInRange.length,
      total: round2(salesInRange.reduce((acc, sale) => acc + safeNumber(sale.total), 0))
    }
  });
});

app.delete('/api/sales/:id', auth, requireAdmin, (req, res) => {
  const db = readDb();
  const sale = db.sales.find((s) => s.id === req.params.id);

  if (!sale) {
    return res.status(404).json({ ok: false, message: 'Venta no encontrada' });
  }

  const productMap = new Map(db.products.map((p) => [p.id, p]));
  for (const item of normalizeTicketItems(sale.items || [])) {
    const product = productMap.get(item.productId);
    if (product) {
      product.stock = roundQuantity(safeNumber(product.stock) + safeNumber(item.quantity));
      product.updatedAt = new Date().toISOString();
    }
  }

  db.sales = db.sales.filter((s) => s.id !== req.params.id);
  db.cashEntries = (db.cashEntries || []).filter((entry) => entry.id !== sale.cashEntryId && entry.saleId !== sale.id);
  db.inventoryAdjustments = (db.inventoryAdjustments || []).filter((entry) => entry.saleId !== sale.id);
  writeDb(db);
  broadcastInventory(db, 'SALE_DELETED', { saleId: req.params.id });

  res.json({ ok: true });
});


app.get('/api/pos/seller-sessions', auth, (req, res) => {
  const db = readDb();
  ensureSellerSessions(db);
  res.json({ ok: true, sessions: db.sellerSessions || [] });
});

app.get('/api/pos/seller-sessions/:operatorId', auth, (req, res) => {
  const db = readDb();
  ensureSellerSessions(db);
  const session = getSellerSessionByOperator(db, req.params.operatorId);
  if (!session) return res.status(404).json({ ok: false, message: 'Sesión de vendedor no encontrada.' });
  res.json({ ok: true, session });
});

app.post('/api/pos/seller-sessions/upsert', auth, (req, res) => {
  try {
    const db = readDb();
    const session = upsertSellerSession(db, req.body || {}, req.user || null);
    writeDb(db);
    res.json({ ok: true, session, sessions: db.sellerSessions || [] });
  } catch (err) {
    res.status(400).json({ ok: false, message: err.message });
  }
});

app.post('/api/pos/seller-sessions/:operatorId/clear', auth, (req, res) => {
  const db = readDb();
  const session = clearSellerSession(db, req.params.operatorId, String(req.body?.status || 'CLEARED').trim() || 'CLEARED');
  if (!session) return res.status(404).json({ ok: false, message: 'Sesión de vendedor no encontrada.' });
  writeDb(db);
  res.json({ ok: true, session, sessions: db.sellerSessions || [] });
});

app.post('/api/pos/seller-sessions/:operatorId/release', auth, requireAdmin, (req, res) => {
  const db = readDb();
  const session = getSellerSessionByOperator(db, req.params.operatorId);
  if (!session) return res.status(404).json({ ok: false, message: 'Sesión de vendedor no encontrada.' });
  session.takenBy = '';
  session.takenByUsername = '';
  session.status = 'IDLE';
  session.updatedAt = new Date().toISOString();
  session.lastActivityAt = new Date().toISOString();
  writeDb(db);
  res.json({ ok: true, session, sessions: db.sellerSessions || [] });
});

app.get('/api/terminal/view-permissions', (req, res) => {
  const db = readDb();
  const nodeId = normalizeNodeId13640(req.headers['x-terminal-node-id'] || '');
  const terminalCode = normalizeTerminalCode13640(req.headers['x-terminal-code'] || '');
  const terminal = findTerminalForNode13640(db, nodeId, terminalCode);
  if (!terminal) {
    return res.status(404).json({ ok: false, code: 'NODE_NOT_REGISTERED', message: 'Terminal no encontrada.' });
  }
  const before = terminal.viewPermissions && typeof terminal.viewPermissions === 'object' ? { ...terminal.viewPermissions } : null;
  const viewPermissions = ensureTerminalViewPermissions13641(terminal);
  if (!before || !terminalViewPermissionsEqual13642(before, viewPermissions)) {
    terminal.permissionsRevision = Math.max(1, Number(terminal.permissionsRevision || 0));
    terminal.permissionsUpdatedAt ||= new Date().toISOString();
    writeDb(db);
  }
  res.json({
    ok: true,
    terminalId: terminal.id || '',
    terminalCode: terminal.terminalCode || '',
    nodeId: terminal.nodeId || '',
    viewPermissions,
    permissionsRevision: Number(terminal.permissionsRevision || 0),
    permissionsUpdatedAt: terminal.permissionsUpdatedAt || ''
  });
});

app.get('/api/terminals', auth, requireAdmin, (_req, res) => {
  const now = Date.now();
  const db = readDb();
  let permissionsNormalized = false;
  const terminals = (db.terminals || []).map((terminal) => {
    const seenMs = Date.parse(terminal.lastSeenAt || '') || 0;
    const online = seenMs > 0 && (now - seenMs) <= 90 * 1000;
    const lastIp = normalizeLanIp(terminal.lastIp || '');
    const fixedIp = normalizeLanIp(terminal.fixedIp || '');
    const pairingExpiresMs = Date.parse(terminal.pairingCodeExpiresAt || '') || 0;
    const isAdminTerminal = String(terminal.terminalType || '').toUpperCase() === 'ADMIN';
    const pairingPending = !isAdminTerminal && (terminal.authStatus === 'PENDING' || terminal.authorized !== true);
    const beforePermissions = terminal.viewPermissions && typeof terminal.viewPermissions === 'object'
      ? { ...terminal.viewPermissions }
      : null;
    const viewPermissions = ensureTerminalViewPermissions13641(terminal);
    if (!beforePermissions || !terminalViewPermissionsEqual13642(beforePermissions, viewPermissions)) {
      permissionsNormalized = true;
      terminal.permissionsUpdatedAt ||= new Date().toISOString();
      terminal.permissionsRevision = Math.max(1, Number(terminal.permissionsRevision || 0));
    }
    const {
      pairingTokenHash: _pairingTokenHash,
      pairingCodeHash: _pairingCodeHash,
      ...safeTerminal
    } = terminal;
    return {
      ...safeTerminal,
      lastIp,
      fixedIp,
      online,
      ipMatches: !fixedIp || !lastIp || fixedIp === lastIp,
      authStatus: isAdminTerminal ? 'ADMIN' : (terminal.authStatus || (terminal.authorized === true ? 'AUTHORIZED' : 'PENDING')),
      pairingRequired: pairingPending,
      viewPermissions,
      pairingCode: pairingPending && pairingExpiresMs > Date.now() ? String(terminal.pairingCode || '') : '',
      pairingCodeExpiresAt: pairingPending ? (terminal.pairingCodeExpiresAt || '') : ''
    };
  });
  // Si una terminal antigua todavía no tenía viewPermissions, persistimos el valor migrado
  // para que el siguiente reinicio no dependa de valores calculados únicamente en memoria.
  if (permissionsNormalized) writeDb(db);
  res.json({ ok: true, terminals });
});

app.get('/api/network/lan-info', auth, requireAdmin, (_req, res) => {
  const interfaces = [];
  const nets = os.networkInterfaces();
  for (const [name, entries] of Object.entries(nets || {})) {
    for (const net of entries || []) {
      if (net.family === 'IPv4' && !net.internal) {
        interfaces.push({ name, address: net.address, url: `http://${net.address}:${PORT}` });
      }
    }
  }
  res.json({
    ok: true,
    adminName: ADMIN_NAME,
    port: PORT,
    discoveryPort: DISCOVERY_PORT,
    interfaces,
    offlineLanSupported: true
  });
});

app.post('/api/terminals/:code/view-permissions', auth, requireAdmin, (req, res) => {
  const db = readDb();
  const identifier = String(req.params.code || '').trim();
  const terminal = findTerminalForPermissions13642(db, identifier);
  if (!terminal) return res.status(404).json({ ok: false, message: 'Terminal no encontrada.' });

  const canonicalId = String(terminal.id || '');
  const canonicalCode = String(terminal.terminalCode || '').trim();
  const canonicalNodeId = String(terminal.nodeId || '').trim();

  if (String(terminal.terminalType || '').toUpperCase() === 'ADMIN' || canonicalCode === 'ADMIN-01') {
    terminal.viewPermissions = defaultTerminalViewPermissions13641('ADMIN');
    terminal.permissionsRevision = Number(terminal.permissionsRevision || 0) + 1;
    terminal.permissionsUpdatedAt = new Date().toISOString();
    writeDb(db);
    return res.json({
      ok: true,
      persisted: true,
      terminalId: canonicalId,
      terminalCode: canonicalCode,
      nodeId: canonicalNodeId,
      viewPermissions: terminal.viewPermissions,
      permissionsRevision: terminal.permissionsRevision,
      permissionsUpdatedAt: terminal.permissionsUpdatedAt,
      protected: true
    });
  }

  const incoming = req.body?.permissions;
  if (!incoming || typeof incoming !== 'object') {
    return res.status(400).json({ ok: false, message: 'Permisos de visualización inválidos.' });
  }
  const normalized = normalizeTerminalViewPermissions13641(incoming, terminal.terminalType || 'POS');
  if (!Object.values(normalized).some(Boolean)) {
    return res.status(400).json({ ok: false, message: 'La terminal debe tener al menos un módulo visible.' });
  }

  terminal.viewPermissions = { ...normalized };
  terminal.permissionsRevision = Number(terminal.permissionsRevision || 0) + 1;
  terminal.permissionsUpdatedAt = new Date().toISOString();
  terminal.updatedAt = terminal.permissionsUpdatedAt;
  db.audit ||= [];
  db.audit.push({
    id: uuidv4(),
    type: 'TERMINAL_VIEW_PERMISSIONS_CHANGED',
    terminalId: canonicalId,
    terminalCode: canonicalCode,
    nodeId: canonicalNodeId,
    terminalName: terminal.terminalName || '',
    permissions: { ...normalized },
    revision: terminal.permissionsRevision,
    changedBy: req.user?.username || '',
    createdAt: terminal.permissionsUpdatedAt
  });

  writeDb(db);

  // Verificación inmediata contra el archivo ya escrito. No reportamos éxito si la
  // configuración no puede recuperarse exactamente desde la base persistente.
  const verifyDb = readDb();
  const verifiedTerminal = findTerminalForPermissions13642(verifyDb, canonicalId || canonicalNodeId || canonicalCode);
  const verifiedPermissions = verifiedTerminal
    ? ensureTerminalViewPermissions13641(verifiedTerminal)
    : null;
  if (!verifiedTerminal || !verifiedPermissions || !terminalViewPermissionsEqual13642(verifiedPermissions, normalized)) {
    return res.status(500).json({
      ok: false,
      code: 'TERMINAL_PERMISSIONS_NOT_PERSISTED',
      message: 'Los permisos no pudieron verificarse después de guardarlos. No se aplicó un éxito falso.'
    });
  }

  const eventPayload = {
    terminalId: canonicalId,
    terminalCode: canonicalCode,
    nodeId: canonicalNodeId,
    viewPermissions: verifiedPermissions,
    permissionsRevision: Number(verifiedTerminal.permissionsRevision || terminal.permissionsRevision || 0),
    permissionsUpdatedAt: verifiedTerminal.permissionsUpdatedAt || terminal.permissionsUpdatedAt
  };
  io.emit('terminal-permissions:update', eventPayload);
  io.emit('terminals:update', {
    ok: true,
    reason: 'VIEW_PERMISSIONS_CHANGED',
    terminalId: canonicalId,
    terminalCode: canonicalCode,
    nodeId: canonicalNodeId,
    permissionsRevision: eventPayload.permissionsRevision
  });
  return res.json({ ok: true, persisted: true, ...eventPayload });
});

app.post('/api/terminals/bindings', auth, requireAdmin, (req, res) => {
  const db = readDb();
  const body = req.body || {};
  const terminalCode = String(body.terminalCode || '').trim();
  const fixedIp = normalizeLanIp(body.fixedIp || '');
  if (!terminalCode) return res.status(400).json({ ok: false, message: 'El código de terminal es obligatorio.' });
  if (!isValidLanIpv4(fixedIp)) return res.status(400).json({ ok: false, message: 'Captura una IPv4 válida para el punto de venta.' });

  let terminal = (db.terminals || []).find((entry) => entry.id === body.id || entry.terminalCode === terminalCode);
  if (!terminal) {
    terminal = {
      id: uuidv4(),
      terminalCode,
      terminalType: 'POS',
      terminalName: String(body.terminalName || terminalCode).trim() || terminalCode,
      lastSeenAt: '',
      lastIp: '',
      fixedIp,
      connectionMode: 'FIXED_IP',
      enabled: body.enabled !== false,
      createdAt: new Date().toISOString()
    };
    ensureTerminalViewPermissions13641(terminal);
    db.terminals.push(terminal);
  } else {
    terminal.terminalCode = terminalCode;
    terminal.terminalName = String(body.terminalName || terminal.terminalName || terminalCode).trim() || terminalCode;
    terminal.fixedIp = fixedIp;
    terminal.connectionMode = 'FIXED_IP';
    terminal.enabled = body.enabled !== false;
    terminal.updatedAt = new Date().toISOString();
    terminal.ipMatches = !terminal.lastIp || normalizeLanIp(terminal.lastIp) === fixedIp;
    ensureTerminalViewPermissions13641(terminal);
  }
  writeDb(db);
  res.json({ ok: true, terminal });
});

app.post('/api/terminals/:code/use-detected-ip', auth, requireAdmin, (req, res) => {
  const db = readDb();
  const code = String(req.params.code || '').trim();
  const terminal = (db.terminals || []).find((entry) => entry.terminalCode === code);
  if (!terminal) return res.status(404).json({ ok: false, message: 'Terminal no encontrada.' });
  const detectedIp = normalizeLanIp(terminal.lastIp || '');
  if (!isValidLanIpv4(detectedIp)) return res.status(400).json({ ok: false, message: 'La terminal todavía no tiene una IP LAN detectada.' });
  terminal.fixedIp = detectedIp;
  terminal.connectionMode = 'FIXED_IP';
  terminal.enabled = true;
  terminal.ipMatches = true;
  terminal.updatedAt = new Date().toISOString();
  writeDb(db);
  res.json({ ok: true, terminal });
});

app.delete('/api/terminals/:code/binding', auth, requireAdmin, (req, res) => {
  const db = readDb();
  const code = String(req.params.code || '').trim();
  const terminal = (db.terminals || []).find((entry) => entry.terminalCode === code);
  if (!terminal) return res.status(404).json({ ok: false, message: 'Terminal no encontrada.' });
  terminal.fixedIp = '';
  terminal.connectionMode = 'AUTO';
  terminal.ipMatches = true;
  terminal.updatedAt = new Date().toISOString();
  writeDb(db);
  res.json({ ok: true, terminal });
});

app.delete('/api/terminals/:terminalCode', auth, requireAdmin, (req, res) => {
  const terminalCode = String(req.params.terminalCode || '').trim();
  if (!terminalCode) {
    return res.status(400).json({ ok: false, message: 'Código de terminal requerido.' });
  }

  const db = readDb();
  db.terminals ||= [];

  const terminal = db.terminals.find((t) => String(t.terminalCode || '') === terminalCode);
  if (!terminal) {
    return res.status(404).json({ ok: false, message: 'La terminal/nodo no existe o ya fue eliminado.' });
  }

  const currentTerminalCode = String(req.headers['x-terminal-code'] || '').trim();
  const isCurrentTerminal = currentTerminalCode && currentTerminalCode === terminalCode;
  const isAdminNode = String(terminal.terminalType || '').toUpperCase() === 'ADMIN';
  const isCoreAdminCode = terminalCode === 'ADMIN-01';

  if (isCurrentTerminal || isCoreAdminCode || isAdminNode) {
    return res.status(400).json({
      ok: false,
      message: 'No se puede eliminar la terminal administradora actual. Solo elimina nodos POS antiguos o duplicados.'
    });
  }

  const before = db.terminals.length;
  db.terminals = db.terminals.filter((t) => String(t.terminalCode || '') !== terminalCode);
  db.saleConflicts = (db.saleConflicts || []).filter((c) => String(c.terminalCode || '') !== terminalCode);
  db.audit ||= [];
  db.audit.push({
    id: uuidv4(),
    type: 'TERMINAL_DELETED',
    terminalCode,
    terminalName: terminal.terminalName || '',
    deletedBy: req.user?.username || '',
    createdAt: new Date().toISOString()
  });
  writeDb(db);

  io.emit('terminals:update', { ok: true, reason: 'TERMINAL_DELETED', terminalCode, terminals: db.terminals });

  return res.json({
    ok: true,
    deleted: before - db.terminals.length,
    terminalCode,
    terminals: db.terminals,
    conflicts: db.saleConflicts || []
  });
});

app.get('/api/conflicts', auth, requireAdmin, (_req, res) => {
  res.json({ ok: true, conflicts: readDb().saleConflicts.slice().reverse() });
});

app.get('/api/security/admin', auth, requireAdmin, (_req, res) => {
  const db = readDb();
  const admin = db.users.find((u) => u.role === 'ADMIN');
  res.json({
    ok: true,
    admin: {
      username: admin?.username || 'admin',
      name: admin?.name || 'Administrador General',
      requireAdminPasswordForEdits: false,
      remoteApiKey: db.security?.remoteAccess?.apiKey || '',
      allowRemoteAdminCredentials: db.security?.remoteAccess?.allowAdminCredentials !== false,
      allowRemoteInventoryWrite: db.security?.remoteAccess?.permissions?.writeInventory !== false,
      restockWhatsappNumber: db.security?.notifications?.whatsappRestockNumber || ''
    }
  });
});

app.post('/api/security/admin', auth, requireAdmin, (req, res) => {
  const {
    currentPassword,
    username,
    newPassword,
    confirmPassword,
    name,
    requireAdminPasswordForEdits,
    remoteApiKey,
    allowRemoteAdminCredentials,
    allowRemoteInventoryWrite,
    restockWhatsappNumber
  } = req.body || {};
  const db = readDb();
  const admin = db.users.find((u) => u.role === 'ADMIN');

  if (!admin) {
    return res.status(404).json({ ok: false, message: 'Administrador no encontrado' });
  }

  // v13.6.27: no se pide contraseña adicional para modificar ajustes desde una sesión ADMIN activa.

  if (newPassword && newPassword !== confirmPassword) {
    return res.status(400).json({ ok: false, message: 'La confirmación de contraseña no coincide' });
  }

  if (username && String(username).trim()) {
    admin.username = String(username).trim();
  }
  if (name && String(name).trim()) {
    admin.name = String(name).trim();
  }
  if (newPassword && String(newPassword).trim()) {
    admin.password = String(newPassword).trim();
  }

  db.security = {
    ...(db.security || {}),
    requireAdminPasswordForEdits: false,
    remoteAccess: {
      apiKey: String(remoteApiKey || db.security?.remoteAccess?.apiKey || '').trim() || `REMOTE-${uuidv4().replace(/-/g, '').slice(0, 24).toUpperCase()}`,
      allowAdminCredentials: allowRemoteAdminCredentials !== false,
      permissions: {
        readSummary: true,
        readInventory: true,
        writeInventory: allowRemoteInventoryWrite !== false,
        syncBridge: true
      }
    },
    notifications: {
      whatsappRestockNumber: String(restockWhatsappNumber || '').trim()
    },
    updatedAt: new Date().toISOString()
  };

  writeDb(db);
  res.json({
    ok: true,
    admin: {
      username: admin.username,
      name: admin.name,
      requireAdminPasswordForEdits: false,
      remoteApiKey: db.security?.remoteAccess?.apiKey || '',
      allowRemoteAdminCredentials: db.security?.remoteAccess?.allowAdminCredentials !== false,
      allowRemoteInventoryWrite: db.security?.remoteAccess?.permissions?.writeInventory !== false,
      restockWhatsappNumber: db.security?.notifications?.whatsappRestockNumber || ''
    }
  });
});


app.get('/api/branch-profile', auth, requireAdmin, (_req, res) => {
  const db = readDb();
  res.json({
    ok: true,
    profile: getBranchProfile(db),
    publication: getPublicationConfig(db),
    mirror: {
      connected: branchRelayClientStatus.connected,
      status: branchRelayClientStatus.status,
      lastError: branchRelayClientStatus.lastError
    }
  });
});

app.post('/api/branch-profile', auth, requireAdmin, (req, res) => {
  const db = readDb();
  const body = req.body || {};
  const currentProfile = getBranchProfile(db);
  const currentPublication = getPublicationConfig(db);
  db.branchProfile = {
    ...currentProfile,
    name: String(body.name || currentProfile.name || ADMIN_NAME).trim() || ADMIN_NAME,
    code: String(body.code || currentProfile.code || 'ADMIN-01').trim() || 'ADMIN-01',
    publicUrl: String(body.publicUrl || currentProfile.publicUrl || '').trim().replace(/\/$/, ''),
    mode: String(body.mode || currentProfile.mode || 'LOCAL').trim() || 'LOCAL',
    description: String(body.description || currentProfile.description || '').trim(),
    updatedAt: new Date().toISOString()
  };
  db.publication = {
    ...currentPublication,
    enabled: body.enabled === true,
    mirrorUrl: String(body.mirrorUrl || currentPublication.mirrorUrl || '').trim().replace(/\/$/, ''),
    authType: 'token',
    apiKey: String(body.apiKey || currentPublication.apiKey || '').trim(),
    remoteAdminUsername: String(body.remoteAdminUsername || currentPublication.remoteAdminUsername || 'admin').trim() || 'admin',
    remoteAdminPassword: String(body.remoteAdminPassword || currentPublication.remoteAdminPassword || '').trim(),
    autoSyncEnabled: body.autoSyncEnabled === true,
    syncIntervalSec: Math.max(15, Math.min(600, safeNumber(body.syncIntervalSec, currentPublication.syncIntervalSec || 60))),
    lastSyncAt: currentPublication.lastSyncAt || '',
    lastSyncStatus: currentPublication.lastSyncStatus || 'IDLE',
    lastSyncError: currentPublication.lastSyncError || '',
    lastCommandPollAt: currentPublication.lastCommandPollAt || '',
    lastCommandApplyAt: currentPublication.lastCommandApplyAt || '',
    updatedAt: new Date().toISOString()
  };
  writeDb(db);
  configureBranchRelayClient(db);
  return res.json({ ok: true, profile: db.branchProfile, publication: db.publication });
});

app.post('/api/branch-profile/test-publication', auth, requireAdmin, async (_req, res) => {
  const db = readDb();
  const publication = getPublicationConfig(db);
  const profile = getBranchProfile(db);
  try {
    if (!publication.mirrorUrl) throw new Error('La URL del puente Render está vacía');
    await waitForBranchRelayConnection(db, 120000);
    const info = await fetchRemoteBranch({
      remoteUrl: publication.mirrorUrl,
      branchCode: profile.code,
      authType: publication.authType,
      apiKey: publication.apiKey,
      remoteAdminUsername: publication.remoteAdminUsername,
      remoteAdminPassword: publication.remoteAdminPassword
    }, '/api/remote/publication-info');
    return res.json({ ok: true, remote: info });
  } catch (err) {
    return res.status(502).json({ ok: false, message: `No se pudo validar la publicación remota: ${err.message}` });
  }
});

app.post('/api/branch-profile/sync-now', auth, requireAdmin, async (_req, res) => {
  const db = readDb();
  const publication = getPublicationConfig(db);
  if (!publication.enabled) {
    return res.status(400).json({ ok: false, message: 'La publicación remota está deshabilitada' });
  }
  if (!publication.mirrorUrl) {
    return res.status(400).json({ ok: false, message: 'Falta la URL del puente Render' });
  }

  try {
    await waitForBranchRelayConnection(db, 120000);
    db.publication.lastSyncAt = new Date().toISOString();
    db.publication.lastSyncStatus = 'ONLINE';
    db.publication.lastSyncError = '';
    db.publication.updatedAt = new Date().toISOString();
    writeDb(db);
    return res.json({ ok: true, connected: true, mirror: { ...branchRelayClientStatus }, publication: db.publication });
  } catch (err) {
    db.publication.lastSyncAt = new Date().toISOString();
    db.publication.lastSyncStatus = 'ERROR';
    db.publication.lastSyncError = err.message;
    db.publication.updatedAt = new Date().toISOString();
    writeDb(db);
    return res.status(502).json({ ok: false, message: `No se pudo conectar al puente Render: ${err.message}`, publication: db.publication });
  }
});


app.get('/api/branches', auth, requireAdmin, (_req, res) => {
  const db = readDb();
  res.json({ ok: true, branches: (db.branches || []).map(getBranchSummarySkeleton) });
});

app.post('/api/branches', auth, requireAdmin, (req, res) => {
  const db = readDb();
  const body = req.body || {};
  const name = String(body.name || '').trim();
  const code = String(body.code || '').trim();
  const remoteUrl = String(body.remoteUrl || '').trim().replace(/\/$/, '');
  const authType = body.authType === 'token' ? 'token' : 'credentials';
  const submittedApiKey = String(body.apiKey || '').trim();
  const remoteAdminUsername = String(body.remoteAdminUsername || '').trim();
  const remoteAdminPassword = String(body.remoteAdminPassword || '').trim();
  const existing = (db.branches || []).find((b) => b.id === body.id);
  const color = /^#[0-9a-f]{6}$/i.test(String(body.color || ''))
    ? String(body.color).toLowerCase()
    : (existing?.color || '#1f6feb');

  if (!name) return res.status(400).json({ ok: false, message: 'El nombre de la sucursal es obligatorio' });
  if (!code) return res.status(400).json({ ok: false, message: 'El código de la sucursal es obligatorio' });
  if (!remoteUrl) return res.status(400).json({ ok: false, message: 'La URL remota es obligatoria' });
  try {
    validateRemoteBranchUrl(remoteUrl);
  } catch (err) {
    return res.status(400).json({ ok: false, message: err.message });
  }
  const apiKey = submittedApiKey || existing?.apiKey || '';
  if (authType === 'token' && !apiKey) return res.status(400).json({ ok: false, message: 'El token/API key remoto es obligatorio' });
  if (authType === 'credentials' && (!remoteAdminUsername || !(remoteAdminPassword || body.id))) {
    return res.status(400).json({ ok: false, message: 'Las credenciales del administrador remoto son obligatorias' });
  }

  const branch = {
    id: existing?.id || uuidv4(),
    name,
    code,
    color,
    remoteUrl,
    authType,
    apiKey: authType === 'token' ? apiKey : '',
    remoteAdminUsername: authType === 'credentials' ? remoteAdminUsername : '',
    remoteAdminPassword: authType === 'credentials'
      ? (remoteAdminPassword || existing?.remoteAdminPassword || '')
      : '',
    mode: body.mode || existing?.mode || 'render',
    active: body.active !== false,
    lastStatus: existing?.lastStatus || 'PENDING',
    lastCheckedAt: existing?.lastCheckedAt || '',
    lastError: existing?.lastError || '',
    lastRemoteSummary: existing?.lastRemoteSummary || null,
    createdAt: existing?.createdAt || new Date().toISOString(),
    updatedAt: new Date().toISOString()
  };

  if (existing) {
    Object.assign(existing, branch);
  } else {
    db.branches.push(branch);
  }

  writeDb(db);
  res.json({ ok: true, branch: getBranchSummarySkeleton(branch) });
});

app.delete('/api/branches/:id', auth, requireAdmin, (req, res) => {
  const db = readDb();
  db.branches = (db.branches || []).filter((b) => b.id !== req.params.id);
  writeDb(db);
  res.json({ ok: true });
});

app.post('/api/branches/test', auth, requireAdmin, async (req, res) => {
  try {
    const existing = (readDb().branches || []).find((row) => row.id === req.body?.id);
    const candidate = {
      remoteUrl: req.body?.remoteUrl,
      branchCode: req.body?.code || existing?.code,
      authType: req.body?.authType,
      apiKey: req.body?.apiKey || existing?.apiKey,
      remoteAdminUsername: req.body?.remoteAdminUsername || existing?.remoteAdminUsername,
      remoteAdminPassword: req.body?.remoteAdminPassword || existing?.remoteAdminPassword
    };
    const data = await fetchRemoteBranchSummary(candidate);
    res.json({ ok: true, remote: data });
  } catch (err) {
    res.status(502).json({ ok: false, message: `No se pudo conectar a la sucursal remota: ${err.message}` });
  }
});

app.get('/api/branches/:id/summary', auth, requireAdmin, async (req, res) => {
  const db = readDb();
  const branch = (db.branches || []).find((b) => b.id === req.params.id);
  if (!branch) return res.status(404).json({ ok: false, message: 'Sucursal no encontrada' });

  try {
    const data = await fetchRemoteBranchSummary(branch);
    branch.lastStatus = 'ONLINE';
    branch.lastCheckedAt = new Date().toISOString();
    branch.lastError = '';
    branch.lastRemoteSummary = data;
    writeDb(db);
    return res.json({ ok: true, branch: getBranchSummarySkeleton(branch), remote: data });
  } catch (err) {
    branch.lastStatus = 'OFFLINE';
    branch.lastCheckedAt = new Date().toISOString();
    branch.lastError = err.message;
    writeDb(db);
    return res.status(502).json({
      ok: false,
      message: `La sucursal no respondió: ${err.message}`,
      branch: getBranchSummarySkeleton(branch)
    });
  }
});


app.get('/api/branches/:id/inventory', auth, requireAdmin, async (req, res) => {
  const db = readDb();
  const branch = (db.branches || []).find((b) => b.id === req.params.id);
  if (!branch) return res.status(404).json({ ok: false, message: 'Sucursal no encontrada' });

  try {
    const query = new URLSearchParams();
    if (req.query?.query) query.set('query', String(req.query.query));
    if (req.query?.category) query.set('category', String(req.query.category));
    if (String(req.query?.lowOnly || '') === '1') query.set('lowOnly', '1');
    const data = await fetchRemoteBranch(branch, `/api/remote/inventory${query.toString() ? `?${query.toString()}` : ''}`);
    branch.lastStatus = 'ONLINE';
    branch.lastCheckedAt = new Date().toISOString();
    branch.lastError = '';
    writeDb(db);
    return res.json({ ok: true, branch: getBranchSummarySkeleton(branch), inventory: data });
  } catch (err) {
    branch.lastStatus = 'OFFLINE';
    branch.lastCheckedAt = new Date().toISOString();
    branch.lastError = err.message;
    writeDb(db);
    return res.status(502).json({ ok: false, message: `No se pudo consultar el inventario remoto: ${err.message}` });
  }
});

app.post('/api/branches/:id/inventory/adjust', auth, requireAdmin, async (req, res) => {
  const db = readDb();
  const branch = (db.branches || []).find((b) => b.id === req.params.id);
  if (!branch) return res.status(404).json({ ok: false, message: 'Sucursal no encontrada' });
  return res.status(403).json({ ok: false, message: 'El puente de sucursales es de solo consulta. Realiza los ajustes en la computadora de esa sucursal.' });
});

app.get('/api/branches/:id/session/connect', auth, requireAdmin, async (req, res) => {
  const db = readDb();
  const branch = (db.branches || []).find((b) => b.id === req.params.id);
  if (!branch) return res.status(404).json({ ok: false, message: 'Sucursal no encontrada' });
  try {
    const data = await fetchRemoteBranchPublicationInfo(branch);
    branch.lastStatus = 'ONLINE';
    branch.lastCheckedAt = new Date().toISOString();
    branch.lastError = '';
    writeDb(db);
    return res.json({ ok: true, branch: getBranchSummarySkeleton(branch), remote: data });
  } catch (err) {
    branch.lastStatus = 'OFFLINE';
    branch.lastCheckedAt = new Date().toISOString();
    branch.lastError = err.message;
    writeDb(db);
    return res.status(502).json({ ok: false, message: `No se pudo conectar con la sucursal: ${err.message}` });
  }
});

app.get('/api/branches/:id/session/cash-overview', auth, requireAdmin, async (req, res) => {
  const db = readDb();
  const branch = (db.branches || []).find((b) => b.id === req.params.id);
  if (!branch) return res.status(404).json({ ok: false, message: 'Sucursal no encontrada' });
  try {
    const params = new URLSearchParams({ date: normalizeDateKey(req.query?.date) });
    const data = await fetchRemoteBranch(branch, `/api/remote/cash-overview?${params.toString()}`);
    branch.lastStatus = 'ONLINE';
    branch.lastCheckedAt = new Date().toISOString();
    branch.lastError = '';
    writeDb(db);
    return res.json({ ok: true, branch: getBranchSummarySkeleton(branch), session: data });
  } catch (err) {
    branch.lastStatus = 'OFFLINE';
    branch.lastCheckedAt = new Date().toISOString();
    branch.lastError = err.message;
    writeDb(db);
    return res.status(502).json({ ok: false, message: `No se pudo consultar Caja en la sucursal: ${err.message}` });
  }
});

app.get('/api/branches/:id/session/users', auth, requireAdmin, async (req, res) => {
  const db = readDb();
  const branch = (db.branches || []).find((b) => b.id === req.params.id);
  if (!branch) return res.status(404).json({ ok: false, message: 'Sucursal no encontrada' });
  try {
    const data = await fetchRemoteBranch(branch, '/api/remote/users');
    branch.lastStatus = 'ONLINE';
    branch.lastCheckedAt = new Date().toISOString();
    branch.lastError = '';
    writeDb(db);
    return res.json({ ok: true, branch: getBranchSummarySkeleton(branch), session: data });
  } catch (err) {
    branch.lastStatus = 'OFFLINE';
    branch.lastCheckedAt = new Date().toISOString();
    branch.lastError = err.message;
    writeDb(db);
    return res.status(502).json({ ok: false, message: `No se pudieron consultar los usuarios de la sucursal: ${err.message}` });
  }
});

app.get('/api/branches/:id/session/sales', auth, requireAdmin, async (req, res) => {
  const db = readDb();
  const branch = (db.branches || []).find((b) => b.id === req.params.id);
  if (!branch) return res.status(404).json({ ok: false, message: 'Sucursal no encontrada' });
  try {
    const params = new URLSearchParams();
    if (req.query?.limit) params.set('limit', String(req.query.limit));
    if (req.query?.query) params.set('query', String(req.query.query));
    if (req.query?.employee) params.set('employee', String(req.query.employee));
    if (req.query?.customer) params.set('customer', String(req.query.customer));
    if (req.query?.range) params.set('range', String(req.query.range));
    if (req.query?.from) params.set('from', String(req.query.from));
    if (req.query?.to) params.set('to', String(req.query.to));
    const data = await fetchRemoteBranch(branch, `/api/remote/sales${params.toString() ? `?${params.toString()}` : ''}`);
    branch.lastStatus = 'ONLINE';
    branch.lastCheckedAt = new Date().toISOString();
    branch.lastError = '';
    writeDb(db);
    return res.json({ ok: true, branch: getBranchSummarySkeleton(branch), session: data });
  } catch (err) {
    branch.lastStatus = 'OFFLINE';
    branch.lastCheckedAt = new Date().toISOString();
    branch.lastError = err.message;
    writeDb(db);
    return res.status(502).json({ ok: false, message: `No se pudieron consultar las ventas remotas: ${err.message}` });
  }
});

app.get('/api/branches/:id/session/inventory', auth, requireAdmin, async (req, res) => {
  const db = readDb();
  const branch = (db.branches || []).find((b) => b.id === req.params.id);
  if (!branch) return res.status(404).json({ ok: false, message: 'Sucursal no encontrada' });
  try {
    const params = new URLSearchParams();
    if (req.query?.query) params.set('query', String(req.query.query));
    if (req.query?.category) params.set('category', String(req.query.category));
    if (req.query?.alertMode) params.set('alertMode', String(req.query.alertMode));
    const data = await fetchRemoteBranch(branch, `/api/remote/inventory${params.toString() ? `?${params.toString()}` : ''}`);
    branch.lastStatus = 'ONLINE';
    branch.lastCheckedAt = new Date().toISOString();
    branch.lastError = '';
    writeDb(db);
    return res.json({ ok: true, branch: getBranchSummarySkeleton(branch), session: data });
  } catch (err) {
    branch.lastStatus = 'OFFLINE';
    branch.lastCheckedAt = new Date().toISOString();
    branch.lastError = err.message;
    writeDb(db);
    return res.status(502).json({ ok: false, message: `No se pudo consultar el inventario remoto: ${err.message}` });
  }
});

app.get('/api/branches/:id/session/employee-report', auth, requireAdmin, async (req, res) => {
  const db = readDb();
  const branch = (db.branches || []).find((b) => b.id === req.params.id);
  if (!branch) return res.status(404).json({ ok: false, message: 'Sucursal no encontrada' });
  try {
    const params = new URLSearchParams();
    if (req.query?.range) params.set('range', String(req.query.range));
    if (req.query?.anchorDate) params.set('anchorDate', String(req.query.anchorDate));
    if (req.query?.from) params.set('from', String(req.query.from));
    if (req.query?.to) params.set('to', String(req.query.to));
    const data = await fetchRemoteBranch(branch, `/api/remote/employee-report${params.toString() ? `?${params.toString()}` : ''}`);
    branch.lastStatus = 'ONLINE';
    branch.lastCheckedAt = new Date().toISOString();
    branch.lastError = '';
    writeDb(db);
    return res.json({ ok: true, branch: getBranchSummarySkeleton(branch), session: data });
  } catch (err) {
    branch.lastStatus = 'OFFLINE';
    branch.lastCheckedAt = new Date().toISOString();
    branch.lastError = err.message;
    writeDb(db);
    return res.status(502).json({ ok: false, message: `No se pudo consultar el reporte remoto por empleado: ${err.message}` });
  }
});

app.get('/api/branches/:id/session/stock-alerts', auth, requireAdmin, async (req, res) => {
  const db = readDb();
  const branch = (db.branches || []).find((b) => b.id === req.params.id);
  if (!branch) return res.status(404).json({ ok: false, message: 'Sucursal no encontrada' });
  try {
    const data = await fetchRemoteBranchStockAlerts(branch);
    branch.lastStatus = 'ONLINE';
    branch.lastCheckedAt = new Date().toISOString();
    branch.lastError = '';
    writeDb(db);
    return res.json({ ok: true, branch: getBranchSummarySkeleton(branch), session: data });
  } catch (err) {
    branch.lastStatus = 'OFFLINE';
    branch.lastCheckedAt = new Date().toISOString();
    branch.lastError = err.message;
    writeDb(db);
    return res.status(502).json({ ok: false, message: `No se pudieron consultar las alertas remotas: ${err.message}` });
  }
});

app.post('/api/branches/refresh-all', auth, requireAdmin, async (_req, res) => {
  const db = readDb();
  const branches = db.branches || [];
  const results = [];

  for (const branch of branches) {
    try {
      const data = await fetchRemoteBranchSummary(branch);
      branch.lastStatus = 'ONLINE';
      branch.lastCheckedAt = new Date().toISOString();
      branch.lastError = '';
      branch.lastRemoteSummary = data;
      results.push({ id: branch.id, ok: true, name: branch.name, checkedAt: branch.lastCheckedAt });
    } catch (err) {
      branch.lastStatus = 'OFFLINE';
      branch.lastCheckedAt = new Date().toISOString();
      branch.lastError = err.message;
      results.push({ id: branch.id, ok: false, name: branch.name, message: err.message, checkedAt: branch.lastCheckedAt });
    }
  }

  writeDb(db);
  res.json({ ok: true, results, branches: branches.map(getBranchSummarySkeleton) });
});

app.get('/api/promotions', auth, (_req, res) => {
  const db = readDb();
  ensurePromotionCollections(db);
  res.json({ ok: true, promotions: db.promotions || [] });
});

app.post('/api/promotions', auth, requireAdmin, (req, res) => {
  const db = readDb();
  ensurePromotionCollections(db);
  const payload = normalizePromotionPayload(req.body || {});
  if (!payload.productIds.length) return res.status(400).json({ ok: false, message: 'Selecciona al menos un producto para la promoción.' });
  const index = (db.promotions || []).findIndex((entry) => entry.id === payload.id);
  if (index >= 0) db.promotions[index] = { ...db.promotions[index], ...payload };
  else db.promotions.unshift(payload);
  writeDb(db);
  res.json({ ok: true, promotion: payload, promotions: db.promotions });
});

app.delete('/api/promotions/:id', auth, requireAdmin, (req, res) => {
  const db = readDb();
  ensurePromotionCollections(db);
  db.promotions = (db.promotions || []).filter((entry) => entry.id !== req.params.id);
  writeDb(db);
  res.json({ ok: true, promotions: db.promotions });
});

app.get('/api/cash/overview', auth, (req, res) => {
  const db = readDb();
  res.json({ ok: true, overview: buildCashOverview(db, normalizeDateKey(req.query.date)) });
});

app.get('/api/cash/daily-report', auth, requireAdmin, (req, res) => {
  const db = readDb();
  const dateKey = normalizeDateKey(req.query?.date);
  const overview = buildCashOverview(db, dateKey);
  const latestClosing = overview.closings[overview.closings.length - 1] || null;
  const refunds = (overview.entries || []).filter((entry) => ['refund', 'sale_refund', 'cancellation_refund'].includes(String(entry.kind || '').toLowerCase()));
  const otherCashEntries = (overview.entries || []).filter((entry) => !['sale_income', 'operational_expense', 'owner_consumption', 'restock_purchase'].includes(entry.kind));
  const cashTenderTotal = round2((overview.sales || []).reduce((sum, sale) => {
    const explicit = safeNumber(sale.amountReceived, NaN);
    if (Number.isFinite(explicit) && explicit > 0) return sum + explicit;
    return sum + (sale.payments || []).filter((payment) => payment.method === 'cash').reduce((amount, payment) => amount + safeNumber(payment.amount), 0);
  }, 0));
  const cashChangeTotal = round2((overview.sales || []).reduce((sum, sale) => sum + safeNumber(sale.changeDue), 0));
  const refundTotal = round2(refunds.reduce((sum, entry) => sum + Math.abs(safeNumber(entry.amountCashImpact ?? entry.accountingAmount ?? entry.amount)), 0));
  const otherCashNet = round2(otherCashEntries.reduce((sum, entry) => sum + safeNumber(entry.amountCashImpact), 0));
  const report = {
    branchName: getBranchProfile(db).name || ADMIN_NAME,
    branchCode: getBranchProfile(db).code || '',
    dateKey,
    period: `${dateKey} 00:00–23:59 (America/Mexico_City)`,
    generatedBy: req.user?.name || req.user?.username || 'Administrador',
    generatedAt: new Date().toISOString(),
    status: latestClosing ? 'CERRADO' : 'ABIERTO',
    reportId: latestClosing?.id || '',
    openingAmount: overview.openingAmount,
    salesCount: overview.salesCount,
    salesTotal: overview.salesTotal,
    payments: overview.payments,
    cashTenderTotal,
    cashChangeTotal,
    refunds,
    refundTotal,
    expensesTotal: overview.expensesTotal,
    expensesCashImpactTotal: overview.expensesCashImpactTotal,
    restockTotal: overview.restockTotal,
    restockCashOutflowTotal: overview.restockCashOutflowTotal,
    restockPurchases: overview.restockPurchases,
    procurementTotal: overview.procurementTotal,
    ownerValueTotal: overview.ownerValueTotal,
    otherCashEntries,
    otherCashNet,
    expectedCash: overview.expectedCash,
    countedCash: latestClosing ? round2(latestClosing.realCash) : null,
    cashDifference: latestClosing ? round2(latestClosing.difference) : null,
    netTotal: round2(overview.salesTotal - overview.expensesTotal - overview.restockTotal - overview.procurementTotal),
    openings: overview.openings,
    closings: overview.closings,
    entries: overview.entries,
    purchases: overview.purchases,
    sales: overview.sales
  };
  return res.json({ ok: true, report });
});

app.post('/api/cash/opening', auth, (req, res) => {
  const db = readDb();
  ensureCashCollections(db);
  const amount = round2(req.body?.amount);
  if (amount < 0) return res.status(400).json({ ok: false, message: 'El fondo de caja no puede ser negativo.' });
  const entry = { id: uuidv4(), dateKey: normalizeDateKey(req.body?.dateKey), amount, terminalCode: String(req.body?.terminalCode || 'POS-01').trim(), userId: req.user?.id || '', userName: req.user?.name || req.user?.username || '', createdAt: new Date().toISOString() };
  db.cashOpenings.push(entry);
  writeDb(db);
  res.json({ ok: true, opening: entry, overview: buildCashOverview(db, entry.dateKey) });
});

app.post('/api/cash/expenses', auth, (req, res) => {
  const db = readDb();
  ensureCashCollections(db);
  const amount = round2(req.body?.amount);
  if (amount <= 0) return res.status(400).json({ ok: false, message: 'El gasto debe ser mayor a 0.' });
  const paymentMethod = normalizeBusinessPaymentMethod(req.body?.paymentMethod || 'other');
  const paymentReference = String(req.body?.paymentReference || '').trim();
  const entry = { id: uuidv4(), kind: 'operational_expense', dateKey: normalizeDateKey(req.body?.dateKey), amountCashImpact: 0, accountingAmount: amount, inventoryValue: 0, paymentMethod, paymentMethodLabel: paymentMethodLabel(paymentMethod), paymentReference, description: String(req.body?.description || '').trim() || 'Gasto operativo', notes: String(req.body?.notes || '').trim(), terminalCode: String(req.body?.terminalCode || 'POS-01').trim(), userId: req.user?.id || '', userName: req.user?.name || req.user?.username || '', printTicketRequested: true, createdAt: new Date().toISOString() };
  db.cashEntries.push(entry);
  writeDb(db);
  res.json({ ok: true, entry, overview: buildCashOverview(db, entry.dateKey) });
});

app.post('/api/cash/owner-consumption', auth, requireAdmin, (req, res) => {
  const db = readDb();
  ensureCashCollections(db);
  db.sales = Array.isArray(db.sales) ? db.sales : [];
  db.inventoryAdjustments = Array.isArray(db.inventoryAdjustments) ? db.inventoryAdjustments : [];
  db.counters = { ...(db.counters || {}) };

  const productMap = new Map((db.products || []).map((product) => [product.id, product]));
  const items = Array.isArray(req.body?.items) ? req.body.items : [];
  const terminalCode = String(req.body?.terminalCode || 'ADMIN-01').trim() || 'ADMIN-01';
  const notes = String(req.body?.notes || '').trim();
  const dateKey = normalizeDateKey(req.body?.dateKey);
  const paymentMethod = normalizeBusinessPaymentMethod(req.body?.paymentMethod || 'internal');
  const paymentReference = String(req.body?.paymentReference || '').trim();

  if (!items.length) {
    return res.status(400).json({ ok: false, message: 'Selecciona al menos un producto para registrar el consumo interno.' });
  }

  const normalizedItems = [];
  for (const item of items) {
    const product = productMap.get(item.productId);
    if (!product) return res.status(400).json({ ok: false, message: 'Producto no encontrado en consumo interno.' });

    const quantity = safeNumber(item.quantity);
    if (quantity <= 0) return res.status(400).json({ ok: false, message: 'La cantidad debe ser mayor a 0.' });
    if (safeNumber(product.stock) < quantity) return res.status(409).json({ ok: false, message: `Stock insuficiente para ${product.name}. Disponible: ${safeNumber(product.stock)} ${product.unit || ''}.` });

    const referenceUnitPrice = round2(safeNumber(product.costPrice || product.retailPrice || product.price));
    const lineTotal = round2(quantity * referenceUnitPrice);
    normalizedItems.push({
      product,
      quantity,
      referenceUnitPrice,
      lineTotal,
      ticketItem: normalizeTicketItems([{
        productId: product.id,
        name: product.name,
        quantity,
        unit: product.unit || 'PZA',
        saleMode: 'INTERNAL',
        unitPrice: referenceUnitPrice,
        originalUnitPrice: referenceUnitPrice,
        discountAmount: 0,
        discountPercent: 0,
        lineTotal
      }])[0]
    });
  }

  const createdAt = new Date().toISOString();
  const inventoryValue = round2(normalizedItems.reduce((acc, item) => acc + item.lineTotal, 0));
  const nextOwnerSequence = safeNumber(db.counters.ownerConsumptionSequence, 0) + 1;
  db.counters.ownerConsumptionSequence = nextOwnerSequence;

  const saleId = uuidv4();
  const entryId = uuidv4();
  const folio = `DUE-${String(nextOwnerSequence).padStart(6, '0')}`;

  for (const item of normalizedItems) {
    item.product.stock = roundQuantity(safeNumber(item.product.stock) - item.quantity);
    item.product.updatedAt = createdAt;
  }

  const saleItems = normalizedItems.map((item) => item.ticketItem);
  const entry = {
    id: entryId,
    kind: 'owner_consumption',
    dateKey,
    amountCashImpact: 0,
    inventoryValue,
    paymentMethod,
    paymentMethodLabel: paymentMethodLabel(paymentMethod),
    paymentReference,
    description: String(req.body?.description || '').trim() || 'Consumo interno del dueño',
    notes,
    terminalCode,
    userId: req.user?.id || '',
    userName: req.user?.name || req.user?.username || '',
    saleId,
    folio,
    items: saleItems.map((item) => ({
      productId: item.productId,
      name: item.name,
      quantity: item.quantity,
      unit: item.unit,
      inventoryValue: round2(safeNumber(item.lineTotal))
    })),
    createdAt
  };

  const internalSale = {
    id: saleId,
    folio,
    terminalCode,
    operatorId: req.user?.id || '',
    operatorName: req.user?.name || req.user?.username || 'Administrador',
    employeeNumber: req.user?.employeeNumber || 'ADMIN',
    customerName: 'Consumo personal del dueño',
    customerReference: 'No cobrado / no afecta caja',
    createdAt,
    clientTxnId: `OWNER-${terminalCode}-${Date.now()}-${String(nextOwnerSequence).padStart(6, '0')}`,
    ticketId: '',
    items: saleItems,
    subtotalBeforeDiscount: inventoryValue,
    discountTotal: 0,
    total: inventoryValue,
    inventoryValue,
    paymentMethod: 'internal',
    referencePaymentMethod: paymentMethod,
    referencePaymentMethodLabel: paymentMethodLabel(paymentMethod),
    paymentReference,
    payments: [],
    amountReceived: 0,
    changeDue: 0,
    saleType: 'owner_consumption',
    transactionType: 'owner_consumption',
    internalType: 'OWNER_CONSUMPTION',
    affectsCash: false,
    affectsSalesTotals: false,
    affectsInventory: true,
    cashEntryId: entryId,
    notes,
    printTicketRequested: true
  };

  db.cashEntries.push(entry);
  db.sales.push(internalSale);
  db.inventoryAdjustments.push({
    id: uuidv4(),
    type: 'owner_consumption',
    saleId,
    cashEntryId: entryId,
    createdAt,
    items: saleItems.map((item) => ({ productId: item.productId, quantity: item.quantity }))
  });

  writeDb(db);
  io.emit('sale:created', internalSale);
  broadcastInventory(db, 'OWNER_CONSUMPTION', { entryId, saleId, folio });
  res.json({ ok: true, entry, sale: internalSale, overview: buildCashOverview(db, entry.dateKey), inventory: db.products });
});

app.post('/api/cash/close', auth, (req, res) => {
  const db = readDb();
  ensureCashCollections(db);
  const dateKey = normalizeDateKey(req.body?.dateKey);
  const overview = buildCashOverview(db, dateKey);
  const closing = { id: uuidv4(), dateKey, openingAmount: overview.openingAmount, salesTotal: overview.salesTotal, salesDiscountTotal: overview.salesDiscountTotal, payments: overview.payments, expensesTotal: overview.expensesTotal, ownerValueTotal: overview.ownerValueTotal, procurementTotal: overview.procurementTotal, acquisitionCost: overview.acquisitionCost, grossProfit: overview.grossProfit, expectedCash: overview.expectedCash, realCash: round2(req.body?.realCash), difference: round2(round2(req.body?.realCash) - overview.expectedCash), closedByUserId: req.user?.id || '', closedByUserName: req.user?.name || req.user?.username || '', createdAt: new Date().toISOString() };
  db.cashClosings.unshift(closing);
  writeDb(db);
  res.json({ ok: true, closing, overview: buildCashOverview(db, dateKey) });
});


// ---------- v13.6.0 módulos empresariales desacoplados ----------
app.get('/api/security/rbac', auth, requireAdmin, (req, res) => {
  const db = readDb();
  ensureRbacCollections(db);
  const users = (db.users || []).map((u) => ({
    id: u.id,
    username: u.username,
    name: u.name,
    role: u.role,
    active: u.active !== false,
    employeeNumber: u.employeeNumber || '',
    profileLabel: u.profileLabel || '',
    isPrincipalAdmin: isPrincipalAdmin(db, u),
    permissions: getUserPermissions(db, u)
  }));
  res.json({ ok: true, modules: RBAC_MODULES, principalProfile: db.security.principalProfile, users });
});

app.post('/api/security/rbac', auth, requireAdmin, (req, res) => {
  const db = readDb();
  if (!isPrincipalAdmin(db, req.user)) return res.status(403).json({ ok: false, message: 'Solo el Administrador Principal puede modificar perfiles y permisos.' });
  ensureRbacCollections(db);
  const body = req.body || {};
  const id = String(body.id || '').trim();
  const username = String(body.username || '').trim().toLowerCase();
  const name = String(body.name || '').trim();
  const role = body.role === 'ADMIN' ? 'ADMIN' : 'POS';
  if (!username) return res.status(400).json({ ok: false, message: 'El usuario es obligatorio.' });
  if (!name) return res.status(400).json({ ok: false, message: 'El nombre es obligatorio.' });
  const duplicate = (db.users || []).find((u) => u.username === username && u.id !== id);
  if (duplicate) return res.status(409).json({ ok: false, message: 'Ya existe un usuario con ese nombre de acceso.' });
  const existing = (db.users || []).find((u) => u.id === id);
  if (existing && isPrincipalAdmin(db, existing)) return res.status(400).json({ ok: false, message: 'El administrador principal no se edita desde perfiles secundarios.' });
  const now = new Date().toISOString();
  const user = existing || { id: uuidv4(), createdAt: now };
  user.username = username;
  user.name = name;
  user.role = role;
  user.employeeNumber = String(body.employeeNumber || user.employeeNumber || '').trim();
  user.profileLabel = String(body.profileLabel || (role === 'ADMIN' ? 'Sub-administrador' : 'Cajero')).trim();
  user.active = body.active !== false;
  if (body.password) user.password = String(body.password).trim();
  if (!user.password) user.password = role === 'ADMIN' ? 'admin123' : (user.employeeNumber || '1234');
  if (body.permissions && typeof body.permissions === 'object') {
    user.permissions = normalizePermissions(body.permissions, role === 'ADMIN' ? 'full' : 'pos');
  } else if (!existing) {
    user.permissions = defaultModulePermissions(role === 'ADMIN' ? 'full' : 'pos');
  } else {
    user.permissions = normalizePermissions(user.permissions || {}, role === 'ADMIN' ? 'full' : 'pos');
  }
  user.updatedAt = now;
  if (!existing) db.users.push(user);
  writeDb(db);
  res.json({ ok: true, user: { id: user.id, username: user.username, name: user.name, role: user.role, active: user.active, permissions: user.permissions } });
});

app.delete('/api/security/rbac/:id', auth, requireAdmin, (req, res) => {
  const db = readDb();
  if (!isPrincipalAdmin(db, req.user)) return res.status(403).json({ ok: false, message: 'Solo el Administrador Principal puede eliminar perfiles.' });
  const user = (db.users || []).find((u) => u.id === req.params.id);
  if (!user) return res.status(404).json({ ok: false, message: 'Usuario no encontrado.' });
  if (isPrincipalAdmin(db, user)) return res.status(400).json({ ok: false, message: 'No se puede eliminar el Administrador Principal.' });
  db.users = db.users.filter((u) => u.id !== req.params.id);
  writeDb(db);
  res.json({ ok: true });
});

app.post('/auth/admin/authorize-discount', auth, (req, res) => {
  // v13.6.27: autorización por sesión iniciada, sin pedir contraseña operativa adicional.
  const db = readDb();
  const actor = req.user || {};
  if (actor.role !== 'ADMIN' && !userHasPermission(db, actor, 'catalog', 'write')) {
    return res.status(403).json({ ok: false, message: 'Este usuario no tiene permiso para aplicar condonaciones o descuentos.' });
  }
  res.json({ ok: true, bypassed: true, authorizedBy: { id: actor.id, username: actor.username, name: actor.name || actor.username || 'Administrador' } });
});

app.get('/api/reports/financial', auth, requireAdmin, (req, res) => {
  const db = readDb();
  res.json(buildFinancialReport(db, req.query || {}));
});

app.post('/api/procurement/purchases', auth, requireAdmin, (req, res) => {
  const db = readDb();
  db.procurementPurchases = Array.isArray(db.procurementPurchases) ? db.procurementPurchases : [];
  db.enterpriseExpenses = Array.isArray(db.enterpriseExpenses) ? db.enterpriseExpenses : [];
  const product = (db.products || []).find((p) => p.id === req.body?.productId);
  if (!product) return res.status(404).json({ ok: false, message: 'Producto no encontrado.' });
  const quantity = safeNumber(req.body?.quantity);
  if (quantity <= 0) return res.status(400).json({ ok: false, message: 'La cantidad comprada debe ser mayor a cero.' });
  const unitCost = round2(req.body?.unitCost !== undefined && req.body?.unitCost !== '' ? req.body.unitCost : product.costPrice);
  if (unitCost < 0) return res.status(400).json({ ok: false, message: 'El costo unitario no puede ser negativo.' });
  const totalCost = round2(quantity * unitCost);
  const paymentMethod = normalizeBusinessPaymentMethod(req.body?.paymentMethod || 'other');
  const paymentReference = String(req.body?.paymentReference || '').trim();
  const now = new Date().toISOString();
  product.stock = roundQuantity(safeNumber(product.stock) + quantity);
  product.costPrice = unitCost || product.costPrice;
  product.updatedAt = now;
  const nextSeq = safeNumber(db.counters?.purchaseSequence, 0) + 1;
  db.counters = { ...(db.counters || {}), purchaseSequence: nextSeq };
  const entry = { id: uuidv4(), folio: `CMP-${String(nextSeq).padStart(6, '0')}`, kind: 'merchandise_purchase', productId: product.id, productName: product.name, quantity, unit: product.unit || 'PZA', unitCost, totalCost, paymentMethod, paymentMethodLabel: paymentMethodLabel(paymentMethod), paymentReference, dateKey: normalizeDateKey(req.body?.dateKey), notes: String(req.body?.notes || '').trim(), terminalCode: String(req.body?.terminalCode || 'ADMIN-01').trim(), userId: req.user?.id || '', userName: req.user?.name || req.user?.username || '', createdAt: now };
  db.procurementPurchases.push(entry);
  db.enterpriseExpenses.push({ ...entry, accountingAmount: totalCost, amountCashImpact: 0, description: `Compra de mercancía: ${product.name}` });
  db.inventoryAdjustments.push({ id: uuidv4(), type: 'procurement_purchase', purchaseId: entry.id, productId: product.id, productName: product.name, previousStock: roundQuantity(safeNumber(product.stock) - quantity), newStock: product.stock, delta: quantity, unit: product.unit || 'PZA', reason: 'Compra de mercancía', createdAt: now });
  writeDb(db);
  broadcastInventory(db, 'PROCUREMENT_PURCHASE', { purchaseId: entry.id });
  res.json({ ok: true, purchase: entry, inventory: db.products });
});

function buildUploadsBackupPayload() {
  const files = [];
  if (!fs.existsSync(UPLOADS_DIR)) return files;
  for (const filename of fs.readdirSync(UPLOADS_DIR)) {
    const filePath = path.join(UPLOADS_DIR, filename);
    if (!fs.statSync(filePath).isFile()) continue;
    files.push({
      name: filename,
      encoding: 'base64',
      data: fs.readFileSync(filePath).toString('base64')
    });
  }
  return files;
}

function restoreUploadsFromBackup(files = []) {
  if (!Array.isArray(files) || !files.length) return 0;
  fs.mkdirSync(UPLOADS_DIR, { recursive: true });
  let restored = 0;
  for (const file of files) {
    const safeName = path.basename(String(file.name || ''));
    if (!safeName || !file.data) continue;
    fs.writeFileSync(path.join(UPLOADS_DIR, safeName), Buffer.from(String(file.data), file.encoding === 'base64' ? 'base64' : 'utf8'));
    restored += 1;
  }
  return restored;
}

app.get('/api/backup/export', auth, requireAdmin, (_req, res) => {
  const db = readDb();
  const payload = {
    app: 'Cremería El Güero',
    version: '13.6.18',
    exportedAt: new Date().toISOString(),
    description: 'Respaldo total: base de datos completa, ventas, tickets, caja, gastos, consumo interno, inventario, facturación, usuarios, permisos, sucursales, terminales y archivos cargados.',
    data: db,
    files: {
      uploads: buildUploadsBackupPayload()
    }
  };
  const filename = `cremeria-el-guero-backup-total-${new Date().toISOString().slice(0,10)}.json`;
  res.setHeader('Content-Type', 'application/json; charset=utf-8');
  res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
  res.send(JSON.stringify(payload, null, 2));
});

app.post('/api/backup/import', auth, requireAdmin, upload.single('file'), (req, res) => {
  const db = readDb();
  if (!isPrincipalAdmin(db, req.user)) return res.status(403).json({ ok: false, message: 'Solo el Administrador Principal puede restaurar respaldos.' });
  if (!req.file) return res.status(400).json({ ok: false, message: 'Selecciona un archivo de respaldo JSON.' });
  let payload;
  try { payload = JSON.parse(fs.readFileSync(req.file.path, 'utf8')); } catch (err) { return res.status(400).json({ ok: false, message: 'El respaldo no es un JSON válido.' }); }
  const nextDb = payload.data && typeof payload.data === 'object' ? payload.data : payload;
  if (!Array.isArray(nextDb.products) || !Array.isArray(nextDb.sales) || !Array.isArray(nextDb.users)) return res.status(400).json({ ok: false, message: 'El archivo no parece ser un respaldo válido de Cremería El Güero.' });
  fs.copyFileSync(DB_FILE, `${DB_FILE}.before-restore-${Date.now()}.json`);
  writeDb(nextDb);
  const restoredUploads = restoreUploadsFromBackup(payload.files?.uploads || []);
if (!RENDER_MIRROR_ONLY) ensureDb();
  res.json({ ok: true, message: 'Respaldo restaurado correctamente.', restoredUploads });
});

app.get('/api/system/network', auth, requireAdmin, (_req, res) => {
  const db = readDb();
  res.json({ ok: true, network: getLocalNetworkInfo(db) });
});

app.post('/api/system/network', auth, requireAdmin, (req, res) => {
  const db = readDb();
  db.systemNetwork = normalizeSystemNetworkConfig({ ...(db.systemNetwork || {}), ...(req.body || {}), updatedAt: new Date().toISOString() });
  const info = getLocalNetworkInfo(db);
  db.systemNetwork.lastKnownUrl = info.localUrl;
  writeDb(db);
  res.json({ ok: true, network: getLocalNetworkInfo(db) });
});

app.get('/api/inventory/export-critical', auth, requireAdmin, (req, res) => {
  const db = readDb();
  const products = (db.products || []).filter((p) => ['missing','near'].includes(getInventoryAlertStatus(p)));
  const lines = products.map((p) => `${p.name} · SKU: ${p.sku || '-'} · Código barras: ${p.barcode || p.codigoBarras || '-'} · Stock: ${p.stock} ${p.unit || ''} · Mínimo: ${p.lowStockThreshold}`).join('\n');
  const phone = String(db.security?.notifications?.whatsappRestockNumber || '').replace(/\D/g, '');
  const text = `Faltantes Cremería El Güero:\n${lines || 'Sin faltantes.'}`;
  const whatsappUrl = phone ? `https://wa.me/${phone}?text=${encodeURIComponent(text)}` : `https://wa.me/?text=${encodeURIComponent(text)}`;
  res.json({ ok: true, products, text, whatsappUrl });
});

app.get('/api/inventory/export-excel', auth, requireAdmin, (_req, res) => {
  const db = readDb();
  const html = buildInventoryExcelHtml(db.products || []);
  const filename = `inventario-cremeria-el-guero-${new Date().toISOString().slice(0,10)}.xls`;
  res.setHeader('Content-Type', 'application/vnd.ms-excel; charset=utf-8');
  res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
  res.send(html);
});

app.get('/api/fiscal/config', auth, requireAdmin, (_req, res) => {
  res.json({ ok: true, config: readDb().fiscalConfig });
});

app.post('/api/fiscal/config', auth, requireAdmin, (req, res) => {
  const db = readDb();
  const incoming = req.body || {};
  db.fiscalConfig = {
    ...db.fiscalConfig,
    ...incoming,
    commercialName: String(incoming.commercialName ?? db.fiscalConfig.commercialName ?? 'Studio Code').trim() || 'Studio Code',
    emitterRfc: String(incoming.emitterRfc ?? db.fiscalConfig.emitterRfc ?? '').trim().toUpperCase(),
    emitterName: String(incoming.emitterName ?? db.fiscalConfig.emitterName ?? '').trim(),
    emitterRegimen: normalizeFiscalRegimeValue(incoming.emitterRegimen ?? db.fiscalConfig.emitterRegimen ?? '', '605'),
    expeditionZip: String(incoming.expeditionZip ?? db.fiscalConfig.expeditionZip ?? '55750').trim(),
    pacProvider: 'FACTORAMA',
    preferredInvoiceProvider: String(incoming.preferredInvoiceProvider ?? db.fiscalConfig.preferredInvoiceProvider ?? 'SAT_ASSISTED').toUpperCase() === 'FACTURAMA' ? 'FACTURAMA' : 'SAT_ASSISTED',
    satPortalUrl: String(incoming.satPortalUrl ?? db.fiscalConfig.satPortalUrl ?? SAT_PORTAL_DEFAULT).trim() || SAT_PORTAL_DEFAULT,
    satInfoUrl: String(incoming.satInfoUrl ?? db.fiscalConfig.satInfoUrl ?? SAT_INFO_PORTAL_DEFAULT).trim() || SAT_INFO_PORTAL_DEFAULT,
    satSimulatorUrl: String(incoming.satSimulatorUrl ?? db.fiscalConfig.satSimulatorUrl ?? SAT_SIMULATOR_DEFAULT).trim() || SAT_SIMULATOR_DEFAULT,
    satCancellationUrl: String(incoming.satCancellationUrl ?? db.fiscalConfig.satCancellationUrl ?? SAT_CANCELLATION_DEFAULT).trim() || SAT_CANCELLATION_DEFAULT,
    satWorkflowMode: String(incoming.satWorkflowMode ?? db.fiscalConfig.satWorkflowMode ?? 'SIMULATOR').toUpperCase() === 'PORTAL' ? 'PORTAL' : 'SIMULATOR',
    pacMode: normalizeFactoramaEnvironment(incoming.factoramaEnvironment ?? db.fiscalConfig.factoramaEnvironment ?? db.fiscalConfig.pacMode).toUpperCase(),
    factoramaEnvironment: normalizeFactoramaEnvironment(incoming.factoramaEnvironment ?? db.fiscalConfig.factoramaEnvironment ?? db.fiscalConfig.pacMode),
    factoramaApiUsername: String(incoming.factoramaApiUsername ?? db.fiscalConfig.factoramaApiUsername ?? '').trim(),
    factoramaApiPassword: String(incoming.factoramaApiPassword ?? db.fiscalConfig.factoramaApiPassword ?? '').trim(),
    defaultSeries: String(incoming.defaultSeries ?? db.fiscalConfig.defaultSeries ?? 'CREG').trim() || 'CREG',
    defaultPaymentForm: normalizePaymentFormValue(incoming.defaultPaymentForm ?? db.fiscalConfig.defaultPaymentForm ?? '01'),
    defaultPaymentMethod: normalizePaymentMethodValue(incoming.defaultPaymentMethod ?? db.fiscalConfig.defaultPaymentMethod ?? 'PUE'),
    defaultUseCfdi: normalizeCfdiUseValue(incoming.defaultUseCfdi ?? db.fiscalConfig.defaultUseCfdi ?? 'S01'),
    defaultCancelMotive: String(incoming.defaultCancelMotive ?? db.fiscalConfig.defaultCancelMotive ?? '02').trim() || '02'
  };
  writeDb(db);
  res.json({ ok: true, config: db.fiscalConfig });
});

app.post('/api/fiscal/test-connection', auth, requireAdmin, async (req, res) => {
  try {
    const db = readDb();
    const incoming = req.body || {};
    const baseConfig = db.fiscalConfig || {};
    const config = {
      ...baseConfig,
      factoramaEnvironment: normalizeFactoramaEnvironment(incoming.factoramaEnvironment ?? baseConfig.factoramaEnvironment ?? baseConfig.pacMode),
      pacMode: normalizeFactoramaEnvironment(incoming.factoramaEnvironment ?? baseConfig.factoramaEnvironment ?? baseConfig.pacMode).toUpperCase(),
      factoramaApiUsername: String(incoming.factoramaApiUsername ?? baseConfig.factoramaApiUsername ?? '').trim(),
      factoramaApiPassword: String(incoming.factoramaApiPassword ?? baseConfig.factoramaApiPassword ?? '').trim(),
      commercialName: String(incoming.commercialName ?? baseConfig.commercialName ?? '').trim(),
      emitterRfc: String(incoming.emitterRfc ?? baseConfig.emitterRfc ?? '').trim().toUpperCase(),
      emitterName: String(incoming.emitterName ?? baseConfig.emitterName ?? '').trim(),
      emitterRegimen: normalizeFiscalRegimeValue(incoming.emitterRegimen ?? baseConfig.emitterRegimen ?? '', '605'),
      expeditionZip: String(incoming.expeditionZip ?? baseConfig.expeditionZip ?? '55750').trim(),
      defaultSeries: String(incoming.defaultSeries ?? baseConfig.defaultSeries ?? 'CREG').trim() || 'CREG',
      defaultPaymentForm: normalizePaymentFormValue(incoming.defaultPaymentForm ?? baseConfig.defaultPaymentForm ?? '01'),
      defaultPaymentMethod: normalizePaymentMethodValue(incoming.defaultPaymentMethod ?? baseConfig.defaultPaymentMethod ?? 'PUE'),
      defaultUseCfdi: normalizeCfdiUseValue(incoming.defaultUseCfdi ?? baseConfig.defaultUseCfdi ?? 'S01'),
      defaultCancelMotive: String(incoming.defaultCancelMotive ?? baseConfig.defaultCancelMotive ?? '02').trim() || '02'
    };
    if (!hasFacturamaCredentials(config)) {
      return res.status(400).json({ ok: false, message: 'Configura primero el usuario y la contraseña API de Facturama.' });
    }
    const info = await factoramaFetchFirst(['/api/Account/UserInfo', '/account/UserInfo'], { config });
    res.json({
      ok: true,
      environment: normalizeFactoramaEnvironment(config.factoramaEnvironment),
      info
    });
  } catch (err) {
    res.status(Number(err.status || 500)).json({ ok: false, message: err.message, detail: err.payload || null });
  }
});

app.get('/api/fiscal/customers', auth, requireAdmin, (_req, res) => {
  res.json({ ok: true, customers: readDb().fiscalCustomers });
});

app.post('/api/fiscal/customers', auth, requireAdmin, (req, res) => {
  const db = readDb();
  const body = req.body || {};
  const requestedRfc = String(body.rfc || '').trim().toUpperCase();
  const existingCustomer = (db.fiscalCustomers || []).find((entry) => entry.id === body.id || (requestedRfc && entry.rfc === requestedRfc)) || null;
  const customer = {
    id: body.id || existingCustomer?.id || uuidv4(),
    name: String(body.name || '').trim().toUpperCase(),
    rfc: requestedRfc,
    regimenFiscal: normalizeFiscalRegimeValue(body.regimenFiscal || '', requestedRfc === 'XAXX010101000' ? '616' : '605'),
    zip: String(body.zip || db.fiscalConfig.expeditionZip || '55750').trim(),
    email: String(body.email || '').trim(),
    useCfdi: normalizeCfdiUseValue(body.useCfdi || existingCustomer?.useCfdi || db.fiscalConfig.defaultUseCfdi || 'S01'),
    lastUseCfdi: normalizeCfdiUseValue(existingCustomer?.lastUseCfdi || body.useCfdi || db.fiscalConfig.defaultUseCfdi || 'S01'),
    lastPaymentMethod: normalizePaymentMethodValue(existingCustomer?.lastPaymentMethod || db.fiscalConfig.defaultPaymentMethod || 'PUE'),
    lastPaymentForm: normalizePaymentFormValue(existingCustomer?.lastPaymentForm || db.fiscalConfig.defaultPaymentForm || '01'),
    lastInvoicedAt: existingCustomer?.lastInvoicedAt || '',
    createdAt: existingCustomer?.createdAt || body.createdAt || new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    isDefault: body.isDefault === true
  };

  if (!customer.name) return res.status(400).json({ ok: false, message: 'El nombre o razón social es obligatorio.' });
  if (!customer.rfc) return res.status(400).json({ ok: false, message: 'El RFC es obligatorio.' });

  const existingIdx = (db.fiscalCustomers || []).findIndex((entry) => entry.id === customer.id || entry.rfc === customer.rfc);
  if (existingIdx >= 0) db.fiscalCustomers[existingIdx] = { ...db.fiscalCustomers[existingIdx], ...customer };
  else db.fiscalCustomers.push(customer);

  if (customer.isDefault === true) {
    db.fiscalCustomers = db.fiscalCustomers.map((entry) => ({
      ...entry,
      isDefault: entry.id === customer.id
    }));
  } else if (db.fiscalCustomers.length === 1) {
    db.fiscalCustomers[0].isDefault = true;
  } else if (!db.fiscalCustomers.some((entry) => entry.isDefault === true)) {
    db.fiscalCustomers[0].isDefault = true;
  }

  writeDb(db);
  res.json({ ok: true, customer: db.fiscalCustomers.find((entry) => entry.id === customer.id) || customer });
});

app.delete('/api/fiscal/customers/:id', auth, requireAdmin, (req, res) => {
  const db = readDb();
  const customerId = String(req.params.id || '').trim();
  if (!customerId) return res.status(400).json({ ok: false, message: 'ID de cliente fiscal inválido.' });

  const existing = (db.fiscalCustomers || []).find((entry) => entry.id === customerId);
  if (!existing) return res.status(404).json({ ok: false, message: 'Cliente fiscal no encontrado.' });

  db.fiscalCustomers = (db.fiscalCustomers || []).filter((entry) => entry.id !== customerId);
  if (db.fiscalCustomers.length && !db.fiscalCustomers.some((entry) => entry.isDefault === true)) {
    db.fiscalCustomers[0].isDefault = true;
  }

  writeDb(db);
  res.json({ ok: true, deletedId: customerId, customers: db.fiscalCustomers });
});

function extractFacturamaErrorMessage(payload = null, fallback = 'La solicitud no es válida.') {
  if (!payload) return fallback;
  if (typeof payload === 'string') return payload || fallback;
  if (Array.isArray(payload)) return payload.map((entry) => extractFacturamaErrorMessage(entry, '')).filter(Boolean).join(' | ') || fallback;
  if (payload.Message && String(payload.Message).trim()) return String(payload.Message);
  if (payload.message) return String(payload.message);
  if (payload.ModelState && typeof payload.ModelState === 'object') {
    const parts = [];
    for (const value of Object.values(payload.ModelState)) {
      if (Array.isArray(value)) parts.push(...value.map(String));
      else if (value) parts.push(String(value));
    }
    if (parts.length) return parts.join(' | ');
  }
  if (payload.Errors && Array.isArray(payload.Errors)) {
    const parts = payload.Errors.map((entry) => entry?.Message || entry?.message || JSON.stringify(entry)).filter(Boolean);
    if (parts.length) return parts.join(' | ');
  }
  try {
    const compact = JSON.stringify(payload);
    if (compact && compact !== '{}' && compact !== '[]') return `${fallback} Detalle: ${compact}`;
  } catch (_) {}
  return fallback;
}

app.post('/api/fiscal/customers/validate', auth, requireAdmin, async (req, res) => {
  try {
    const db = readDb();
    const customer = req.body || {};
    if (!hasFacturamaCredentials(db.fiscalConfig || {})) {
      return res.json({ ok: true, localOnly: true, valid: true, customer });
    }
    const receiverName = String(customer.name || '').trim().replace(/\s+/g, ' ').toUpperCase();
    const data = await factoramaFetchFirst(['/api/customers/validate', '/customers/validate'], {
      method: 'POST',
      config: db.fiscalConfig || {},
      body: {
        Rfc: String(customer.rfc || '').trim().toUpperCase(),
        Name: receiverName,
        FiscalRegime: normalizeFiscalRegimeValue(customer.regimenFiscal || '', '616'),
        TaxZipCode: String(customer.zip || '').trim(),
        CfdiUse: normalizeCfdiUseValue(customer.useCfdi || 'S01')
      }
    });
    res.json({ ok: true, validation: data });
  } catch (err) {
    res.status(Number(err.status || 500)).json({ ok: false, message: err.message, detail: err.payload || null });
  }
});


function updateFiscalCustomerInvoicePreferences(db, customerId, { useCfdi, paymentMethod, paymentForm } = {}) {
  const customer = (db.fiscalCustomers || []).find((entry) => entry.id === customerId);
  if (!customer) return null;
  customer.lastUseCfdi = normalizeCfdiUseValue(useCfdi || customer.lastUseCfdi || customer.useCfdi || db.fiscalConfig?.defaultUseCfdi || 'S01');
  customer.lastPaymentMethod = normalizePaymentMethodValue(paymentMethod || customer.lastPaymentMethod || db.fiscalConfig?.defaultPaymentMethod || 'PUE');
  customer.lastPaymentForm = normalizePaymentFormValue(paymentForm || customer.lastPaymentForm || db.fiscalConfig?.defaultPaymentForm || '01');
  customer.useCfdi = customer.lastUseCfdi;
  customer.lastInvoicedAt = new Date().toISOString();
  customer.updatedAt = new Date().toISOString();
  return customer;
}

function addFiscalEvent(db, type, detail = {}) {
  db.fiscalEvents = Array.isArray(db.fiscalEvents) ? db.fiscalEvents : [];
  const event = {
    id: uuidv4(),
    type: String(type || 'FISCAL_EVENT').toUpperCase(),
    createdAt: new Date().toISOString(),
    ...detail
  };
  db.fiscalEvents.unshift(event);
  if (db.fiscalEvents.length > 2000) db.fiscalEvents = db.fiscalEvents.slice(0, 2000);
  return event;
}

function isValidRfcBasic(value = '') {
  return /^[A-ZÑ&]{3,4}\d{6}[A-Z0-9]{3}$/i.test(String(value || '').trim());
}

function buildSatAssistedPreparation({ db, sale, customer, useCfdi, paymentMethod, paymentForm }) {
  const normalizedUse = normalizeCfdiUseValue(useCfdi || customer?.lastUseCfdi || customer?.useCfdi || db.fiscalConfig?.defaultUseCfdi || 'S01');
  const normalizedMethod = normalizePaymentMethodValue(paymentMethod || customer?.lastPaymentMethod || db.fiscalConfig?.defaultPaymentMethod || 'PUE');
  let normalizedForm = normalizePaymentFormValue(paymentForm || customer?.lastPaymentForm || db.fiscalConfig?.defaultPaymentForm || '01');
  if (normalizedMethod === 'PPD') normalizedForm = '99';

  const receiver = {
    id: customer.id,
    name: String(customer.name || '').trim().replace(/\s+/g, ' ').toUpperCase(),
    rfc: String(customer.rfc || '').trim().toUpperCase(),
    regimenFiscal: normalizeFiscalRegimeValue(customer.regimenFiscal || '', '616'),
    zip: String(customer.zip || '').trim(),
    email: String(customer.email || '').trim(),
    useCfdi: normalizedUse
  };

  const emitter = {
    commercialName: db.fiscalConfig?.commercialName || ADMIN_NAME,
    rfc: String(db.fiscalConfig?.emitterRfc || '').trim().toUpperCase(),
    name: String(db.fiscalConfig?.emitterName || '').trim(),
    regimenFiscal: normalizeFiscalRegimeValue(db.fiscalConfig?.emitterRegimen || '', '605'),
    expeditionZip: String(db.fiscalConfig?.expeditionZip || '').trim()
  };

  const blockers = [];
  const warnings = [];
  if (!emitter.rfc) blockers.push('Falta RFC del emisor en Configuración fiscal.');
  else if (!isValidRfcBasic(emitter.rfc)) warnings.push('El RFC del emisor no supera la validación de estructura básica; confirma el dato antes de emitir.');
  if (!emitter.name) blockers.push('Falta nombre o razón social del emisor.');
  if (!emitter.regimenFiscal) blockers.push('Falta régimen fiscal del emisor.');
  if (!/^\d{5}$/.test(emitter.expeditionZip)) blockers.push('El CP de expedición debe tener 5 dígitos.');

  if (!receiver.name) blockers.push('Falta nombre o razón social del receptor.');
  if (!receiver.rfc) blockers.push('Falta RFC del receptor.');
  else if (!isValidRfcBasic(receiver.rfc)) warnings.push('El RFC del receptor no supera la validación de estructura básica; confirma el dato antes de emitir.');
  if (!receiver.regimenFiscal) blockers.push('Falta régimen fiscal del receptor.');
  if (!/^\d{5}$/.test(receiver.zip)) blockers.push('El CP fiscal del receptor debe tener 5 dígitos.');
  if (!normalizedUse) blockers.push('Falta Uso CFDI.');
  if (!Array.isArray(sale.items) || !sale.items.length) blockers.push('La venta no contiene conceptos para facturar.');

  const concepts = (Array.isArray(sale.items) ? sale.items : []).map((item, index) => {
    const product = (db.products || []).find((entry) => entry.id === item.productId) || null;
    const explicitProductCode = String(product?.fiscalProductCode || item.fiscalProductCode || '').trim();
    const explicitUnitCode = String(product?.fiscalUnitCode || item.fiscalUnitCode || '').trim();
    const explicitTaxObject = String(product?.fiscalTaxObject || item.fiscalTaxObject || '').trim();
    const fiscal = getDefaultItemFiscalData(item, product);
    const label = item.name || product?.name || `Concepto ${index + 1}`;
    if (!explicitProductCode) blockers.push(`El concepto “${label}” no tiene ClaveProdServ SAT configurada.`);
    if (!explicitUnitCode) blockers.push(`El concepto “${label}” no tiene ClaveUnidad SAT configurada.`);
    if (!explicitTaxObject) warnings.push(`El concepto “${label}” no tiene ObjetoImp explícito; revisa su configuración fiscal.`);
    const quantity = safeNumber(item.quantity, 1);
    const unitPrice = round2(safeNumber(item.unitPrice, 0));
    const discount = round2(safeNumber(item.discountAmount ?? item.discount ?? 0, 0));
    const lineTotal = round2(safeNumber(item.lineTotal, quantity * unitPrice - discount));
    return {
      productId: item.productId || '',
      name: label,
      quantity,
      unitPrice,
      discount,
      lineTotal,
      productCode: explicitProductCode || fiscal.ProductCode,
      unitCode: explicitUnitCode || fiscal.UnitCode,
      unitName: fiscal.Unit,
      description: fiscal.Description,
      taxObject: explicitTaxObject || fiscal.TaxObject,
      vatRate: safeNumber(fiscal.VatRate, 0),
      fiscalConfigured: Boolean(explicitProductCode && explicitUnitCode)
    };
  });

  const preparation = {
    provider: 'SAT',
    mode: 'SAT_ASSISTED_PRO',
    workflowMode: String(db.fiscalConfig?.satWorkflowMode || 'SIMULATOR').toUpperCase() === 'PORTAL' ? 'PORTAL' : 'SIMULATOR',
    satPortalUrl: String(db.fiscalConfig?.satPortalUrl || SAT_PORTAL_DEFAULT).trim() || SAT_PORTAL_DEFAULT,
    satSimulatorUrl: String(db.fiscalConfig?.satSimulatorUrl || SAT_SIMULATOR_DEFAULT).trim() || SAT_SIMULATOR_DEFAULT,
    satCancellationUrl: String(db.fiscalConfig?.satCancellationUrl || SAT_CANCELLATION_DEFAULT).trim() || SAT_CANCELLATION_DEFAULT,
    sale: {
      id: sale.id,
      folio: sale.folio || '',
      createdAt: sale.createdAt || '',
      subtotal: round2(safeNumber(sale.subtotal ?? sale.total)),
      discountTotal: round2(safeNumber(sale.discountTotal || 0)),
      total: round2(safeNumber(sale.total))
    },
    emitter,
    receiver,
    payment: {
      useCfdi: normalizedUse,
      paymentMethod: normalizedMethod,
      paymentForm: normalizedForm,
      currency: 'MXN',
      cfdiType: 'I',
      exportation: '01'
    },
    concepts,
    blockers: [...new Set(blockers)],
    warnings: [...new Set(warnings)]
  };
  preparation.readyForPortal = preparation.blockers.length === 0;
  preparation.copyText = [
    `VENTA: ${preparation.sale.folio}`,
    `TOTAL: $${preparation.sale.total.toFixed(2)} MXN`,
    '',
    'EMISOR',
    `RFC: ${emitter.rfc}`,
    `NOMBRE: ${emitter.name}`,
    `REGIMEN: ${emitter.regimenFiscal}`,
    `CP EXPEDICION: ${emitter.expeditionZip}`,
    '',
    'RECEPTOR',
    `RFC: ${receiver.rfc}`,
    `NOMBRE: ${receiver.name}`,
    `CP FISCAL: ${receiver.zip}`,
    `REGIMEN: ${receiver.regimenFiscal}`,
    `USO CFDI: ${normalizedUse}`,
    `METODO PAGO: ${normalizedMethod}`,
    `FORMA PAGO: ${normalizedForm}`,
    '',
    'CONCEPTOS',
    ...concepts.map((item, index) => `${index + 1}. ${item.quantity} x ${item.description} | SAT ${item.productCode || '[FALTA]'} | UNIDAD ${item.unitCode || '[FALTA]'} | OBJETO ${item.taxObject || '[REVISAR]'} | $${item.unitPrice.toFixed(2)} | TOTAL $${item.lineTotal.toFixed(2)}`)
  ].join('\n');
  return preparation;
}

function decodeXmlEntities(value = '') {
  return String(value || '')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&');
}

function xmlTagAttributes(xml = '', localName = '') {
  const re = new RegExp(`<(?:(?:[A-Za-z0-9_\\-]+):)?${localName}\\b([^>]*)>`, 'i');
  const match = String(xml || '').match(re);
  return match ? match[1] : '';
}

function xmlAttr(attrs = '', name = '') {
  const re = new RegExp(`\\b${name}\\s*=\\s*["']([^"']*)["']`, 'i');
  const match = String(attrs || '').match(re);
  return match ? decodeXmlEntities(match[1]) : '';
}

function parseSatCfdiXml(xmlText = '') {
  const xml = String(xmlText || '').replace(/^\uFEFF/, '').trim();
  if (!xml || !/<(?:(?:[A-Za-z0-9_\-]+):)?Comprobante\b/i.test(xml)) {
    const err = new Error('El archivo seleccionado no parece ser un CFDI XML válido.');
    err.status = 400;
    throw err;
  }
  const comprobante = xmlTagAttributes(xml, 'Comprobante');
  const receptor = xmlTagAttributes(xml, 'Receptor');
  const emisor = xmlTagAttributes(xml, 'Emisor');
  const timbre = xmlTagAttributes(xml, 'TimbreFiscalDigital');
  const uuid = xmlAttr(timbre, 'UUID').toUpperCase();
  if (!uuid) {
    const err = new Error('El XML no contiene UUID de Timbre Fiscal Digital. Verifica que sea una factura ya timbrada.');
    err.status = 400;
    throw err;
  }
  return {
    uuid,
    serie: xmlAttr(comprobante, 'Serie'),
    folio: xmlAttr(comprobante, 'Folio'),
    fecha: xmlAttr(comprobante, 'Fecha'),
    fechaTimbrado: xmlAttr(timbre, 'FechaTimbrado'),
    total: round2(safeNumber(xmlAttr(comprobante, 'Total'))),
    subtotal: round2(safeNumber(xmlAttr(comprobante, 'SubTotal'))),
    paymentForm: normalizePaymentFormValue(xmlAttr(comprobante, 'FormaPago') || '99'),
    paymentMethod: normalizePaymentMethodValue(xmlAttr(comprobante, 'MetodoPago') || 'PUE'),
    emitterRfc: String(xmlAttr(emisor, 'Rfc')).toUpperCase(),
    emitterName: xmlAttr(emisor, 'Nombre'),
    receiverRfc: String(xmlAttr(receptor, 'Rfc')).toUpperCase(),
    receiverName: xmlAttr(receptor, 'Nombre'),
    receiverZip: xmlAttr(receptor, 'DomicilioFiscalReceptor'),
    receiverRegimen: xmlAttr(receptor, 'RegimenFiscalReceptor'),
    useCfdi: normalizeCfdiUseValue(xmlAttr(receptor, 'UsoCFDI') || 'S01'),
    xml
  };
}

app.post('/api/invoices/sat/prepare', auth, requireAdmin, (req, res) => {
  try {
    const db = readDb();
    const { saleId, customerId, useCfdi, paymentMethod, paymentForm } = req.body || {};
    const sale = (db.sales || []).find((entry) => entry.id === saleId);
    if (!sale) return res.status(404).json({ ok: false, message: 'Venta no encontrada.' });
    const customer = (db.fiscalCustomers || []).find((entry) => entry.id === customerId);
    if (!customer) return res.status(404).json({ ok: false, message: 'Cliente fiscal no encontrado.' });

    const existingInvoice = (db.invoices || []).find((entry) => entry.saleId === sale.id && !String(entry.status || '').includes('CANCEL'));
    if (existingInvoice) {
      return res.status(409).json({ ok: false, message: `La venta ${sale.folio || ''} ya está vinculada al CFDI ${existingInvoice.uuid || existingInvoice.id}.` });
    }

    const preparation = buildSatAssistedPreparation({ db, sale, customer, useCfdi, paymentMethod, paymentForm });
    updateFiscalCustomerInvoicePreferences(db, customer.id, preparation.payment);
    const now = new Date().toISOString();
    db.fiscalDrafts = Array.isArray(db.fiscalDrafts) ? db.fiscalDrafts : [];
    let draft = db.fiscalDrafts.find((entry) => entry.saleId === sale.id && entry.customerId === customer.id && !['IMPORTED','CANCELED'].includes(String(entry.status || '').toUpperCase()));
    if (!draft) {
      draft = { id: uuidv4(), saleId: sale.id, customerId: customer.id, createdAt: now };
      db.fiscalDrafts.unshift(draft);
    }
    Object.assign(draft, {
      saleFolio: sale.folio || '',
      status: preparation.readyForPortal ? 'READY' : 'NEEDS_REVIEW',
      workflowMode: preparation.workflowMode,
      preparation,
      updatedAt: now
    });
    addFiscalEvent(db, 'SAT_PREPARATION_CREATED', { draftId: draft.id, saleId: sale.id, saleFolio: sale.folio || '', customerId: customer.id, status: draft.status });
    writeDb(db);
    res.json({ ok: true, preparation: { ...preparation, draftId: draft.id, draftStatus: draft.status } });
  } catch (err) {
    res.status(Number(err.status || 500)).json({ ok: false, message: err.message });
  }
});

app.post('/api/invoices/sat/import', auth, requireAdmin, (req, res) => {
  try {
    const db = readDb();
    const { saleId, customerId, xmlText } = req.body || {};
    const sale = (db.sales || []).find((entry) => entry.id === saleId);
    if (!sale) return res.status(404).json({ ok: false, message: 'Venta no encontrada.' });
    const customer = (db.fiscalCustomers || []).find((entry) => entry.id === customerId);
    if (!customer) return res.status(404).json({ ok: false, message: 'Cliente fiscal no encontrado.' });

    const parsed = parseSatCfdiXml(xmlText);
    const expectedRfc = String(customer.rfc || '').trim().toUpperCase();
    if (expectedRfc && parsed.receiverRfc && parsed.receiverRfc !== expectedRfc) {
      return res.status(400).json({ ok: false, message: `El XML pertenece al RFC ${parsed.receiverRfc}, pero seleccionaste ${expectedRfc}.` });
    }

    const duplicateUuid = (db.invoices || []).find((entry) => String(entry.uuid || entry.id).toUpperCase() === parsed.uuid);
    const otherSaleInvoice = (db.invoices || []).find((entry) => entry.saleId === sale.id && !String(entry.status || '').includes('CANCEL') && String(entry.uuid || entry.id).toUpperCase() !== parsed.uuid);
    if (otherSaleInvoice) {
      return res.status(409).json({ ok: false, message: `La venta ${sale.folio || ''} ya está vinculada a otro CFDI: ${otherSaleInvoice.uuid || otherSaleInvoice.id}.` });
    }

    const warnings = [];
    if (Math.abs(round2(parsed.total) - round2(safeNumber(sale.total))) > 0.02) {
      warnings.push(`El total del XML (${parsed.total.toFixed(2)}) no coincide exactamente con el total de la venta (${round2(safeNumber(sale.total)).toFixed(2)}).`);
    }

    const invoiceRecord = {
      id: parsed.uuid,
      uuid: parsed.uuid,
      provider: 'SAT',
      source: 'SAT_ASSISTED',
      apiMode: 'SAT_PORTAL',
      status: 'ISSUED',
      saleId: sale.id,
      saleFolio: sale.folio || '',
      customerId: customer.id,
      receiverRfc: parsed.receiverRfc || expectedRfc,
      receiverName: parsed.receiverName || customer.name,
      receiverZip: parsed.receiverZip || customer.zip,
      receiverRegimen: parsed.receiverRegimen || customer.regimenFiscal,
      total: parsed.total || round2(safeNumber(sale.total)),
      subtotal: parsed.subtotal || round2(safeNumber(sale.subtotal ?? sale.total)),
      useCfdi: parsed.useCfdi,
      paymentMethod: parsed.paymentMethod,
      paymentForm: parsed.paymentForm,
      serie: parsed.serie || '',
      folio: parsed.folio || '',
      createdAt: parsed.fechaTimbrado || parsed.fecha || new Date().toISOString(),
      importedAt: new Date().toISOString(),
      xmlBase64: Buffer.from(parsed.xml, 'utf8').toString('base64'),
      pdfBase64: duplicateUuid?.pdfBase64 || '',
      lineItems: buildInvoiceLineItemsSnapshot(sale),
      commercialName: db.fiscalConfig?.commercialName || ADMIN_NAME,
      emitterName: parsed.emitterName || db.fiscalConfig?.emitterName || '',
      emitterRfc: parsed.emitterRfc || db.fiscalConfig?.emitterRfc || '',
      warnings,
      cancellation: duplicateUuid?.cancellation || null
    };

    if (duplicateUuid) Object.assign(duplicateUuid, invoiceRecord);
    else db.invoices.push(invoiceRecord);
    updateFiscalCustomerInvoicePreferences(db, customer.id, {
      useCfdi: parsed.useCfdi,
      paymentMethod: parsed.paymentMethod,
      paymentForm: parsed.paymentForm
    });
    const matchingDraft = (db.fiscalDrafts || []).find((entry) => entry.saleId === sale.id && entry.customerId === customer.id && !['IMPORTED','CANCELED'].includes(String(entry.status || '').toUpperCase()));
    if (matchingDraft) {
      matchingDraft.status = 'IMPORTED';
      matchingDraft.invoiceId = invoiceRecord.id;
      matchingDraft.uuid = invoiceRecord.uuid;
      matchingDraft.updatedAt = new Date().toISOString();
    }
    addFiscalEvent(db, 'SAT_XML_IMPORTED', { invoiceId: invoiceRecord.id, uuid: invoiceRecord.uuid, saleId: sale.id, saleFolio: sale.folio || '', customerId: customer.id, warnings });
    writeDb(db);
    res.json({ ok: true, invoice: invoiceRecord, warnings });
  } catch (err) {
    res.status(Number(err.status || 500)).json({ ok: false, message: err.message });
  }
});


app.get('/api/fiscal/drafts', auth, requireAdmin, (_req, res) => {
  const db = readDb();
  res.json({ ok: true, drafts: db.fiscalDrafts || [], events: (db.fiscalEvents || []).slice(0, 250) });
});

app.get('/api/fiscal/readiness', auth, requireAdmin, (_req, res) => {
  const db = readDb();
  const cfg = db.fiscalConfig || {};
  const checks = [
    { key: 'emitterRfc', label: 'RFC emisor', ok: isValidRfcBasic(cfg.emitterRfc || ''), value: cfg.emitterRfc || '' },
    { key: 'emitterName', label: 'Nombre / razón social emisor', ok: Boolean(String(cfg.emitterName || '').trim()), value: cfg.emitterName || '' },
    { key: 'emitterRegimen', label: 'Régimen fiscal emisor', ok: Boolean(String(cfg.emitterRegimen || '').trim()), value: cfg.emitterRegimen || '' },
    { key: 'expeditionZip', label: 'CP de expedición', ok: /^\d{5}$/.test(String(cfg.expeditionZip || '')), value: cfg.expeditionZip || '' }
  ];
  const products = (db.products || []).map((product) => ({
    id: product.id,
    name: product.name,
    sku: product.sku || '',
    productCode: product.fiscalProductCode || '',
    unitCode: product.fiscalUnitCode || '',
    taxObject: product.fiscalTaxObject || '',
    vatRate: safeNumber(product.fiscalVatRate, 0),
    ready: Boolean(String(product.fiscalProductCode || '').trim() && String(product.fiscalUnitCode || '').trim())
  }));
  res.json({ ok: true, checks, products, ready: checks.every((x) => x.ok) && products.every((x) => x.ready), simulatorUrl: cfg.satSimulatorUrl || SAT_SIMULATOR_DEFAULT, portalUrl: cfg.satPortalUrl || SAT_PORTAL_DEFAULT });
});

app.post('/api/fiscal/products/:id', auth, requireAdmin, (req, res) => {
  const db = readDb();
  const product = (db.products || []).find((entry) => entry.id === req.params.id);
  if (!product) return res.status(404).json({ ok: false, message: 'Producto no encontrado.' });
  const body = req.body || {};
  product.fiscalProductCode = String(body.fiscalProductCode ?? product.fiscalProductCode ?? '').trim();
  product.fiscalUnitCode = String(body.fiscalUnitCode ?? product.fiscalUnitCode ?? '').trim();
  product.fiscalUnitName = String(body.fiscalUnitName ?? product.fiscalUnitName ?? '').trim();
  product.fiscalDescription = String(body.fiscalDescription ?? product.fiscalDescription ?? product.name ?? '').trim();
  product.fiscalTaxObject = String(body.fiscalTaxObject ?? product.fiscalTaxObject ?? '').trim();
  if (body.fiscalVatRate !== undefined && body.fiscalVatRate !== '') product.fiscalVatRate = safeNumber(body.fiscalVatRate, 0);
  product.updatedAt = new Date().toISOString();
  addFiscalEvent(db, 'PRODUCT_FISCAL_PROFILE_UPDATED', { productId: product.id, productName: product.name });
  writeDb(db);
  res.json({ ok: true, product });
});

app.post('/api/invoices/:id/sat/attach-pdf', auth, requireAdmin, (req, res) => {
  const db = readDb();
  const invoice = (db.invoices || []).find((entry) => entry.id === req.params.id || entry.uuid === req.params.id);
  if (!invoice) return res.status(404).json({ ok: false, message: 'Factura no encontrada.' });
  const base64 = String(req.body?.contentBase64 || '').trim();
  if (!base64) return res.status(400).json({ ok: false, message: 'No se recibió PDF.' });
  if (base64.length > 14 * 1024 * 1024) return res.status(413).json({ ok: false, message: 'El PDF excede el tamaño permitido.' });
  invoice.pdfBase64 = base64;
  invoice.pdfAttachedAt = new Date().toISOString();
  invoice.pdfFileName = String(req.body?.fileName || 'representacion-impresa.pdf').trim();
  addFiscalEvent(db, 'SAT_PDF_ATTACHED', { invoiceId: invoice.id, uuid: invoice.uuid || invoice.id, saleId: invoice.saleId || '' });
  writeDb(db);
  res.json({ ok: true, invoice });
});

app.post('/api/invoices/:id/sat/cancellation/start', auth, requireAdmin, (req, res) => {
  const db = readDb();
  const invoice = (db.invoices || []).find((entry) => entry.id === req.params.id || entry.uuid === req.params.id);
  if (!invoice) return res.status(404).json({ ok: false, message: 'Factura no encontrada.' });
  const motive = String(req.body?.motive || '02').trim();
  if (!['01','02','03','04'].includes(motive)) return res.status(400).json({ ok: false, message: 'Motivo de cancelación inválido.' });
  const uuidReplacement = String(req.body?.uuidReplacement || '').trim().toUpperCase();
  if (motive === '01' && !uuidReplacement) return res.status(400).json({ ok: false, message: 'El motivo 01 requiere UUID del CFDI sustituto.' });
  invoice.cancellation = {
    ...(invoice.cancellation || {}),
    motive,
    uuidReplacement,
    internalNote: String(req.body?.internalNote || '').trim(),
    status: 'PENDING_PORTAL',
    requestedAt: new Date().toISOString(),
    updatedAt: new Date().toISOString()
  };
  addFiscalEvent(db, 'SAT_CANCELLATION_PREPARED', { invoiceId: invoice.id, uuid: invoice.uuid || invoice.id, motive, uuidReplacement });
  writeDb(db);
  res.json({ ok: true, invoice, cancellation: invoice.cancellation, portalUrl: db.fiscalConfig?.satCancellationUrl || SAT_CANCELLATION_DEFAULT });
});

app.post('/api/invoices/:id/sat/cancellation/result', auth, requireAdmin, (req, res) => {
  const db = readDb();
  const invoice = (db.invoices || []).find((entry) => entry.id === req.params.id || entry.uuid === req.params.id);
  if (!invoice) return res.status(404).json({ ok: false, message: 'Factura no encontrada.' });
  const status = String(req.body?.status || '').trim().toUpperCase();
  const allowed = ['PENDING_ACCEPTANCE','CANCELED','REJECTED','NOT_CANCELED'];
  if (!allowed.includes(status)) return res.status(400).json({ ok: false, message: 'Estado de cancelación inválido.' });
  invoice.cancellation = {
    ...(invoice.cancellation || {}),
    status,
    resultRegisteredAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    acuseFileName: String(req.body?.fileName || invoice.cancellation?.acuseFileName || '').trim(),
    acuseMimeType: String(req.body?.mimeType || invoice.cancellation?.acuseMimeType || '').trim(),
    acuseBase64: String(req.body?.contentBase64 || invoice.cancellation?.acuseBase64 || '').trim()
  };
  if (status === 'CANCELED') invoice.status = 'CANCELED';
  addFiscalEvent(db, 'SAT_CANCELLATION_RESULT', { invoiceId: invoice.id, uuid: invoice.uuid || invoice.id, status });
  writeDb(db);
  res.json({ ok: true, invoice, cancellation: invoice.cancellation });
});

app.post('/api/invoices/generate', auth, requireAdmin, async (req, res) => {
  try {
    const db = readDb();
    const { saleId, customerId, useCfdi, paymentMethod, paymentForm } = req.body || {};
    const sale = (db.sales || []).find((s) => s.id === saleId);
    if (!sale) return res.status(404).json({ ok: false, message: 'Venta no encontrada' });

    const customer = customerId
      ? (db.fiscalCustomers || []).find((c) => c.id === customerId)
      : (db.fiscalCustomers || []).find((c) => c.isDefault) || db.fiscalCustomers?.[0] || defaultFiscalTestCustomer();

    if (!customer) return res.status(404).json({ ok: false, message: 'Cliente fiscal no encontrado' });

    const normalizedCustomer = {
      ...customer,
      name: String(customer.name || '').trim().replace(/\s+/g, ' ').toUpperCase(),
      rfc: String(customer.rfc || '').trim().toUpperCase(),
      regimenFiscal: normalizeFiscalRegimeValue(customer.regimenFiscal || '', '616'),
      zip: String(customer.zip || '').trim(),
      useCfdi: normalizeCfdiUseValue(useCfdi || customer.useCfdi || db.fiscalConfig?.defaultUseCfdi || 'S01')
    };

    if (normalizedCustomer.rfc === 'XAXX010101000' && normalizedCustomer.name === 'PUBLICO EN GENERAL') {
      return res.status(400).json({ ok: false, message: 'Para PUBLICO EN GENERAL debes emitir una factura global con GlobalInformation. Usa un cliente fiscal real para esta prueba.' });
    }

    try {
      await factoramaFetchFirst(['/api/customers/validate', '/customers/validate'], {
        method: 'POST',
        config: db.fiscalConfig || {},
        body: {
          Rfc: normalizedCustomer.rfc,
          Name: normalizedCustomer.name,
          FiscalRegime: normalizedCustomer.regimenFiscal,
          TaxZipCode: normalizedCustomer.zip,
          CfdiUse: normalizedCustomer.useCfdi
        }
      });
    } catch (validationErr) {
      const msg = extractFacturamaErrorMessage(validationErr.payload, validationErr.message || 'El cliente fiscal no es válido.');
      return res.status(Number(validationErr.status || 400)).json({ ok: false, message: `Cliente fiscal inválido: ${msg}`, detail: validationErr.payload || null });
    }

    const payload = buildFacturamaInvoicePayload({
      db,
      config: db.fiscalConfig || {},
      sale,
      customer: normalizedCustomer,
      overrides: { useCfdi: normalizedCustomer.useCfdi, paymentMethod, paymentForm }
    });

    let issued;
    try {
      issued = await factoramaFetchFirst(['/3/cfdis', '/api/3/cfdis'], {
        method: 'POST',
        config: db.fiscalConfig || {},
        body: payload
      });
    } catch (issueErr) {
      issueErr.requestPayload = payload;
      throw issueErr;
    }

    const invoiceRecord = {
      ...normalizeFacturamaInvoiceRecord(issued, sale),
      saleId: sale.id,
      saleFolio: sale.folio,
      customerId: customer.id,
      receiverRfc: normalizedCustomer.rfc,
      receiverName: normalizedCustomer.name,
      total: round2(safeNumber(issued.Total ?? sale.total)),
      useCfdi: normalizedCustomer.useCfdi,
      paymentMethod: normalizePaymentMethodValue(paymentMethod || db.fiscalConfig.defaultPaymentMethod || 'PUE'),
      paymentForm: normalizePaymentFormValue(paymentForm || db.fiscalConfig.defaultPaymentForm || '01'),
      createdAt: issued.Date || new Date().toISOString(),
      apiMode: normalizeFactoramaEnvironment(db.fiscalConfig.factoramaEnvironment || 'sandbox'),
      requestPayload: payload,
      lineItems: buildInvoiceLineItemsSnapshot(sale),
      commercialName: db.fiscalConfig.commercialName || ADMIN_NAME,
      emitterName: db.fiscalConfig.emitterName || '',
      emitterRfc: db.fiscalConfig.emitterRfc || ''
    };

    const existingIdx = (db.invoices || []).findIndex((entry) => entry.id === invoiceRecord.id || (entry.saleId === sale.id && entry.status !== 'CANCELED'));
    if (existingIdx >= 0) db.invoices[existingIdx] = { ...db.invoices[existingIdx], ...invoiceRecord };
    else db.invoices.push(invoiceRecord);
    updateFiscalCustomerInvoicePreferences(db, customer.id, {
      useCfdi: normalizedCustomer.useCfdi,
      paymentMethod: invoiceRecord.paymentMethod,
      paymentForm: invoiceRecord.paymentForm
    });
    writeDb(db);

    res.json({ ok: true, invoice: invoiceRecord });
  } catch (err) {
    const message = extractFacturamaErrorMessage(err.payload, err.message);
    res.status(Number(err.status || 500)).json({ ok: false, message, detail: err.payload || null, requestPayload: err.requestPayload || null });
  }
});

app.post('/api/invoices/:id/cancel', auth, requireAdmin, async (req, res) => {
  try {
    const db = readDb();
    const invoice = (db.invoices || []).find((entry) => entry.id === req.params.id);
    if (!invoice) return res.status(404).json({ ok: false, message: 'Factura no encontrada' });
    if (String(invoice.provider || '').toUpperCase() === 'SAT') {
      return res.status(400).json({ ok: false, message: 'Este CFDI fue importado desde el SAT. La cancelación debe realizarse en el portal oficial del SAT y después actualizarse manualmente en el sistema.' });
    }
    if (!invoice.id) return res.status(400).json({ ok: false, message: 'La factura no tiene identificador CFDI válido para cancelar.' });

    const motive = String(req.body?.motive || db.fiscalConfig.defaultCancelMotive || '02').trim() || '02';
    const uuidReplacement = String(req.body?.uuidReplacement || '').trim();
    if (motive === '01' && !uuidReplacement) {
      return res.status(400).json({ ok: false, message: 'El motivo 01 requiere UUID de reemplazo.' });
    }

    const result = await factoramaFetchFirst([
      `/cfdi/${invoice.id}?type=issued&motive=${encodeURIComponent(motive)}${uuidReplacement ? `&uuidReplacement=${encodeURIComponent(uuidReplacement)}` : ''}`,
      `/api/cfdi/${invoice.id}?type=issued&motive=${encodeURIComponent(motive)}${uuidReplacement ? `&uuidReplacement=${encodeURIComponent(uuidReplacement)}` : ''}`
    ], {
      method: 'DELETE',
      config: db.fiscalConfig || {}
    });

    invoice.status = String(result.Status || 'canceled').toUpperCase();
    invoice.cancellation = {
      motive,
      uuidReplacement,
      requestedAt: new Date().toISOString(),
      status: result.Status || 'canceled',
      responseDate: result.ResponseDate || '',
      requestDate: result.RequestDate || '',
      acuseXmlBase64: result.AcuseXmlBase64 || ''
    };
    writeDb(db);

    res.json({ ok: true, invoice, cancellation: invoice.cancellation });
  } catch (err) {
    res.status(Number(err.status || 500)).json({ ok: false, message: err.message, detail: err.payload || null });
  }
});

app.get('/api/invoices/:id/file/:format', auth, requireAdmin, async (req, res) => {
  try {
    const format = String(req.params.format || '').trim().toLowerCase();
    if (!['pdf', 'xml', 'acuse'].includes(format)) {
      return res.status(400).json({ ok: false, message: 'Formato no soportado. Usa pdf, xml o acuse.' });
    }
    const db = readDb();
    const invoice = (db.invoices || []).find((entry) => entry.id === req.params.id);
    if (!invoice) return res.status(404).json({ ok: false, message: 'Factura no encontrada' });

    let contentBase64 = format === 'pdf' ? (invoice.pdfBase64 || '') : format === 'xml' ? (invoice.xmlBase64 || '') : (invoice.cancellation?.acuseBase64 || '');
    if (!contentBase64 && String(invoice.provider || '').toUpperCase() === 'SAT') {
      const msg = format === 'pdf' ? 'Este CFDI todavía no tiene PDF adjunto.' : format === 'xml' ? 'No hay XML guardado para este CFDI.' : 'No hay acuse de cancelación adjunto.';
      return res.status(404).json({ ok: false, message: msg });
    }
    if (!contentBase64) {
      if (format === 'acuse') return res.status(404).json({ ok: false, message: 'No hay acuse de cancelación disponible.' });
      const raw = await factoramaFetchFirst([
        `/api/Cfdi/${format}/issued/${invoice.id}`,
        `/cfdi/${format}/issued/${invoice.id}`
      ], {
        method: 'GET',
        config: db.fiscalConfig || {},
        expectJson: false
      });
      contentBase64 = extractBase64Content(raw);
      if (!contentBase64) {
        return res.status(502).json({ ok: false, message: `Facturama no devolvió contenido ${format.toUpperCase()} utilizable.` });
      }
      if (format === 'pdf') invoice.pdfBase64 = contentBase64;
      else invoice.xmlBase64 = contentBase64;
      writeDb(db);
    }

    res.json({
      ok: true,
      fileName: format === 'acuse' ? (invoice.cancellation?.acuseFileName || `acuse-cancelacion-${invoice.uuid || invoice.id}.xml`) : `${invoice.serie || 'CFDI'}-${invoice.folio || invoice.uuid || invoice.id}.${format}`,
      mimeType: format === 'pdf' ? 'application/pdf' : format === 'xml' ? 'application/xml' : (invoice.cancellation?.acuseMimeType || 'application/octet-stream'),
      contentBase64
    });
  } catch (err) {
    res.status(Number(err.status || 500)).json({ ok: false, message: err.message, detail: err.payload || null });
  }
});

app.post('/api/print-jobs', auth, (req, res) => {
  const html = String(req.body?.html || '').trim();
  if (!html) return res.status(400).json({ ok: false, message: 'Ticket vacío.' });
  if (Buffer.byteLength(html, 'utf8') > 2 * 1024 * 1024) {
    return res.status(413).json({ ok: false, message: 'El ticket excede el tamaño permitido.' });
  }

  const adminRoom = io.sockets.adapter.rooms.get('admin-printers');
  if (!adminRoom || adminRoom.size < 1) {
    return res.status(503).json({ ok: false, message: 'El nodo ADMIN no está conectado para recibir impresiones.' });
  }

  const now = new Date().toISOString();
  const job = {
    id: uuidv4(),
    html,
    kind: String(req.body?.kind || 'ticket').trim() || 'ticket',
    sourceTerminal: String(req.body?.sourceTerminal || '').trim(),
    sourceName: String(req.body?.sourceName || '').trim(),
    status: 'QUEUED',
    message: '',
    createdAt: now,
    updatedAt: now
  };
  lanPrintJobs.set(job.id, job);
  io.to('admin-printers').emit('print:job', {
    id: job.id,
    html: job.html,
    kind: job.kind,
    sourceTerminal: job.sourceTerminal,
    sourceName: job.sourceName,
    createdAt: job.createdAt
  });
  res.json({ ok: true, job: { ...job, html: undefined } });
});

app.get('/api/print-jobs/:id/status', auth, (req, res) => {
  const job = lanPrintJobs.get(req.params.id);
  if (!job) return res.status(404).json({ ok: false, message: 'Trabajo de impresión no encontrado.' });
  res.json({ ok: true, job: { ...job, html: undefined } });
});

app.post('/api/print-jobs/:id/result', auth, (req, res) => {
  const job = lanPrintJobs.get(req.params.id);
  if (!job) return res.status(404).json({ ok: false, message: 'Trabajo de impresión no encontrado.' });
  job.status = req.body?.success === true ? 'DONE' : 'FAILED';
  job.message = String(req.body?.message || '').trim();
  job.updatedAt = new Date().toISOString();
  lanPrintJobs.set(job.id, job);
  setTimeout(() => lanPrintJobs.delete(job.id), 5 * 60 * 1000).unref?.();
  res.json({ ok: true, job: { ...job, html: undefined } });
});

app.get('/api/invoices', auth, requireAdmin, (_req, res) => {
  res.json({ ok: true, invoices: readDb().invoices.slice().reverse() });
});

// Rutas reservadas para futura sincronización entre sucursales vía Render.
// app.get('/api/branches/:branchId/summary', auth, requireAdmin, (_req, res) => {
//   res.status(501).json({ ok: false, message: 'Módulo de sucursales aún no habilitado' });
// });

io.on('connection', (socket) => {
  if (RENDER_MIRROR_ONLY) {
    const branchCode = socket.data.branchCode;
    const previous = activeBranchRelays.get(branchCode);
    if (previous && previous.id !== socket.id) previous.disconnect(true);
    activeBranchRelays.set(branchCode, socket);
    const waitingQueries = branchRelayConnectionWaiters.get(branchCode) || [];
    branchRelayConnectionWaiters.delete(branchCode);
    waitingQueries.forEach((resumeQuery) => resumeQuery(socket));
    socket.on('disconnect', () => {
      if (activeBranchRelays.get(branchCode)?.id === socket.id) activeBranchRelays.delete(branchCode);
    });
    return;
  }
  socket.on('terminal:hello', (payload) => {
    const terminalType = String(payload?.terminalType || '').toUpperCase();
    const db = readDb();

    if (terminalType !== 'ADMIN') {
      const nodeId = normalizeNodeId13640(payload?.nodeId || '');
      const token = String(payload?.pairingToken || '').trim();
      const terminal = findTerminalForNode13640(db, nodeId, payload?.terminalCode || '');
      if (!terminalTokenValid13640(terminal, nodeId, token)) {
        socket.emit('terminal:authorization-required', {
          code: 'NODE_NOT_AUTHORIZED',
          pairingRequired: true,
          message: 'La terminal fue eliminada, revocada o aún no ha sido autorizada.'
        });
        return;
      }
    } else {
      socket.join('admin-printers');
    }

    touchTerminal(db, { ...(payload || {}), ipAddress: socket.handshake?.address, lastIp: normalizeLanIp(socket.handshake?.address || socket.conn?.remoteAddress || '') });
    writeDb(db);

    socket.emit('inventory:update', {
      products: db.products,
      lowStock: db.products.filter(isLowStock).map((p) => p.id),
      reason: 'HELLO_SYNC'
    });

    socket.emit('tickets:update', {
      tickets: db.openTickets || [],
      reason: 'HELLO_SYNC'
    });
  });
});

server.listen(PORT, '0.0.0.0', () => {
  console.log(`LAN backend escuchando en 0.0.0.0:${PORT}`);
  if (!RENDER_MIRROR_ONLY) configureBranchRelayClient();
});

const udp = RENDER_MIRROR_ONLY ? null : dgram.createSocket('udp4');

udp?.on('message', (msg, rinfo) => {
  try {
    const data = JSON.parse(String(msg));
    if (data.type !== 'DISCOVER_LAN_ADMIN') return;

    const reply = Buffer.from(
      JSON.stringify({ type: 'LAN_ADMIN_HERE', adminName: ADMIN_NAME, port: PORT })
    );

    udp.send(reply, 0, reply.length, rinfo.port, rinfo.address);
  } catch {}
});

udp?.bind(DISCOVERY_PORT, '0.0.0.0', () => {
  console.log(`Descubrimiento UDP en ${DISCOVERY_PORT}`);
});

// ---------- v13.6.13: normalización robusta de unidades importadas ----------
function normalizeUnitBackend13613(raw = '') {
  return String(raw || '')
    .trim()
    .toUpperCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[.]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}
function compactUnitBackend13613(raw = '') {
  return normalizeUnitBackend13613(raw).replace(/\s+/g, '');
}
function isPieceUnitBackend13613(raw = '') {
  return ['PZA','PZ','PIEZA','PIEZAS','UNIDAD','UNIDADES','UND','UD','H87','CAJA','CAJAS','PAQ','PAQUETE','PAQUETES','BOLSA','BOLSAS','BOTELLA','BOTELLAS','LATA','LATAS'].includes(compactUnitBackend13613(raw));
}
function isMeasuredUnitBackend13613(raw = '') {
  const unit = normalizeUnitBackend13613(raw);
  const compact = unit.replace(/\s+/g, '');
  if (['KG','KGS','KILO','KILOS','KILOGRAMO','KILOGRAMOS','KGM','G','GR','GRS','GRAMO','GRAMOS','GM','GRM','L','LT','LTS','LITRO','LITROS','ML','MLS','MILILITRO','MILILITROS'].includes(compact)) return true;
  if (/\b(KG|KILO|KILOS|KILOGRAMO|KILOGRAMOS|KGM|GRAMO|GRAMOS|GRANEL|PESO|LT|LITRO|LITROS|MILILITRO|MILILITROS)\b/.test(unit)) return true;
  return false;
}
function canonicalUnitBackend13613(raw = '') {
  const compact = compactUnitBackend13613(raw);
  if (['KG','KGS','KILO','KILOS','KILOGRAMO','KILOGRAMOS','KGM'].includes(compact)) return 'KG';
  if (['G','GR','GRS','GRAMO','GRAMOS','GM','GRM'].includes(compact)) return 'GR';
  if (['L','LT','LTS','LITRO','LITROS'].includes(compact)) return 'LT';
  if (['ML','MLS','MILILITRO','MILILITROS'].includes(compact)) return 'ML';
  if (isPieceUnitBackend13613(raw)) return 'PZA';
  return normalizeUnitBackend13613(raw) || 'PZA';
}

const __origNormalizeInventoryImportRow13613 = normalizeInventoryImportRow;
normalizeInventoryImportRow = function(row = {}) {
  const parsed = __origNormalizeInventoryImportRow13613(row);
  parsed.unit = canonicalUnitBackend13613(parsed.unit);
  if (isMeasuredUnitBackend13613(parsed.unit)) {
    parsed.allowsFraction = true;
    parsed.saleByAmountEnabled = true;
  } else if (isPieceUnitBackend13613(parsed.unit)) {
    parsed.allowsFraction = false;
    parsed.saleByAmountEnabled = false;
  }
  return parsed;
};

const __origNormalizeProductPayload13613 = normalizeProductPayload;
normalizeProductPayload = function(body = {}, imageUrl = '') {
  const nextBody = { ...(body || {}) };
  nextBody.unit = canonicalUnitBackend13613(nextBody.unit || nextBody.unidad || 'PZA');
  if (isMeasuredUnitBackend13613(nextBody.unit)) nextBody.allowsFraction = true;
  if (isPieceUnitBackend13613(nextBody.unit)) nextBody.allowsFraction = false;
  const product = __origNormalizeProductPayload13613(nextBody, imageUrl);
  product.unit = canonicalUnitBackend13613(product.unit);
  product.allowsFraction = isMeasuredUnitBackend13613(product.unit) ? true : (isPieceUnitBackend13613(product.unit) ? false : !!product.allowsFraction);
  product.saleByAmountEnabled = isMeasuredUnitBackend13613(product.unit);
  return product;
};

// ---------- v13.6.17: promociones robustas por ID, SKU, código y nombre ----------
function normalizePromoKeyBackend13617(value = '') {
  return String(value ?? '')
    .trim()
    .toUpperCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^A-Z0-9]+/g, '')
    .trim();
}
function collectPromotionTargetKeysBackend13617(promotion = {}) {
  const keys = new Set();
  const add = (value) => {
    const key = normalizePromoKeyBackend13617(value);
    if (key) keys.add(key);
  };
  const addArray = (values) => {
    if (!Array.isArray(values)) return;
    values.forEach((value) => {
      if (value && typeof value === 'object') {
        add(value.id); add(value.productId); add(value.sku); add(value.barcode); add(value.codigoBarras); add(value.name); add(value.productName);
      } else add(value);
    });
  };
  add(promotion.productId);
  add(promotion.product_id);
  add(promotion.sku);
  add(promotion.productSku);
  add(promotion.productName);
  add(promotion.nameProduct);
  add(promotion.barcode);
  add(promotion.productBarcode);
  addArray(promotion.productIds);
  addArray(promotion.productSkus);
  addArray(promotion.productNames);
  addArray(promotion.productBarcodes);
  addArray(promotion.productRefs);
  addArray(promotion.products);
  return keys;
}
function collectProductKeysBackend13617(product = {}, productId = '') {
  const keys = new Set();
  [productId, product.id, product.productId, product.sku, product.barcode, product.codigoBarras, product.name].forEach((value) => {
    const key = normalizePromoKeyBackend13617(value);
    if (key) keys.add(key);
  });
  return keys;
}
function promotionMatchesProductBackend13617(promotion = {}, product = {}, productId = '') {
  if (promotion.appliesToAll === true || promotion.allProducts === true || promotion.productScope === 'all') return true;
  const targetKeys = collectPromotionTargetKeysBackend13617(promotion);
  if (!targetKeys.size) return false;
  const productKeys = collectProductKeysBackend13617(product, productId);
  for (const key of targetKeys) if (productKeys.has(key)) return true;
  return false;
}
function uniqueStringsBackend13617(values = []) {
  return Array.from(new Set((values || []).map((value) => String(value || '').trim()).filter(Boolean)));
}

const __origNormalizePromotionPayload13617 = normalizePromotionPayload;
normalizePromotionPayload = function(body = {}) {
  const payload = __origNormalizePromotionPayload13617(body || {});
  const refs = Array.isArray(body.productRefs) ? body.productRefs : [];
  payload.productIds = uniqueStringsBackend13617([...(payload.productIds || []), body.productId, ...(Array.isArray(body.productIds) ? body.productIds : [])]);
  payload.productSkus = uniqueStringsBackend13617([body.productSku, body.sku, ...(Array.isArray(body.productSkus) ? body.productSkus : []), ...refs.map((ref) => ref?.sku)]);
  payload.productNames = uniqueStringsBackend13617([body.productName, ...(Array.isArray(body.productNames) ? body.productNames : []), ...refs.map((ref) => ref?.name || ref?.productName)]);
  payload.productBarcodes = uniqueStringsBackend13617([body.productBarcode, body.barcode, body.codigoBarras, ...(Array.isArray(body.productBarcodes) ? body.productBarcodes : []), ...refs.map((ref) => ref?.barcode || ref?.codigoBarras)]);
  payload.productRefs = refs
    .filter((ref) => ref && typeof ref === 'object')
    .map((ref) => ({
      id: String(ref.id || ref.productId || '').trim(),
      sku: String(ref.sku || '').trim(),
      name: String(ref.name || ref.productName || '').trim(),
      barcode: String(ref.barcode || ref.codigoBarras || '').trim()
    }))
    .filter((ref) => ref.id || ref.sku || ref.name || ref.barcode);
  return payload;
};

getApplicablePromotionsForProduct = function(db, productId, atDate = new Date()) {
  ensurePromotionCollections(db);
  const now = atDate instanceof Date ? atDate : new Date(atDate || Date.now());
  const product = (db.products || []).find((p) => String(p.id) === String(productId))
    || (db.products || []).find((p) => collectProductKeysBackend13617(p, p.id).has(normalizePromoKeyBackend13617(productId)))
    || { id: productId };
  return (db.promotions || []).filter((promotion) => {
    if (promotion.isActive === false) return false;
    if (!promotionMatchesProductBackend13617(promotion, product, productId)) return false;
    if (promotion.startAt) {
      const start = new Date(promotion.startAt);
      if (!Number.isNaN(start.getTime()) && now < start) return false;
    }
    if (promotion.endAt) {
      const end = new Date(promotion.endAt);
      if (!Number.isNaN(end.getTime()) && now > end) return false;
    }
    return true;
  });
};

// ---------- v13.6.19: promociones robustas también para importe exacto ----------
function normalizePromoKeyBackend13619(value = '') {
  return String(value ?? '')
    .trim()
    .toUpperCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/Ñ/g, 'N')
    .replace(/[^A-Z0-9]+/g, '')
    .trim();
}

function collectPromotionTargetKeysBackend13619(promotion = {}) {
  const keys = new Set();
  const add = (value) => {
    const key = normalizePromoKeyBackend13619(value);
    if (key) keys.add(key);
  };
  const addArray = (values) => {
    if (!Array.isArray(values)) return;
    values.forEach((value) => {
      if (value && typeof value === 'object') {
        [value.id, value.productId, value.sku, value.barcode, value.codigoBarras, value.name, value.productName, value.descripcion, value.description].forEach(add);
      } else {
        add(value);
      }
    });
  };
  [promotion.productId, promotion.product_id, promotion.sku, promotion.productSku, promotion.productName, promotion.nameProduct, promotion.barcode, promotion.productBarcode, promotion.codigoBarras].forEach(add);
  addArray(promotion.productIds);
  addArray(promotion.productSkus);
  addArray(promotion.productNames);
  addArray(promotion.productBarcodes);
  addArray(promotion.productRefs);
  addArray(promotion.products);
  return keys;
}

function collectProductKeysBackend13619(product = {}, productId = '') {
  const keys = new Set();
  [productId, product.id, product.productId, product.sku, product.barcode, product.codigoBarras, product.name, product.productName, product.descripcion, product.description].forEach((value) => {
    const key = normalizePromoKeyBackend13619(value);
    if (key) keys.add(key);
  });
  return keys;
}

function promotionMatchesProductBackend13619(promotion = {}, product = {}, productId = '') {
  if (promotion.appliesToAll === true || promotion.allProducts === true || promotion.productScope === 'all') return true;
  const targetKeys = collectPromotionTargetKeysBackend13619(promotion);
  if (!targetKeys.size) return false;
  const productKeys = collectProductKeysBackend13619(product, productId);
  for (const key of targetKeys) if (productKeys.has(key)) return true;
  return false;
}

const __origNormalizePromotionPayload13619 = normalizePromotionPayload;
normalizePromotionPayload = function(body = {}) {
  const payload = __origNormalizePromotionPayload13619(body || {});
  const refs = Array.isArray(body.productRefs) ? body.productRefs : [];
  const addUnique = (list = []) => Array.from(new Set(list.map((value) => String(value || '').trim()).filter(Boolean)));
  payload.productIds = addUnique([...(payload.productIds || []), body.productId, ...(Array.isArray(body.productIds) ? body.productIds : []), ...refs.map((ref) => ref?.id || ref?.productId)]);
  payload.productSkus = addUnique([body.productSku, body.sku, ...(Array.isArray(body.productSkus) ? body.productSkus : []), ...refs.map((ref) => ref?.sku)]);
  payload.productNames = addUnique([body.productName, ...(Array.isArray(body.productNames) ? body.productNames : []), ...refs.map((ref) => ref?.name || ref?.productName)]);
  payload.productBarcodes = addUnique([body.productBarcode, body.barcode, body.codigoBarras, ...(Array.isArray(body.productBarcodes) ? body.productBarcodes : []), ...refs.map((ref) => ref?.barcode || ref?.codigoBarras)]);
  payload.productRefs = refs
    .filter((ref) => ref && typeof ref === 'object')
    .map((ref) => ({
      id: String(ref.id || ref.productId || '').trim(),
      sku: String(ref.sku || '').trim(),
      name: String(ref.name || ref.productName || '').trim(),
      barcode: String(ref.barcode || ref.codigoBarras || '').trim()
    }))
    .filter((ref) => ref.id || ref.sku || ref.name || ref.barcode);
  return payload;
};

getApplicablePromotionsForProduct = function(db, productId, atDate = new Date()) {
  ensurePromotionCollections(db);
  const now = atDate instanceof Date ? atDate : new Date(atDate || Date.now());
  const normalizedProductRef = normalizePromoKeyBackend13619(productId);
  const product = (db.products || []).find((p) => String(p.id) === String(productId))
    || (db.products || []).find((p) => collectProductKeysBackend13619(p, p.id).has(normalizedProductRef))
    || { id: productId };
  return (db.promotions || []).filter((promotion) => {
    if (promotion.isActive === false) return false;
    if (!promotionMatchesProductBackend13619(promotion, product, productId)) return false;
    if (promotion.startAt) {
      const start = new Date(promotion.startAt);
      if (!Number.isNaN(start.getTime()) && now < start) return false;
    }
    if (promotion.endAt) {
      const end = new Date(promotion.endAt);
      if (!Number.isNaN(end.getTime()) && now > end) return false;
    }
    return true;
  });
};

// ---------- v13.6.20: Registro de merma dentro de Caja ----------
function ensureWasteCollections13620(db) {
  ensureCashCollections(db);
  db.inventoryAdjustments = Array.isArray(db.inventoryAdjustments) ? db.inventoryAdjustments : [];
  db.products = Array.isArray(db.products) ? db.products : [];
  db.counters = { ...(db.counters || {}) };
  if (!Number.isFinite(Number(db.counters.wasteSequence))) db.counters.wasteSequence = 0;
}

function buildWasteEntry13620(db, body = {}, user = {}) {
  ensureWasteCollections13620(db);
  const productId = String(body.productId || '').trim();
  const product = (db.products || []).find((row) => String(row.id) === productId);
  if (!product) throw new Error('Selecciona un producto válido para registrar la merma.');
  const quantity = roundQuantity(safeNumber(body.quantity));
  if (quantity <= 0) throw new Error('La cantidad de merma debe ser mayor a 0.');
  const motive = String(body.motive || body.reason || '').trim();
  if (!motive) throw new Error('Escribe el motivo de la merma.');
  if (safeNumber(product.stock) < quantity) {
    throw new Error(`Stock insuficiente para ${product.name}. Disponible: ${safeNumber(product.stock)} ${product.unit || ''}.`);
  }

  const adjustment = applyInventoryAdjustment(db, {
    productId: product.id,
    operation: 'subtract',
    quantity,
    reason: `Merma: ${motive}`,
    type: 'waste',
    sourceType: 'LOCAL_ADMIN_WASTE',
    actorName: user?.name || user?.username || 'Administrador',
    actorUsername: user?.username || 'admin',
    actorBranch: ADMIN_NAME
  });

  db.counters.wasteSequence = safeNumber(db.counters.wasteSequence, 0) + 1;
  const createdAt = new Date().toISOString();
  const unitCost = round2(safeNumber(product.costPrice || product.retailPrice || product.price));
  const inventoryValue = round2(quantity * unitCost);
  const entry = {
    id: uuidv4(),
    kind: 'waste',
    folio: `MER-${String(db.counters.wasteSequence).padStart(6, '0')}`,
    dateKey: normalizeDateKey(body.dateKey),
    amountCashImpact: 0,
    accountingAmount: inventoryValue,
    inventoryValue,
    paymentMethod: 'internal',
    paymentMethodLabel: 'Merma / sin cobro',
    paymentReference: '',
    description: `Merma: ${product.name}`,
    notes: motive,
    motive,
    terminalCode: String(body.terminalCode || 'ADMIN-01').trim() || 'ADMIN-01',
    userId: user?.id || '',
    userName: user?.name || user?.username || 'Administrador',
    productId: product.id,
    productName: product.name,
    sku: product.sku || '',
    quantity,
    unit: product.unit || 'PZA',
    previousStock: adjustment.movement.previousStock,
    newStock: adjustment.movement.newStock,
    inventoryAdjustmentId: adjustment.movement.id,
    printTicketRequested: false,
    createdAt
  };
  db.cashEntries.unshift(entry);
  return { entry, product: adjustment.product, movement: adjustment.movement };
}

const __origBuildCashOverview13620 = buildCashOverview;
buildCashOverview = function(db, dateKey = normalizeDateKey()) {
  const overview = __origBuildCashOverview13620(db, dateKey);
  const wasteEntries = (overview.entries || []).filter((entry) => entry.kind === 'waste');
  const wasteValueTotal = round2(wasteEntries.reduce((acc, entry) => acc + safeNumber(entry.inventoryValue ?? entry.accountingAmount), 0));
  return {
    ...overview,
    wasteValueTotal,
    wasteCount: wasteEntries.length,
    wasteEntries
  };
};

app.post('/api/cash/waste', auth, requireAdmin, (req, res) => {
  const db = readDb();
  try {
    const result = buildWasteEntry13620(db, req.body || {}, req.user || {});
    writeDb(db);
    broadcastInventory(db, 'WASTE', { productId: result.product.id, movement: result.movement, entry: result.entry });
    res.json({ ok: true, entry: result.entry, product: result.product, movement: result.movement, overview: buildCashOverview(db, result.entry.dateKey), inventory: db.products });
  } catch (err) {
    res.status(400).json({ ok: false, message: err.message || 'No se pudo registrar la merma.' });
  }
});
