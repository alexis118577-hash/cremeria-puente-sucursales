const crypto = require('crypto');
const { execFileSync } = require('child_process');

const SECRET_PREFIX = 'CGSECRET:v1:';
const KEY_PREFIX = 'CGDPAPI:v1:';
let migrationPending = false;
const keyCache = new Map();

function runDpapi(action, value) {
  if (process.platform !== 'win32') {
    throw new Error('La protección de credenciales requiere Windows DPAPI en el nodo administrador.');
  }

  const script = action === 'protect'
    ? "$ErrorActionPreference='Stop'; Add-Type -AssemblyName System.Security; $plain=[Convert]::FromBase64String([Console]::In.ReadToEnd()); $protected=[System.Security.Cryptography.ProtectedData]::Protect($plain,$null,[System.Security.Cryptography.DataProtectionScope]::CurrentUser); [Console]::Out.Write([Convert]::ToBase64String($protected))"
    : "$ErrorActionPreference='Stop'; Add-Type -AssemblyName System.Security; $protected=[Convert]::FromBase64String([Console]::In.ReadToEnd()); $plain=[System.Security.Cryptography.ProtectedData]::Unprotect($protected,$null,[System.Security.Cryptography.DataProtectionScope]::CurrentUser); [Console]::Out.Write([Convert]::ToBase64String($plain))";

  try {
    const result = execFileSync('powershell.exe', ['-NoLogo', '-NoProfile', '-NonInteractive', '-Command', script], {
      input: String(value || ''),
      encoding: 'utf8',
      windowsHide: true,
      timeout: 10000,
      maxBuffer: 1024 * 1024
    }).trim();
    if (!result) throw new Error('empty result');
    return result;
  } catch (_) {
    throw new Error('No se pudo acceder a la protección de credenciales de Windows.');
  }
}

function secretRefs(db) {
  const refs = [];
  for (const user of Array.isArray(db.users) ? db.users : []) refs.push([user, 'password']);
  for (const branch of Array.isArray(db.branches) ? db.branches : []) {
    refs.push([branch, 'apiKey'], [branch, 'remoteAdminPassword']);
  }
  refs.push([db.security?.remoteAccess, 'apiKey']);
  refs.push([db.publication, 'apiKey'], [db.publication, 'remoteAdminPassword']);
  refs.push([db.fiscalConfig, 'factoramaApiPassword']);
  return refs.filter(([parent, key]) => parent && typeof parent === 'object' && typeof parent[key] === 'string' && parent[key]);
}

function getProtectedKey(db, create = true) {
  db.security ||= {};
  db.security.secretStore ||= {};
  let protectedKey = String(db.security.secretStore.protectedKey || '');
  if (!protectedKey && create) {
    const rawKey = crypto.randomBytes(32).toString('base64');
    protectedKey = KEY_PREFIX + runDpapi('protect', rawKey);
    db.security.secretStore = { version: 1, protectedKey };
  }
  if (!protectedKey.startsWith(KEY_PREFIX)) {
    if (!protectedKey) return null;
    throw new Error('La clave protegida de la base no tiene un formato reconocido.');
  }
  if (keyCache.has(protectedKey)) return keyCache.get(protectedKey);
  const rawKey = Buffer.from(runDpapi('unprotect', protectedKey.slice(KEY_PREFIX.length)), 'base64');
  if (rawKey.length !== 32) throw new Error('La clave protegida de la base no es válida.');
  keyCache.set(protectedKey, rawKey);
  return rawKey;
}

function encryptValue(value, key) {
  const iv = crypto.randomBytes(12);
  const cipher = crypto.createCipheriv('aes-256-gcm', key, iv);
  const encrypted = Buffer.concat([cipher.update(String(value), 'utf8'), cipher.final()]);
  return SECRET_PREFIX + Buffer.concat([iv, cipher.getAuthTag(), encrypted]).toString('base64');
}

function decryptValue(value, key) {
  try {
    const payload = Buffer.from(String(value).slice(SECRET_PREFIX.length), 'base64');
    if (payload.length < 29) throw new Error('invalid');
    const iv = payload.subarray(0, 12);
    const tag = payload.subarray(12, 28);
    const encrypted = payload.subarray(28);
    const decipher = crypto.createDecipheriv('aes-256-gcm', key, iv);
    decipher.setAuthTag(tag);
    return Buffer.concat([decipher.update(encrypted), decipher.final()]).toString('utf8');
  } catch (_) {
    throw new Error('No se pudo descifrar una credencial protegida de la base.');
  }
}

function decryptDatabaseSecrets(db) {
  const refs = secretRefs(db);
  const protectedKey = db.security?.secretStore?.protectedKey;
  if (!refs.length && !protectedKey) return false;
  const key = getProtectedKey(db, true);
  let changed = false;
  for (const [parent, field] of refs) {
    const value = parent[field];
    if (value.startsWith(SECRET_PREFIX)) {
      parent[field] = decryptValue(value, key);
    } else {
      changed = true;
    }
  }
  if (changed) migrationPending = true;
  return changed;
}

function encryptDatabaseSecrets(db) {
  const refs = secretRefs(db);
  if (!refs.length) return db;
  const key = getProtectedKey(db, true);
  for (const [parent, field] of refs) {
    if (!parent[field].startsWith(SECRET_PREFIX)) parent[field] = encryptValue(parent[field], key);
  }
  return db;
}

function isSecretMigrationPending() {
  return migrationPending;
}

function clearSecretMigrationPending() {
  migrationPending = false;
}

module.exports = {
  decryptDatabaseSecrets,
  encryptDatabaseSecrets,
  isSecretMigrationPending,
  clearSecretMigrationPending
};
