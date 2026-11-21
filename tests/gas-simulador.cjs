// Simulador mínimo de Apps Script para probar enviarOT_ / prepararEnvioOT_.
const fs = require('fs'), vm = require('vm');
function crear(opts = {}) {
  const sheets = {};
  const fmt = v => v instanceof Date ? v.toISOString() : String(v == null ? '' : v);
  function hoja(nombre, filas) {
    const data = filas || [];
    const self = {
      _data: data, getName: () => nombre,
      getLastRow: () => data.length, getMaxColumns: () => 30,
      appendRow: r => { data.push(r.slice()); return self; },
      getDataRange: () => self.getRange(1, 1, Math.max(1, data.length), Math.max(1, ...data.map(r => r.length))),
      getRange: (r, c, nr = 1, nc = 1) => ({
        getValues: () => Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => (data[r - 1 + i] || [])[c - 1 + j] ?? '')),
        getDisplayValues() { return this.getValues().map(x => x.map(fmt)); },
        setValues: v => { v.forEach((row, i) => { data[r - 1 + i] = data[r - 1 + i] || []; row.forEach((x, j) => data[r - 1 + i][c - 1 + j] = x); }); },
        setValue: x => { data[r - 1] = data[r - 1] || []; data[r - 1][c - 1] = x; },
        setFontWeight() { return this; }, setBackground() { return this; }, setFontColor() { return this; }
      }),
      setFrozenRows() {}, setColumnWidth() {}
    };
    return self;
  }
  const ss = { getSheetByName: n => sheets[n] || null, insertSheet: n => (sheets[n] = hoja(n)) };
  const folders = {};
  let fid = 0;
  function carpeta(nombre, padre) {
    const id = 'F' + (++fid) + 'xxxxxxxxxxxxxxxxxxxxxxxxx';
    const files = [], hijos = [];
    const f = { getId: () => id, getName: () => nombre, getUrl: () => 'https://drive.google.com/drive/folders/' + id,
      getParents: () => { let d = false; return { hasNext: () => !d && !!padre, next: () => { d = true; return padre; } }; },
      getFoldersByName: n => { const l = hijos.filter(h => h.getName() === n); return { hasNext: () => l.length > 0, next: () => l.shift() }; },
      getFolders: () => { const l = hijos.slice(); return { hasNext: () => l.length > 0, next: () => l.shift() }; },
      createFolder: n => { const h = carpeta(n, f); hijos.push(h); return h; },
      getFilesByName: n => { const l = files.filter(x => x.getName() === n); return { hasNext: () => l.length > 0, next: () => l.shift() }; },
      createFile: b => { const bytes = b.getBytes(); const name = b.getName(); const fl = { getName: () => name, getSize: () => bytes.length,
        getBlob: () => ({ getBytes: () => bytes }), getUrl: () => 'https://drive.google.com/file/d/' + name }; files.push(fl); return fl; },
      _files: files };
    folders[id] = f;
    return f;
  }
  const raiz = carpeta('RAIZ');
  const enviados = [];
  let fallarGmail = opts.fallarGmail || false;
  const blob = (bytes, mime, name) => ({ getBytes: () => bytes, getName: () => name, setName(n) { name = n; return this; },
    copyBlob() { return blob(bytes, mime, name); }, getContentType: () => mime });
  const ctx = vm.createContext({
    console, Logger: { log: () => {} },
    SpreadsheetApp: { openById: () => ss, flush() {} },
    DriveApp: { getFolderById: id => { if (!folders[id]) throw new Error('no folder'); return folders[id]; } },
    GmailApp: { getAliases: () => opts.aliases || ['operaciones@ejecutivambiental.com'],
      sendEmail: (to, subject, body, o) => { if (fallarGmail) throw new Error('Service unavailable'); enviados.push({ to, subject, body, o }); } },
    MailApp: { getRemainingDailyQuota: () => 100 },
    Session: { getEffectiveUser: () => ({ getEmail: () => 'ejecutiva@gmail.com' }) },
    CacheService: { getScriptCache: () => ({ get: () => null, put() {} }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) },
    Utilities: { getUuid: () => Math.random().toString(36).slice(2), formatDate: d => d.toISOString().slice(0, 19).replace('T', ' '),
      base64Decode: s => Array.from(Buffer.from(s, 'base64')).map(b => b > 127 ? b - 256 : b), newBlob: blob,
      computeDigest: (_a, bytes) => Array.from(require('crypto').createHash('md5').update(Buffer.from(bytes.map(b => b & 255))).digest()),
      DigestAlgorithm: { MD5: 'MD5' } },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    ContentService: { createTextOutput: t => ({ t, setMimeType() { return this; } }), MimeType: { JSON: 'json' } }
  });
  vm.runInContext(fs.readFileSync(require('path').join(__dirname, '..', 'BACKEND_FIXES.gs'), 'utf8'), ctx, { filename: 'BACKEND_FIXES.gs' });
  // Datos: cliente con carpeta RFC → sucursal, y una OT registrada.
  vm.runInContext(`CONFIG.FOLDER_ID = ${JSON.stringify(raiz.getId())};`, ctx);
  const padre = raiz.createFolder('EMP010101AB1 - EMPRESA DEMO');
  const suc = padre.createFolder('Planta Norte');
  const filaCli = Array(22).fill(''); Object.assign(filaCli, { 0: '01/11/2026', 1: 'EMPRESA DEMO SA DE CV', 2: 'Planta Norte', 3: 'EMP010101AB1', 8: 'compras@cliente-demo.com', 20: suc.getUrl() });
  sheets.CLIENTES_MAESTRO = hoja('CLIENTES_MAESTRO', [Array(22).fill('h'), filaCli]);
  const filaOt = ['01/11/2026', 'OTB2611-0012', 'OTB', 'NOM-035-STPS-2018', 'EMPRESA DEMO SA DE CV', 'Planta Norte', 'EMP010101AB1',
    'Martín Luna', '2026-11-25', '', '', 'NO INICIADO', suc.getUrl(), '', '', '', ''];
  sheets.ORDENES_TRABAJO = hoja('ORDENES_TRABAJO', [Array(17).fill('h'), filaOt]);
  return { ctx, run: c => vm.runInContext(c, ctx), sheets, enviados, suc, raiz, set fallarGmail(v) { fallarGmail = v; } };
}
module.exports = { crear };
