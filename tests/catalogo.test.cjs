// Catálogo de servicios y reglas de serie compartidas entre SEAOT y el backend.
// Ejecutar: node --test tests/catalogo.test.cjs
const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs'), path = require('path'), vm = require('vm');
const raiz = path.join(__dirname, '..');

function catalogo() {
  const ctx = vm.createContext({ window: {} });
  vm.runInContext(fs.readFileSync(path.join(raiz, 'catalogo-servicios.js'), 'utf8'), ctx);
  return ctx.window.CATALOGO_SERVICIOS;
}
function backend() {
  const ctx = vm.createContext({ console });
  vm.runInContext(fs.readFileSync(path.join(raiz, 'BACKEND_FIXES.gs'), 'utf8'), ctx);
  return ctx;
}

test('las NOM sinInforme del catálogo coinciden con NOM_SIN_INFORME_ del backend', () => {
  const cat = catalogo().filter(s => s.sinInforme).map(s => s.nombre).sort().join('|');
  const gs = vm.runInContext('NOM_SIN_INFORME_.slice().sort().join("|")', backend());
  assert.equal(cat, gs);
});

test('las NOM sinInforme van en OTC y no se rechazan como norma con informe', () => {
  const b = backend();
  for (const s of catalogo().filter(x => x.sinInforme)) {
    assert.equal(s.serie, 'OTC', s.nombre);
    assert.equal(b.serviciosNomEnLista_(s.nombre).length, 0, s.nombre);
  }
  assert.deepEqual([...b.serviciosNomEnLista_('NOM-043-SEMARNAT-1993, NOM-025-STPS-2008')], ['NOM-025-STPS-2008']);
});

test('el QR codifica la liga de WhatsApp de Atención a Clientes con el folio', () => {
  const html = fs.readFileSync(path.join(raiz, 'SEAOT.html'), 'utf8');
  assert.match(html, /WHATSAPP_ATENCION_ = '525652821561'/);
  const ctx = vm.createContext({});
  vm.runInContext(fs.readFileSync(path.join(raiz, 'qr-codigo.js'), 'utf8'), ctx);
  const m = ctx.QRCodigo.matriz('https://wa.me/525652821561?text=' + encodeURIComponent('Hola, tengo una duda sobre la Orden de Trabajo OTB2611-0318'));
  assert.ok(m.length >= 21 && m.length % 4 === 1, 'tamaño de QR válido');
  // patrones de posición en las tres esquinas
  for (const [r, c] of [[0, 0], [0, m.length - 7], [m.length - 7, 0]]) {
    assert.equal(m[r][c] && m[r + 6][c + 6] && !m[r + 1][c + 1] && m[r + 3][c + 3], true);
  }
});
