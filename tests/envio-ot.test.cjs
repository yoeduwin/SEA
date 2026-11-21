// Flujo completo de prepararEnvioOT_ / enviarOT_ sobre un simulador de Apps Script.
// Ejecutar: node --test tests/envio-ot.test.cjs
const { test } = require('node:test');
const assert = require('node:assert/strict');
const { crear } = require('./gas-simulador.cjs');
const PDF = Buffer.from('%PDF-1.4\n% prueba\n').toString('base64');
const base = { folio: 'OTB2611-12', rfc: 'EMP010101AB1', cliente: 'EMPRESA DEMO SA DE CV', sucursal: 'Planta Norte', servicio: 'NOM-035-STPS-2018',
  cuerpo: { giro: 'Alimentos', direccion: 'Av. Reforma 1, Puebla', servicios: [{ servicio: 'NOM-035-STPS-2018', fecha: '2026-11-25', cantidad: '10 puntos' }] } };
const env = (g, extra) => g.ctx.enviarOT_(Object.assign({}, base, { para: 'compras@cliente-demo.com', pdf_base64: PDF }, extra), 'capturista@ejecutivambiental.com');

test('preparar devuelve remitente, CC fijas, asunto y correos del perfil', () => {
  const g = crear();
  const r = g.ctx.prepararEnvioOT_(base);
  assert.equal(r.success, true);
  assert.equal(r.folio, 'OTB2611-0012');
  assert.equal(r.remitente, 'operaciones@ejecutivambiental.com');
  assert.deepEqual([...r.cc_fijos], ['aclientes@ejecutivambiental.com', 'direccion.general@ejecutivambiental.com']);
  assert.equal(r.asunto, 'Orden de Trabajo OTB2611-0012 - EMPRESA DEMO SA DE CV - Ejecutiva Ambiental');
  assert.deepEqual([...r.correos_perfil], ['compras@cliente-demo.com']);
  assert.equal(r.nombre_pdf, 'OT_OTB2611-0012_EMPRESA DEMO SA DE CV.pdf');
  assert.equal(r.ya_enviada, false);
});

test('envío completo: Gmail, Drive, registro y bloqueo de doble envío', () => {
  const g = crear();
  const r = env(g, { cc_extra: 'Jefe@Cliente-Demo.com, jefe@cliente-demo.com' });
  assert.equal(r.success, true, r.error);
  assert.equal(g.enviados.length, 1);
  const m = g.enviados[0];
  assert.equal(m.to, 'compras@cliente-demo.com');
  assert.equal(m.o.from, 'operaciones@ejecutivambiental.com');
  assert.equal(m.o.name, 'Ejecutiva Ambiental - Operaciones');
  assert.equal(m.o.replyTo, 'operaciones@ejecutivambiental.com');
  assert.equal(m.o.cc, 'aclientes@ejecutivambiental.com,direccion.general@ejecutivambiental.com,jefe@cliente-demo.com');
  assert.equal(m.o.attachments[0].getName(), 'OT_OTB2611-0012_EMPRESA DEMO SA DE CV.pdf');
  assert.ok(m.o.htmlBody.includes('Alimentos') && m.body.includes('Fecha: 25/11/2026'));
  const sub = g.suc.getFoldersByName('ORDENES DE TRABAJO').next();
  assert.equal(sub._files.length, 1, 'PDF guardado en Drive');
  const fila = g.sheets.ENVIOS_OT._data[1];
  assert.equal(fila[2], 'OTB2611-0012'); assert.equal(fila[3], 'ENVÍO'); assert.equal(fila[13], 'Enviada');
  assert.equal(fila[11], 'capturista@ejecutivambiental.com');
  assert.ok(fila[12].includes('drive.google.com'));
  // Doble envío
  const r2 = env(g);
  assert.equal(r2.code, 'YA_ENVIADA'); assert.equal(g.enviados.length, 1);
  assert.equal(g.ctx.prepararEnvioOT_(base).ya_enviada, true);
  // Reenvío explícito con motivo
  assert.equal(env(g, { reenviar: true, motivo_reenvio: 'corto' }).code, 'MOTIVO_REENVIO');
  const r3 = env(g, { reenviar: true, motivo_reenvio: 'El cliente no recibió el correo' });
  assert.equal(r3.success, true); assert.equal(r3.tipo, 'REENVÍO');
  assert.equal(g.sheets.ENVIOS_OT._data[2][3], 'REENVÍO');
  assert.equal(g.sheets.ENVIOS_OT._data[2][14], 'El cliente no recibió el correo');
  assert.equal(sub._files.length, 1, 'mismo PDF: no se duplica en Drive');
  // ORDENES_TRABAJO no se modificó
  assert.equal(g.sheets.ORDENES_TRABAJO._data.length, 2);
  assert.equal(g.sheets.ORDENES_TRABAJO._data[1].length, 17);
});

test('falla de Gmail queda como "Error de envío" y permite reintentar', () => {
  const g = crear({ fallarGmail: true });
  const r = env(g);
  assert.equal(r.success, false); assert.equal(r.estatus, 'Error de envío');
  assert.match(r.error, /Gmail no envió/);
  assert.equal(g.sheets.ENVIOS_OT._data[1][13], 'Error de envío');
  g.fallarGmail = false;
  assert.equal(env(g).success, true);
});

test('sin alias "Enviar como" no se envía y el error es claro', () => {
  const g = crear({ aliases: [] });
  const r = env(g);
  assert.equal(r.success, false); assert.match(r.error, /Enviar como/);
  assert.equal(g.enviados.length, 0);
});

test('reglas: OT no registrada, otro cliente, sin Para, PDF falso', () => {
  const g = crear();
  assert.equal(g.ctx.prepararEnvioOT_(Object.assign({}, base, { folio: 'OTB2611-0099' })).code, 'NO_REGISTRADA');
  assert.equal(g.ctx.prepararEnvioOT_(Object.assign({}, base, { rfc: 'OTR010101AB1', cliente: 'OTRA' })).code, 'FOLIO_DE_OTRA_OT');
  assert.equal(env(g, { para: '' }).code, 'SIN_DESTINATARIO');
  assert.equal(env(g, { pdf_base64: Buffer.from('hola').toString('base64') }).code, 'PDF_INVALIDO');
  g.sheets.ORDENES_TRABAJO._data[1][11] = 'CANCELADO';
  assert.equal(g.ctx.prepararEnvioOT_(base).code, 'CANCELADA');
  assert.equal(g.enviados.length, 0);
});

test('prueba PRUEBA-0001: no toca ORDENES_TRABAJO y solo va a internos', () => {
  const g = crear();
  const p = Object.assign({}, base, { folio: 'prueba-1' });
  const prep = g.ctx.prepararEnvioOT_(p);
  assert.equal(prep.success, true); assert.equal(prep.prueba, true); assert.match(prep.asunto, /^\[PRUEBA\]/);
  const ext = g.ctx.enviarOT_(Object.assign({}, p, { para: 'compras@cliente-demo.com', pdf_base64: PDF }), 'u');
  assert.equal(ext.code, 'PRUEBA_EXTERNA');
  const ok = g.ctx.enviarOT_(Object.assign({}, p, { para: 'operaciones@ejecutivambiental.com', pdf_base64: PDF, actualizar_perfil: true }), 'u');
  assert.equal(ok.success, true, ok.error); assert.equal(ok.tipo, 'PRUEBA'); assert.equal(ok.cambio_perfil, '');
  assert.equal(g.ctx.enviarOT_(Object.assign({}, p, { para: 'operaciones@ejecutivambiental.com', pdf_base64: PDF }), 'u').code, 'YA_ENVIADA');
  assert.equal(g.sheets.ORDENES_TRABAJO._data.length, 2);
  assert.ok(g.raiz.getFoldersByName('_PRUEBAS_ENVIO_OT').hasNext());
});

test('actualizar perfil: va a CAMBIOS_PENDIENTES, nunca directo al maestro', () => {
  const g = crear();
  const r = env(g, { para: 'nuevo@cliente-demo.com, compras@cliente-demo.com', actualizar_perfil: true });
  assert.equal(r.success, true); assert.equal(r.cambio_perfil, 'PENDIENTE');
  const c = g.sheets.CAMBIOS_PENDIENTES._data[1];
  assert.deepEqual([c[5], c[7], c[8], c[9], c[11]], ['correo_informe', 'compras@cliente-demo.com', 'nuevo@cliente-demo.com, compras@cliente-demo.com', 'SEAOT', 'PENDIENTE']);
  assert.equal(g.sheets.CLIENTES_MAESTRO._data[1][8], 'compras@cliente-demo.com');
});

test('doPost: acciones registradas y protegidas por sesión de Google', () => {
  const g = crear();
  const out = g.ctx.doPost({ postData: { contents: JSON.stringify({ action: 'enviarOT' }) } });
  assert.match(out.t, /Se requiere iniciar sesión/);
  assert.equal(g.run("AUTH_MODE.prepararEnvioOT + ACTION_MODULE.enviarOT"), 'GOOGLESEAOT');
});

test('Codex P1: una liga vieja a la raíz no desvía el PDF fuera de la sucursal', () => {
  const g = crear();
  g.sheets.ORDENES_TRABAJO._data[1][12] = g.raiz.getUrl();
  const r = env(g);
  assert.equal(r.success, true, r.error);
  assert.equal(g.raiz.getFoldersByName('ORDENES DE TRABAJO').hasNext(), false, 'nada en la raíz');
  assert.equal(g.suc.getFoldersByName('ORDENES DE TRABAJO').next()._files.length, 1, 'PDF en la sucursal');
});

test('Codex P2: si falla el registro después de enviar, un reintento queda bloqueado', () => {
  const g = crear();
  const original = g.ctx.cerrarEnvioOT_;
  g.ctx.cerrarEnvioOT_ = () => { throw new Error('Sheets no disponible'); };
  assert.equal(env(g).success, true);
  g.ctx.cerrarEnvioOT_ = original;
  assert.equal(g.sheets.ENVIOS_OT._data[1][13], 'Enviando');
  const r = env(g);
  assert.equal(r.code, 'YA_ENVIADA');
  assert.match(r.error, /no se pudo confirmar/);
  assert.equal(g.enviados.length, 1, 'no hay segundo correo');
  assert.equal(env(g, { reenviar: true, motivo_reenvio: 'Verificado: no salió el primero' }).success, true);
});

test('el intento queda escrito antes de llamar a Gmail', () => {
  const g = crear();
  let estatusAlEnviar = null;
  const send = g.ctx.GmailApp.sendEmail;
  g.ctx.GmailApp.sendEmail = (...a) => { estatusAlEnviar = g.sheets.ENVIOS_OT._data[1][13]; return send(...a); };
  assert.equal(env(g).success, true);
  assert.equal(estatusAlEnviar, 'Enviando');
  assert.equal(g.sheets.ENVIOS_OT._data[1][13], 'Enviada');
  assert.equal(g.sheets.ENVIOS_OT._data.length, 2, 'una sola fila por intento');
});

test('Codex: reintento de un reenvío cuya respuesta se perdió no manda otro correo', () => {
  const g = crear();
  assert.equal(env(g, { intento_id: 'primer-intento-0001' }).success, true);
  const reenvio = { reenviar: true, motivo_reenvio: 'El cliente pidió que se reenviara', intento_id: 'reenvio-intento-0002' };
  assert.equal(env(g, reenvio).success, true);          // Gmail lo mandó; la respuesta "se perdió"
  assert.equal(g.enviados.length, 2);
  // La recuperación consulta ESTE intento, no el historial de la OT.
  assert.equal(g.ctx.prepararEnvioOT_(Object.assign({}, base, { intento_id: 'reenvio-intento-0002' })).estado_intento, 'Enviada');
  const r = env(g, reenvio);                              // clic en "Reintentar envío" con el reenvío aún marcado
  assert.equal(r.success, true); assert.equal(r.duplicado, true);
  assert.equal(g.enviados.length, 2, 'no hay tercer correo');
});

test('un intento que quedó en "Enviando" no se reintenta con el mismo id', () => {
  const g = crear();
  const original = g.ctx.cerrarEnvioOT_;
  g.ctx.cerrarEnvioOT_ = () => { throw new Error('Sheets no disponible'); };
  env(g, { intento_id: 'intento-sin-cierre-01' });
  g.ctx.cerrarEnvioOT_ = original;
  assert.equal(g.ctx.prepararEnvioOT_(Object.assign({}, base, { intento_id: 'intento-sin-cierre-01' })).estado_intento, 'Enviando');
  assert.equal(env(g, { intento_id: 'intento-sin-cierre-01', reenviar: true, motivo_reenvio: 'Reintento automático' }).code, 'SIN_CONFIRMAR');
  assert.equal(g.enviados.length, 1);
});

test('tras un error de Gmail, el mismo intento sí puede reintentarse', () => {
  const g = crear({ fallarGmail: true });
  assert.equal(env(g, { intento_id: 'intento-con-error-01' }).estatus, 'Error de envío');
  g.fallarGmail = false;
  assert.equal(env(g, { intento_id: 'intento-con-error-01' }).success, true);
  assert.equal(g.enviados.length, 1);
});
