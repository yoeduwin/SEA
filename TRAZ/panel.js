'use strict';
// Las acciones disponibles continúan limitadas a consultas del backend existente.
const API_URL = 'https://script.google.com/macros/s/AKfycbw9A7lgw0zI4f3V7YLVnuplLoShTS-5Qfprtw-E-iKTTMNd-F9Sh6_u3y2d3Xg0o3s/exec';
let RESUMEN = [], HUERFANOS = [], seleccion = null, ultimaActualizacion = null;
let autorizado = false, cargando = false, generacionDetalle = 0, serieActiva = '', focoAnterior = null;
let clavePreferencias = null, preferenciasCargadas = false, recientes = [];
const el = id => document.getElementById(id);
function prepararPreferencias(usuario) {
  const identidad = usuario && (usuario.sub || usuario.email);
  clavePreferencias = identidad ? 'seatraz.consulta.v1:' + identidad : null;
  preferenciasCargadas = false;
  recientes = [];
}
function restaurarPreferencias() {
  if(preferenciasCargadas) return;
  preferenciasCargadas = true;
  if(!clavePreferencias) return;
  try {
    const prefs = JSON.parse(localStorage.getItem(clavePreferencias) || 'null');
    if(!prefs || prefs.version !== 1) return;
    const filtros = prefs.filtros || {};
    for(const id of ['buscar','estado','servicio','personal']) {
      const valor = typeof filtros[id] === 'string' ? filtros[id].slice(0,300) : '';
      if(id === 'buscar' || [...el(id).options].some(o => o.value === valor)) el(id).value = valor;
    }
    serieActiva = ['OT','OTB','OTC'].includes(filtros.serie) ? filtros.serie : '';
    document.querySelectorAll('input[name="serie"]').forEach(r => r.checked = r.value === serieActiva);
    recientes = Array.isArray(prefs.recientes) ? [...new Set(prefs.recientes.filter(f => typeof f === 'string' && f.length <= 120))].slice(0,6) : [];
    el('preferences-note').textContent = ['buscar','estado','servicio','personal'].some(id => el(id).value) || serieActiva ? 'Se recuperaron tus últimos filtros.' : '';
  } catch(_) { /* La consulta funciona aunque el navegador bloquee almacenamiento. */ }
}
function guardarPreferencias() {
  if(!clavePreferencias || !preferenciasCargadas) return;
  const filtros = {serie:serieActiva};
  ['buscar','estado','servicio','personal'].forEach(id => filtros[id] = el(id).value.slice(0,300));
  try { localStorage.setItem(clavePreferencias, JSON.stringify({version:1,filtros,recientes:recientes.slice(0,6)})); } catch(_) {}
}
function pintarRecientes() {
  const visibles = recientes.map(f => RESUMEN.find(o => o.ot === f)).filter(Boolean);
  el('recientes').innerHTML = visibles.length ? '<div class="recent-label">Consultadas recientemente</div><div class="recent-links">' + visibles.map(o => '<button class="recent-order" data-ot="' + esc(o.ot) + '"><strong>' + esc(o.ot) + '</strong><span>' + esc(o.cliente || o.nom || '') + '</span></button>').join('') + '</div><button class="clear-recent" id="clear-recent">Limpiar recientes</button>' : '';
  el('recientes').hidden = !visibles.length;
}
function registrarConsulta(ot) {
  recientes = [ot,...recientes.filter(f => f !== ot)].slice(0,6);
  guardarPreferencias();
  pintarRecientes();
}

const estadoDe = o => o.estado_traz || o.estatus_ot || 'SIN ESTADO';
const normalizar = s => String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
// Coincidencia exacta: OT nunca debe incluir OTB/OTC. No inventar una serie si falta.
function serieDe(o) {
  const tipo = String(o.tipo || '').trim().toUpperCase();
  if(['OT','OTB','OTC'].includes(tipo)) return tipo;
  const m = String(o.ot || '').trim().toUpperCase().match(/^(OTB|OTC|OT)(?=[\s\d-]|$)/);
  return m ? m[1] : '';
}
const alertasDe = o => [!o.tiene_carpeta && 'Sin carpeta de OT', !Number(o.num_informes) && serieDe(o) !== 'OTC' && 'Sin informe relacionado'].filter(Boolean);
function filtrados() {
  const q = normalizar(el('buscar').value.trim());
  return RESUMEN.filter(o =>
    (!q || normalizar([o.ot,o.cliente,o.sucursal,o.nom,o.personal,estadoDe(o),...(o.folios_informe || [])].join(' ')).includes(q)) &&
    (!serieActiva || serieDe(o) === serieActiva) &&
    (!el('estado').value || estadoDe(o) === el('estado').value) &&
    (!el('servicio').value || o.nom === el('servicio').value) &&
    (!el('personal').value || o.personal === el('personal').value));
}
function actualizarOpciones(id, valores) {
  const selector = el(id), previo = selector.value;
  while(selector.options.length > 1) selector.remove(1);
  [...new Set(valores.filter(Boolean))].sort((a,b) => a.localeCompare(b,'es')).forEach(v => selector.add(new Option(v,v)));
  selector.value = [...selector.options].some(o => o.value === previo) ? previo : '';
}
function banner(texto, error = false) {
  el('notice').textContent = texto;
  el('notice').className = texto ? (error ? 'notice error-notice' : 'notice') : '';
}
async function consultar(action, extra = {}) {
  if(!['trazResumen','trazDetalle'].includes(action)) throw new Error('Acción fuera del modo consulta.');
  return api({...extra, action});
}
async function cargarResumen() {
  if(!autorizado || cargando) return;
  cargando = true;
  el('btnRefresh').disabled = true;
  el('sync').textContent = 'Consultando SEA…';
  banner('Consultando los registros existentes…');
  try {
    const r = await consultar('trazResumen');
    if(!r.success || !Array.isArray(r.data)) throw new Error(r.message || r.error || 'Respuesta de consulta no válida.');
    RESUMEN = r.data.slice().reverse();
    HUERFANOS = Array.isArray(r.huerfanos) ? r.huerfanos : [];
    ultimaActualizacion = new Intl.DateTimeFormat('es-MX',{timeZone:'America/Mexico_City',dateStyle:'medium',timeStyle:'short'}).format(new Date());
    el('sync').textContent = 'Última consulta: ' + ultimaActualizacion + ' · CDMX';
    actualizarOpciones('estado', RESUMEN.map(estadoDe));
    actualizarOpciones('servicio', RESUMEN.map(o => o.nom));
    actualizarOpciones('personal', RESUMEN.map(o => o.personal));
    restaurarPreferencias();
    recientes = recientes.filter(f => RESUMEN.some(o => o.ot === f));
    guardarPreferencias();
    banner('');
    render();
    pintarRecientes();
    if(seleccion) {
      if(RESUMEN.some(o => o.ot === seleccion)) await abrir(seleccion, false);
      else cerrarDetalle();
    }
  } catch(e) {
    el('sync').textContent = ultimaActualizacion ? 'Consulta anterior: ' + ultimaActualizacion + ' · CDMX' : 'Sin consulta disponible';
    banner((ultimaActualizacion ? 'No se pudo actualizar. Se conserva la consulta anterior. ' : 'No se pudo cargar SEATRAZ. ') + e.message, true);
  } finally {
    cargando = false;
    el('btnRefresh').disabled = false;
  }
}
function celdaInformes(o) {
  const folios = o.folios_informe || [];
  const sinFolio = Math.max(0,(Number(o.num_informes)||0)-folios.length);
  if(!folios.length && !sinFolio) return serieDe(o) === 'OTC' ? '<span class="cell-sub">No aplica · OTC</span>' : '<span class="alert-chip">Sin informe relacionado</span>';
  return folios.map(f => `<button class="report-folio report-link" data-ot="${esc(o.ot)}" data-report="${esc(f)}" aria-label="Consultar informe ${esc(f)}">${esc(f)}</button>`).join('') + (sinFolio ? `<span class="alert-chip">${sinFolio} registro(s) sin folio</span>` : '');
}
function tabla(items) {
  if(!items.length) return '<div class="empty">No hay resultados con estos filtros. <button class="open-ot" data-clear>Limpiar filtros</button></div>';
  return `<div class="table-scroll"><table class="orders"><thead><tr><th>Orden</th><th>Cliente / sucursal</th><th>Servicio</th><th>Estado TRAZ</th><th>Informes</th><th>Expediente</th></tr></thead><tbody>${items.map(o => `<tr><td><button class="order-link" data-ot="${esc(o.ot)}">${esc(o.ot)}</button><span class="cell-sub">${esc(o.fecha_visita || 'Sin fecha de visita')}</span></td><td><strong>${esc(o.cliente || 'Sin cliente registrado')}</strong><span class="cell-sub">${esc(o.sucursal || 'Sin sucursal registrada')}</span></td><td>${esc(o.nom || 'Sin servicio')}</td><td><span class="state-mini ${estadoClase(estadoDe(o))}">${esc(estadoDe(o))}</span><button class="change-link" data-ot="${esc(o.ot)}" data-changes>Ver cambios</button></td><td>${celdaInformes(o)}</td><td><button class="open-ot" data-ot="${esc(o.ot)}" data-documents>Ver documentos</button>${!o.tiene_carpeta ? '<span class="alert-chip">Sin carpeta de OT</span>' : ''}</td></tr>`).join('')}</tbody></table></div>`;
}
function huerfanos() {
  if(!HUERFANOS.length) return '';
  return `<details class="orphan-section"><summary>Informes con OT no encontrada <span class="count-pill">${HUERFANOS.length}</span></summary><p class="section-note">Consulta global, independiente de los filtros.</p><div class="table-scroll"><table class="orders"><thead><tr><th>Informe</th><th>OT referenciada</th></tr></thead><tbody>${HUERFANOS.map(h => `<tr><td>${esc(h.num_informe || 'Sin folio')}</td><td>${esc(h.ot)}</td></tr>`).join('')}</tbody></table></div></details>`;
}
function render() {
  const items = filtrados();
  el('extra-label').textContent = el('personal').value ? 'Más filtros · 1 activo' : 'Más filtros';
  el('results').innerHTML = `<section class="card result-card"><div class="card-head"><h3>Órdenes e informes</h3><span>${items.length} de ${RESUMEN.length} órdenes</span></div>${tabla(items)}</section>${huerfanos()}`;
}
function cerrarDetalle() {
  seleccion = null;
  generacionDetalle++;
  const drawer = el('detail-shell');
  if(drawer.open) drawer.close();
  document.body.classList.remove('drawer-open');
  el('detalle').innerHTML = '';
  if(focoAnterior && focoAnterior.isConnected) focoAnterior.focus({preventScroll:true});
  focoAnterior = null;
}
async function abrir(ot, enfocar = true, destino = '') {
  if(destino === true) destino = 'documentos';
  seleccion = ot;
  const version = ++generacionDetalle, drawer = el('detail-shell');
  if(!drawer.open) {
    focoAnterior = document.activeElement;
    drawer.showModal();
    document.body.classList.add('drawer-open');
  }
  el('detail-title').textContent = 'Orden ' + ot;
  el('detalle').innerHTML = '<div class="loading" role="status">Consultando el servicio…</div>';
  if(enfocar) {drawer.scrollTop = 0; el('close-detail').focus({preventScroll:true});}
  try {
    const r = await consultar('trazDetalle',{ot});
    if(version !== generacionDetalle || ot !== seleccion) return;
    if(!r.success) throw new Error(r.message || r.error || 'No se pudo consultar la orden.');
    // OTC es una serie sin informe. No presentar su ausencia como incidencia.
    const sinInformeOTC = serieDe({tipo:r.ot?.tipo,ot:r.ot?.folio}) === 'OTC' && !(r.informes || []).length;
    const detalle = sinInformeOTC ? {...r,
      advertencias:(r.advertencias || []).filter(w => w !== 'La OT no tiene ningún informe relacionado.'),
      linea_tiempo:(r.linea_tiempo || []).map(p => p.clave === 'informe' ? {...p,titulo:'Informe no aplica',extra:'Serie OTC',estado:'no-aplica'} : p)
    } : r;
    pintarDetalle(detalle);
    registrarConsulta(ot);
    if(destino === 'documentos') el('documentos-servicio')?.focus();
    else if(destino === 'cambios') el('ultimo-cambio')?.focus();
    else if(destino.startsWith('informe:')) {
      const tab = el('tab-informes');
      if(tab) mostrarTab({currentTarget:tab},'informes');
      const folio = destino.slice(8);
      const informe = [...document.querySelectorAll('[data-report-folio]')].find(n => n.dataset.reportFolio === folio);
      if(informe) informe.focus();
      else tab?.focus();
    }
  } catch(e) {
    if(version !== generacionDetalle) return;
    el('detalle').innerHTML = '<div class="empty" role="alert">' + esc(e.message) + '</div>';
  }
}
function limpiarFiltros() {
  ['buscar','estado','servicio','personal'].forEach(id => el(id).value = '');
  serieActiva = '';
  document.querySelectorAll('input[name="serie"]').forEach(r => r.checked = r.value === '');
  el('preferences-note').textContent = '';
  guardarPreferencias();
  if(autorizado && ultimaActualizacion) render();
}
document.querySelectorAll('input[name="serie"]').forEach(r => r.addEventListener('change', () => {
  serieActiva = r.value;
  el('preferences-note').textContent = '';
  guardarPreferencias();
  if(autorizado && ultimaActualizacion) render();
}));
['buscar','estado','servicio','personal'].forEach(id => el(id).addEventListener(id === 'buscar' ? 'input' : 'change', () => {
  el('preferences-note').textContent = '';
  guardarPreferencias();
  if(autorizado && ultimaActualizacion) render();
}));
el('clear').addEventListener('click',limpiarFiltros);
el('results').addEventListener('click', e => {
  if(e.target.closest('[data-clear]')) {limpiarFiltros(); return;}
  const b = e.target.closest('button[data-ot]');
  if(b && autorizado) abrir(b.dataset.ot,true,b.hasAttribute('data-documents') ? 'documentos' : b.hasAttribute('data-changes') ? 'cambios' : b.hasAttribute('data-report') ? 'informe:' + b.dataset.report : '');
});
el('recientes').addEventListener('click',e => {
  if(e.target.closest('#clear-recent')) {recientes = []; guardarPreferencias(); pintarRecientes(); return;}
  const b = e.target.closest('button[data-ot]');
  if(b && autorizado) abrir(b.dataset.ot);
});
el('close-detail').addEventListener('click',cerrarDetalle);
el('detail-shell').addEventListener('cancel',e => {e.preventDefault(); cerrarDetalle();});
el('btnRefresh').addEventListener('click',cargarResumen);
SEAAuth.init((_token,usuario) => { prepararPreferencias(usuario); autorizado = true; cargarResumen(); },{gasUrl:API_URL});
