'use strict';
// Conexión existente. Únicamente se permiten las dos acciones de consulta.
const API_URL = 'https://script.google.com/macros/s/AKfycbw9A7lgw0zI4f3V7YLVnuplLoShTS-5Qfprtw-E-iKTTMNd-F9Sh6_u3y2d3Xg0o3s/exec';
let RESUMEN = [], HUERFANOS = [], seleccion = null, ultimaActualizacion = null;
let vista = 'resumen', autorizado = false, cargando = false, generacionDetalle = 0;
const el = id => document.getElementById(id);
const titles = {
  resumen: ['Resumen general', 'De la orden de trabajo al informe. Toda la ruta, en una sola consulta.'],
  ordenes: ['Órdenes de trabajo', 'Encuentra una orden y consulta su expediente, informes e historia.'],
  informes: ['Informes relacionados', 'Folios asociados a las OT. Abre una orden para consultar el expediente y su estado.'],
  alertas: ['Alertas de trazabilidad', 'Datos faltantes y relaciones por revisar en los módulos de origen.']
};
const estadoDe = o => o.estado_traz || o.estatus_ot || 'SIN ESTADO';
const cancelada = o => estadoDe(o).toUpperCase() === 'CANCELADO';
const entregada = o => o.entregado === true;
const abierta = o => !entregada(o) && !cancelada(o);
const alertasDe = o => [!o.tiene_carpeta && 'Sin carpeta de OT', !Number(o.num_informes) && 'Sin informe relacionado'].filter(Boolean);
const normalizar = s => String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();

function filtrados() {
  const q = normalizar(el('buscar').value.trim());
  return RESUMEN.filter(o =>
    (!q || normalizar([o.ot,o.cliente,o.sucursal,o.nom,o.personal,estadoDe(o),...(o.folios_informe || [])].join(' ')).includes(q)) &&
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
    banner('');
    render();
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
function tarjeta(numero, titulo, nota, tono = '') {
  return `<div class="metric ${tono}"><div>${esc(titulo)}</div><strong>${esc(numero)}</strong><small>${esc(nota)}</small></div>`;
}
function resumen(items) {
  const estados = new Map();
  items.forEach(o => estados.set(estadoDe(o), (estados.get(estadoDe(o)) || 0) + 1));
  const dist = [...estados].sort((a,b) => b[1]-a[1]);
  return `<div class="metrics">${tarjeta(items.length,'Órdenes consultadas',RESUMEN.length + ' OT en la fuente')}${tarjeta(items.filter(abierta).length,'Órdenes abiertas','Excluye entregadas y canceladas')}${tarjeta(items.filter(entregada).length,'Entregadas / finalizadas','Según estado informado por TRAZ','green')}${tarjeta(items.filter(o => alertasDe(o).length).length,'OT con datos faltantes','Carpeta de OT o informe relacionado','amber')}</div>
    <div class="overview-grid"><section class="card distribution"><div class="card-head"><h3>Distribución por estado</h3><span>Según filtros</span></div><div class="bars">${dist.length ? dist.map(([s,n]) => `<div class="bar-row"><div><span>${esc(s)}</span><strong>${n}</strong></div><div class="bar-track"><span style="width:${Math.round(n/items.length*100)}%"></span></div></div>`).join('') : '<p>Sin registros para esta selección.</p>'}</div></section>
    <section class="consult-note"><div class="eyebrow">CONTINUIDAD DOCUMENTAL</div><h2>Una orden.<br>Su historia completa.</h2><p>Selecciona una OT para ver la ruta del servicio, el expediente y los cambios registrados en AUDITORIA.</p><div class="mini-route"><span>Orden</span><span>→</span><span>Expediente</span><span>→</span><span>Informe</span></div><small>Una carpeta vinculada no confirma que su contenido esté completo.</small></section></div>`;
}
function tabla(items, soloAlertas = false) {
  if(!items.length) return '<div class="empty">No hay órdenes que coincidan con esta consulta.</div>';
  return `<div class="table-scroll"><table class="orders"><thead><tr><th>Orden / cliente</th><th>Servicio / sucursal</th><th>Estado TRAZ</th><th>Personal / visita OT</th><th>${soloAlertas ? 'Datos por revisar' : 'Documentación'}</th><th>Consulta</th></tr></thead><tbody>${items.map(o => `<tr><td><strong>${esc(o.ot)}</strong><span class="cell-sub">${esc(o.cliente || 'Sin cliente registrado')}</span></td><td>${esc(o.nom || 'Sin servicio')}<span class="cell-sub">${esc(o.sucursal || 'Sin sucursal registrada')}</span></td><td><span class="state-mini ${estadoClase(estadoDe(o))}">${esc(estadoDe(o))}</span></td><td>${esc(o.personal || 'Sin personal registrado')}<span class="cell-sub">${esc(o.fecha_visita || 'Sin fecha registrada')}</span></td><td>${soloAlertas ? alertasDe(o).map(a => `<span class="alert-chip">${esc(a)}</span>`).join('') : `<span class="cell-sub">${o.tiene_carpeta ? 'Carpeta de OT vinculada' : 'Sin carpeta de OT'}</span><span class="cell-sub">${Number(o.num_informes) || 0} informe(s)</span>`}</td><td><button class="open-ot" data-ot="${esc(o.ot)}">Ver ruta ↗</button></td></tr>`).join('')}</tbody></table></div>`;
}
function informes(items) {
  const rows = items.flatMap(o => (o.folios_informe || []).map(f => ({o, f})));
  const sinFolio = items.reduce((s,o) => s + Math.max(0,(Number(o.num_informes)||0) - (o.folios_informe || []).length),0);
  return `<p class="section-note">El resumen expone los folios y el estado de la OT. El detalle de cada informe se consulta dentro de su orden.${sinFolio ? ' Hay ' + sinFolio + ' registro(s) de informe sin folio dentro de las OT filtradas.' : ''}</p>` +
    (rows.length ? `<div class="table-scroll"><table class="orders"><thead><tr><th>Folio de informe</th><th>Orden</th><th>Cliente</th><th>Servicio</th><th>Consulta</th></tr></thead><tbody>${rows.map(({o,f}) => `<tr><td><strong>${esc(f)}</strong></td><td>${esc(o.ot)}</td><td>${esc(o.cliente)}</td><td>${esc(o.nom)}</td><td><button class="open-ot" data-ot="${esc(o.ot)}">Ver detalle ↗</button></td></tr>`).join('')}</tbody></table></div>` : '<div class="empty">No hay folios de informe para esta selección.</div>');
}
function huerfanos() {
  return `<section class="orphan-section"><h3>Informes con OT no encontrada <span class="count-pill">${HUERFANOS.length}</span></h3><p class="section-note">Consulta global, independiente de los filtros. Se muestran las relaciones que reporta el backend.</p>${HUERFANOS.length ? `<div class="table-scroll"><table class="orders"><thead><tr><th>Informe</th><th>OT referenciada</th><th>Relación</th></tr></thead><tbody>${HUERFANOS.map(h => `<tr><td>${esc(h.num_informe || 'Sin folio')}</td><td>${esc(h.ot)}</td><td>OT no encontrada en ORDENES_TRABAJO</td></tr>`).join('')}</tbody></table></div>` : '<div class="good">No se reportan informes con OT inexistente en esta consulta.</div>'}</section>`;
}
function render() {
  const items = filtrados();
  const afectados = items.filter(o => alertasDe(o).length);
  el('overview').innerHTML = vista === 'resumen' ? resumen(items) : '';
  let contenido = vista === 'informes' ? informes(items) : tabla(vista === 'alertas' ? afectados : items, vista === 'alertas');
  if(vista === 'alertas') contenido = '<p class="section-note">Estos indicadores señalan datos faltantes; no determinan incumplimientos. Abre la OT para consultar las advertencias completas de su expediente.</p>' + contenido + huerfanos();
  el('results').innerHTML = `<section class="card result-card"><div class="card-head"><h3>${vista === 'informes' ? 'Índice de informes' : vista === 'alertas' ? 'Órdenes por revisar' : 'Órdenes de trabajo'}</h3><span>${vista === 'alertas' ? afectados.length : items.length} OT · ${RESUMEN.length} en la fuente</span></div>${contenido}</section>`;
}
function cerrarDetalle() {
  seleccion = null;
  generacionDetalle++;
  el('detail-shell').hidden = true;
  el('detalle').innerHTML = '';
  el('results').hidden = false;
  el('overview').hidden = false;
}
async function abrir(ot, enfocar = true) {
  seleccion = ot;
  const version = ++generacionDetalle;
  el('results').hidden = true;
  el('overview').hidden = true;
  el('detail-shell').hidden = false;
  el('detalle').innerHTML = '<div class="loading" role="status">Consultando la ruta del servicio…</div>';
  if(enfocar) el('detalle').focus();
  try {
    const r = await consultar('trazDetalle',{ot});
    if(version !== generacionDetalle || ot !== seleccion) return;
    if(!r.success) throw new Error(r.message || r.error || 'No se pudo consultar la orden.');
    pintarDetalle(r);
    el('detalle').insertAdjacentHTML('beforeend','<p class="section-note">Los hitos de ejecución provienen de las fechas registradas en SEA; no constituyen una verificación independiente de campo. Revisión y emisión solo se muestran si están registradas en AUDITORIA.</p>');
  } catch(e) {
    if(version !== generacionDetalle) return;
    el('detalle').innerHTML = '<div class="empty" role="alert">' + esc(e.message) + '</div>';
  }
}
document.querySelectorAll('[data-view]').forEach(b => b.addEventListener('click', () => {
  vista = b.dataset.view;
  cerrarDetalle();
  document.querySelectorAll('[data-view]').forEach(n => {n.classList.toggle('selected',n === b); n.setAttribute('aria-current',n === b ? 'page' : 'false');});
  el('view-title').textContent = titles[vista][0];
  el('view-subtitle').textContent = titles[vista][1];
  if(autorizado && ultimaActualizacion) render();
}));
['buscar','estado','servicio','personal'].forEach(id => el(id).addEventListener(id === 'buscar' ? 'input' : 'change', () => {
  cerrarDetalle();
  if(autorizado && ultimaActualizacion) render();
}));
el('clear').addEventListener('click', () => {
  ['buscar','estado','servicio','personal'].forEach(id => el(id).value = '');
  cerrarDetalle();
  if(autorizado && ultimaActualizacion) render();
});
el('results').addEventListener('click', e => {
  const b = e.target.closest('button[data-ot]');
  if(b && autorizado) abrir(b.dataset.ot);
});
el('close-detail').addEventListener('click',cerrarDetalle);
el('btnRefresh').addEventListener('click',cargarResumen);
SEAAuth.init(() => { autorizado = true; cargarResumen(); },{gasUrl:API_URL});
