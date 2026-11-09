// Renderizadores aislados del panel. No modifica el frontend original.
  function esc(s){
    return String(s == null ? '' : s).replace(/[&<>"]/g, c => ({
      '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'
    }[c]));
  }
  function safeUrl(u){
    return (typeof u === 'string' && /^https:\/\//i.test(u)) ? u : '';
  }
  function estadoClase(s){
    s = String(s || '').toUpperCase();
    if(s === 'CANCELADO') return 'bad';
    if(s === 'EN PAUSA' || s === 'PROGRAMADO' || s.indexOf('PEND') >= 0) return 'warn';
    if(['ENTREGADO','FINALIZADO','PARA REVISION','PARA REVISIÓN','PARA IMPRESION','PARA IMPRESIÓN','EN PROCESO','SERVICIO EJECUTADO'].includes(s)) return 'ok';
    return '';
  }
  function formatNow(){
    try{
      return new Intl.DateTimeFormat('es-MX',{hour:'2-digit',minute:'2-digit'}).format(new Date());
    }catch(_){ return ''; }
  }
  function fechaMs(v){
    const s = String(v || '').trim();
    if(!s) return null;
    let m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
    if(m) return new Date(+m[3],+m[2]-1,+m[1],+(m[4]||0),+(m[5]||0),+(m[6]||0)).getTime();
    m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T\s](\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
    if(m) return new Date(+m[1],+m[2]-1,+m[3],+(m[4]||0),+(m[5]||0),+(m[6]||0)).getTime();
    const d = Date.parse(s);
    return Number.isFinite(d) ? d : null;
  }

  async function api(params){
    const qs = new URLSearchParams(params);
    const resp = await SEAAuth.wrapFetch(API_URL + '?' + qs.toString(), {cache:'no-store'});
    if(!resp.ok) throw new Error('HTTP ' + resp.status);
    const data = await resp.json();
    if(data && data.error === 'AUTH_REQUIRED'){
      SEAAuth.logout();
      throw new Error(data.message || 'Acceso no autorizado.');
    }
    return data;
  }

  function evento(fecha,titulo,detalle,fuente,ordenExtra){
    const ms = fechaMs(fecha);
    return {
      fecha: fecha || '',
      titulo: titulo || '',
      detalle: detalle || '',
      fuente: fuente || 'DATO',
      ms: ms,
      ordenExtra: ordenExtra || 0
    };
  }

  function construirHistoria(r){
    const ot = r.ot || {};
    const informes = r.informes || [];
    const bitacora = r.bitacora || [];
    const eventos = [];

    if(ot.fecha_alta){
      eventos.push(evento(ot.fecha_alta,'Orden de trabajo registrada',
        [ot.folio,ot.nom].filter(Boolean).join(' · '),'OT',1));
    }

    informes.forEach((inf,idx) => {
      if(inf.timestamp){
        eventos.push(evento(inf.timestamp,'Expediente / informe registrado',
          [inf.num_informe,inf.responsable ? 'Responsable: ' + inf.responsable : ''].filter(Boolean).join(' · '),'INFORMES',10+idx));
      }
    });

    const ejec = (r.linea_tiempo || []).find(x => x.clave === 'ejecucion');
    if(ejec && ejec.fecha){
      eventos.push(evento(ejec.fecha,
        ejec.estado === 'ok' ? 'Servicio ejecutado' : 'Servicio programado',
        ejec.estado === 'ok' ? 'Fecha de servicio reconocida por TRAZ.' : 'La fecha registrada aún no se considera ejecutada.',
        'TRAZ',30));
    }

    bitacora.forEach((b,idx) => {
      let det = [b.accion,b.campo].filter(Boolean).join(' · ');
      if(b.valor_anterior || b.valor_nuevo){
        det += (det ? ' · ' : '') + (b.valor_anterior || '—') + ' → ' + (b.valor_nuevo || '—');
      }
      eventos.push(evento(b.timestamp,'Cambio registrado en AUDITORIA',det,'AUDITORIA',40+idx));
    });

    if(ot.fecha_real_entrega){
      eventos.push(evento(ot.fecha_real_entrega,'Entrega registrada',
        'La OT cuenta con fecha real de entrega.','OT',90));
    }

    eventos.sort((a,b) => {
      if(a.ms !== null && b.ms !== null && a.ms !== b.ms) return a.ms - b.ms;
      if(a.ms !== null && b.ms === null) return -1;
      if(a.ms === null && b.ms !== null) return 1;
      return a.ordenExtra - b.ordenExtra;
    });
    return eventos;
  }

  function pintarHistoria(r){
    const eventos = construirHistoria(r);
    if(!eventos.length) return '<div class="no-history">No hay hitos suficientes para reconstruir la historia de esta OT.</div>';
    return '<div class="history">' + eventos.map(e => {
      const audit = e.fuente === 'AUDITORIA';
      return '<div class="event ' + (audit ? 'audit' : '') + '">' +
        '<div class="event-time">' + esc(e.fecha || '—') + '</div>' +
        '<div class="rail"><span class="mark"></span></div>' +
        '<div class="event-body"><div class="event-title"><strong>' + esc(e.titulo) + '</strong>' +
        '<span class="source ' + (audit ? 'audit' : '') + '">' + esc(e.fuente) + '</span></div>' +
        (e.detalle ? '<p class="' + (audit ? 'change' : '') + '">' + esc(e.detalle) + '</p>' : '') +
        '</div></div>';
    }).join('') + '</div>';
  }

  function pintarRuta(r){
    return '<div class="steps">' + (r.linea_tiempo || []).map(p => {
      const ok = p.estado === 'ok';
      const noAplica = p.estado === 'no-aplica';
      const fecha = p.fecha || p.extra || '';
      let titulo = p.titulo || '';
      if(p.clave === 'ejecucion' && !ok){
        titulo = fecha ? 'Servicio programado' : 'Servicio pendiente';
      }
      return '<div class="step ' + (noAplica ? 'not-applicable' : ok ? 'ok' : 'pending') + '">' +
        '<div class="dot">' + (noAplica ? '—' : ok ? '✓' : '•') + '</div>' +
        '<div class="step-title">' + esc(titulo) + '</div>' +
        '<div class="step-date">' + esc(fecha) + '</div>' +
      '</div>';
    }).join('') + '</div>';
  }

  function pintarAcciones(r){
    const ot = r.ot || {};
    const exp = safeUrl(r.expediente);
    const cliente = safeUrl(ot.link_drive);
    const informes = r.informes || [];
    let fechaServicio = '';
    for(let i = informes.length - 1; i >= 0; i--){
      if(String(informes[i].fecha_servicio || '').trim()){
        fechaServicio = informes[i].fecha_servicio;
        break;
      }
    }
    if(!fechaServicio) fechaServicio = ot.fecha_visita || '';

    let html = '<div class="actions">';
    html += exp
      ? '<a class="action-primary" href="' + esc(exp) + '" target="_blank" rel="noopener"><span>Abrir expediente</span><span>↗</span></a>'
      : '<div class="action-secondary disabled"><span>Sin expediente disponible</span><span>—</span></div>';

    if(cliente){
      html += '<a class="action-secondary" href="' + esc(cliente) + '" target="_blank" rel="noopener"><span>Carpeta de la OT</span><span>↗</span></a>';
    }

    informes.forEach(inf => {
      const u = safeUrl(inf.link_drive);
      if(!u) return;
      html += '<a class="action-secondary" href="' + esc(u) + '" target="_blank" rel="noopener"><span>' +
        esc(inf.num_informe || 'Informe') +
        '<div class="action-meta">' + esc(inf.estatus || 'Sin estado') + '</div></span><span>↗</span></a>';
    });
    html += '</div>';

    html += '<div class="facts">' +
      '<div class="fact"><span>Responsable</span><span>' + esc(ot.personal || '—') + '</span></div>' +
      '<div class="fact"><span>Fecha servicio</span><span>' + esc(fechaServicio || '—') + '</span></div>' +
      '<div class="fact"><span>RFC</span><span>' + esc(ot.rfc || '—') + '</span></div>' +
      '<div class="fact"><span>Serie</span><span>' + esc(ot.tipo || '—') + '</span></div>' +
    '</div>';
    return html;
  }

  function pintarAdvertencias(r){
    const w = r.advertencias || [];
    if(!w.length) return '<div class="good">Sin inconsistencias detectadas con los datos disponibles.</div>';
    return w.map(x => '<div class="warning">' + esc(x) + '</div>').join('');
  }

  function pintarInformes(informes){
    if(!informes.length) return '<div class="audit-empty">No hay informes relacionados con esta OT.</div>';
    return '<div class="report-list">' + informes.map(inf => {
      const u = safeUrl(inf.link_drive);
      return '<div class="report" tabindex="-1" data-report-folio="' + esc(inf.num_informe || '') + '"><div><strong>' + esc(inf.num_informe || '(sin folio)') + '</strong>' +
        '<small>' + esc(inf.estatus || 'Sin estado') + ' · Servicio: ' + esc(inf.fecha_servicio || '—') +
        ' · Responsable: ' + esc(inf.responsable || '—') + '</small></div>' +
        (u ? '<a href="' + esc(u) + '" target="_blank" rel="noopener">Abrir expediente ↗</a>' : '') +
      '</div>';
    }).join('') + '</div>';
  }

  function pintarAuditoria(bit){
    if(!bit.length) return '<div class="audit-empty">Sin eventos registrados en AUDITORIA para esta OT.</div>';
    return '<table><thead><tr><th>Fecha</th><th>Usuario</th><th>Acción</th><th>Campo</th><th>Antes → Después</th></tr></thead><tbody>' +
      bit.map(b => '<tr><td>' + esc(b.timestamp) + '</td><td>' + esc(b.usuario) + '</td><td>' + esc(b.accion) +
        '</td><td>' + esc(b.campo) + '</td><td>' + esc(b.valor_anterior || '—') + ' → ' + esc(b.valor_nuevo || '—') + '</td></tr>').join('') +
      '</tbody></table>';
  }

  function ultimoCambioRegistrado(bitacora) {
    // Solo AUDITORIA. Fechas de visita, informes y entregas no son cambios auditados.
    return (bitacora || []).map((b,i) => ({...b,ms:fechaMs(b.timestamp),indice:i}))
      .filter(b => b.ms !== null)
      .sort((a,b) => b.ms - a.ms || b.indice - a.indice)[0] || null;
  }
  function pintarUltimoCambio(bitacora) {
    const ultimo = ultimoCambioRegistrado(bitacora);
    const sinFecha = (bitacora || []).filter(b => fechaMs(b.timestamp) === null).length;
    let cuerpo = !bitacora.length ? '<p>No hay cambios registrados en AUDITORIA para esta orden.</p>' :
      !ultimo ? '<p>Los eventos registrados no tienen una fecha reconocible. Revisa la bitácora.</p>' :
      '<div class="audit-when">' + esc(ultimo.timestamp) + '</div><strong class="audit-action">' + esc([ultimo.accion,ultimo.campo].filter(Boolean).join(' · ') || 'Cambio registrado') + '</strong>' +
      '<p class="audit-change">' + esc(ultimo.valor_anterior || '—') + ' → ' + esc(ultimo.valor_nuevo || '—') + '</p><p>Registrado por ' + esc(ultimo.usuario || 'Usuario no registrado') + '</p>';
    if(sinFecha && ultimo) cuerpo += '<p>' + sinFecha + ' evento(s) sin fecha reconocible en la bitácora.</p>';
    return '<section class="card latest-change" id="ultimo-cambio" tabindex="-1"><div class="card-head"><h3>Último cambio con fecha</h3><span>AUDITORIA</span></div><div class="audit-brief">' + cuerpo +
      (bitacora.length ? '<button type="button" class="quiet" onclick="abrirBitacora()">Ver bitácora completa</button>' : '') + '</div></section>';
  }
  function abrirBitacora() {
    const tab = document.getElementById('tab-auditoria');
    if(tab) {mostrarTab({currentTarget:tab},'auditoria'); document.getElementById('panel-auditoria').focus();}
  }

  function pintarDetalle(r){
    const ot = r.ot || {};
    const estado = r.estado_traz || ot.estatus || 'SIN ESTADO';
    const informes = r.informes || [];
    const bit = r.bitacora || [];

    document.getElementById('detalle').innerHTML =
      '<div class="kicker-row"><span class="kicker">Trazabilidad del servicio</span><span class="updated">' +
        ('Consultado ' + esc(new Intl.DateTimeFormat('es-MX',{timeZone:'America/Mexico_City',timeStyle:'short'}).format(new Date())) + ' · CDMX') + '</span></div>' +

      '<section class="hero"><div><h2>' + esc(ot.folio || '') + '</h2>' +
        '<div class="hero-client">' + esc(ot.cliente || '—') + '</div>' +
        '<div class="hero-sub">' + esc(ot.nom || '—') + (ot.sucursal ? ' · ' + esc(ot.sucursal) : '') + '</div></div>' +
        '<div class="status-box"><span class="status-pill ' + estadoClase(estado) + '">' + esc(estado) + '</span>' +
        '<div class="sea-status">Estatus SEA: ' + esc(ot.estatus || '—') + '</div></div></section>' +

      '<div class="quick-detail"><section class="card" id="documentos-servicio" tabindex="-1"><div class="card-head"><h3>Documentos del servicio</h3><span>Expediente e informes</span></div>' +
        pintarAcciones(r) + '</section>' + pintarUltimoCambio(bit) + '</div>' +
      '<div class="detail-warnings">' + pintarAdvertencias(r) + '</div>' +
      '<section class="route"><div class="route-head"><strong>Ruta del servicio</strong>' +
        '<span>Según los registros disponibles</span></div>' + pintarRuta(r) + '</section>' +
      '<details class="history-disclosure"><summary>Historia completa del servicio</summary><section class="card">' + pintarHistoria(r) + '</section></details>' +

      '<div class="tabs">' +
        '<button class="tab active" onclick="mostrarTab(event,\'datos\')">Datos del servicio</button>' +
        '<button class="tab" id="tab-informes" onclick="mostrarTab(event,\'informes\')">Informes</button>' +
        '<button class="tab" id="tab-auditoria" onclick="mostrarTab(event,\'auditoria\')">AUDITORIA</button>' +
      '</div>' +

      '<div class="panel active" id="panel-datos"><div class="data-card">' +
        '<div class="data-item"><small>OT</small><strong>' + esc(ot.folio || '—') + '</strong></div>' +
        '<div class="data-item"><small>Tipo</small><strong>' + esc(ot.tipo || '—') + '</strong></div>' +
        '<div class="data-item"><small>Servicio</small><strong>' + esc(ot.nom || '—') + '</strong></div>' +
        '<div class="data-item"><small>Cliente</small><strong>' + esc(ot.cliente || '—') + '</strong></div>' +
        '<div class="data-item"><small>Sucursal</small><strong>' + esc(ot.sucursal || '—') + '</strong></div>' +
        '<div class="data-item"><small>Personal</small><strong>' + esc(ot.personal || '—') + '</strong></div>' +
        '<div class="data-item"><small>Fecha visita OT</small><strong>' + esc(ot.fecha_visita || '—') + '</strong></div>' +
        '<div class="data-item"><small>Entrega límite</small><strong>' + esc(ot.fecha_entrega_limite || '—') + '</strong></div>' +
        '<div class="data-item"><small>Entrega real</small><strong>' + esc(ot.fecha_real_entrega || '—') + '</strong></div>' +
        '<div class="data-item"><small>RFC</small><strong>' + esc(ot.rfc || '—') + '</strong></div>' +
        '<div class="data-item"><small>Estado TRAZ</small><strong>' + esc(estado) + '</strong></div>' +
        '<div class="data-item"><small>Estatus SEA</small><strong>' + esc(ot.estatus || '—') + '</strong></div>' +
      '</div></div>' +

      '<div class="panel" id="panel-informes">' + pintarInformes(informes) + '</div>' +
      '<div class="panel" id="panel-auditoria" tabindex="-1">' + pintarAuditoria(bit) + '</div>';
  }

  function mostrarTab(ev,nombre){
    document.querySelectorAll('.tab').forEach(x => x.classList.remove('active'));
    document.querySelectorAll('.panel').forEach(x => x.classList.remove('active'));
    ev.currentTarget.classList.add('active');
    const p = document.getElementById('panel-' + nombre);
    if(p) p.classList.add('active');
  }
