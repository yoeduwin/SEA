from pathlib import Path
import re, subprocess

paic_path = Path('PAIC.html')
backend_path = Path('BACKEND_FIXES.gs')
paic = paic_path.read_text(encoding='utf-8')
backend = backend_path.read_text(encoding='utf-8')
base_paic = subprocess.check_output(['git','show','origin/claude/zen-mayer-ocbfln:PAIC.html'], text=True)


def replace_once(text, old, new, label):
    count = text.count(old)
    if count != 1:
        raise SystemExit(f'{label}: se esperaba 1 coincidencia, encontradas {count}')
    return text.replace(old, new, 1)


def function_block(text, name, next_name=None):
    start = text.find(f'function {name}(')
    if start < 0:
        raise SystemExit(f'No se encontró función {name}')
    if next_name:
        end = text.find(f'\nfunction {next_name}(', start)
        if end < 0:
            raise SystemExit(f'No se encontró función siguiente {next_name}')
        return text[start:end]
    m = re.search(r'\nfunction [A-Za-z0-9_]+\(', text[start+1:])
    return text[start:] if not m else text[start:start+1+m.start()]


protected_profile_before = function_block(backend, 'generarPerfilSheet', 'enviarNotificacionRobusta')

# ---- PAIC: alta inicial por asesor/intermediario, no mantenimiento ----
paic = replace_once(
    paic,
    'Esta plataforma permite registrar o actualizar los datos maestros de cada cliente y sucursal que trabaja con Ejecutiva Ambiental.',
    'Esta plataforma permite a asesores, intermediarios y consultorías registrar por primera vez los datos del cliente final y de la sucursal con la que Ejecutiva Ambiental trabajará.',
    'welcome copy')
paic = replace_once(
    paic,
    '<strong>1 registro = 1 cliente / sucursal.</strong> Las órdenes de trabajo y los servicios se generan posteriormente por Ejecutiva Ambiental. No es necesario volver a registrar al cliente cada vez que se realice un servicio si sus datos siguen vigentes.',
    '<strong>1 registro = 1 alta inicial de cliente / sucursal.</strong> PAIC se utiliza cuando existe un asesor, intermediario o consultoría entre el cliente final y Ejecutiva Ambiental. Si la combinación RFC + sucursal ya está registrada, no debe registrarse nuevamente. Las órdenes de trabajo se generan posteriormente y por separado en SEAOT.',
    'welcome rule')
paic = replace_once(paic, '¿Cliente ya registrado?', '¿RFC ya registrado?', 'header button')
paic = replace_once(
    paic,
    'Busque por RFC para actualizar una sucursal existente o registrar una nueva sucursal del mismo cliente.',
    'Busque por RFC para consultar las sucursales ya registradas o dar de alta una nueva sucursal del mismo cliente. Las sucursales existentes no se actualizan desde PAIC.',
    'returning modal copy')
paic = replace_once(
    paic,
    '<div class="banner-text"><strong>Cliente identificado</strong><br>Hemos prellenado los datos de esta sucursal. Revise y actualice solo lo que haya cambiado.</div>',
    '<div class="banner-text"><strong>RFC identificado</strong><br>Capture los datos de la nueva sucursal. Las sucursales ya registradas no se modifican desde PAIC.</div>',
    'returning banner')
paic = replace_once(
    paic,
    '<div class="warning-box"><strong>Importante:</strong> Este envío registra o actualiza al cliente/sucursal. No genera una orden de trabajo ni un servicio.</div>',
    '<div class="warning-box"><strong>Importante:</strong> Este envío realiza el alta inicial del cliente/sucursal a través del asesor o intermediario. No genera una orden de trabajo. La OT se gestiona por separado en SEAOT.</div>',
    'confirmation copy')
paic = replace_once(
    paic,
    '<div class="message success-message" id="successMsg"><strong>Información registrada correctamente.</strong><br>El perfil del cliente/sucursal quedó disponible para futuras órdenes de trabajo.</div>',
    '<div class="message success-message" id="successMsg"><strong>Cliente/sucursal registrado correctamente.</strong><br>El alta inicial quedó registrada. La orden de trabajo se gestiona por separado en SEAOT.</div>',
    'success copy')

# ---- Restaurar NOM-020 y PIPC: indicación/documentación del alta ----
def extract_section(source, marker):
    pos = source.find(marker)
    if pos < 0:
        raise SystemExit(f'No se encontró marker: {marker}')
    start = source.rfind('        <div class="section">', 0, pos)
    if start < 0:
        raise SystemExit(f'No se encontró inicio de sección: {marker}')
    nxt = source.find('\n        <div class="section">', pos + len(marker))
    if nxt < 0:
        nxt = source.find('\n        <div class="submit-section">', pos + len(marker))
    if nxt < 0:
        raise SystemExit(f'No se encontró fin de sección: {marker}')
    return source[start:nxt]

nom020 = extract_section(base_paic, 'INSPECCI&Oacute;N &mdash; NOM-020-STPS')
pipc = extract_section(base_paic, 'Programa Interno de Protección Civil (PIPC)')
nom020 = nom020.replace(
    '¿El servicio es para evaluación de NOM-020-STPS-2011 (Recipientes sujetos a presión, recipientes criogénicos y generadores de vapor o calderas)?',
    '¿Este registro requiere incluir la indicación y documentación para NOM-020-STPS-2011 (recipientes sujetos a presión, recipientes criogénicos y generadores de vapor o calderas)?')
nom020 = nom020.replace('No, no aplica para nosotros', 'No, no se requiere en este registro')
pipc = pipc.replace(
    '¿El servicio es para realizar su Programa Interno de Protección Civil?',
    '¿Este registro requiere incluir la indicación y documentación para Programa Interno de Protección Civil?')
pipc = pipc.replace('Sí, requiero PIPC', 'Sí, incluir PIPC en este registro')
pipc = pipc.replace('No, no aplica', 'No, no se requiere en este registro')

if 'name="aplica_nom020"' in paic or 'name="requiere_pipc"' in paic:
    raise SystemExit('PAIC ya contiene NOM-020/PIPC; abortando para evitar duplicados')
anchor = '        <div class="submit-section">'
if paic.count(anchor) != 1:
    raise SystemExit('submit-section no único')
paic = paic.replace(anchor, nom020 + '\n\n' + pipc + '\n\n' + anchor, 1)

# CSS mínimo para secciones condicionales restauradas.
css_anchor = '    .returning-banner { display:none;'
if css_anchor not in paic:
    raise SystemExit('No se encontró anchor CSS')
extra_css = '''    .radio-group { display:flex; gap:20px; margin-top:10px; flex-wrap:wrap; }\n    .radio-option { display:flex; align-items:center; gap:8px; }\n    .radio-option label { margin:0; cursor:pointer; text-transform:none; font-weight:normal; }\n    .conditional-section { display:none; margin-top:20px; padding:20px; background:#fff3cd; border-left:4px solid #ffc107; border-radius:5px; }\n    .conditional-section.active { display:block; }\n    .conditional-header { font-weight:bold; color:#856404; margin-bottom:15px; font-size:12px; }\n'''
paic = paic.replace(css_anchor, extra_css + css_anchor, 1)

# ---- JS condicional y reset ----
js_anchor = "    const rfcInput = document.querySelector('input[name=\"rfc\"]');"
if paic.count(js_anchor) != 1:
    raise SystemExit('No se encontró anchor RFC')
conditional_js = r'''    function resetConditionalRegistrationSections() {
      document.querySelectorAll('input[name="aplica_nom020"], input[name="requiere_pipc"]').forEach(r => { r.checked=false; });
      document.getElementById('documentacionLegalSection')?.classList.remove('active');
      document.getElementById('documentacionPIPCSection')?.classList.remove('active');
      document.querySelectorAll('[data-conditional-required="nom020"], [data-conditional-required="pipc"]').forEach(f => { f.required=false; });
      const sinCal=document.getElementById('check_sin_calibracion'); if (sinCal) sinCal.checked=false;
      const calib=document.querySelector('[data-file-input="calibracion_valvula"]'); if (calib) { calib.style.display=''; calib.required=false; }
      const ast=document.getElementById('asterisco_calib'); if (ast) ast.style.display='';
      const help=document.getElementById('help_text_calib'); if (help) help.textContent='Adjunte el archivo PDF de su última calibración.';
    }

    function resetUploadState() {
      document.querySelectorAll('.file-on-record').forEach(i => i.remove());
      document.querySelectorAll('[data-file-input]').forEach(input => { input.value=''; input.style.display=''; input.required=input.name==='planos'; });
      document.querySelectorAll('.file-preview').forEach(p => p.classList.remove('active'));
      document.querySelectorAll('.file-upload-group').forEach(g => g.classList.remove('has-file'));
      resetConditionalRegistrationSections();
    }

    document.querySelectorAll('input[name="aplica_nom020"]').forEach(r => r.addEventListener('change', function() {
      const on=this.value==='si';
      document.getElementById('documentacionLegalSection')?.classList.toggle('active',on);
      document.querySelectorAll('[data-conditional-required="nom020"]').forEach(f => { f.required=on; });
      if (!on) document.querySelectorAll('[data-conditional-required="nom020"]').forEach(f => {
        if (f.type==='file' && f.dataset.fileInput) clearFileUI(f.dataset.fileInput); else if (f.type!=='radio' && f.type!=='checkbox') f.value='';
      });
    }));

    document.querySelectorAll('input[name="requiere_pipc"]').forEach(r => r.addEventListener('change', function() {
      const on=this.value==='si';
      document.getElementById('documentacionPIPCSection')?.classList.toggle('active',on);
      document.querySelectorAll('[data-conditional-required="pipc"]').forEach(f => { f.required=on; });
      if (!on) document.querySelectorAll('[data-conditional-required="pipc"]').forEach(f => {
        if (f.type==='file' && f.dataset.fileInput) clearFileUI(f.dataset.fileInput); else if (f.type!=='radio' && f.type!=='checkbox') f.value='';
      });
    }));

    document.getElementById('check_sin_calibracion')?.addEventListener('change', function() {
      const calib=document.querySelector('[data-file-input="calibracion_valvula"]');
      const ast=document.getElementById('asterisco_calib');
      const help=document.getElementById('help_text_calib');
      if (this.checked) {
        if (calib) { calib.required=false; calib.style.display='none'; clearFileUI('calibracion_valvula'); }
        if (ast) ast.style.display='none';
        if (help) help.textContent='Ejecutiva Ambiental puede realizar la calibración el día de la dictaminación.';
      } else {
        const nomOn=document.querySelector('input[name="aplica_nom020"]:checked')?.value==='si';
        if (calib) { calib.style.display=''; calib.required=nomOn; }
        if (ast) ast.style.display='';
        if (help) help.textContent='Adjunte el archivo PDF de su última calibración.';
      }
    });

'''
paic = paic.replace(js_anchor, conditional_js + js_anchor, 1)

old_reset = """        document.querySelectorAll('.file-preview').forEach(p => p.classList.remove('active'));
        document.querySelectorAll('.file-upload-group').forEach(g => g.classList.remove('has-file'));
        document.querySelectorAll('.file-on-record').forEach(i => i.remove());
        document.querySelectorAll('[data-file-input]').forEach(input => { input.style.display=''; input.required=input.name==='planos'; });"""
paic = replace_once(paic, old_reset, '        resetUploadState();', 'success reset')

# ---- RFC existente: solo consulta y nueva sucursal ----
pattern = re.compile(r"    async function searchClientByRFC\(\) \{.*?    window\.applyClientData=applyClientData;\n", re.S)
if len(pattern.findall(paic)) != 1:
    raise SystemExit(f'Bloque returning inesperado: {len(pattern.findall(paic))}')
new_returning = r'''    async function searchClientByRFC() {
      const input=document.getElementById('rfcSearch');
      const resultDiv=document.getElementById('rfcSearchResult');
      const btn=document.getElementById('btnSearchRFC');
      const rfc=String(input?.value || '').trim().toUpperCase();
      if (rfc.length < 12) { resultDiv.className='rfc-search-result error'; resultDiv.style.display='block'; resultDiv.textContent='RFC no válido.'; return; }
      btn.disabled=true; btn.textContent='Buscando...';
      try {
        const response=await SEARecaptcha.wrapFetch(`${SCRIPT_URL}?action=buscarClienteRFC&rfc=${encodeURIComponent(rfc)}`);
        const data=await response.json();
        if (data.found && Array.isArray(data.sucursales) && data.sucursales.length) {
          window._returningClientBranches=data.sucursales; window._returningRazonSocial=data.razon_social; window._returningRfc=rfc;
          const branches=data.sucursales.map(s => `<li>${escHtml(s.sucursal || 'Matriz')} — ya registrada</li>`).join('');
          resultDiv.className='rfc-search-result found'; resultDiv.style.display='block';
          resultDiv.innerHTML=`<strong>Cliente encontrado:</strong> ${escHtml(data.razon_social)}<br><br><strong>Sucursales ya registradas:</strong><ul style="margin:8px 0 15px 20px;">${branches}</ul><p style="margin-bottom:12px;">PAIC no modifica sucursales existentes. Si necesita dar de alta otra ubicación del mismo RFC, continúe con una nueva sucursal.</p><button type="button" class="btn-search" style="width:100%;" onclick="applyNewBranchData()">Registrar nueva sucursal</button>`;
        } else {
          window._returningClientBranches=[]; window._returningRazonSocial=''; window._returningRfc=rfc;
          resultDiv.className='rfc-search-result not-found'; resultDiv.style.display='block';
          resultDiv.innerHTML='<strong>RFC no encontrado.</strong><br>Este cliente puede registrarse por primera vez desde el formulario.';
        }
      } catch (_) {
        resultDiv.className='rfc-search-result error'; resultDiv.style.display='block'; resultDiv.textContent='Error de conexión. Intente nuevamente.';
      } finally { btn.disabled=false; btn.textContent='Buscar'; }
    }
    window.searchClientByRFC=searchClientByRFC;

    function setField(name,value) {
      const el=document.querySelector(`[name="${name}"]`); if (el) el.value=value == null ? '' : value;
    }

    function applyNewBranchData() {
      const base=window._returningClientBranches?.[0] || {};
      const razon=base.razon_social || window._returningRazonSocial || '';
      const rfc=base.rfc || window._returningRfc || '';
      ['sucursal','direccion_evaluacion','telefono_empresa','representante_legal','responsable','telefono_responsable','correo_informe','nombre_dirigido','puesto_dirigido','giro','actividad_principal','descripcion_proceso','registro_patronal','capacidad_instalada','capacidad_operacion','dias_turnos_horarios'].forEach(k => setField(k,''));
      resetUploadState();
      setField('razon_social',razon); setField('rfc',rfc);
      const banner=document.getElementById('returningBanner');
      if (banner) { banner.classList.add('active'); banner.querySelector('.banner-text').innerHTML='<strong>RFC identificado</strong><br>Capture los datos completos de la nueva sucursal. Las sucursales ya registradas no se modifican desde PAIC.'; }
      if (rfcInput) { rfcInput.value=rfc; rfcInput.dispatchEvent(new Event('input')); }
      closeReturningModal(); document.querySelector('.form-content')?.scrollIntoView({behavior:'smooth'});
    }
    window.applyNewBranchData=applyNewBranchData;
'''
paic = pattern.sub(new_returning, paic, count=1)

show_pattern = re.compile(r"\n    function showOnRecordIndicators\(\) \{.*?    window\.showFileInput=showFileInput;\n", re.S)
if len(show_pattern.findall(paic)) != 1:
    raise SystemExit('showOnRecordIndicators/showFileInput no único')
paic = show_pattern.sub('\n', paic, count=1)

clear_pattern = re.compile(r"    function clearPrefill\(\) \{.*?    window\.clearPrefill=clearPrefill;", re.S)
if len(clear_pattern.findall(paic)) != 1:
    raise SystemExit('clearPrefill no único')
paic = clear_pattern.sub(r'''    function clearPrefill() {
      document.getElementById('paicForm')?.reset();
      document.getElementById('returningBanner')?.classList.remove('active');
      resetUploadState();
      if (rfcInput) rfcInput.style.borderColor='';
    }
    window.clearPrefill=clearPrefill;''', paic, count=1)

# ---- Backend: alta inicial PAIC y separación de correos ----
backend_anchor = """    const companyClean = cleanCompanyName(data.razon_social || 'Cliente');
    const branchClean = sanitizeFileName(data.sucursal || 'Matriz');
    const timestamp = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyMMdd');
    // 1. LÓGICA DE CARPETA PADRE (Empresa)"""
backend_insert = """    const companyClean = cleanCompanyName(data.razon_social || 'Cliente');
    const branchClean = sanitizeFileName(data.sucursal || 'Matriz');
    const timestamp = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyMMdd');
    const esPaic = String((data && data.portal_origen) || '').toUpperCase() === 'PAIC';

    if (esPaic) {
      const correoAcuse = String(data.correo_acuse || '').trim();
      if (!correoAcuse) return { success: false, error: 'El correo del asesor/intermediario es obligatorio en PAIC.' };
      if (['si','no'].indexOf(String(data.aplica_nom020 || '').toLowerCase()) === -1 ||
          ['si','no'].indexOf(String(data.requiere_pipc || '').toLowerCase()) === -1) {
        return { success: false, error: 'Indique si este registro requiere NOM-020 y PIPC.' };
      }
      const existingRows = sheet.getDataRange().getValues();
      for (let i = existingRows.length - 1; i >= 1; i--) {
        const sameRfc = String(existingRows[i][CL.RFC] || '').toUpperCase().trim() === rfcClean;
        const sameBranch = sanitizeFileName(String(existingRows[i][CL.SUCURSAL] || 'Matriz')).toLowerCase() === branchClean.toLowerCase();
        if (sameRfc && sameBranch) {
          return { success: false, code: 'CLIENTE_SUCURSAL_YA_REGISTRADO', error: 'Este cliente y sucursal ya están registrados. PAIC solo permite el alta inicial; no vuelva a registrar ni actualizar esta sucursal desde este portal.' };
        }
      }
    }

    // 1. LÓGICA DE CARPETA PADRE (Empresa)"""
backend = replace_once(backend, backend_anchor, backend_insert, 'duplicate guard')

old_robusta = function_block(backend, 'enviarNotificacionRobusta')
new_robusta = r'''function enviarNotificacionRobusta(data, files, carpetaCliente, sheetUrl, addLog) {
  let lastError = null;
  const esPaic = String((data && data.portal_origen) || '').toUpperCase() === 'PAIC';
  for (let attempt = 1; attempt <= CONFIG.EMAIL_RETRY_ATTEMPTS; attempt++) {
    try {
      if (esPaic) enviarNotificacionEquipoPaic(data, files, carpetaCliente, sheetUrl);
      else enviarNotificacionEquipo(data, files, carpetaCliente, sheetUrl);
      Utilities.sleep(1000);
      if (esPaic) enviarConfirmacionPaic(data, carpetaCliente, files);
      else enviarConfirmacionCliente(data, carpetaCliente, files);
      return { success: true };
    } catch (error) {
      lastError = error;
      if (attempt < CONFIG.EMAIL_RETRY_ATTEMPTS) Utilities.sleep(CONFIG.EMAIL_RETRY_DELAY_MS);
    }
  }
  try {
    if (esPaic) enviarEmailSimpleFallbackPaic(data, carpetaCliente, sheetUrl);
    else enviarEmailSimpleFallback(data, carpetaCliente, sheetUrl);
    return { success: true, usedFallback: true };
  } catch (fallbackError) { return { success: false, error: fallbackError.toString() }; }
}
'''
backend = backend.replace(old_robusta, new_robusta, 1)

recipient_old = """function enviarConfirmacionCliente(data, carpetaCliente, files) {
  const esPaic = String((data && data.portal_origen) || '').toUpperCase() === 'PAIC';
  const emailCliente = esPaic ? String(data.correo_acuse || '').trim() : String(data.correo_informe || '').trim();
  if (!emailCliente) return;"""
recipient_new = """function enviarConfirmacionCliente(data, carpetaCliente, files) {
  const emailCliente = data.correo_informe;
  if (!emailCliente || emailCliente.trim() === '') return;"""
backend = replace_once(backend, recipient_old, recipient_new, 'SEAPD confirmation restore')

mail_anchor = 'function enviarNotificacionEquipo(data, files, carpetaCliente, sheetUrl) {'
if backend.count(mail_anchor) != 1:
    raise SystemExit('enviarNotificacionEquipo no único')
paic_mail = r'''function valorSiNoPaic_(value) {
  return String(value || '').toLowerCase() === 'si' ? 'SÍ' : 'NO';
}

function enviarNotificacionEquipoPaic(data, files, carpetaCliente, sheetUrl) {
  const timestamp = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd/MM/yyyy HH:mm');
  const docs = (files || []).map(f => `• ${f.label || f.name || 'Documento'}`).join('\n') || '• Sin archivos adjuntos';
  const subject = `Alta PAIC · ${data.razon_social || 'Cliente'} — ${data.sucursal || 'Sucursal'}`;
  const plain = [
    'NUEVA ALTA DE CLIENTE VÍA PAIC','',
    `Asesor / intermediario: ${data.nombre_solicitante || '-'}`,
    `Consultoría / despacho: ${data.asesor_consultor || '-'}`,
    `Correo de acuse del asesor: ${data.correo_acuse || '-'}`,'',
    `Cliente final: ${data.razon_social || '-'}`,
    `Sucursal: ${data.sucursal || '-'}`,
    `RFC: ${data.rfc || '-'}`,
    `Responsable en sitio: ${data.responsable || '-'}`,
    `Correo del cliente para informes: ${data.correo_informe || '-'}`,'',
    `Indicación NOM-020 en este registro: ${valorSiNoPaic_(data.aplica_nom020)}`,
    `Indicación PIPC en este registro: ${valorSiNoPaic_(data.requiere_pipc)}`,'',
    'Documentación recibida:',docs,'',
    `Carpeta Drive: ${carpetaCliente.getUrl()}`,
    `Perfil de datos: ${sheetUrl}`,'',
    'Este mensaje confirma un alta de cliente/sucursal vía asesor. No significa que exista una OT; las órdenes de trabajo se generan por separado en SEAOT.'
  ].join('\n');
  const html = envolturaEmail_(
    `Alta PAIC · ${escHtml_(data.razon_social || '')} · ${escHtml_(data.sucursal || '')}`,
    encabezadoEmail_({ kicker:'Alta vía PAIC', titulo:escHtml_(data.razon_social || 'Cliente final'), subtitulo:`Sucursal: ${valorOGuion_(data.sucursal)}`, metas:[{etiqueta:'Recibido',valor:timestamp},{etiqueta:'RFC',valor:valorOGuion_(data.rfc)}] }) +
    `<tr><td style="padding:26px 30px;">
      ${etiquetaEmail_('Asesor / intermediario')}
      <table width="100%" border="0" cellpadding="0" cellspacing="0">${filaEmail_('Nombre',valorOGuion_(data.nombre_solicitante))}${filaEmail_('Consultoría / despacho',valorOGuion_(data.asesor_consultor))}${filaEmail_('Correo de acuse',valorOGuion_(data.correo_acuse))}</table>
      ${separadorEmail_()}${etiquetaEmail_('Cliente final')}
      <table width="100%" border="0" cellpadding="0" cellspacing="0">${filaEmail_('Sucursal',valorOGuion_(data.sucursal))}${filaEmail_('RFC',valorOGuion_(data.rfc))}${filaEmail_('Responsable en sitio',valorOGuion_(data.responsable))}${filaEmail_('Correo para informes',valorOGuion_(data.correo_informe))}</table>
      ${separadorEmail_()}${etiquetaEmail_('Indicaciones del registro')}
      <p style="font-size:14px;line-height:1.7;margin:0;">NOM-020: <strong>${valorSiNoPaic_(data.aplica_nom020)}</strong><br>PIPC: <strong>${valorSiNoPaic_(data.requiere_pipc)}</strong></p>
      ${separadorEmail_()}<table border="0" cellpadding="0" cellspacing="0"><tr>${botonEmail_('Ver carpeta en Drive',carpetaCliente.getUrl(),'secundario')}${botonEmail_('Ver perfil de datos',sheetUrl,'secundario')}</tr></table>
      <p style="margin:18px 0 0;font-size:12px;color:${EMAIL_COLORS_.textoSuave};">Alta vía PAIC. No implica la creación de una orden de trabajo; SEAOT gestiona las OT por separado.</p>
    </td></tr>` + pieEmail_([`<strong style="color:#5A665A;">${CONFIG.COMPANY_NAME}</strong> · Registro PAIC`,'Mensaje automático para seguimiento interno.'])
  );
  GmailApp.sendEmail(CONFIG.EMAIL_TO.join(','), subject, plain, { htmlBody:html, name:CONFIG.COMPANY_NAME });
}

function enviarConfirmacionPaic(data, carpetaCliente, files) {
  const emailAsesor = String(data.correo_acuse || '').trim();
  if (!emailAsesor) return;
  const timestamp = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd/MM/yyyy HH:mm');
  const docsCount = (files || []).length;
  const plain = [
    'Recibimos el registro del cliente.','',
    `Cliente final: ${data.razon_social || '-'}`,`Sucursal: ${data.sucursal || '-'}`,`RFC: ${data.rfc || '-'}`,
    `Fecha de recepción: ${timestamp}`,`Documentos recibidos: ${docsCount}`,
    `Indicación NOM-020: ${valorSiNoPaic_(data.aplica_nom020)}`,`Indicación PIPC: ${valorSiNoPaic_(data.requiere_pipc)}`,'',
    'Este acuse confirma el alta de datos y documentación realizada a través de PAIC. No confirma ni genera una orden de trabajo; las OT se gestionan por separado con Ejecutiva Ambiental.','',
    `Atención a Clientes: ${CONFIG.SUPPORT_EMAIL} · ${CONFIG.SUPPORT_PHONE}`
  ].join('\n');
  const html = envolturaEmail_(
    `Recibimos el alta PAIC · ${escHtml_(data.razon_social || '')}`,
    encabezadoEmail_({ kicker:'Acuse PAIC', titulo:'Registro recibido', subtitulo:'Recibimos los datos del cliente final y de la sucursal.', metas:[{etiqueta:'Fecha de recepción',valor:timestamp}] }) +
    `<tr><td style="padding:26px 30px;"><p style="margin:0 0 18px;font-size:15px;line-height:1.7;color:${EMAIL_COLORS_.texto};">Este mensaje confirma la recepción del alta realizada por usted como asesor/intermediario.</p><table width="100%" border="0" cellpadding="0" cellspacing="0">${filaEmail_('Cliente final',valorOGuion_(data.razon_social))}${filaEmail_('Sucursal',valorOGuion_(data.sucursal))}${filaEmail_('RFC',valorOGuion_(data.rfc))}${filaEmail_('Documentos recibidos',String(docsCount))}${filaEmail_('Indicación NOM-020',valorSiNoPaic_(data.aplica_nom020))}${filaEmail_('Indicación PIPC',valorSiNoPaic_(data.requiere_pipc))}</table>${separadorEmail_()}<div style="background-color:${EMAIL_COLORS_.panel};border:1px solid ${EMAIL_COLORS_.borde};border-radius:8px;padding:16px;font-size:13px;line-height:1.6;color:${EMAIL_COLORS_.textoSuave};">El alta PAIC no genera una orden de trabajo. Las OT se gestionan posteriormente y por separado con Ejecutiva Ambiental.</div></td></tr>` +
    pieEmail_([`<strong style="color:#5A665A;">${CONFIG.COMPANY_NAME}</strong> · Portal de Asesores, Intermediarios y Consultorías`,`Para aclaraciones: ${escHtml_(CONFIG.SUPPORT_EMAIL)}`])
  );
  GmailApp.sendEmail(emailAsesor, `Registro PAIC recibido · ${data.razon_social || 'Cliente'} — ${data.sucursal || 'Sucursal'}`, plain, { htmlBody:html, name:CONFIG.COMPANY_NAME });
}

function enviarEmailSimpleFallbackPaic(data, carpetaCliente, sheetUrl) {
  const subject=`Alta PAIC · ${data.razon_social || 'Cliente'} — ${data.sucursal || 'Sucursal'}`;
  const body=`Alta de cliente/sucursal vía PAIC.\nCliente: ${data.razon_social || '-'}\nSucursal: ${data.sucursal || '-'}\nRFC: ${data.rfc || '-'}\nAsesor: ${data.nombre_solicitante || '-'}\nCorreo asesor: ${data.correo_acuse || '-'}\nCarpeta: ${carpetaCliente.getUrl()}\nPerfil: ${sheetUrl}\n\nNo implica creación de OT.`;
  GmailApp.sendEmail(CONFIG.EMAIL_TO.join(','),subject,body);
  const asesor=String(data.correo_acuse || '').trim();
  if (asesor) GmailApp.sendEmail(asesor,`Registro PAIC recibido · ${data.razon_social || 'Cliente'}`,`Recibimos el alta del cliente ${data.razon_social || '-'} / ${data.sucursal || '-'}. Este acuse no genera una OT; la OT se gestiona por separado.`);
}

'''
backend = backend.replace(mail_anchor, paic_mail + mail_anchor, 1)

# ---- Invariantes ----
protected_profile_after = function_block(backend, 'generarPerfilSheet', 'enviarNotificacionRobusta')
if protected_profile_after != protected_profile_before:
    raise SystemExit('PROHIBIDO: generarPerfilSheet() cambió')
if 'estudio_laboratorio' in paic or 'estudio_higiene' in paic or '1 registro = 1 estudio' in paic.lower():
    raise SystemExit('Se reintrodujo lógica de estudio prohibida')
if 'applyClientData' in paic or 'Actualizar ${escHtml' in paic:
    raise SystemExit('Quedó lógica de actualización de sucursal')
if 'name="aplica_nom020"' not in paic or 'name="requiere_pipc"' not in paic:
    raise SystemExit('No se restauraron NOM-020/PIPC')
if 'CLIENTE_SUCURSAL_YA_REGISTRADO' not in backend:
    raise SystemExit('No quedó guard de duplicado PAIC')

paic_path.write_text(paic, encoding='utf-8')
backend_path.write_text(backend, encoding='utf-8')
