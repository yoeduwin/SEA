from pathlib import Path
import re

path = Path('PAIC.html')
s = path.read_text(encoding='utf-8')

# Remove the two legacy duplicate calibration containers; the visible calibration input + checkbox remain.
def remove_balanced_div(text, marker):
    start = text.find(marker)
    if start < 0:
        return text, False
    token_re = re.compile(r'<div\b|</div>', re.I)
    depth = 0
    began = False
    for m in token_re.finditer(text, start):
        if m.group(0).lower().startswith('<div'):
            depth += 1
            began = True
        else:
            depth -= 1
            if began and depth == 0:
                end = m.end()
                # consume one following newline only
                if end < len(text) and text[end:end+1] == '\n':
                    end += 1
                return text[:start] + text[end:], True
    raise SystemExit(f'No se pudo balancear {marker}')

for marker in ['<div id="seccion_archivo_calibracion"', '<div id="seccion_mensaje_calibracion"']:
    s, removed = remove_balanced_div(s, marker)
    if not removed:
        raise SystemExit(f'No se encontró bloque legado {marker}')

# Improve modal wording: it now includes general + conditional registration documents.
s = s.replace('<h3>Documentación general (${files.length})</h3>', '<h3>Documentación adjunta (${files.length})</h3>', 1)

# Add a reusable conditional-file reset helper immediately after clearFileUI.
anchor = """    document.querySelectorAll('input[type=\"file\"]').forEach(input => {"""
helper = r'''    function clearConditionalFilesByAttr(attrValue) {
      document.querySelectorAll(`[data-conditional-required="${attrValue}"]`).forEach(field => {
        field.required = false;
        if (field.type === 'file') {
          const key = field.dataset.fileInput;
          if (key) clearFileUI(key);
          else field.value = '';
        }
      });
    }

'''
if s.count(anchor) != 1:
    raise SystemExit('Anchor de file handlers inesperado')
s = s.replace(anchor, helper + anchor, 1)

# Expand reset to clear/hide PIPC subconditionals as well.
old_reset = r'''    function resetConditionalRegistrationSections() {
      document.querySelectorAll('input[name="aplica_nom020"], input[name="requiere_pipc"]').forEach(r => { r.checked=false; });
      document.getElementById('documentacionLegalSection')?.classList.remove('active');
      document.getElementById('documentacionPIPCSection')?.classList.remove('active');
      document.querySelectorAll('[data-conditional-required="nom020"], [data-conditional-required="pipc"]').forEach(f => { f.required=false; });
      const sinCal=document.getElementById('check_sin_calibracion'); if (sinCal) sinCal.checked=false;
      const calib=document.querySelector('[data-file-input="calibracion_valvula"]'); if (calib) { calib.style.display=''; calib.required=false; }
      const ast=document.getElementById('asterisco_calib'); if (ast) ast.style.display='';
      const help=document.getElementById('help_text_calib'); if (help) help.textContent='Adjunte el archivo PDF de su última calibración.';
    }
'''
new_reset = r'''    function resetConditionalRegistrationSections() {
      document.querySelectorAll('input[name="aplica_nom020"], input[name="requiere_pipc"], input[name^="pipc_tiene_"]').forEach(r => { r.checked=false; });
      document.getElementById('documentacionLegalSection')?.classList.remove('active');
      document.getElementById('documentacionPIPCSection')?.classList.remove('active');
      ['pipc_medidas_preventivas_section','pipc_gas_natural_section','pipc_sustancias_quimicas_section','pipc_dc3_operadores_section'].forEach(id => {
        const el=document.getElementById(id); if (el) el.style.display='none';
      });
      ['nom020','pipc','pipc_medidas','pipc_gas','pipc_quimicos','pipc_montacargas'].forEach(clearConditionalFilesByAttr);
      const sinCal=document.getElementById('check_sin_calibracion'); if (sinCal) sinCal.checked=false;
      const calib=document.querySelector('[data-file-input="calibracion_valvula"]'); if (calib) { calib.style.display=''; calib.required=false; }
      const ast=document.getElementById('asterisco_calib'); if (ast) ast.style.display='';
      const help=document.getElementById('help_text_calib'); if (help) help.textContent='Adjunte el archivo PDF de su última calibración.';
    }
'''
if s.count(old_reset) != 1:
    raise SystemExit('resetConditionalRegistrationSections inesperado')
s = s.replace(old_reset, new_reset, 1)

# Replace the simple PIPC top-level handler with complete nested conditional behavior.
old_handler = r'''    document.querySelectorAll('input[name="requiere_pipc"]').forEach(r => r.addEventListener('change', function() {
      const on=this.value==='si';
      document.getElementById('documentacionPIPCSection')?.classList.toggle('active',on);
      document.querySelectorAll('[data-conditional-required="pipc"]').forEach(f => { f.required=on; });
      if (!on) document.querySelectorAll('[data-conditional-required="pipc"]').forEach(f => {
        if (f.type==='file' && f.dataset.fileInput) clearFileUI(f.dataset.fileInput); else if (f.type!=='radio' && f.type!=='checkbox') f.value='';
      });
    }));
'''
new_handler = r'''    document.querySelectorAll('input[name="requiere_pipc"]').forEach(r => r.addEventListener('change', function() {
      const on=this.value==='si';
      document.getElementById('documentacionPIPCSection')?.classList.toggle('active',on);
      document.querySelectorAll('[data-conditional-required="pipc"]').forEach(f => { f.required=on; });
      if (!on) resetPIPCConditionals();
    }));

    function resetPIPCConditionals() {
      document.querySelectorAll('input[name^="pipc_tiene_"]').forEach(r => { r.checked=false; });
      ['pipc_medidas_preventivas_section','pipc_gas_natural_section','pipc_sustancias_quimicas_section','pipc_dc3_operadores_section'].forEach(id => {
        const el=document.getElementById(id); if (el) el.style.display='none';
      });
      ['pipc','pipc_medidas','pipc_gas','pipc_quimicos','pipc_montacargas'].forEach(clearConditionalFilesByAttr);
    }

    function bindPipcConditional(radioName, sectionId, requiredKey) {
      document.querySelectorAll(`input[name="${radioName}"]`).forEach(radio => radio.addEventListener('change', function() {
        const section=document.getElementById(sectionId);
        const on=this.value==='si';
        if (section) section.style.display=on ? 'block' : 'none';
        document.querySelectorAll(`[data-conditional-required="${requiredKey}"]`).forEach(f => { f.required=on; });
        if (!on) clearConditionalFilesByAttr(requiredKey);
      }));
    }
    bindPipcConditional('pipc_tiene_medidas','pipc_medidas_preventivas_section','pipc_medidas');
    bindPipcConditional('pipc_tiene_gas','pipc_gas_natural_section','pipc_gas');
    bindPipcConditional('pipc_tiene_quimicos','pipc_sustancias_quimicas_section','pipc_quimicos');
    bindPipcConditional('pipc_tiene_montacargas','pipc_dc3_operadores_section','pipc_montacargas');
'''
if s.count(old_handler) != 1:
    raise SystemExit('Handler PIPC simple inesperado')
s = s.replace(old_handler, new_handler, 1)

# Invariants for the frontend only.
assert s.count('name="calibracion_valvula"') == 1, 'Debe quedar un solo input calibracion_valvula'
assert 'estudio_laboratorio' not in s and 'estudio_higiene' not in s
assert 'applyClientData' not in s
assert 'bindPipcConditional' in s
assert 'name="aplica_nom020"' in s and 'name="requiere_pipc"' in s

path.write_text(s, encoding='utf-8')
