const {test} = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

function entorno() {
  const elements = new Map();
  const get = id => {
    if(!elements.has(id)) elements.set(id, {value:'',innerHTML:'',textContent:'',hidden:false,open:false,showModal(){this.open=true;},close(){this.open=false;},addEventListener(){},focus(){},insertAdjacentHTML(_where,html){this.innerHTML += html;}});
    return elements.get(id);
  };
  const llamadas = [];
  const ctx = vm.createContext({
    URLSearchParams, Intl, Date, console,
    document: {getElementById:get, querySelectorAll:()=>[],body:{classList:{add(){},remove(){}}}},
    SEAAuth:{init(){},logout(){},wrapFetch:async url => {
      llamadas.push(new URL(url).searchParams.get('action'));
      return {ok:true,json:async()=>({success:true})};
    }}
  });
  for(const file of ['panel-detalle.js','panel.js']) vm.runInContext(fs.readFileSync(path.join(__dirname,'..',file),'utf8'),ctx);
  return {ctx,get,llamadas,run:code=>vm.runInContext(code,ctx)};
}

test('el cliente solo permite consultar resumen y detalle',async()=>{
  const e = entorno();
  await e.run("consultar('trazResumen')");
  await e.run("consultar('trazDetalle',{ot:'PRUEBA'})");
  await assert.rejects(e.run("consultar('actualizarOT')"),/modo consulta/);
  assert.deepEqual(e.llamadas,['trazResumen','trazDetalle']);
});

test('combina búsqueda y filtros sin confundir carpetas de OT con informes',()=>{
  const e = entorno();
  e.run(`RESUMEN=[
    {ot:'PRUEBA-1',cliente:'Plásticos',nom:'NOM-025',personal:'Equipo A',estado_traz:'EN PROCESO',folios_informe:['INF-1'],tiene_carpeta:true,num_informes:0},
    {ot:'PRUEBA-2',cliente:'Otra empresa',nom:'NOM-011',personal:'Equipo B',estado_traz:'CANCELADO',entregado:false,folios_informe:[],tiene_carpeta:false,num_informes:1}
  ]`);
  e.get('buscar').value = 'plasticos';
  e.get('servicio').value = 'NOM-025';
  assert.equal(e.run('filtrados().length'),1);
  e.get('estado').value = 'CANCELADO';
  assert.equal(e.run('filtrados().length'),0);
  assert.equal(e.run('alertasDe(RESUMEN[0])[0]'),'Sin informe relacionado');
  assert.equal(e.run('alertasDe(RESUMEN[1])[0]'),'Sin carpeta de OT');
});

test('escapa datos de Sheets y rechaza enlaces ejecutables',()=>{
  const e = entorno();
  assert.equal(e.run(`safeUrl('javascript:alert(1)')`),'');
  assert.equal(e.run(`safeUrl('https://drive.google.com/folder')`),'https://drive.google.com/folder');
  assert.equal(e.run(`esc('<img src=x onerror="alert(1)">')`),'&lt;img src=x onerror=&quot;alert(1)&quot;&gt;');
  const html = e.run(`tabla([{ot:'" onclick="alert(1)',cliente:'<script>bad</script>',folios_informe:[]}])`);
  assert.ok(!html.includes('<script>'));
  assert.ok(html.includes('&quot; onclick=&quot;'));
});

test('descarta un detalle tardío cuando se vuelve a resultados',async()=>{
  const e = entorno();
  let resolver;
  e.ctx.SEAAuth.wrapFetch = () => new Promise(resolve=> {resolver = resolve;});
  const pendiente = e.run("abrir('PRUEBA-1')");
  e.run('cerrarDetalle()');
  resolver({ok:true,json:async()=>({success:true,ot:{folio:'PRUEBA-1'}})});
  await pendiente;
  assert.equal(e.get('detail-shell').open,false);
  assert.equal(e.get('detalle').innerHTML,'');
});


test('filtro de serie exacto: OT no incluye OTB/OTC y no inventa series',()=>{
  const e = entorno();
  e.run(`RESUMEN=[
    {ot:'2611-01',tipo:'OT'},
    {ot:'OTB2611-02',tipo:''},
    {ot:'OTC 2611-03',tipo:'OTC'},
    {ot:'2611-04',tipo:'OTB'},
    {ot:'2611-05',tipo:''}
  ]; serieActiva='OT'`);
  assert.equal(e.run('filtrados().length'),1);
  e.run("serieActiva='OTB'");
  assert.equal(e.run('filtrados().length'),2);
  e.run("serieActiva='OTC'");
  assert.equal(e.run('filtrados().length'),1);
  assert.equal(e.run("serieDe({ot:'OTROS'})"),'');
  assert.equal(e.run("serieDe({ot:'2611-05'})"),'');
});

test('varios informes permanecen en una fila y OTC sin informe no dispara alerta',()=>{
  const e = entorno();
  const html=e.run(`tabla([{ot:'OTB2611-01',tipo:'OTB',folios_informe:['INF-1','INF-2'],num_informes:2,tiene_carpeta:true}])`);
  assert.equal((html.match(/<tbody><tr>/g)||[]).length,1);
  assert.equal((html.match(/<\/tr>/g)||[]).length,2);
  assert.ok(html.includes('INF-1') && html.includes('INF-2'));
  assert.equal(e.run("alertasDe({tipo:'OTC',num_informes:0,tiene_carpeta:true}).length"),0);
  assert.ok(e.run("celdaInformes({tipo:'OTC',num_informes:0})").includes('No aplica'));
  assert.ok(e.run("celdaInformes({tipo:'OTC',num_informes:1,folios_informe:['INF-EXISTENTE']})").includes('INF-EXISTENTE'));
});

test('abrir y cerrar detalle conserva la tabla y los filtros',async()=>{
  const e=entorno();
  e.get('results').innerHTML='tabla original';
  e.get('buscar').value='empresa';
  const focus={isConnected:true,focus(){this.restaurado=true;}};
  e.ctx.document.activeElement=focus;
  e.ctx.SEAAuth.wrapFetch=async()=>({ok:true,json:async()=>({success:true,ot:{folio:'OTC2611-01',tipo:'OTC'},informes:[],linea_tiempo:[],advertencias:['La OT no tiene ningún informe relacionado.']})});
  await e.run("abrir('OTC2611-01')");
  assert.equal(e.get('detail-shell').open,true);
  assert.ok(!e.get('detalle').innerHTML.includes('La OT no tiene ningún informe relacionado.'));
  e.run('cerrarDetalle()');
  assert.equal(e.get('results').innerHTML,'tabla original');
  assert.equal(e.get('results').hidden,false);
  assert.equal(e.get('buscar').value,'empresa');
  assert.equal(focus.restaurado,true);
});
