const SHEET_CIERRE = 'Cierre caja';
const SHEET_CONFIG = 'Configuración';
const SHEET_CAJA = 'Caja 3 meses';
const SHEET_GASTOS = 'Gastos';
const SHEET_DISTRIBUIDORES = 'Distribuidores';
const SHEET_MOVIMIENTOS = 'Gastos';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Caja')
    .addItem('Pasar cierre', 'pasarCierre')
    .addItem('Abrir aplicación de cierre', 'abrirAplicacion')
    .addToUi();
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Cierre de caja')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function abrirAplicacion() {
  const url = ScriptApp.getService().getUrl();
  if (!url) {
    SpreadsheetApp.getUi().alert('Primero despliega el proyecto como aplicación web.');
    return;
  }

  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput(
      '<p style="font-family:Arial">Abre la aplicación desde este enlace:</p>' +
      '<p><a target="_blank" href="' + url + '">' + url + '</a></p>'
    ).setWidth(500).setHeight(180),
    'Aplicación de cierre'
  );
}


/* =========================================================
   DISTRIBUIDORES
   ========================================================= */

function obtenerDistribuidores_() {
  const ss = SpreadsheetApp.getActive();
  const lista = [];

  const sh = ss.getSheetByName(SHEET_DISTRIBUIDORES);
  if (sh && sh.getLastRow() >= 1) {
    sh.getRange(1,1,sh.getLastRow(),1).getDisplayValues()
      .flat().map(x=>String(x).trim()).filter(Boolean)
      .forEach(x=>{ if(!lista.some(y=>y.toLowerCase()===x.toLowerCase())) lista.push(x); });
  }

  // También recuperamos distribuidores históricos de Gastos!B:B.
  const gastos = ss.getSheetByName(SHEET_MOVIMIENTOS);
  if (gastos && gastos.getLastRow() >= 2) {
    const vals = gastos.getRange(2,2,gastos.getLastRow()-1,1).getDisplayValues().flat();
    vals.map(x=>String(x).trim()).filter(Boolean).forEach(x=>{
      if(x.toUpperCase()==='INGRESO' || x.toLowerCase()==='distribuidor') return;
      if(!lista.some(y=>y.toLowerCase()===x.toLowerCase())) lista.push(x);
    });
  }

  return lista;
}


/* =========================================================
   CONCEPTOS DE INGRESOS
   ========================================================= */

function obtenerConceptosIngresos_() {
  const ss = SpreadsheetApp.getActive();
  const base = ['Cupones', 'Dinero prestado', 'Devolución', 'Otros ingresos'];
  const sh = ss.getSheetByName(SHEET_MOVIMIENTOS);
  const conceptosSh = ss.getSheetByName('Conceptos ingresos');

  let guardados = [];
  if (conceptosSh && conceptosSh.getLastRow() >= 2) {
    guardados = conceptosSh.getRange(2,1,conceptosSh.getLastRow()-1,1)
      .getDisplayValues().flat().map(x=>String(x).trim()).filter(Boolean);
  }

  let encontrados = [];
  if (sh && sh.getLastRow() >= 5) {
    const vals = sh.getRange(5,1,sh.getLastRow()-4,4).getDisplayValues();
    encontrados = vals
      .filter(r => String(r[1]).trim().toUpperCase() === 'INGRESO')
      .map(r => String(r[2]).trim())
      .filter(Boolean);
  }

  return [...new Set(base.concat(guardados, encontrados))];
}

function añadirConceptoIngresoDesdeApp(nombre) {
  nombre = texto(nombre).trim();
  if (!nombre) throw new Error('Escribe el concepto.');

  let sh = SpreadsheetApp.getActive().getSheetByName('Conceptos ingresos');
  if (!sh) {
    sh = SpreadsheetApp.getActive().insertSheet('Conceptos ingresos');
    sh.getRange('A1').setValue('Concepto');
  }

  const lista = obtenerConceptosIngresos_();
  if (!lista.some(x=>x.toLowerCase()===nombre.toLowerCase())) {
    sh.getRange(sh.getLastRow()+1,1).setValue(nombre);
  }
  return obtenerConceptosIngresos_();
}


/* =========================================================
   DATOS INICIALES
   ========================================================= */

function obtenerDatosInicialesApp() {
  const cierre = SpreadsheetApp.getActive().getSheetByName(SHEET_CIERRE);
  if (!cierre) throw new Error('No encuentro la hoja "Cierre caja".');

  const result = {
    dianas: {},
    cambioCajon: {},
    cambioCaja: {},
    ingresos: {recreativos:'', otros:''},
    gastos: {},
    cierre: {},
    calculados: {},
    distribuidores: obtenerDistribuidores_(),
    conceptosIngresos: obtenerConceptosIngresos_(),
    balance: obtenerBalanceMovimientos()
  };

  const d = cierre.getRange('A3:D3').getValues()[0];
  result.dianas = {diana1:d[0],diana2:d[1],diana3:d[2],diana4:d[3]};

  for(let r=7;r<=17;r++){
    result.cambioCajon['A'+r]=cierre.getRange('A'+r).getDisplayValue();
    result.cambioCajon['B'+r]=cierre.getRange('B'+r).getValue();
  }
  for(let r=20;r<=29;r++){
    result.cambioCaja['A'+r]=cierre.getRange('A'+r).getDisplayValue();
    result.cambioCaja['B'+r]=cierre.getRange('B'+r).getValue();
  }

  result.cierre = {
    santander:cierre.getRange('A34').getValue(),
    caixa:cierre.getRange('B34').getValue(),
    sumup:cierre.getRange('C34').getValue(),
    ticketTarjetas:cierre.getRange('D34').getValue()
  };

  result.calculados = leerCalculados_(cierre);

  // Tickets introducidos por el usuario.
  // B32 = Ticket metálico
  // D34 = Ticket tarjetas
  result.calculados.ticketMetalico = cierre.getRange('B32').getValue();
  result.calculados.ticketTarjetas = cierre.getRange('D34').getValue();

  return result;
}

function obtenerMovimientosDelDia() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_MOVIMIENTOS);
  const config = SpreadsheetApp.getActive().getSheetByName(SHEET_CONFIG);
  const result = {gastos:[],ingresos:[]};
  if(!sh || sh.getLastRow()<5) return result;

  const fecha = config ? config.getRange('B3').getValue() : new Date();
  if(!(fecha instanceof Date)) return result;

  const key = Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const vals = sh.getRange(5,1,sh.getLastRow()-4,4).getValues();

  vals.forEach(r=>{
    if(!(r[0] instanceof Date)) return;
    if(Utilities.formatDate(r[0], Session.getScriptTimeZone(), 'yyyy-MM-dd')!==key) return;

    const distribuidor=String(r[1]||'').trim();
    const concepto=String(r[2]||'').trim();

    // Formato nuevo: A Fecha, B Distribuidor/INGRESO, C Concepto, D Importe.
    // Formato antiguo: A Fecha, B Distribuidor, C Importe.
    let importe=Number(r[3])||0;
    let conceptoFinal=concepto;
    if(!importe && r[2] !== '' && r[2] !== null && r[2] !== undefined){
      const antiguo=Number(r[2]);
      if(!isNaN(antiguo) && antiguo!==0){
        importe=antiguo;
        conceptoFinal='';
      }
    }
    if(!importe) return;

    if(distribuidor.toUpperCase()==='INGRESO'){
      result.ingresos.push({importe,concepto:conceptoFinal});
    }else{
      result.gastos.push({importe,distribuidor,concepto:conceptoFinal});
    }
  });
  return result;
}


/* =========================================================
   VALORES CALCULADOS
   ========================================================= */

function leerCalculados_(cierre) {
  SpreadsheetApp.flush();

  // A32 y D32 son valores calculados por la hoja.
  // La aplicación SOLO los lee: nunca los escribe.
  return {
    metalico: cierre.getRange('A32').getDisplayValue(),
    totalCaja: cierre.getRange('D32').getDisplayValue(),
    ticketMetalico: cierre.getRange('B32').getDisplayValue(),
    cuadraCaja: cierre.getRange('C32').getDisplayValue(),
    ticketTarjetas: cierre.getRange('D34').getDisplayValue(),
    cuadraTarjeta: cierre.getRange('E34').getDisplayValue()
  };
}


/* =========================================================
   MOVIMIENTOS DINÁMICOS
   ========================================================= */

function guardarMovimientosDesdeApp(tipo, movimientos) {
  if(tipo!=='gasto' && tipo!=='ingreso') throw new Error('Tipo de movimiento no válido.');
  if(!Array.isArray(movimientos) || !movimientos.length) throw new Error('No hay movimientos para guardar.');

  const ss=SpreadsheetApp.getActive();
  const sh=ss.getSheetByName(SHEET_MOVIMIENTOS);
  if(!sh) throw new Error('No encuentro la hoja "Gastos".');

  // La hoja puede venir del formato antiguo de 3 columnas.
  // No bloqueamos el guardado: D se crea/usa automáticamente.
  if(sh.getLastRow()===0){
    sh.getRange('A1:D1').setValues([['Fecha','Distribuidor','Concepto','Importe']]);
  }

  const config=ss.getSheetByName(SHEET_CONFIG);
  const fecha=config ? config.getRange('B3').getValue() : new Date();
  const filas=[];

  movimientos.forEach(m=>{
    const importe=numero(m.importe);
    const concepto=texto(m.concepto).trim();
    const distribuidor=texto(m.distribuidor).trim();

    if(importe!=='' && importe>0){
      if(tipo==='gasto' && !distribuidor) throw new Error('Selecciona un distribuidor.');
      if(tipo==='ingreso' && !concepto) throw new Error('Selecciona o escribe un concepto.');

      filas.push([fecha, tipo==='ingreso'?'INGRESO':distribuidor, concepto, importe]);

      if(tipo==='gasto') asegurarLista_(SHEET_DISTRIBUIDORES,distribuidor,'Distribuidor');
      if(tipo==='ingreso') asegurarLista_('Conceptos ingresos',concepto,'Concepto');
    }
  });

  if(!filas.length) throw new Error('Introduce un importe válido.');

  const inicio=sh.getLastRow()+1;
  sh.getRange(inicio,1,filas.length,4).setValues(filas);
  sh.getRange(inicio,1,filas.length,1).setNumberFormat('dd/MM/yyyy');
  sh.getRange(inicio,4,filas.length,1).setNumberFormat('#,##0.00');

  actualizarE5DesdeGastos_();
  SpreadsheetApp.flush();

  return {
    movimientos: obtenerMovimientosDelDia(),
    balance: obtenerBalanceMovimientos(),
    distribuidores: obtenerDistribuidores_(),
    conceptosIngresos: obtenerConceptosIngresos_(),
    mensaje: tipo==='gasto'?'Gasto guardado correctamente.':'Ingreso guardado correctamente.'
  };
}

function obtenerBalanceMovimientos() {
  const mov=obtenerMovimientosDelDia();
  const gastos=mov.gastos.reduce((a,m)=>a+(Number(m.importe)||0),0);
  const ingresos=mov.ingresos.reduce((a,m)=>a+(Number(m.importe)||0),0);
  return {gastos,ingresos,balance:gastos-ingresos};
}

function actualizarE5DesdeGastos_(){
  const cierre=SpreadsheetApp.getActive().getSheetByName(SHEET_CIERRE);
  if(!cierre) throw new Error('No encuentro la hoja "Cierre caja".');
  const b=obtenerBalanceMovimientos();
  cierre.getRange('E5').setValue(b.balance);
  cierre.getRange('E5').setNumberFormat('#,##0.00');
  return b.balance;
}

function asegurarLista_(sheetName, valor, encabezado){
  const ss=SpreadsheetApp.getActive();
  let sh=ss.getSheetByName(sheetName);
  if(!sh){
    sh=ss.insertSheet(sheetName);
    sh.getRange('A1').setValue(encabezado);
  }
  const limpio=String(valor).trim();
  if(!limpio) return;
  const vals=sh.getLastRow()?sh.getRange(1,1,sh.getLastRow(),1).getDisplayValues().flat().map(x=>String(x).trim().toLowerCase()):[];
  if(vals.indexOf(limpio.toLowerCase())===-1){
    sh.getRange(Math.max(sh.getLastRow()+1,2),1).setValue(limpio);
  }
}


/* =========================================================
   ACTUALIZAR VALORES
   ========================================================= */

function actualizarValoresDesdeApp(datos) {
  escribirDatosApp_(datos);

  SpreadsheetApp.flush();
  Utilities.sleep(500);
  SpreadsheetApp.flush();

  const cierre = SpreadsheetApp.getActive().getSheetByName(SHEET_CIERRE);

  return {
    ok: true,
    calculados: leerCalculados_(cierre),
    mensaje: 'Valores actualizados correctamente.'
  };
}


/* =========================================================
   GUARDAR Y CERRAR CAJA
   ========================================================= */

function guardarCierreDesdeApp(datos) {
  escribirDatosApp_(datos);

  SpreadsheetApp.flush();
  Utilities.sleep(500);
  SpreadsheetApp.flush();

  const cierre = SpreadsheetApp.getActive().getSheetByName(SHEET_CIERRE);

  guardarGastosHistoricos_(cierre);

  // Primero hacemos el cierre y el reseteo de la hoja.
  const resultado = pasarCierreInterno_();

  // Ahora volvemos a leer los valores DESPUÉS del cierre.
  // Así la aplicación recibe el estado real ya reseteado.
  SpreadsheetApp.flush();
  resultado.calculados = leerCalculados_(cierre);

  return resultado;
}


/* =========================================================
   ESCRIBIR DATOS DEL CIERRE
   ========================================================= */

function escribirDatosApp_(datos) {
  const cierre=SpreadsheetApp.getActive().getSheetByName(SHEET_CIERRE);
  if(!cierre) throw new Error('No encuentro la hoja "Cierre caja".');

  cierre.getRange('A3:D3').setValues([[
    numero(datos.diana1),numero(datos.diana2),numero(datos.diana3),numero(datos.diana4)
  ]]);

  const cajon=[]; for(let r=7;r<=17;r++) cajon.push([numero(datos['cajon'+r])]);
  cierre.getRange('B7:B17').setValues(cajon);

  const caja=[]; for(let r=20;r<=29;r++) caja.push([numero(datos['caja'+r])]);
  cierre.getRange('B20:B29').setValues(caja);

  cierre.getRange('A34:C34').setValues([[
    numero(datos.santander),numero(datos.caixa),numero(datos.sumup)
  ]]);
  // Ticket tarjetas D34:
  // solo se escribe si la app trae realmente un valor.
  // Si viene vacío, conservamos el valor que ya existe en la hoja.
  if (
    datos.ticketTarjetas !== undefined &&
    datos.ticketTarjetas !== null &&
    String(datos.ticketTarjetas).trim() !== ''
  ) {
    cierre.getRange('D34').setValue(numero(datos.ticketTarjetas));
  }

  // Ticket metálico B32:
  // solo se escribe si la app trae realmente un valor.
  // Si viene vacío, conservamos el valor que ya existe en la hoja.
  if (
    datos.ticketMetalico !== undefined &&
    datos.ticketMetalico !== null &&
    String(datos.ticketMetalico).trim() !== ''
  ) {
    cierre.getRange('B32').setValue(numero(datos.ticketMetalico));
  }

  // A32, C32 y E34 son calculados por las fórmulas de la hoja.
  // No los escribimos desde la aplicación.

  // E5 se recalcula exclusivamente desde la hoja Gastos.
  actualizarE5DesdeGastos_();
  SpreadsheetApp.flush();
}


/* =========================================================
   HISTÓRICO ANTIGUO DE GASTOS
   ========================================================= */

function guardarGastosHistoricos_(cierre) {
  // La hoja Gastos ya es el histórico definitivo. No se copia nada desde Cierre caja.
}


/* =========================================================
   PASAR CIERRE
   ========================================================= */

function pasarCierre() {
  try {
    SpreadsheetApp.getUi().alert(
      pasarCierreInterno_().mensaje
    );
  } catch (error) {
    SpreadsheetApp.getUi().alert(
      'No se ha podido hacer el cierre:\n\n' +
      error.message
    );
  }
}


function pasarCierreInterno_() {
  const ss = SpreadsheetApp.getActive();

  const cierre = ss.getSheetByName(SHEET_CIERRE);
  const config = ss.getSheetByName(SHEET_CONFIG);
  const caja = ss.getSheetByName(SHEET_CAJA);

  if (!cierre || !config || !caja) {
    throw new Error(
      'Falta alguna hoja: Cierre caja, Configuración o Caja 3 meses.'
    );
  }

  const fecha = config.getRange('B3').getValue();

  if (!(fecha instanceof Date)) {
    throw new Error(
      'La fecha de Configuración!B3 no es válida.'
    );
  }

  const valores = [[
    cierre.getRange('A32').getValue(),
    cierre.getRange('A34').getValue(),
    cierre.getRange('C34').getValue(),
    cierre.getRange('B34').getValue()
  ]];

  const n = Math.max(caja.getLastRow() - 5, 0);

  const fechas = n
    ? caja.getRange(6, 1, n, 1).getValues()
    : [];

  let fila = -1;

  for (let i = 0; i < fechas.length; i++) {
    const d = fechas[i][0];

    if (
      d instanceof Date &&
      d.getFullYear() === fecha.getFullYear() &&
      d.getMonth() === fecha.getMonth() &&
      d.getDate() === fecha.getDate()
    ) {
      fila = i + 6;
      break;
    }
  }

  if (fila === -1) {
    throw new Error(
      'No encuentro esa fecha en Caja 3 meses. Comprueba el trimestre/mes.'
    );
  }

  const diaSemana = caja
    .getRange(fila, 2)
    .getDisplayValue()
    .toLowerCase();

  if (diaSemana.startsWith('lun')) {
    throw new Error(
      'Ese día es lunes y está reservado para el resumen semanal.'
    );
  }

  const actuales = caja
    .getRange(fila, 3, 1, 4)
    .getValues()[0];

  if (actuales.some(v => v !== '' && v !== null)) {
    throw new Error(
      'Ese día ya tiene datos. Para evitar sobrescribirlos, revisa el cierre desde la hoja.'
    );
  }

  caja.getRange(fila, 3, 1, 4).setValues(valores);

  SpreadsheetApp.flush();

  cierre.getRange('D9')
    .setValue(cierre.getRange('D11').getValue());

  cierre.getRange('D25')
    .setValue(cierre.getRange('A30').getValue());

  cierre.getRange('B32').clearContent();
  cierre.getRange('A34:D34').clearContent();

  SpreadsheetApp.flush();

  const siguiente = new Date(fecha);

  if (siguiente.getDay() === 0) {
    siguiente.setDate(siguiente.getDate() + 2);
  } else {
    siguiente.setDate(siguiente.getDate() + 1);
  }

  config.getRange('B3')
    .setValue(siguiente)
    .setNumberFormat('dd/mm/yyyy');

  SpreadsheetApp.flush();
  actualizarE5DesdeGastos_();

  const fechaTexto = Utilities.formatDate(
    siguiente,
    Session.getScriptTimeZone(),
    'dd/MM/yyyy'
  );

  return {
    ok: true,
    mensaje:
      'Cierre pasado correctamente. La caja se ha reseteado y la próxima fecha es ' +
      fechaTexto +
      '.'
  };
}


/* =========================================================
   UTILIDADES
   ========================================================= */

function numero(v) {
  if (v === '' || v === null || v === undefined) {
    return '';
  }

  if (typeof v === 'number') {
    return v;
  }

  const n = Number(
    String(v).replace(',', '.')
  );

  return isNaN(n) ? '' : n;
}


function texto(v) {
  if (v === null || v === undefined) {
    return '';
  }

  return String(v);
}
