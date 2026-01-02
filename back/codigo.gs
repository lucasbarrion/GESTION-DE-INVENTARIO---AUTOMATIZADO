const CONFIG = {
  SHEET_PRODUCTOS: 'Productos',
  SHEET_MOVIMIENTOS: 'MOVIMIENTOS',
  SHEET_GASTOS: 'GASTOS PROOVEDOR',
  CATEGORIAS: ['Mate', 'Bombilla', 'Yerba', 'Accesorio'],
  LOGO_FILE_ID: 'inserte url de tu hoja de sheets',
  STOCK_BAJO_UMBRAL: 3
};


function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('IDOS DEL MATE SYSTEM')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function asegurarHoja(nombre, headers, formatos = {}) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(nombre);

  if (!sh) {
    sh = ss.insertSheet(nombre);
  }

  if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
    sh.setFrozenRows(1);

    Object.entries(formatos).forEach(([col, fmt]) => {
      sh.getRange(col + ':' + col).setNumberFormat(fmt);
    });
  }

  return sh;
}
function crearHojasSistema() {

  asegurarHoja(
    CONFIG.SHEET_PRODUCTOS,
    ['ID_PRODUCTO','NOMBRE','CATEGORIA','STOCK','COSTO_UNITARIO'],
    { D: '0', E: '#,##0.00' }
  );

  asegurarHoja(
    CONFIG.SHEET_MOVIMIENTOS,
    [
      'ID_MOVIMIENTO','FECHA','MATE','BOMBILLA','YERBA','ACCESORIO',
      'GRABADO','TIPO','MONTO','OBS','PERSONA'
    ],
    { B: 'yyyy-mm-dd', I: '#,##0.00' }
  );

  asegurarHoja(
    CONFIG.SHEET_GASTOS,
    ['FECHA','MONTO'],
    { A: 'yyyy-mm-dd', B: '#,##0.00' }
  );

  return { ok: true };
}
function formatearEncabezadosSistema() {
  const ss = SpreadsheetApp.getActive();

  const CONFIG_ENCABEZADOS = [
    {
      hoja: 'Productos',
      rango: 'A1:E1'
    },
    {
      hoja: 'MOVIMIENTOS',
      rango: 'A1:K1'
    },
    {
      hoja: 'GASTOS PROOVEDOR',
      rango: 'A1:B1'
    }
  ];

  CONFIG_ENCABEZADOS.forEach(cfg => {
    const sh = ss.getSheetByName(cfg.hoja);
    if (!sh) return;

    const header = sh.getRange(cfg.rango);

    header
      .setBackground('#2ecc71')   // verde atractivo (no chillón)
      .setFontColor('#ffffff')    // blanco
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true);

    sh.setRowHeight(1, 34);
  });
}


function aplicarFormatoStock() {
  const sh = SpreadsheetApp.getActive().getSheetByName('Productos');
  if (!sh) throw new Error('No existe la hoja Productos');

  const rangoStock = sh.getRange('D2:D'); // STOCK

  const reglas = [

    // 🔴 0, 1 o 2 → ROJO
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThanOrEqualTo(2)
      .setBackground('#dc2626')
      .setFontColor('#ffffff')
      .setRanges([rangoStock])
      .build(),

    // 🟠 3 o 4 → NARANJA
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(3, 4)
      .setBackground('#f59e0b')
      .setFontColor('#000000')
      .setRanges([rangoStock])
      .build(),

    // 🟢 5 o más → VERDE
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(5)
      .setBackground('#16a34a')
      .setFontColor('#ffffff')
      .setRanges([rangoStock])
      .build()
  ];

  sh.setConditionalFormatRules(reglas);
}

function agregarProducto(p) {
  validarTexto(p?.id, 'ID_PRODUCTO');
  validarTexto(p?.nombre, 'NOMBRE');
  validarTexto(p?.categoria, 'CATEGORIA');
  validarNumeroNoNegativo(p?.costo, 'COSTO_UNITARIO');

  if (!CONFIG.CATEGORIAS.includes(p.categoria)) {
    throw new Error('Categoría inválida');
  }

  const sh = asegurarHoja(
    CONFIG.SHEET_PRODUCTOS,
    ['ID_PRODUCTO','NOMBRE','CATEGORIA','STOCK','COSTO_UNITARIO'],
    { D: '0', E: '#,##0.00' }
  );

  const ids = sh.getLastRow() > 1
    ? sh.getRange(2,1,sh.getLastRow()-1,1).getValues().flat()
    : [];

  if (ids.includes(p.id)) {
    throw new Error('ID duplicado');
  }

  sh.appendRow([
    p.id,
    p.nombre,
    p.categoria,
    0,
    Number(p.costo)
  ]);

  return { ok: true };
}

function actualizarCosto(data) {
  if (!data) throw new Error('Datos inexistentes');

  const id = String(data.idProducto || '').trim();
  if (!id) throw new Error('Producto obligatorio');

  const nuevoCosto = Number(data.nuevoCosto);
  if (!isFinite(nuevoCosto) || nuevoCosto < 0) {
    throw new Error('Costo inválido');
  }

  const sh = SpreadsheetApp
    .getActive()
    .getSheetByName(CONFIG.SHEET_PRODUCTOS);

  if (!sh || sh.getLastRow() < 2) {
    throw new Error('No hay productos cargados');
  }

  const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();

  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === id) {
      const fila = i + 2;
      sh.getRange(fila, 5).setValue(nuevoCosto); // COSTO_UNITARIO
      return { ok: true };
    }
  }

  throw new Error('Producto no encontrado');
}


function guardarMovimiento(m) {
  if (!m) throw new Error('Datos de movimiento inexistentes');

  if (!m.fecha) throw new Error('FECHA obligatoria');
  if (!isFinite(Number(m.monto)) || Number(m.monto) === 0) {
    throw new Error('MONTO inválido');
  }
  if (!m.persona) throw new Error('PERSONA obligatoria');

  const sh = asegurarHoja(
    CONFIG.SHEET_MOVIMIENTOS,
    [
      'ID_MOVIMIENTO','FECHA','MATE','BOMBILLA','YERBA','ACCESORIO',
      'GRABADO','TIPO','MONTO','OBS','PERSONA'
    ],
    {
      B: 'yyyy-mm-dd',
      I: '#,##0.00'
    }
  );

  const fecha = new Date(m.fecha);
  fecha.setHours(0, 0, 0, 0);

  sh.appendRow([
    'MOV-' + Date.now(),  
    fecha,                
    m.mate || '',          
    m.bombilla || '',      
    m.yerba || '',         
    m.accesorio || '',     
    m.grabado || 'NO',     
    m.tipo || '',          
    Number(m.monto),       
    m.obs || '',           
    m.persona             
  ]);

  return { ok: true };
}


function ajustarStock(data) {
  if (!data || !data.idProducto) {
    throw new Error('Producto obligatorio');
  }

  const cantidad = Number(data.cantidad);
  if (!Number.isInteger(cantidad) || cantidad === 0) {
    throw new Error('Cantidad inválida');
  }

  const sh = SpreadsheetApp
    .getActive()
    .getSheetByName(CONFIG.SHEET_PRODUCTOS);

  if (!sh || sh.getLastRow() < 2) {
    throw new Error('No hay productos cargados');
  }

  const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();

  for (let i = 0; i < ids.length; i++) {
    if (ids[i][0] === data.idProducto) {
      const fila = i + 2;
      const celdaStock = sh.getRange(fila, 4);
      const stockActual = Number(celdaStock.getValue()) || 0;
      const nuevoStock = stockActual + cantidad;

      if (nuevoStock < 0) {
        throw new Error('El stock no puede ser negativo');
      }

      celdaStock.setValue(nuevoStock);
      return { ok: true };
    }
  }

  throw new Error('Producto no encontrado');
}


function registrarGasto(g) {
  if (!g) throw new Error('Datos de gasto inexistentes');

  if (!g.fecha) {
    throw new Error('FECHA obligatoria');
  }

  const monto = Number(g.monto);
  if (!isFinite(monto) || monto <= 0) {
    throw new Error('MONTO inválido');
  }

  const ss = SpreadsheetApp.getActive(); 
  const sh = asegurarHoja(
    CONFIG.SHEET_GASTOS,
    ['FECHA', 'MONTO'],
    { A: 'yyyy-mm-dd', B: '#,##0.00' }
  );

  const fecha = new Date(g.fecha);
  fecha.setHours(0, 0, 0, 0);

  
  sh.appendRow([
    fecha,
    -Math.abs(monto)
  ]);

  return { ok: true };
}




function getProductos() {
  const sh = SpreadsheetApp
    .getActive()
    .getSheetByName(CONFIG.SHEET_PRODUCTOS);

  if (!sh || sh.getLastRow() < 2) return [];

  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();

  return data.map(r => ({
    id: r[0],
    nombre: r[1],
    categoria: r[2],
    stock: Number(r[3] || 0),
    costo: Number(r[4] || 0)
  }));
}

/************ DASHBOARD ************/
function getDashboardData(filtro) {
  const desde = filtro?.desde ? new Date(filtro.desde) : null;
  const hasta = filtro?.hasta ? new Date(filtro.hasta) : null;
  if (hasta) hasta.setHours(23, 59, 59, 999);

  const ss = SpreadsheetApp.getActive();
  const shP = ss.getSheetByName(CONFIG.SHEET_PRODUCTOS);
  const shM = ss.getSheetByName(CONFIG.SHEET_MOVIMIENTOS);
  const shG = ss.getSheetByName(CONFIG.SHEET_GASTOS);


  const productos = shP && shP.getLastRow() >= 2
    ? shP.getRange(2, 1, shP.getLastRow() - 1, 5).getValues()
    : [];


  const movimientos = shM && shM.getLastRow() >= 2
    ? shM.getRange(2, 1, shM.getLastRow() - 1, 11).getValues()
    : [];

  
  let valorInventario = 0;
  let sinStock = 0;
  let stockBajo = 0;

  productos.forEach(r => {
    const stock = Number(r[3] || 0);
    const costo = Number(r[4] || 0);

    valorInventario += stock * costo;

    if (stock === 0) sinStock++;
    if (stock > 0 && stock <= CONFIG.STOCK_BAJO_UMBRAL) stockBajo++;
  });

 
  let ingresosMes = 0;
  let egresosMovimientos = 0;
  let movimientosFiltrados = 0;

  movimientos.forEach(r => {
    const fecha = new Date(r[1]);
    if (desde && fecha < desde) return;
    if (hasta && fecha > hasta) return;

    movimientosFiltrados++;

    const tipo = String(r[7] || '');
    const monto = Number(r[8] || 0);

    if (tipo === 'Ingreso') ingresosMes += monto;
    if (tipo === 'Egreso') egresosMovimientos += monto;
  });

  const gastos = shG && shG.getLastRow() >= 2
    ? shG.getRange(2, 1, shG.getLastRow() - 1, 2).getValues()
    : [];

  let costoVentaMes = 0;

  gastos.forEach(r => {
    const fecha = new Date(r[0]);
    if (desde && fecha < desde) return;
    if (hasta && fecha > hasta) return;

    const monto = Number(r[1] || 0);
    costoVentaMes += Math.abs(monto); 
  });


  costoVentaMes += egresosMovimientos;

 
  const utilidadBrutaMes = ingresosMes - costoVentaMes;
  const margenBruto = ingresosMes > 0
    ? (utilidadBrutaMes / ingresosMes) * 100
    : 0;

  return {
    totalProductos: productos.length,
    totalMovimientos: movimientosFiltrados,

    sinStock,
    stockBajo,

    valorInventario,

    ingresosMes,
    costoVentaMes,
    utilidadBrutaMes,
    margenBruto
  };
}



function getMovimientos(filtro) {
Logger.log('🔥🔥🔥 ENTRE A getMovimientos 🔥🔥🔥');
  filtro = filtro || {}; // 

  Logger.log('getMovimientos llamado');
  Logger.log(JSON.stringify(filtro));


  Logger.log('filtro normalizado: ' + JSON.stringify(filtro));
  
  const sh = SpreadsheetApp
    .getActive()
    .getSheetByName(CONFIG.SHEET_MOVIMIENTOS);

  if (!sh || sh.getLastRow() < 2) return [];

  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 11).getValues();

  const desde = filtro.desde && filtro.desde !== ''
    ? new Date(filtro.desde)
    : null;

  const hasta = filtro.hasta && filtro.hasta !== ''
    ? new Date(filtro.hasta)
    : null;

  if (desde) desde.setHours(0,0,0,0);
  if (hasta) hasta.setHours(23,59,59,999);

  const filtrados = data.filter(r => {

    if (!r[1]) return false;

    let fecha;
if (r[1] instanceof Date) {
  fecha = new Date(r[1]);
} else {
  fecha = new Date(String(r[1]));
}

if (isNaN(fecha)) return false;

fecha.setHours(0,0,0,0);


    if (desde && fecha < desde) return false;
    if (hasta && fecha > hasta) return false;

    if (filtro.persona && !String(r[10]).toLowerCase().includes(filtro.persona.toLowerCase())) return false;
    if (filtro.mate && !String(r[2]).toLowerCase().includes(filtro.mate.toLowerCase())) return false;
    if (filtro.bombilla && !String(r[3]).toLowerCase().includes(filtro.bombilla.toLowerCase())) return false;
    if (filtro.yerba && !String(r[4]).toLowerCase().includes(filtro.yerba.toLowerCase())) return false;
    if (filtro.accesorio && !String(r[5]).toLowerCase().includes(filtro.accesorio.toLowerCase())) return false;

    return true;
  });

  Logger.log('Cantidad devuelta: ' + filtrados.length);
Logger.log('ANTES DEL RETURN (array crudo):');
Logger.log(JSON.stringify(filtrados));

const resultado = filtrados.map(r => ({
  id: r[0],
  fecha: Utilities.formatDate(
  new Date(r[1]),
  Session.getScriptTimeZone(),
  'yyyy-MM-dd'
),
  mate: r[2],
  bombilla: r[3],
  yerba: r[4],
  accesorio: r[5],
  grabado: r[6],
  tipo: r[7],
  monto: Number(r[8] || 0),
  obs: r[9],
  persona: r[10]
}));

Logger.log('RETURN FINAL (array de objetos):');
Logger.log(JSON.stringify(resultado));

return resultado;
}


function DEBUG_movimientos() {
  const sh = SpreadsheetApp.getActive().getSheetByName('MOVIMIENTOS');
  const data = sh.getRange(2,1,sh.getLastRow()-1,11).getValues();
  Logger.log(data);
}


function exportarMovimientosCSV(filtro) {
  const rows = getMovimientos(filtro);
  if (!rows.length) throw new Error('No hay datos para exportar');

  const header = [
    'ID','Fecha','Mate','Bombilla','Yerba','Accesorio',
    'Grabado','Tipo','Monto','Obs','Persona'
  ];

  const data = rows.map(r => [
    r.id, r.fecha, r.mate, r.bombilla, r.yerba, r.accesorio,
    r.grabado, r.tipo, r.monto, r.obs, r.persona
  ]);

  const csv = [header, ...data]
    .map(r => r.map(v => `"${String(v).replace(/"/g,'""')}"`).join(','))
    .join('\n');

  return csv;
}

function getInventario() {
  const sh = SpreadsheetApp
    .getActive()
    .getSheetByName(CONFIG.SHEET_PRODUCTOS);

  if (!sh || sh.getLastRow() < 2) return [];

  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();

  return data.map(r => ({
    id: r[0],
    nombre: r[1],
    categoria: r[2],
    stock: Number(r[3] || 0),
    costo: Number(r[4] || 0)
  }));
}

function getAlertasStock() {
  const sh = SpreadsheetApp
    .getActive()
    .getSheetByName(CONFIG.SHEET_PRODUCTOS);

  if (!sh || sh.getLastRow() < 2) return [];

  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();

  return data
    .map(r => ({
      id: r[0],
      nombre: r[1],
      categoria: r[2],
      stock: Number(r[3] || 0),
      costo: Number(r[4] || 0)
    }))
    .filter(p =>
      p.stock === 0 ||
      p.stock <= CONFIG.STOCK_BAJO_UMBRAL
    );
}


function buscarProductosBackend(texto) {
  const sh = SpreadsheetApp
    .getActive()
    .getSheetByName(CONFIG.SHEET_PRODUCTOS);

  if (!sh || sh.getLastRow() < 2) return [];

  const data = sh.getRange(2,1,sh.getLastRow()-1,5).getValues();
  const q = String(texto || '').toLowerCase();

  return data
    .filter(r =>
      String(r[0]).toLowerCase().includes(q) || // ID
      String(r[1]).toLowerCase().includes(q) || // Nombre
      String(r[2]).toLowerCase().includes(q)    // Categoria
    )
    .map(r => ({
      id: r[0],
      nombre: r[1],
      categoria: r[2],
      stock: r[3],
      costo: r[4]
    }));
}

function getInventarioSoloAlertas() {
  const sh = SpreadsheetApp
    .getActive()
    .getSheetByName(CONFIG.SHEET_PRODUCTOS);

  if (!sh || sh.getLastRow() < 2) return [];

  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();

  return data
    .map(r => ({
      id: r[0],
      nombre: r[1],
      categoria: r[2],
      stock: Number(r[3] || 0),
      costo: Number(r[4] || 0)
    }))
    .filter(p =>
      p.stock === 0 ||
      p.stock <= CONFIG.STOCK_BAJO_UMBRAL
    );
}




function validarTexto(v,campo){
  if (!v || String(v).trim()==='') throw new Error(`${campo} obligatorio`);
}

function validarNumeroNoNegativo(v, campo){
  if (!isFinite(Number(v)) || Number(v) < 0) {
    throw new Error(`${campo} inválido`);
  }
}
