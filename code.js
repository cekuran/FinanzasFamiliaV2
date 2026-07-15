/** @OnlyCurrentDoc */
// Emerald Ledger — backend Google Apps Script (Sheets + WebApp).
// Multi-usuario: cada fila lleva owner_email; CRUD siempre filtra por usuario activo.

const SCHEMA = {
  Cuentas:       ['owner_email','id','nombre','tipo','moneda','icono','saldo_inicial','orden','oculta','fecha_creacion'],
  Categorias:    ['owner_email','id','nombre','color','tipo','orden'],
  Transacciones: ['owner_email','id','fecha','tipo','importe','moneda','cuenta_id','cuenta_destino_id','importe_destino','ratio_conversion','categoria_id','descripcion','estado','recurrente_id','fecha_pago','conciliada_con','notas','fecha_creacion'],
  Recurrentes:   ['owner_email','id','plantilla','ultima_generacion','activa'],
  Presupuestos:  ['owner_email','id','anio','mes','categoria_id','importe_esperado'],
  Conciliaciones:['owner_email','id','fecha','cuenta_id','saldo_sistema','saldo_banco','diferencia','notas'],
  TiposCambio:   ['owner_email','id','fecha','base','destino','ratio']
};

const HOJAS = Object.keys(SCHEMA);

const SEMILLA = {
  Cuentas: [
    { nombre: 'BBVA Nómina', tipo: 'activo', moneda: 'EUR', icono: 'account_balance', saldo_inicial: 0, orden: 1 },
    { nombre: 'Caja de Ahorro', tipo: 'activo', moneda: 'EUR', icono: 'savings', saldo_inicial: 0, orden: 2 },
    { nombre: 'Tarjeta Visa', tipo: 'pasivo', moneda: 'EUR', icono: 'credit_card', saldo_inicial: 0, orden: 3 }
  ],
  Categorias: [
    { nombre: 'Vivienda',       color: '#00613e', tipo: 'gasto',  orden: 1 },
    { nombre: 'Alimentación',   color: '#0c68db', tipo: 'gasto',  orden: 2 },
    { nombre: 'Transporte',     color: '#ba1a1a', tipo: 'gasto',  orden: 3 },
    { nombre: 'Ocio',           color: '#b13c68', tipo: 'gasto',  orden: 4 },
    { nombre: 'Salud',          color: '#d27b1b', tipo: 'gasto',  orden: 5 },
    { nombre: 'Suscripciones',  color: '#0050af', tipo: 'gasto',  orden: 6 },
    { nombre: 'Nómina',         color: '#006c46', tipo: 'ingreso',orden: 7 },
    { nombre: 'Ingresos extra', color: '#107c52', tipo: 'ingreso',orden: 8 }
  ]
};

// ───────── Helpers de sesión y ss ─────────
const PROP_SHEET_ID = 'SHEET_ID';

function ss_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty(PROP_SHEET_ID);
  if (id) {
    try { return SpreadsheetApp.openById(id); }
    catch (e) { /* id inválido, recrear */ }
  }
  // Primera vez (o id corrupto): crear spreadsheet propio para el usuario
  const ss = SpreadsheetApp.create('Emerald Ledger · ' + email_());
  props.setProperty(PROP_SHEET_ID, ss.getId());
  // Mover la hoja "Hoja 1" por defecto al final, queda fuera de la vista
  const porDefecto = ss.getSheets()[0];
  if (porDefecto && ss.getSheets().length === 1) ss.renameSheet(porDefecto, '_log');
  return ss;
}

function resetSheet() {
  PropertiesService.getScriptProperties().deleteProperty(PROP_SHEET_ID);
  ss_();
  return 'ok';
}

function email_() {
  const e = Session.getActiveUser().getEmail();
  return e || ('anon-' + Session.getTemporaryActiveUserKeys().join('')) || 'anon';
}
function uid_(prefixo) { return (prefixo || 'id') + '_' + Utilities.getUuid().slice(0, 8); }
function isoHoy_() { return new Date().toISOString().slice(0, 10); }
function isoAhora_() { return new Date().toISOString(); }

function asegurarHoja(nombre) {
  const ss = ss_();
  let h = ss.getSheetByName(nombre);
  if (!h) {
    h = ss.insertSheet(nombre);
    h.getRange(1, 1, 1, SCHEMA[nombre].length).setValues([SCHEMA[nombre]]).setFontWeight('bold');
    h.setFrozenRows(1);
  } else if (h.getLastRow() === 0) {
    h.getRange(1, 1, 1, SCHEMA[nombre].length).setValues([SCHEMA[nombre]]).setFontWeight('bold');
    h.setFrozenRows(1);
  }
  return h;
}

function asegurarEsquema() {
  HOJAS.forEach(asegurarHoja);
}

function leerHoja(nombre) {
  asegurarHoja(nombre);
  const h = ss_().getSheetByName(nombre);
  const valores = h.getDataRange().getValues();
  if (valores.length < 2) return [];
  const cab = valores[0];
  return valores.slice(1).map(fila => {
    const o = {};
    cab.forEach((k, i) => (o[k] = fila[i]));
    return o;
  });
}

function escribirHoja(nombre, filas) {
  const h = ss_().getSheetByName(nombre);
  const cab = SCHEMA[nombre];
  h.clearContents();
  const matriz = [cab].concat(filas.map(f => cab.map(k => f[k] != null ? f[k] : '')));
  if (matriz.length) h.getRange(1, 1, matriz.length, cab.length).setValues(matriz);
  h.setFrozenRows(1);
}

function upsertFila(nombre, fila) {
  const datos = leerHoja(nombre);
  const idx = datos.findIndex(f => f.id === fila.id && f.owner_email === fila.owner_email);
  if (idx >= 0) datos[idx] = Object.assign({}, datos[idx], fila);
  else datos.push(fila);
  escribirHoja(nombre, datos);
  return fila;
}

function eliminarFila(nombre, owner, id) {
  const datos = leerHoja(nombre).filter(f => !(f.owner_email === owner && f.id === id));
  escribirHoja(nombre, datos);
}

function sembrar(owner) {
  HOJAS.forEach(asegurarHoja);
  // Si ya tiene algo, no duplicar
  if (leerHoja('Cuentas').some(c => c.owner_email === owner)) return;

  const cuentas = SEMILLA.Cuentas.map(c => Object.assign({
    owner_email: owner, id: uid_('cta'), oculta: false, fecha_creacion: isoAhora_()
  }, c));
  escribirHoja('Cuentas', leerHoja('Cuentas').concat(cuentas));

  const cats = SEMILLA.Categorias.map(c => Object.assign({
    owner_email: owner, id: uid_('cat')
  }, c));
  escribirHoja('Categorias', leerHoja('Categorias').concat(cats));
}

// ───────── Bootstrap público ─────────
function doGet() {
  const owner = email_();
  sembrar(owner);
  generarRecurrentesPendientes_(owner, new Date());
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('Emerald Ledger')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function bootstrap() {
  const owner = email_();
  sembrar(owner);
  generarRecurrentesPendientes_(owner, new Date());
  return {
    sesion: { email: owner },
    cuentas: obtenerCuentas(),
    categorias: obtenerCategorias(),
    transacciones: obtenerTransacciones({}),
    recurrentes: obtenerRecurrentes(),
    presupuestos: obtenerPresupuestos(),
    resumen: obtenerResumen()
  };
}

function incluir(html) { return HtmlService.createHtmlOutputFromFile(html).getContent(); }

// ───────── Cuentas ─────────
function obtenerCuentas() {
  const owner = email_();
  const cuentas = leerHoja('Cuentas').filter(c => c.owner_email === owner && !c.oculta);
  const txs = leerHoja('Transacciones').filter(t => t.owner_email === owner);
  const tipos = leerHoja('TiposCambio').filter(t => t.owner_email === owner);

  return cuentas.map(c => {
    const saldosPorMes = {}; // 'YYYY-MM' -> acumulado
    let base = Number(c.saldo_inicial || 0);
    txs.filter(t => t.cuenta_id === c.id || t.cuenta_destino_id === c.id)
      .sort((a, b) => String(a.fecha).localeCompare(String(b.fecha)))
      .forEach(t => {
        const mes = String(t.fecha).slice(0, 7);
        if (t.cuenta_id === c.id) {
          const signo = t.tipo === 'ingreso' ? 1 : (t.tipo === 'gasto' ? -1 : (t.tipo === 'transferencia' ? -1 : 0));
          base += signo * Number(t.importe || 0);
        } else if (t.tipo === 'transferencia') {
          // destino de transferencia: entra el importe destino (si hay conversión) o el mismo importe
          const importe = Number(t.importe_destino || t.importe || 0);
          base += importe;
        }
        saldosPorMes[mes] = base;
      });
    const hoy = new Date();
    const evolucion = [];
    for (let i = 11; i >= 0; i--) {
      const d = new Date(hoy.getFullYear(), hoy.getMonth() - i, 1);
      const k = Utilities.formatDate(d, ss_().getSpreadsheetTimeZone(), 'yyyy-MM');
      evolucion.push({ mes: k, saldo: saldosPorMes[k] != null ? saldosPorMes[k] : base });
    }
    return {
      id: c.id, nombre: c.nombre, tipo: c.tipo, moneda: c.moneda, icono: c.icono,
      saldo_inicial: Number(c.saldo_inicial || 0),
      saldo: base, evolucion: evolucion
    };
  });
}

function guardarCuenta(cuenta) {
  const owner = email_();
  if (!cuenta || !cuenta.nombre) throw new Error('Nombre de cuenta obligatorio');
  if (!['activo', 'pasivo'].includes(cuenta.tipo)) throw new Error('Tipo inválido');
  if (!cuenta.moneda) throw new Error('Moneda obligatoria');
  const fila = Object.assign({
    owner_email: owner,
    id: cuenta.id || uid_('cta'),
    nombre: String(cuenta.nombre).trim(),
    tipo: cuenta.tipo,
    moneda: String(cuenta.moneda).toUpperCase(),
    icono: cuenta.icono || 'account_balance_wallet',
    saldo_inicial: Number(cuenta.saldo_inicial || 0),
    orden: cuenta.orden || 99,
    oculta: !!cuenta.oculta,
    fecha_creacion: isoAhora_()
  }, cuenta);
  upsertFila('Cuentas', fila);
  return obtenerCuentas();
}

function eliminarCuenta(id) {
  const owner = email_();
  const txs = leerHoja('Transacciones').filter(t => t.owner_email === owner && (t.cuenta_id === id || t.cuenta_destino_id === id));
  if (txs.length) throw new Error('La cuenta tiene ' + txs.length + ' movimientos. Reasígnalos o elimínalos primero.');
  eliminarFila('Cuentas', owner, id);
  return obtenerCuentas();
}

// ───────── Categorías ─────────
function obtenerCategorias() {
  const owner = email_();
  return leerHoja('Categorias').filter(c => c.owner_email === owner).sort((a, b) => a.orden - b.orden);
}

function guardarCategoria(cat) {
  const owner = email_();
  const fila = Object.assign({
    owner_email: owner, id: cat.id || uid_('cat'),
    nombre: String(cat.nombre).trim(), color: cat.color || '#00613e',
    tipo: cat.tipo || 'gasto', orden: cat.orden || 99
  }, cat);
  upsertFila('Categorias', fila);
  return obtenerCategorias();
}

function eliminarCategoria(id) {
  const owner = email_();
  eliminarFila('Categorias', owner, id);
  return obtenerCategorias();
}

// ───────── Transacciones ─────────
function parseFecha(s) {
  if (!s) return null;
  if (Object.prototype.toString.call(s) === '[object Date]') return s;
  const d = new Date(s);
  return isNaN(d) ? null : d;
}
function iso_(d) { return Utilities.formatDate(new Date(d), ss_().getSpreadsheetTimeZone(), 'yyyy-MM-dd'); }

function obtenerTransacciones(filtro) {
  const owner = email_();
  filtro = filtro || {};
  let txs = leerHoja('Transacciones').filter(t => t.owner_email === owner);
  if (filtro.cuenta_id) txs = txs.filter(t => t.cuenta_id === filtro.cuenta_id || t.cuenta_destino_id === filtro.cuenta_id);
  if (filtro.tipo) txs = txs.filter(t => t.tipo === filtro.tipo);
  if (filtro.estado) txs = txs.filter(t => t.estado === filtro.estado);
  if (filtro.categoria_id) txs = txs.filter(t => t.categoria_id === filtro.categoria_id);
  if (filtro.desde) txs = txs.filter(t => String(t.fecha) >= filtro.desde);
  if (filtro.hasta) txs = txs.filter(t => String(t.fecha) <= filtro.hasta);
  if (filtro.q) {
    const q = String(filtro.q).toLowerCase();
    txs = txs.filter(t => String(t.descripcion || '').toLowerCase().includes(q) || String(t.notas || '').toLowerCase().includes(q));
  }
  return txs
    .map(t => Object.assign({}, t, {
      importe: Number(t.importe || 0),
      importe_destino: Number(t.importe_destino || 0),
      ratio_conversion: Number(t.ratio_conversion || 0)
    }))
    .sort((a, b) => String(b.fecha).localeCompare(String(a.fecha)));
}

function guardarTransaccion(tx) {
  const owner = email_();
  if (!tx) throw new Error('Transacción vacía');
  const tipo = tx.tipo;
  if (!['gasto', 'ingreso', 'transferencia'].includes(tipo)) throw new Error('Tipo inválido');
  if (!tx.cuenta_id) throw new Error('Cuenta obligatoria');
  if (!(Number(tx.importe) > 0)) throw new Error('Importe debe ser > 0');
  const fecha = parseFecha(tx.fecha);
  if (!fecha) throw new Error('Fecha inválida');
  const fila = {
    owner_email: owner,
    id: tx.id || uid_('tx'),
    fecha: iso_(fecha),
    tipo,
    importe: Number(tx.importe),
    moneda: String(tx.moneda || 'EUR').toUpperCase(),
    cuenta_id: tx.cuenta_id,
    cuenta_destino_id: tx.cuenta_destino_id || '',
    importe_destino: tx.importe_destino ? Number(tx.importe_destino) : '',
    ratio_conversion: tx.ratio_conversion ? Number(tx.ratio_conversion) : '',
    categoria_id: tx.categoria_id || '',
    descripcion: String(tx.descripcion || '').trim(),
    estado: tx.estado || 'pendiente',
    recurrente_id: tx.recurrente_id || '',
    fecha_pago: tx.fecha_pago ? iso_(tx.fecha_pago) : '',
    conciliada_con: tx.conciliada_con || '',
    notas: tx.notas || '',
    fecha_creacion: tx.fecha_creacion || isoAhora_()
  };
  if (fila.tipo === 'transferencia' && !fila.cuenta_destino_id) throw new Error('Cuenta destino obligatoria en transferencia');
  if (fila.tipo === 'transferencia' && fila.ratio_conversion && !(fila.importe_destino > 0)) {
    throw new Error('Falta importe destino o ratio de conversión');
  }
  upsertFila('Transacciones', fila);
  if (tx.recurrente_plantilla) upsertRecurrenteBase_(owner, tx);
  return fila;
}

function eliminarTransaccion(id) {
  const owner = email_();
  eliminarFila('Transacciones', owner, id);
  return { ok: true };
}

// ───────── Recurrentes ─────────
function obtenerRecurrentes() {
  const owner = email_();
  return leerHoja('Recurrentes').filter(r => r.owner_email === owner).map(r => ({
    id: r.id, plantilla: r.plantilla, ultima_generacion: r.ultima_generacion, activa: r.activa === true || r.activa === 'true'
  }));
}

function guardarRecurrente(rec) {
  const owner = email_();
  if (!rec.plantilla) throw new Error('Falta plantilla');
  const fila = {
    owner_email: owner,
    id: rec.id || uid_('rec'),
    plantilla: typeof rec.plantilla === 'string' ? rec.plantilla : JSON.stringify(rec.plantilla),
    ultima_generacion: rec.ultima_generacion || '',
    activa: rec.activa !== false
  };
  upsertFila('Recurrentes', fila);
  return obtenerRecurrentes();
}

function eliminarRecurrente(id) {
  const owner = email_();
  eliminarFila('Recurrentes', owner, id);
  return obtenerRecurrentes();
}

function upsertRecurrenteBase_(owner, tx) {
  const plantilla = {
    tipo: tx.tipo, importe: Number(tx.importe), cuenta_id: tx.cuenta_id,
    categoria_id: tx.categoria_id || '', descripcion: tx.descripcion || '',
    periodicidad: tx.recurrente_periodicidad || 'mensual', dia_mes: tx.recurrente_dia || Number(String(tx.fecha).slice(8, 10)),
    inicio: iso_(tx.fecha), fin: tx.recurrente_fin || ''
  };
  const fila = {
    owner_email: owner, id: uid_('rec'),
    plantilla: JSON.stringify(plantilla), ultima_generacion: iso_(tx.fecha), activa: true
  };
  // append sin duplicar
  const datos = leerHoja('Recurrentes');
  datos.push(fila);
  escribirHoja('Recurrentes', datos);
}

function generarRecurrentesPendientes_(owner, fechaCorte) {
  const recs = leerHoja('Recurrentes').filter(r => r.owner_email === owner && r.activa);
  const txs = leerHoja('Transacciones').filter(t => t.owner_email === owner);
  let cambios = false;
  recs.forEach(r => {
    const p = JSON.parse(r.plantilla);
    const ultima = r.ultima_generacion ? new Date(r.ultima_generacion) : null;
    const inicio = new Date(p.inicio);
    let cursor = ultima && ultima > inicio ? new Date(ultima) : new Date(inicio);
    cursor.setDate(Number(p.dia_mes || 1));
    while (cursor <= fechaCorte) {
      const isoCursor = iso_(cursor);
      const ya = txs.some(t => t.recurrente_id === r.id && String(t.fecha) === isoCursor);
      if (!ya) {
        txs.push({
          owner_email: owner, id: uid_('tx'),
          fecha: isoCursor, tipo: p.tipo, importe: Number(p.importe), moneda: 'EUR',
          cuenta_id: p.cuenta_id, cuenta_destino_id: '', importe_destino: '', ratio_conversion: '',
          categoria_id: p.categoria_id || '', descripcion: p.descripcion || '',
          estado: 'pendiente', recurrente_id: r.id, fecha_pago: '', conciliada_con: '', notas: '',
          fecha_creacion: isoAhora_()
        });
        cambios = true;
      }
      cursor = siguienteCursor_(cursor, p.periodicidad);
    }
    r.ultima_generacion = iso_(fechaCorte);
  });
  if (cambios) escribirHoja('Transacciones', txs);
  escribirHoja('Recurrentes', leerHoja('Recurrentes').map(r => {
    if (r.owner_email !== owner) return r;
    const nuevo = recs.find(x => x.id === r.id);
    return nuevo || r;
  }));
}

function siguienteCursor_(d, periodicidad) {
  const f = new Date(d);
  if (periodicidad === 'diario') f.setDate(f.getDate() + 1);
  else if (periodicidad === 'semanal') f.setDate(f.getDate() + 7);
  else if (periodicidad === 'anual') f.setFullYear(f.getFullYear() + 1);
  else f.setMonth(f.getMonth() + 1);
  return f;
}

function generarRecurrentesPendientes(fechaCorte) {
  generarRecurrentesPendientes_(email_(), parseFecha(fechaCorte) || new Date());
  return obtenerTransacciones({});
}

// ───────── Presupuestos ─────────
function obtenerPresupuestos(anio, mes) {
  const owner = email_();
  let ps = leerHoja('Presupuestos').filter(p => p.owner_email === owner);
  if (anio) ps = ps.filter(p => Number(p.anio) === Number(anio));
  if (mes) ps = ps.filter(p => Number(p.mes) === Number(mes));
  return ps;
}

function guardarPresupuesto(p) {
  const owner = email_();
  if (!p.anio || !p.mes) throw new Error('Año y mes requeridos');
  const fila = {
    owner_email: owner, id: p.id || uid_('ppto'),
    anio: Number(p.anio), mes: Number(p.mes),
    categoria_id: p.categoria_id || '', importe_esperado: Number(p.importe_esperado || 0)
  };
  upsertFila('Presupuestos', fila);
  return obtenerPresupuestos(p.anio, p.mes);
}

function eliminarPresupuesto(id) {
  const owner = email_();
  eliminarFila('Presupuestos', owner, id);
  return obtenerPresupuestos();
}

// ───────── Conciliación ─────────
function conciliar(cuenta_id, saldo_banco, fecha) {
  const owner = email_();
  const cuentas = leerHoja('Cuentas');
  const cuenta = cuentas.find(c => c.owner_email === owner && c.id === cuenta_id);
  if (!cuenta) throw new Error('Cuenta no encontrada');
  const txs = leerHoja('Transacciones').filter(t => t.owner_email === owner && (t.cuenta_id === cuenta_id || t.cuenta_destino_id === cuenta_id));
  let saldo = Number(cuenta.saldo_inicial || 0);
  txs.sort((a, b) => String(a.fecha).localeCompare(String(b.fecha))).forEach(t => {
    if (t.estado === 'conciliado') return;
    if (t.cuenta_id === cuenta_id) {
      const signo = t.tipo === 'ingreso' ? 1 : (t.tipo === 'gasto' ? -1 : (t.tipo === 'transferencia' ? -1 : 0));
      saldo += signo * Number(t.importe || 0);
    } else if (t.tipo === 'transferencia') {
      saldo += Number(t.importe_destino || t.importe || 0);
    }
  });
  const sistema = saldo;
  const banco = Number(saldo_banco);
  const diferencia = +(banco - sistema).toFixed(2);
  // Marcar movimientos pendientes como conciliados
  const nuevasTxs = txs.map(t => Object.assign({}, t, { estado: 'conciliado', fecha_pago: fecha || isoHoy_() }));
  const todas = leerHoja('Transacciones').map(t => {
    const nueva = nuevasTxs.find(n => n.id === t.id);
    return nueva || t;
  });
  escribirHoja('Transacciones', todas);
  // Registrar evento
  const eventos = leerHoja('Conciliaciones');
  eventos.push({
    owner_email: owner, id: uid_('con'),
    fecha: fecha || isoHoy_(), cuenta_id,
    saldo_sistema: sistema, saldo_banco: banco, diferencia, notas: ''
  });
  escribirHoja('Conciliaciones', eventos);
  return { sistema, banco, diferencia, fecha: fecha || isoHoy_() };
}

// ───────── Resumen y evolución ─────────
function obtenerResumen(anio, mes) {
  const owner = email_();
  const a = anio || new Date().getFullYear();
  const m = mes || (new Date().getMonth() + 1);
  const txs = leerHoja('Transacciones').filter(t => t.owner_email === owner);
  const enMes = txs.filter(t => {
    const f = new Date(t.fecha);
    return f.getFullYear() === Number(a) && (f.getMonth() + 1) === Number(m);
  });
  const ingresos = enMes.filter(t => t.tipo === 'ingreso').reduce((s, t) => s + Number(t.importe || 0), 0);
  const gastos = enMes.filter(t => t.tipo === 'gasto').reduce((s, t) => s + Number(t.importe || 0), 0);
  const pendiente = enMes.filter(t => t.estado === 'pendiente').length;
  const vencido = enMes.filter(t => t.estado === 'vencido').length;
  // Evolución últimos 12 meses
  const evol = [];
  for (let i = 11; i >= 0; i--) {
    const d = new Date(a, m - 1 - i, 1);
    const k = Utilities.formatDate(d, ss_().getSpreadsheetTimeZone(), 'yyyy-MM');
    const en = txs.filter(t => String(t.fecha).slice(0, 7) === k);
    evol.push({
      mes: k,
      ingresos: en.filter(t => t.tipo === 'ingreso').reduce((s, t) => s + Number(t.importe || 0), 0),
      gastos: en.filter(t => t.tipo === 'gasto').reduce((s, t) => s + Number(t.importe || 0), 0)
    });
  }
  // Próximos recurrentes
  const recs = leerHoja('Recurrentes').filter(r => r.owner_email === owner && r.activa);
  const proximos = recs.map(r => {
    const p = JSON.parse(r.plantilla);
    return { id: r.id, descripcion: p.descripcion, importe: Number(p.importe), dia_mes: p.dia_mes, periodicidad: p.periodicidad };
  });
  return { anio: a, mes: m, ingresos, gastos, neto: ingresos - gastos, pendiente, vencido, evol, proximos };
}

function obtenerCategoriasResumen(anio, mes) {
  const owner = email_();
  const a = anio || new Date().getFullYear();
  const m = mes || (new Date().getMonth() + 1);
  const txs = leerHoja('Transacciones').filter(t => t.owner_email === owner && t.tipo === 'gasto');
  const ps = leerHoja('Presupuestos').filter(p => p.owner_email === owner && Number(p.anio) === Number(a) && Number(p.mes) === Number(m));
  const cats = obtenerCategorias();
  return cats.map(c => {
    const esperado = ps.filter(p => p.categoria_id === c.id).reduce((s, p) => s + Number(p.importe_esperado || 0), 0);
    const real = txs.filter(t => {
      if (t.categoria_id !== c.id) return false;
      const f = new Date(t.fecha);
      return f.getFullYear() === Number(a) && (f.getMonth() + 1) === Number(m);
    }).reduce((s, t) => s + Number(t.importe || 0), 0);
    return { id: c.id, nombre: c.nombre, color: c.color, esperado, real, diferencia: real - esperado };
  });
}

function guardarTipoCambio(base, destino, ratio) {
  const owner = email_();
  if (!(ratio > 0)) throw new Error('Ratio debe ser > 0');
  const fila = {
    owner_email: owner, id: uid_('tc'),
    fecha: isoHoy_(), base: String(base).toUpperCase(), destino: String(destino).toUpperCase(),
    ratio: Number(ratio)
  };
  upsertFila('TiposCambio', fila);
  return fila;
}

// ───────── Self-test mínimo ─────────
function __selfTest() {
  const owner = email_();
  sembrar(owner);
  const cuentas = obtenerCuentas();
  const cat = obtenerCategorias();
  if (!cuentas.length) throw new Error('Sin cuentas');
  if (!cat.length) throw new Error('Sin categorías');
  const t = guardarTransaccion({
    tipo: 'gasto', importe: 12.5, cuenta_id: cuentas[0].id, categoria_id: cat[0].id,
    descripcion: 'self-test', fecha: isoHoy_()
  });
  if (!t.id) throw new Error('Fallo al guardar transacción');
  eliminarTransaccion(t.id);
  const conciliado = conciliar(cuentas[0].id, cuentas[0].saldo || 0, isoHoy_());
  if (typeof conciliado.diferencia !== 'number') throw new Error('Fallo en conciliación');
  return 'ok ' + cuentas.length + ' cuentas, ' + cat.length + ' categorías, diferencia=' + conciliado.diferencia;
}
