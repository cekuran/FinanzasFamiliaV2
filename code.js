/** @OnlyCurrentDoc */
// Finanzas Familia — backend Google Apps Script (Sheets + WebApp).
// Multi-usuario: cada fila lleva owner_email; CRUD siempre filtra por usuario activo.

const SCHEMA = {
  Cuentas:       ['owner_email','id','parent_id','nombre','tipo','moneda','icono','saldo_inicial','orden','oculta','fecha_creacion'],
  Categorias:    ['owner_email','id','nombre','color','icono','tipo','orden'],
  Transacciones: ['owner_email','id','fecha','tipo','importe','moneda','cuenta_id','subcuenta_id','cuenta_destino_id','subcuenta_destino_id','importe_destino','ratio_conversion','reparto_destino','categoria_id','descripcion','estado','recurrente_id','fecha_pago','conciliada_con','notas','fecha_creacion'],
  Recurrentes:   ['owner_email','id','plantilla','ultima_generacion','activa'],
  Presupuestos:  ['owner_email','id','anio','mes','categoria_id','importe_esperado'],
  Conciliaciones:['owner_email','id','fecha','cuenta_id','saldo_sistema','saldo_banco','diferencia','notas'],
  TiposCambio:   ['owner_email','id','fecha','base','destino','ratio']
};

const HOJAS = Object.keys(SCHEMA);

const SEMILLA = {
  Cuentas: [
    {
      nombre: 'BBVA Nómina', tipo: 'activo', moneda: 'EUR', icono: 'account_balance', saldo_inicial: 0, orden: 1,
      subcuentas: [
        { nombre: 'Gastos diarios', saldo_inicial: 0, orden: 1 },
        { nombre: 'Recibos', saldo_inicial: 0, orden: 2 }
      ]
    },
    { nombre: 'Caja de Ahorro', tipo: 'activo', moneda: 'EUR', icono: 'savings', saldo_inicial: 0, orden: 2 },
    { nombre: 'Tarjeta Visa', tipo: 'pasivo', moneda: 'EUR', icono: 'credit_card', saldo_inicial: 0, orden: 3 }
  ],
  Categorias: [
    { nombre: 'Vivienda',       color: '#00613e', icono: 'home',            tipo: 'gasto',  orden: 1 },
    { nombre: 'Alimentación',   color: '#0c68db', icono: 'shopping_cart',   tipo: 'gasto',  orden: 2 },
    { nombre: 'Transporte',     color: '#ba1a1a', icono: 'directions_car',  tipo: 'gasto',  orden: 3 },
    { nombre: 'Ocio',           color: '#b13c68', icono: 'movie',           tipo: 'gasto',  orden: 4 },
    { nombre: 'Salud',          color: '#d27b1b', icono: 'fitness_center',  tipo: 'gasto',  orden: 5 },
    { nombre: 'Suscripciones',  color: '#0050af', icono: 'subscriptions',   tipo: 'gasto',  orden: 6 },
    { nombre: 'Nómina',         color: '#006c46', icono: 'payments',        tipo: 'ingreso',orden: 7 },
    { nombre: 'Ingresos extra', color: '#107c52', icono: 'work',            tipo: 'ingreso',orden: 8 }
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
  const ss = SpreadsheetApp.create('Finanzas Familia · ' + email_());
  props.setProperty(PROP_SHEET_ID, ss.getId());
  // Mover la hoja "Hoja 1" por defecto al final, queda fuera de la vista
  const porDefecto = ss.getSheets()[0];
  if (porDefecto && ss.getSheets().length === 1) porDefecto.setName('_log');
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

// Añade columnas nuevas a hojas existentes sin tocar filas. Idempotente.
function migrarEsquema() {
  HOJAS.forEach(asegurarHoja);
  Object.keys(SCHEMA).forEach(nombre => {
    const h = ss_().getSheetByName(nombre);
    if (!h) return;
    const cab = h.getRange(1, 1, 1, h.getLastColumn()).getValues()[0];
    SCHEMA[nombre].forEach((col, i) => {
      if (!cab.includes(col)) h.getRange(1, Math.max(h.getLastColumn(), i) + 1).setValue(col);
    });
  });
}

// Sheets convierte automáticamente las fechas en formato texto ('yyyy-MM-dd')
// a objetos Date. Al leer las normalizamos de vuelta a string para que
// String(fecha).slice(0,7) y las comparaciones de fechas funcionen igual
// que cuando se escribieron.
function normalizarValor_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) {
    return Utilities.formatDate(v, ss_().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  }
  return v;
}

function leerHoja(nombre) {
  asegurarHoja(nombre);
  const h = ss_().getSheetByName(nombre);
  const valores = h.getDataRange().getValues();
  if (valores.length < 2) return [];
  const cab = valores[0];
  return valores.slice(1).map(fila => {
    const o = {};
    cab.forEach((k, i) => (o[k] = normalizarValor_(fila[i])));
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

  const cuentas = SEMILLA.Cuentas.flatMap(c => {
    const parent = Object.assign({
      owner_email: owner, id: uid_('cta'), parent_id: '', oculta: false, fecha_creacion: isoAhora_()
    }, c);
    return [parent].concat((c.subcuentas || []).map((s, i) => Object.assign({
      owner_email: owner, id: uid_('cta'), parent_id: parent.id,
      tipo: parent.tipo, moneda: parent.moneda, icono: 'savings',
      saldo_inicial: 0, orden: i + 1, oculta: false, fecha_creacion: isoAhora_()
    }, s)));
  });
  escribirHoja('Cuentas', leerHoja('Cuentas').concat(cuentas));

  const cats = SEMILLA.Categorias.map(c => Object.assign({
    owner_email: owner, id: uid_('cat')
  }, c));
  escribirHoja('Categorias', leerHoja('Categorias').concat(cats));
}

// ───────── Bootstrap público ─────────
function doGet() {
  const owner = email_();
  migrarEsquema();
  sembrar(owner);
  normalizarSubcuentasHuerfanas_(owner);
  generarRecurrentesPendientes_(owner, new Date());
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('Finanzas Familia')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function bootstrap() {
  const base = bootstrapBase();
  base.transacciones = obtenerTransacciones({});
  return base;
}

// Bootstrap ligero para carga progresiva en el cliente.
// Devuelve todo menos transacciones para que la UI pinte antes.
function bootstrapBase() {
  const owner = email_();
  migrarEsquema();
  sembrar(owner);
  normalizarSubcuentasHuerfanas_(owner);
  generarRecurrentesPendientes_(owner, new Date());
  return {
    sesion: { email: owner },
    version: PropertiesService.getScriptProperties().getProperty('VERSION') || '',
    cuentas: obtenerCuentas(),
    categorias: obtenerCategorias(),
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
  const fmtMes = d => Utilities.formatDate(d, ss_().getSpreadsheetTimeZone(), 'yyyy-MM');

  const top = cuentas.filter(c => !c.parent_id);
  return top.map(c => {
    const subs = cuentas.filter(x => x.parent_id === c.id);
    const subIds = new Set(subs.map(x => x.id));
    const parentInitial = Number(c.saldo_inicial || 0);

    // Txs del padre: golpean esta cuenta y NO están asignadas a una subcuenta
    // propia de esta cuenta (ni como origen ni como destino de transferencia).
    // Una subcuenta "huérfana" (de otra cuenta) se trata como movimiento del
    // padre, no se atribuye a la cuenta ajena.
    const parentTxs = txs.filter(t => {
      const origen = t.cuenta_id === c.id;
      const destino = t.cuenta_destino_id === c.id;
      if (!origen && !destino) return false;
      if (origen && t.subcuenta_id && subIds.has(t.subcuenta_id)) return false;
      if (destino) {
        // Cualquier subcuenta del reparto que sea hija de este padre cuenta
        // como sub-tx (no como tx del padre).
        const reparto = parseRepartoDestino_(t.reparto_destino);
        const subDest = reparto.length
          ? reparto.map(r => r.subcuenta_id)
          : (t.subcuenta_destino_id ? [t.subcuenta_destino_id] : []);
        if (subDest.some(id => subIds.has(id))) return false;
      }
      return true;
    });

    const subcuentas = subs.map(s => {
      const subInitial = Number(s.saldo_inicial || 0);
      // Cuenta como movimiento de la subcuenta cuando es su origen (cuenta_id)
      // o el destino de una transferencia (cuenta_destino_id o reparto_destino).
      // En ambos casos la cuenta implicada debe ser este mismo padre.
      const subDelta = txs
        .filter(t => {
          const origenOk = t.subcuenta_id === s.id && t.cuenta_id === c.id;
          if (origenOk) return true;
          if (t.cuenta_destino_id !== c.id || t.tipo !== 'transferencia') return false;
          const reparto = parseRepartoDestino_(t.reparto_destino);
          if (reparto.length) return reparto.some(r => r.subcuenta_id === s.id);
          return t.subcuenta_destino_id === s.id;
        })
        .reduce((sum, t) => sum + deltaSubcuenta_(t, s.id), 0);
      return {
        id: s.id, nombre: s.nombre, parent_id: c.id,
        saldo_inicial: subInitial,
        saldo: subInitial + subDelta,
        orden: Number(s.orden || 99)
      };
    }).sort((a, b) => a.orden - b.orden);

    // Modelo de partición: las subcuentas reparten el saldo del padre, no lo
    // aumentan. Su saldo_inicial ya forma parte del saldo_inicial del padre, así
    // que al total del padre solo se le suma la variación por movimientos de
    // subcuenta (subDelta), nunca su saldo inicial (evita doble conteo).
    const subDeltaTotal = subcuentas.reduce((s, x) => s + (x.saldo - x.saldo_inicial), 0);
    const parentDelta = parentTxs.reduce((s, t) => s + deltaCuenta_(t, c.id), 0);

    const hoy = new Date();
    const evolucion = [];
    for (let i = 11; i >= 0; i--) {
      const d = new Date(hoy.getFullYear(), hoy.getMonth() - i, 1);
      const k = fmtMes(d);
      const parentDeltaK = parentTxs
        .filter(t => String(t.fecha).slice(0, 7) <= k)
        .reduce((s, t) => s + deltaCuenta_(t, c.id), 0);
      const subDeltaK = subcuentas.reduce((s, sub) => {
        const dK = txs
          .filter(t => {
            const fechaOk = String(t.fecha).slice(0, 7) <= k;
            if (!fechaOk) return false;
            if (t.subcuenta_id === sub.id && t.cuenta_id === c.id) return true;
            if (t.cuenta_destino_id !== c.id || t.tipo !== 'transferencia') return false;
            const reparto = parseRepartoDestino_(t.reparto_destino);
            if (reparto.length) return reparto.some(r => r.subcuenta_id === sub.id);
            return t.subcuenta_destino_id === sub.id;
          })
          .reduce((sd, t) => sd + deltaSubcuenta_(t, sub.id), 0);
        return s + dK;
      }, 0);
      evolucion.push({ mes: k, saldo: parentInitial + parentDeltaK + subDeltaK });
    }

    return {
      id: c.id, nombre: c.nombre, tipo: c.tipo, moneda: c.moneda, icono: c.icono,
      saldo_inicial: parentInitial,
      saldo: parentInitial + parentDelta + subDeltaTotal,
      evolucion,
      subcuentas
    };
  });
}

// Variación de saldo que aporta una transacción a una cuenta concreta.
function deltaCuenta_(t, cuentaId) {
  if (t.cuenta_id === cuentaId) {
    if (t.tipo === 'ingreso') return Number(t.importe || 0);
    if (t.tipo === 'gasto' || t.tipo === 'transferencia') return -Number(t.importe || 0);
    return 0;
  }
  if (t.cuenta_destino_id === cuentaId && t.tipo === 'transferencia') {
    return Number(t.importe_destino || t.importe || 0);
  }
  return 0;
}

// Variación de saldo que aporta una transacción a una subcuenta concreta.
// Gasto/ingreso usan subcuenta_id. Las transferencias restan de la subcuenta
// de origen (subcuenta_id) y suman a la subcuenta de destino: si hay
// reparto_destino (JSON con varias subcuentas) se reparte; si no, se usa el
// legado subcuenta_destino_id único.
function deltaSubcuenta_(t, subId) {
  let d = 0;
  if (t.subcuenta_id === subId) {
    if (t.tipo === 'ingreso') d += Number(t.importe || 0);
    else if (t.tipo === 'gasto' || t.tipo === 'transferencia') d += -Number(t.importe || 0);
  }
  if (t.tipo === 'transferencia') {
    const reparto = parseRepartoDestino_(t.reparto_destino);
    if (reparto.length) {
      const r = reparto.find(x => x.subcuenta_id === subId);
      if (r) d += Number(r.importe || 0);
    } else if (t.subcuenta_destino_id === subId) {
      d += Number(t.importe_destino || t.importe || 0);
    }
  }
  return d;
}

// Sanea referencias subcuenta_id huérfanas: si una transacción apunta a una
// subcuenta que no pertenece a su propia cuenta (datos heredados o corruptos),
// se borra la subcuenta_id para que el movimiento quede solo en su cuenta.
// Idempotente: solo escribe cuando encuentra algo que corregir.
function normalizarSubcuentasHuerfanas_(owner) {
  const cuentas = leerHoja('Cuentas').filter(c => c.owner_email === owner);
  const validas = new Set(
    cuentas.filter(c => c.parent_id).map(c => c.parent_id + '|' + c.id)
  );
  let cambios = false;
  const txs = leerHoja('Transacciones').map(t => {
    if (t.owner_email !== owner) return t;
    let nuevo = t;
    if (t.subcuenta_id && !validas.has(t.cuenta_id + '|' + t.subcuenta_id)) {
      cambios = true;
      nuevo = Object.assign({}, nuevo, { subcuenta_id: '' });
    }
    if (t.subcuenta_destino_id && !validas.has(t.cuenta_destino_id + '|' + t.subcuenta_destino_id)) {
      cambios = true;
      nuevo = Object.assign({}, nuevo, { subcuenta_destino_id: '' });
    }
    // Reparto destino: cualquier entrada cuya subcuenta no pertenezca a la
    // cuenta destino se descarta. Si tras limpiar queda vacío o el total no
    // cuadra, se elimina el reparto entero.
    const reparto = parseRepartoDestino_(t.reparto_destino);
    if (reparto.length && t.cuenta_destino_id) {
      const filtrado = reparto.filter(r => validas.has(t.cuenta_destino_id + '|' + r.subcuenta_id));
      const suma = filtrado.reduce((s, r) => s + r.importe, 0);
      const totalEsperado = t.importe_destino ? Number(t.importe_destino) : Number(t.importe || 0);
      const ok = filtrado.length > 0 && Math.abs(suma - totalEsperado) <= 0.01;
      const nuevoJson = ok ? JSON.stringify(filtrado) : '';
      if (nuevoJson !== (t.reparto_destino || '')) {
        cambios = true;
        nuevo = Object.assign({}, nuevo, { reparto_destino: nuevoJson });
      }
    }
    return nuevo;
  });
  if (cambios) escribirHoja('Transacciones', txs);
  return cambios;
}

function guardarCuenta(cuenta) {
  const owner = email_();
  if (!cuenta || !cuenta.nombre) throw new Error('Nombre de cuenta obligatorio');
  const todas = leerHoja('Cuentas');
  let parent = null;
  if (cuenta.parent_id) {
    parent = todas.find(c => c.owner_email === owner && c.id === cuenta.parent_id);
    if (!parent) throw new Error('Cuenta padre no encontrada');
    if (parent.parent_id) throw new Error('Una subcuenta no puede tener subcuentas');
    // Subcuentas heredan tipo y moneda del padre; ignoramos lo que mande el cliente.
    cuenta.tipo = parent.tipo;
    cuenta.moneda = parent.moneda;
  } else {
    if (!['activo', 'pasivo'].includes(cuenta.tipo)) throw new Error('Tipo inválido');
    if (!cuenta.moneda) throw new Error('Moneda obligatoria');
  }
  // Al editar, preservamos metadatos que el cliente no envía (orden, icono y
  // fecha de creación) para no reiniciarlos y romper el orden de subcuentas.
  const existente = cuenta.id ? todas.find(c => c.owner_email === owner && c.id === cuenta.id) : null;
  const fila = {
    owner_email: owner,
    id: cuenta.id || uid_('cta'),
    parent_id: cuenta.parent_id || '',
    nombre: String(cuenta.nombre).trim(),
    tipo: cuenta.tipo,
    moneda: String(cuenta.moneda).toUpperCase(),
    icono: cuenta.icono || (existente && existente.icono) || 'account_balance_wallet',
    saldo_inicial: Number(cuenta.saldo_inicial || 0),
    orden: cuenta.orden || (existente && existente.orden) || 99,
    oculta: !!cuenta.oculta,
    fecha_creacion: (existente && existente.fecha_creacion) || isoAhora_()
  };
  upsertFila('Cuentas', fila);
  return obtenerCuentas();
}

function eliminarCuenta(id) {
  const owner = email_();
  const todas = leerHoja('Cuentas').filter(c => c.owner_email === owner);
  const cuenta = todas.find(c => c.id === id);
  if (!cuenta) throw new Error('Cuenta no encontrada');
  const esSub = !!cuenta.parent_id;
  // ponytail: el chequeo de "tiene subcuentas" solo aplica a top-level.
  // Las subcuentas no admiten anidamiento en el modelo, así que el check
  // nunca debería disparar para ellas (si dispara, hay corrupción de datos,
  // pero bloqueamos igual para que el usuario vea el problema).
  if (!esSub) {
    const tieneSubs = todas.some(c => c.parent_id === id);
    if (tieneSubs) throw new Error('La cuenta tiene subcuentas. Elimínalas primero.');
  }
  const txs = leerHoja('Transacciones').filter(t => t.owner_email === owner && (
    (t.cuenta_id === id || t.cuenta_destino_id === id) ||
    (esSub && t.subcuenta_id === id && t.cuenta_id === cuenta.parent_id) ||
    (esSub && t.subcuenta_destino_id === id && t.cuenta_destino_id === cuenta.parent_id)
  ));
  if (txs.length) throw new Error((esSub ? 'La subcuenta' : 'La cuenta') + ' tiene ' + txs.length + ' movimientos. Reasígnalos o elimínalos primero.');
  eliminarFila('Cuentas', owner, id);
  return obtenerCuentas();
}

function reordenarSubcuentas(parentId, ids) {
  const owner = email_();
  if (!parentId || !Array.isArray(ids)) throw new Error('Parámetros inválidos');
  const todas = leerHoja('Cuentas');
  const padre = todas.find(c => c.owner_email === owner && c.id === parentId);
  if (!padre || padre.parent_id) throw new Error('Cuenta padre inválida');
  ids.forEach(id => {
    if (!todas.some(c => c.owner_email === owner && c.id === id && c.parent_id === parentId)) {
      throw new Error('Subcuenta ' + id + ' no pertenece a la cuenta padre');
    }
  });
  ids.forEach((id, i) => {
    const idx = todas.findIndex(c => c.owner_email === owner && c.id === id);
    todas[idx].orden = i + 1;
  });
  escribirHoja('Cuentas', todas);
  return obtenerCuentas();
}

// ───────── Categorías ─────────
function obtenerCategorias() {
  const owner = email_();
  return leerHoja('Categorias').filter(c => c.owner_email === owner).sort((a, b) => a.orden - b.orden);
}

function guardarCategoria(cat) {
  const owner = email_();
  const existente = cat.id ? leerHoja('Categorias').find(c => c.owner_email === owner && c.id === cat.id) : null;
  const fila = {
    owner_email: owner, id: cat.id || uid_('cat'),
    nombre: String(cat.nombre).trim(), color: cat.color || '#00613e',
    icono: cat.icono || (existente && existente.icono) || 'category',
    tipo: cat.tipo || 'gasto', orden: cat.orden || (existente && existente.orden) || 99
  };
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

// Reparto destino: array de {subcuenta_id, importe} guardado como JSON en
// Transacciones.reparto_destino. Acepta string, array o null; devuelve array
// filtrando entradas inválidas. Vacío o null = reparto a nivel de cuenta.
function parseRepartoDestino_(raw) {
  if (!raw) return [];
  let parsed = raw;
  if (typeof raw === 'string') {
    try { parsed = JSON.parse(raw); } catch (e) { return []; }
  }
  if (!Array.isArray(parsed)) return [];
  return parsed
    .map(r => ({ subcuenta_id: r && r.subcuenta_id, importe: Number(r && r.importe || 0) }))
    .filter(r => r.subcuenta_id && r.importe > 0);
}

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
  let subcuenta_id = '';
  if (tx.subcuenta_id) {
    const sub = leerHoja('Cuentas').find(c => c.owner_email === owner && c.id === tx.subcuenta_id && c.parent_id === tx.cuenta_id);
    if (!sub) throw new Error('Subcuenta no encontrada o no pertenece a la cuenta');
    subcuenta_id = tx.subcuenta_id;
  }
  let subcuenta_destino_id = '';
  let repartoDestinoJson = '';
  if (tipo === 'transferencia' && tx.cuenta_destino_id) {
    const repartoRaw = Array.isArray(tx.reparto_destino) ? tx.reparto_destino : null;
    const cuentasHoja = leerHoja('Cuentas');
    if (repartoRaw && repartoRaw.length) {
      const subValidas = new Set(
        cuentasHoja.filter(c => c.owner_email === owner && c.parent_id === tx.cuenta_destino_id).map(c => c.id)
      );
      const visto = new Set();
      const normalizado = [];
      let suma = 0;
      repartoRaw.forEach(r => {
        const sid = r && r.subcuenta_id;
        if (!sid) throw new Error('Cada destino requiere subcuenta');
        if (!subValidas.has(sid)) throw new Error('Subcuenta destino no pertenece a la cuenta destino');
        if (visto.has(sid)) throw new Error('Subcuenta destino duplicada');
        const imp = Number(r.importe);
        if (!(imp > 0)) throw new Error('Importe de destino debe ser > 0');
        visto.add(sid);
        normalizado.push({ subcuenta_id: sid, importe: imp });
        suma += imp;
      });
      const totalEsperado = tx.importe_destino ? Number(tx.importe_destino) : Number(tx.importe);
      if (Math.abs(suma - totalEsperado) > 0.01) {
        throw new Error('La suma del reparto (' + suma.toFixed(2) + ') no coincide con el importe destino (' + totalEsperado.toFixed(2) + ')');
      }
      repartoDestinoJson = JSON.stringify(normalizado);
      // Backfill del campo legacy para que filtros por subcuenta sigan
      // encontrando la tx mientras conviven datos nuevos y viejos.
      subcuenta_destino_id = normalizado[0].subcuenta_id;
    } else if (tx.subcuenta_destino_id) {
      const subd = cuentasHoja.find(c => c.owner_email === owner && c.id === tx.subcuenta_destino_id && c.parent_id === tx.cuenta_destino_id);
      if (!subd) throw new Error('Subcuenta destino no encontrada o no pertenece a la cuenta destino');
      subcuenta_destino_id = tx.subcuenta_destino_id;
    }
  } else if (tx.subcuenta_destino_id) {
    throw new Error('Subcuenta destino solo aplica a transferencias');
  }
  const fila = {
    owner_email: owner,
    id: tx.id || uid_('tx'),
    fecha: iso_(fecha),
    tipo,
    importe: Number(tx.importe),
    moneda: String(tx.moneda || 'EUR').toUpperCase(),
    cuenta_id: tx.cuenta_id,
    subcuenta_id: subcuenta_id,
    cuenta_destino_id: tx.cuenta_destino_id || '',
    subcuenta_destino_id: subcuenta_destino_id,
    importe_destino: tx.importe_destino ? Number(tx.importe_destino) : '',
    ratio_conversion: tx.ratio_conversion ? Number(tx.ratio_conversion) : '',
    reparto_destino: repartoDestinoJson,
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
  if (fila.tipo === 'transferencia' && fila.cuenta_destino_id === fila.cuenta_id) throw new Error('La cuenta destino no puede ser la misma que la cuenta origen');
  if (fila.tipo === 'transferencia' && fila.ratio_conversion && !(fila.importe_destino > 0)) {
    throw new Error('Falta importe destino o ratio de conversión');
  }
  upsertFila('Transacciones', fila);
  if (tx.recurrente_plantilla) fila.recurrente_id = upsertRecurrenteBase_(owner, tx);
  if (fila.recurrente_id) upsertFila('Transacciones', fila);
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
    subcuenta_id: tx.subcuenta_id || '',
    cuenta_destino_id: tx.cuenta_destino_id || '',
    subcuenta_destino_id: tx.subcuenta_destino_id || '',
    importe_destino: tx.importe_destino ? Number(tx.importe_destino) : '',
    ratio_conversion: tx.ratio_conversion ? Number(tx.ratio_conversion) : '',
    reparto_destino: Array.isArray(tx.reparto_destino) ? JSON.stringify(tx.reparto_destino) : (tx.reparto_destino || ''),
    categoria_id: tx.categoria_id || '', descripcion: tx.descripcion || '',
    periodicidad: tx.recurrente_periodicidad || 'mensual', dia_mes: tx.recurrente_dia || Number(String(tx.fecha).slice(8, 10)),
    inicio: iso_(tx.fecha), fin: tx.recurrente_fin || ''
  };
  const id = uid_('rec');
  const fila = {
    owner_email: owner, id,
    plantilla: JSON.stringify(plantilla), ultima_generacion: iso_(tx.fecha), activa: true
  };
  const datos = leerHoja('Recurrentes');
  datos.push(fila);
  escribirHoja('Recurrentes', datos);
  return id;
}

function generarRecurrentesPendientes_(owner, fechaCorte) {
  const recs = leerHoja('Recurrentes').filter(r => r.owner_email === owner && r.activa);
  const txs = leerHoja('Transacciones').filter(t => t.owner_email === owner);
  let cambios = false;
  recs.forEach(r => {
    try {
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
            cuenta_id: p.cuenta_id,
            subcuenta_id: p.subcuenta_id || '',
            cuenta_destino_id: p.cuenta_destino_id || '',
            subcuenta_destino_id: p.subcuenta_destino_id || '',
            importe_destino: p.importe_destino || '',
            ratio_conversion: p.ratio_conversion || '',
            reparto_destino: p.reparto_destino || '',
            categoria_id: p.categoria_id || '', descripcion: p.descripcion || '',
            estado: 'pendiente', recurrente_id: r.id, fecha_pago: '', conciliada_con: '', notas: '',
            fecha_creacion: isoAhora_()
          });
          cambios = true;
        }
        cursor = siguienteCursor_(cursor, p.periodicidad);
      }
      r.ultima_generacion = iso_(fechaCorte);
    } catch (e) {
      // ponytail: plantilla corrupta no debe romper el bootstrap
      Logger.log('plantilla corrupta ' + r.id + ': ' + e.message);
    }
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
// ponytail: one budget per category, same every month
function obtenerPresupuestos() {
  const owner = email_();
  const all = leerHoja('Presupuestos').filter(p => p.owner_email === owner);
  const seen = new Map();
  all.forEach(p => seen.set(p.categoria_id, p));
  return [...seen.values()];
}

function guardarPresupuesto(p) {
  const owner = email_();
  if (!p.categoria_id) throw new Error('Categoría requerida');
  const fila = {
    owner_email: owner, id: p.id || uid_('ppto'),
    categoria_id: p.categoria_id, importe_esperado: Number(p.importe_esperado || 0)
  };
  upsertFila('Presupuestos', fila);
  return obtenerPresupuestos();
}

function eliminarPresupuesto(id) {
  const owner = email_();
  eliminarFila('Presupuestos', owner, id);
  return obtenerPresupuestos();
}

// ───────── Conciliación ─────────
function conciliar(cuenta_id, saldo_banco, fecha) {
  const owner = email_();
  const cuenta = leerHoja('Cuentas').find(c => c.owner_email === owner && c.id === cuenta_id);
  if (!cuenta) throw new Error('Cuenta no encontrada');
  if (cuenta.parent_id) throw new Error('Solo se pueden conciliar cuentas, no subcuentas');
  // Saldo canónico (mismo cálculo que la UI).
  const cta = obtenerCuentas().find(c => c.id === cuenta_id);
  const sistema = cta ? cta.saldo : 0;
  const banco = Number(saldo_banco);
  const diferencia = +(banco - sistema).toFixed(2);
  const txs = leerHoja('Transacciones').filter(t => t.owner_email === owner && (t.cuenta_id === cuenta_id || t.cuenta_destino_id === cuenta_id));
  const nuevasTxs = txs.map(t => Object.assign({}, t, { estado: 'conciliado', fecha_pago: fecha || isoHoy_() }));
  const todas = leerHoja('Transacciones').map(t => {
    const nueva = nuevasTxs.find(n => n.id === t.id);
    return nueva || t;
  });
  escribirHoja('Transacciones', todas);
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
    // ponytail: usa componentes locales (script tz) — formatDate con ss_tz podía devolver
    // el mes anterior si la hoja estaba en UTC; los slice(0,7) de txs almacenados no matcheaban
    const k = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
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
    try {
      const p = JSON.parse(r.plantilla);
      return { id: r.id, descripcion: p.descripcion, importe: Number(p.importe), dia_mes: p.dia_mes, periodicidad: p.periodicidad };
    } catch (e) {
      return null;
    }
  }).filter(Boolean);
  return { anio: a, mes: m, ingresos, gastos, neto: ingresos - gastos, pendiente, vencido, evol, proximos };
}

function obtenerCategoriasResumen(anio, mes) {
  const owner = email_();
  const a = anio || new Date().getFullYear();
  const m = mes || (new Date().getMonth() + 1);
  const txs = leerHoja('Transacciones').filter(t => t.owner_email === owner && ['gasto', 'transferencia'].includes(t.tipo));
  const ps = leerHoja('Presupuestos').filter(p => p.owner_email === owner);
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

  // ponytail: smoke-test subcuentas + reorder + tx con subcuenta_id
  const parent = cuentas[0];
  const subName = 'self-sub-' + Utilities.getUuid().slice(0, 8);
  const withSub = guardarCuenta({
    parent_id: parent.id, nombre: subName, saldo_inicial: 100
  });
  const subId = withSub.find(c => c.id === parent.id).subcuentas.find(s => s.nombre === subName).id;
  const txSub = guardarTransaccion({
    tipo: 'gasto', importe: 30, cuenta_id: parent.id, subcuenta_id: subId,
    categoria_id: cat[0].id, descripcion: 'self-sub-tx', fecha: isoHoy_()
  });
  const fresh = obtenerCuentas().find(c => c.id === parent.id);
  const sub = (fresh.subcuentas || []).find(s => s.id === subId);
  if (!sub || sub.saldo !== 70) throw new Error('Saldo subcuenta incorrecto: ' + (sub && sub.saldo));
  // Modelo de partición: el saldo_inicial (100) de la subcuenta ya forma parte del
  // saldo del padre, así que crearla no cambia el total; solo el gasto (30) lo baja.
  if (fresh.saldo !== parent.saldo - 30) throw new Error('Saldo padre incorrecto: ' + fresh.saldo);
  // Reorder y segunda alta sin reemplazar la primera
  const anotherName = 'self-sub-' + Utilities.getUuid().slice(0, 8);
  const withAnother = guardarCuenta({
    parent_id: parent.id, nombre: anotherName, saldo_inicial: 0
  });
  const parentWithBoth = withAnother.find(c => c.id === parent.id);
  const anotherId = parentWithBoth.subcuentas.find(s => s.nombre === anotherName).id;
  if (!parentWithBoth.subcuentas.some(s => s.id === subId)) throw new Error('Crear una subcuenta reemplazó la anterior');
  // La validación rechaza mismo origen/destino, así que creamos una cuenta
  // destino separada con sus propias subcuentas para probar transferencias
  // y repartos entre cuentas distintas.
  const destParent = guardarCuenta({
    nombre: 'self-dest-' + Utilities.getUuid().slice(0, 8),
    tipo: 'activo', moneda: 'EUR', saldo_inicial: 0
  }).find(c => c.nombre.startsWith('self-dest-'));
  const destSubName = 'self-dest-sub-' + Utilities.getUuid().slice(0, 8);
  const withDestSub = guardarCuenta({ parent_id: destParent.id, nombre: destSubName, saldo_inicial: 0 });
  const destSubId = withDestSub.find(c => c.id === destParent.id).subcuentas.find(s => s.nombre === destSubName).id;
  const realAntesTransferencia = obtenerCategoriasResumen().find(c => c.id === cat[0].id).real;
  const txTransfer = guardarTransaccion({
    tipo: 'transferencia', importe: 20,
    cuenta_id: parent.id, subcuenta_id: subId,
    cuenta_destino_id: destParent.id, subcuenta_destino_id: destSubId,
    categoria_id: cat[0].id, descripcion: 'self-sub-transfer', fecha: isoHoy_()
  });
  const realDespuesTransferencia = obtenerCategoriasResumen().find(c => c.id === cat[0].id).real;
  if (realDespuesTransferencia !== realAntesTransferencia + 20) throw new Error('Transferencia no sumó al presupuesto: ' + realDespuesTransferencia);
  const targetAfter = obtenerCuentas().find(c => c.id === destParent.id).subcuentas.find(s => s.id === destSubId);
  if (targetAfter.saldo !== 20) throw new Error('Transferencia no acreditó subcuenta destino: ' + targetAfter.saldo);
  const after = reordenarSubcuentas(parent.id, [anotherId, subId]);
  const parentAfter = after.find(c => c.id === parent.id);
  const testOrder = parentAfter.subcuentas.filter(s => s.id === anotherId || s.id === subId);
  if (testOrder.length !== 2 || testOrder[0].id !== anotherId) throw new Error('Reorder falló');

  // Transferencia con reparto a varias subcuentas destino: la suma del reparto
  // debe coincidir con el importe total, y cada subcuenta destino recibe su parte.
  const splitSubAName = 'self-split-a-' + Utilities.getUuid().slice(0, 8);
  const splitSubBName = 'self-split-b-' + Utilities.getUuid().slice(0, 8);
  guardarCuenta({ parent_id: destParent.id, nombre: splitSubAName, saldo_inicial: 0 });
  guardarCuenta({ parent_id: destParent.id, nombre: splitSubBName, saldo_inicial: 0 });
  const destCtas = obtenerCuentas().find(c => c.id === destParent.id);
  const splitSubA = destCtas.subcuentas.find(s => s.nombre === splitSubAName).id;
  const splitSubB = destCtas.subcuentas.find(s => s.nombre === splitSubBName).id;
  const txSplit = guardarTransaccion({
    tipo: 'transferencia', importe: 50,
    cuenta_id: parent.id,
    cuenta_destino_id: destParent.id,
    reparto_destino: [
      { subcuenta_id: splitSubA, importe: 30 },
      { subcuenta_id: splitSubB, importe: 20 }
    ],
    descripcion: 'self-split-transfer', fecha: isoHoy_()
  });
  const afterSplit = obtenerCuentas().find(c => c.id === destParent.id).subcuentas;
  const subAAfter = afterSplit.find(s => s.id === splitSubA);
  const subBAfter = afterSplit.find(s => s.id === splitSubB);
  if (!subAAfter || subAAfter.saldo !== 30) throw new Error('Reparto no acreditó subA: ' + (subAAfter && subAAfter.saldo));
  if (!subBAfter || subBAfter.saldo !== 20) throw new Error('Reparto no acreditó subB: ' + (subBAfter && subBAfter.saldo));
  // Validación: la suma debe ser exactamente el total del importe.
  let fallido = null;
  try {
    guardarTransaccion({
      tipo: 'transferencia', importe: 50,
      cuenta_id: parent.id, cuenta_destino_id: destParent.id,
      reparto_destino: [
        { subcuenta_id: splitSubA, importe: 30 },
        { subcuenta_id: splitSubB, importe: 15 }
      ],
      descripcion: 'self-split-bad', fecha: isoHoy_()
    });
  } catch (e) { fallido = e.message; }
  if (!fallido || !/suma del reparto/.test(fallido)) throw new Error('No se rechazó reparto con suma incorrecta: ' + fallido);

  // Cleanup
  eliminarTransaccion(txSub.id);
  eliminarTransaccion(txTransfer.id);
  eliminarTransaccion(txSplit.id);
  eliminarCuenta(subId);
  eliminarCuenta(anotherId);
  eliminarCuenta(destSubId);
  eliminarCuenta(splitSubA);
  eliminarCuenta(splitSubB);
  eliminarCuenta(destParent.id);

  return 'ok ' + cuentas.length + ' cuentas, ' + cat.length + ' categorías, diferencia=' + conciliado.diferencia;
}
