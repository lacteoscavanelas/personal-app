/**
 * PERSONAL · Pedidos y licencias
 * App web de Google Apps Script para pedidos del personal: vacaciones, días,
 * francos compensatorios, adelantos, avisos de ausencia y licencias.
 *
 * Primera vez:
 *   1. Ejecutá setup() una sola vez (crea la planilla, carga la nómina y
 *      genera usuarios y contraseñas temporales). Mirá el registro: ahí
 *      aparece el usuario y la contraseña de Recursos Humanos.
 *   2. Implementar > Nueva implementación > Aplicación web
 *      Ejecutar como: Yo · Quién tiene acceso: Cualquier persona.
 *   3. Abrí la URL de la app y entrá con el usuario de RRHH.
 * Si ya habías ejecutado setup() con la versión anterior (links personales),
 * ejecutá migrar() en lugar de setup().
 */

const HOJAS = {
  Nomina: ['Clave','Apellido','Nombre','DNI','CUIL','Ingreso','Tarea','Sector','Encargado','Salario','VacDias','Telefono','Usuario','TempPass','PassHash','Salt','CarpetaId'],
  Admins: ['Usuario','Nombre','TempPass','PassHash','Salt'],
  Encargados: ['Nombre','Telefono'],
  Pedidos: ['Id','Clave','Empleado','Tipo','Creado','Desde','Hasta','Dias','Trabajado','Monto','Devolucion','Periodo','Modalidad','Categoria','Certificado','Motivo','Estado','Nota','Respondido','Archivos','EventoId','Trello'],
  Banco: ['Id','Clave','Empleado','Fecha','Movimiento','Dias','Motivo','Pedido','Cargado','Por'],
  VacPendientes: ['Id','Clave','Empleado','Temporada','Dias','Vence','Nota','Cargado','Por'],
  Config: ['Clave','Valor']
};
const TIPOS = {
  vacaciones: 'Vacaciones', dia: 'Pedido de día', franco: 'Franco compensatorio',
  adelanto: 'Adelanto de sueldo', ausencia: 'Aviso de ausencia', licencia: 'Licencia especial'
};
const A_CUENTA_VAC = 'A cuenta de vacaciones';
const MOTIVOS_CREDITO = ['Feriado trabajado', 'Franco trabajado', 'Domingo trabajado', 'Horas extra acumuladas', 'Otro'];
const MOTIVOS_USO = ['Franco tomado', 'Pagado en liquidación', 'Ajuste'];
const TZ = 'America/Argentina/Buenos_Aires';

/* ================= Instalación ================= */

function setup() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('SHEET_ID')) throw new Error('La app ya está instalada. Si venís de la versión con links personales, ejecutá migrar().');
  const ss = SpreadsheetApp.create('Personal - Pedidos y licencias');
  Object.keys(HOJAS).forEach((nombre, i) => {
    const sh = i === 0 ? ss.getSheets()[0].setName(nombre) : ss.insertSheet(nombre);
    sh.getRange(1, 1, sh.getMaxRows(), HOJAS[nombre].length).setNumberFormat('@');
    sh.getRange(1, 1, 1, HOJAS[nombre].length).setValues([HOJAS[nombre]]).setFontWeight('bold');
    sh.setFrozenRows(1);
  });
  // Clave..Telefono (12 columnas) + Usuario, TempPass, PassHash, Salt, CarpetaId
  const nomina = NOMINA_INICIAL.map(r => [...r.map(String), '', '', usuarioInicial_(r[1], r[2], r[3]), passTemporal_(), '', '', '']);
  ss.getSheetByName('Nomina').getRange(2, 1, nomina.length, HOJAS.Nomina.length).setValues(nomina);
  ss.getSheetByName('Encargados').getRange(2, 1, ENCARGADOS_INICIALES.length, 2).setValues(ENCARGADOS_INICIALES.map(r => r.map(String)));
  const tempAdmin = passTemporal_();
  ss.getSheetByName('Admins').getRange(2, 1, 1, 5).setValues([['rrhh', 'Recursos Humanos', tempAdmin, '', '']]);
  let legajos;
  try { legajos = DriveApp.getFolderById(LEGAJOS_ID_INICIAL).getId(); }
  catch (e) { legajos = DriveApp.createFolder('Legajos').getId(); }
  ss.getSheetByName('Config').getRange(2, 1, 2, 2).setValues([['topePct', String(TOPE_ADELANTO_PCT_INICIAL)], ['legajosId', legajos]]);
  props.setProperty('SHEET_ID', ss.getId());
  asegurarCalendario_();
  Logger.log('Planilla creada: ' + ss.getUrl());
  Logger.log('Usuario de Recursos Humanos: rrhh · Contraseña temporal: ' + tempAdmin + ' (te va a pedir cambiarla al entrar).');
  Logger.log('Ahora implementá la app web: Implementar > Nueva implementación > Aplicación web.');
}

/** Para quien ya instaló la versión con links: agrega usuarios y contraseñas sin perder datos. */
function migrar() {
  const ss = ss_();
  Object.keys(HOJAS).forEach(nombre => {
    let sh = ss.getSheetByName(nombre);
    if (!sh) { sh = ss.insertSheet(nombre); sh.getRange(1, 1, sh.getMaxRows(), HOJAS[nombre].length).setNumberFormat('@'); sh.setFrozenRows(1); }
    const h = sh.getLastColumn() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] : [];
    HOJAS[nombre].filter(k => h.indexOf(k) < 0).forEach(k => {
      const col = sh.getLastColumn() + 1;
      sh.getRange(1, col).setValue(k).setFontWeight('bold');
      sh.getRange(2, col, sh.getMaxRows() - 1, 1).setNumberFormat('@');
    });
  });
  leer_('Nomina').forEach(p => { if (!p.Usuario) escribir_('Nomina', p._fila, { Usuario: usuarioInicial_(p.Apellido, p.Nombre, p.DNI), TempPass: passTemporal_() }); });
  if (!leer_('Admins').length) {
    const t = passTemporal_(); agregar_('Admins', { Usuario: 'rrhh', Nombre: 'Recursos Humanos', TempPass: t });
    Logger.log('Usuario de Recursos Humanos: rrhh · Contraseña temporal: ' + t);
  }
  asegurarCalendario_();
  Logger.log('Listo. Los usuarios y contraseñas temporales de cada empleado están en la pestaña Personal de la app.');
}

/* ---- Usuarios y contraseñas ---- */

function sinAcentos_(t) { return String(t || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/ñ/gi, 'n'); }
function usuarioInicial_(apellido, nombre, dni) {
  if (dni) return String(dni).trim();
  return (sinAcentos_(apellido).trim().split(/\s+/)[0] + '.' + sinAcentos_(nombre).trim().split(/\s+/)[0]).toLowerCase();
}
function passTemporal_() { return String(Math.floor(100000 + Math.random() * 900000)); }
function hash_(pass, salt) {
  let h = salt + '|' + pass;
  for (let i = 0; i < 200; i++) h = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h + salt, Utilities.Charset.UTF_8));
  return h;
}
function passOk_(fila, pass) {
  if (fila.TempPass && pass === fila.TempPass) return { ok: true, cambiar: true };
  if (fila.PassHash && fila.Salt && hash_(pass, fila.Salt) === fila.PassHash) return { ok: true, cambiar: false };
  return { ok: false };
}
const DIAS_SESION = 30;
function crearSesion_(rol, id) {
  const props = PropertiesService.getScriptProperties(), ahora = Date.now();
  Object.entries(props.getProperties()).forEach(([k, v]) => { if (k.indexOf('s_') === 0) { try { if (JSON.parse(v).x < ahora) props.deleteProperty(k); } catch (e) { props.deleteProperty(k); } } });
  const ses = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  props.setProperty('s_' + ses, JSON.stringify({ r: rol, c: id, x: ahora + DIAS_SESION * 864e5 }));
  return ses;
}
function sesion_(ses) {
  const v = ses && PropertiesService.getScriptProperties().getProperty('s_' + ses);
  if (!v) throw new Error('SESION: Tu sesión venció. Volvé a entrar.');
  const o = JSON.parse(v); if (o.x < Date.now()) throw new Error('SESION: Tu sesión venció. Volvé a entrar.');
  return o;
}
function cerrarSesionesDe_(rol, id) {
  const props = PropertiesService.getScriptProperties();
  Object.entries(props.getProperties()).forEach(([k, v]) => { if (k.indexOf('s_') === 0) { try { const o = JSON.parse(v); if (o.r === rol && o.c === id) props.deleteProperty(k); } catch (e) {} } });
}
function buscarUsuario_(usuario) {
  const u = String(usuario || '').trim().toLowerCase();
  const a = leer_('Admins').find(x => x.Usuario.trim().toLowerCase() === u);
  if (a) return { rol: 'a', hoja: 'Admins', fila: a, id: a.Usuario };
  const p = leer_('Nomina').find(x => x.Usuario.trim().toLowerCase() === u);
  if (p) return { rol: 'e', hoja: 'Nomina', fila: p, id: p.Clave };
  return null;
}

function entrar(usuario, pass) {
  const cache = CacheService.getScriptCache(), k = 'f_' + String(usuario || '').trim().toLowerCase();
  const fallos = Number(cache.get(k) || 0);
  if (fallos >= 5) throw new Error('Demasiados intentos. Esperá 15 minutos o pedile a Recursos Humanos una contraseña nueva.');
  const u = buscarUsuario_(usuario), r = u ? passOk_(u.fila, String(pass || '')) : { ok: false };
  if (!r.ok) {
    cache.put(k, String(fallos + 1), 900);
    if (fallos + 1 === 5) slack_('⚠️ *Usuario bloqueado 15 minutos* por 5 intentos fallidos: `' + String(usuario || '').trim() + '`' + (u ? ' (' + (u.rol === 'a' ? 'RRHH' : nombreDe_(u.fila)) + ')' : ' (no existe)'));
    throw new Error('Usuario o contraseña incorrectos.');
  }
  cache.remove(k);
  return { sesion: crearSesion_(u.rol, u.id), rol: u.rol === 'a' ? 'admin' : 'empleado', cambiar: r.cambiar };
}

function cambiarContrasena(ses, actual, nueva) {
  const s = sesion_(ses);
  nueva = String(nueva || '');
  if (nueva.length < 6) throw new Error('La contraseña nueva tiene que tener al menos 6 caracteres.');
  return conLock_(() => {
    const hoja = s.r === 'a' ? 'Admins' : 'Nomina';
    const fila = leer_(hoja).find(x => (s.r === 'a' ? x.Usuario : x.Clave) === s.c);
    if (!fila || !passOk_(fila, String(actual || '')).ok) throw new Error('La contraseña actual no es correcta.');
    if (nueva === fila.TempPass) throw new Error('Elegí una contraseña distinta de la temporal.');
    const salt = Utilities.getUuid();
    escribir_(hoja, fila._fila, { PassHash: hash_(nueva, salt), Salt: salt, TempPass: '' });
    if (fila.TempPass) slack_('🙋 *' + (s.r === 'a' ? 'El usuario de RRHH ' + fila.Usuario : nombreDe_(fila)) + '* entró a la app por primera vez y eligió su contraseña.');
    return true;
  });
}

function salir(ses) { if (ses) PropertiesService.getScriptProperties().deleteProperty('s_' + ses); return true; }

/* ================= Web ================= */

function doGet(e) {
  const t = HtmlService.createTemplateFromFile('Index');
  t.pedido = (e && e.parameter && /^[a-f0-9]{12}$/.test(e.parameter.p || '')) ? e.parameter.p : '';
  return t.evaluate()
    .setTitle('Personal · Pedidos y licencias')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setFaviconUrl('https://fonts.gstatic.com/s/e/notoemoji/latest/1f4cb/72.png');
}

/* ================= Utilidades ================= */

function ss_() {
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!id) throw new Error('La app no está instalada: ejecutá setup().');
  return SpreadsheetApp.openById(id);
}
function token_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 12); }
function hoy_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function ahora_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm'); }

function leer_(hoja) {
  const sh = ss_().getSheetByName(hoja);
  if (!sh) return [];
  const v = sh.getDataRange().getDisplayValues();
  const h = v.shift();
  return v.map((r, i) => { const o = { _fila: i + 2 }; h.forEach((k, j) => o[k] = r[j]); return o; })
          .filter(o => o[h[0]] !== '');
}
function escribir_(hoja, fila, obj) {
  const sh = ss_().getSheetByName(hoja);
  const h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  Object.keys(obj).forEach(k => { const j = h.indexOf(k); if (j >= 0) sh.getRange(fila, j + 1).setValue(obj[k] === null || obj[k] === undefined ? '' : String(obj[k])); });
}
function agregar_(hoja, obj) {
  const sh = ss_().getSheetByName(hoja);
  const h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const rg = sh.getRange(sh.getLastRow() + 1, 1, 1, h.length);
  rg.setNumberFormat('@').setValues([h.map(k => obj[k] === undefined || obj[k] === null ? '' : String(obj[k]))]);
}
function config_() { const o = {}; leer_('Config').forEach(r => o[r.Clave] = r.Valor); return o; }
function conLock_(fn) { const l = LockService.getScriptLock(); l.waitLock(20000); try { return fn(); } finally { l.releaseLock(); } }

function persona_(ses) {
  const s = sesion_(ses);
  if (s.r !== 'e') throw new Error('Esta cuenta es de Recursos Humanos.');
  const p = leer_('Nomina').find(r => r.Clave === s.c);
  if (!p) throw new Error('SESION: Tu usuario ya no está en la nómina.');
  return p;
}
function admin_(ses) {
  const s = sesion_(ses);
  if (s.r !== 'a') throw new Error('Acceso no autorizado.');
  return s.c;
}
const nombreDe_ = p => (p.Nombre + ' ' + p.Apellido).trim();

/* ---- Reglas: vacaciones (LCT arts. 150 y 154), período de liquidación 26 al 25 ---- */

function temporadaDe_(iso) { const y = +iso.slice(0, 4), m = +iso.slice(5, 7); return m >= 5 ? y : y - 1; }
function diasEntre_(a, b) { return Math.round((Date.parse(b) - Date.parse(a)) / 864e5) + 1; }
function vacCorresponde_(ingreso, temp) {
  if (!ingreso) return null;
  const fin = Date.UTC(temp, 11, 31), ini = Date.parse(ingreso);
  if (ini > fin) return 0;
  const anios = (fin - ini) / (365.25 * 864e5);
  if (anios < 0.5) return Math.floor(((fin - ini) / 864e5 * 5 / 7) / 20);
  return anios < 5 ? 14 : anios < 10 ? 21 : anios < 20 ? 28 : 35;
}
/**
 * Vacaciones de una persona en una temporada, más los días pendientes de temporadas
 * anteriores que cargó RRHH. Los pedidos usan primero los pendientes que vencen antes
 * (si el pedido empieza antes del vencimiento) y después los de la temporada.
 */
function resumenVac_(p, pedidos, temp, pendientes) {
  const total = p.VacDias !== '' ? Number(p.VacDias) : vacCorresponde_(p.Ingreso, temp);
  const hoy = hoy_();
  const ant = (pendientes || []).filter(x => x.Clave === p.Clave)
    .map(x => ({ id: x.Id, temporada: x.Temporada, dias: Number(x.Dias) || 0, vence: x.Vence, nota: x.Nota, usado: 0, pend: 0 }))
    .sort((a, b) => a.vence.localeCompare(b.vence));
  let usados = 0, pend = 0;
  pedidos.filter(r => r.Clave === p.Clave && r.Desde && (r.Estado === 'Aprobado' || r.Estado === 'Pendiente') &&
                      (r.Tipo === 'vacaciones' || (r.Tipo === 'dia' && r.Modalidad === A_CUENTA_VAC)))
    .sort((a, b) => (a.Desde + a.Creado).localeCompare(b.Desde + b.Creado))
    .forEach(r => {
      let n = Number(r.Dias) || 0; const k = r.Estado === 'Aprobado' ? 'usado' : 'pend';
      for (const a of ant) {
        if (n <= 0) break;
        if (r.Desde > a.vence) continue;
        const libre = a.dias - a.usado - a.pend; if (libre <= 0) continue;
        const t = Math.min(libre, n); a[k] += t; n -= t;
      }
      if (n > 0 && temporadaDe_(r.Desde) === temp) { if (k === 'usado') usados += n; else pend += n; }
    });
  ant.forEach(a => { a.restante = a.dias - a.usado - a.pend; a.vencido = a.vence < hoy; });
  const anteriorDisponible = ant.filter(a => !a.vencido).reduce((t, a) => t + Math.max(0, a.restante), 0);
  return { temporada: temp, total, usados, pend, anteriores: ant, anteriorDisponible };
}
function periodoDe_(fechaHora) {
  let y = +fechaHora.slice(0, 4), m = +fechaHora.slice(5, 7); const d = +fechaHora.slice(8, 10);
  if (d >= 26) { m++; if (m > 12) { m = 1; y++; } }
  return y + '-' + String(m).padStart(2, '0');
}
function topeAdelanto_(p, cfg) {
  const s = Number(String(p.Salario).replace(/[^\d.]/g, '')), pc = Number(cfg.topePct);
  return s > 0 && pc > 0 ? Math.round(s * pc / 100) : null;
}
function adelantosPeriodo_(clave, pedidos, periodo, excluirId) {
  return pedidos.filter(r => r.Clave === clave && r.Tipo === 'adelanto' && r.Periodo === periodo && r.Id !== excluirId && r.Estado !== 'Rechazado' && r.Estado !== 'Cancelado')
                .reduce((t, r) => t + (Number(r.Monto) || 0), 0);
}
function encargadoDe_(p) {
  const e = leer_('Encargados').find(x => x.Nombre.trim().toLowerCase() === String(p.Encargado).trim().toLowerCase());
  return e && e.Telefono ? { nombre: e.Nombre, tel: e.Telefono } : null;
}

/* ---- Banco de días a favor ---- */

function resumenBanco_(clave, banco, pedidos) {
  let favor = 0, usado = 0;
  banco.forEach(m => { if (m.Clave !== clave) return; const d = Number(m.Dias) || 0; if (m.Movimiento === 'Crédito') favor += d; else usado += d; });
  const pedidosPend = pedidos.filter(r => r.Clave === clave && r.Tipo === 'franco' && r.Estado === 'Pendiente').reduce((t, r) => t + (Number(r.Dias) || 0), 0);
  const saldo = Math.round((favor - usado) * 2) / 2;
  return { favor, usado, saldo, pedidosPend, disponible: saldo - pedidosPend };
}
function borrarFilas_(hoja, filas) {
  const sh = ss_().getSheetByName(hoja);
  filas.sort((a, b) => b - a).forEach(f => sh.deleteRow(f));
}
/** Al aprobar un franco se descuenta del banco; si se vuelve atrás, se devuelve. */
function sincronizarFranco_(r, estado, por) {
  const usos = leer_('Banco').filter(m => m.Pedido === r.Id);
  if (estado === 'Aprobado') {
    if (!usos.length) agregar_('Banco', { Id: token_(), Clave: r.Clave, Empleado: r.Empleado, Fecha: r.Desde, Movimiento: 'Uso', Dias: r.Dias || 1,
      Motivo: 'Franco tomado', Pedido: r.Id, Cargado: ahora_(), Por: por });
  } else if (usos.length) borrarFilas_('Banco', usos.map(m => m._fila));
}

/* ================= Empleado ================= */

function datosEmpleado(ses) {
  const p = persona_(ses), cfg = config_(), pedidos = leer_('Pedidos'), banco = leer_('Banco');
  const per = periodoDe_(ahora_());
  return {
    nombre: nombreDe_(p), sector: p.Sector, encargado: encargadoDe_(p),
    vac: resumenVac_(p, pedidos, temporadaDe_(hoy_()), leer_('VacPendientes')),
    adelanto: { tope: topeAdelanto_(p, cfg), usado: adelantosPeriodo_(p.Clave, pedidos, per), periodo: per, pct: cfg.topePct },
    banco: Object.assign(resumenBanco_(p.Clave, banco, pedidos), { movimientos: banco.filter(m => m.Clave === p.Clave).map(limpiar_).reverse() }),
    pedidos: pedidos.filter(r => r.Clave === p.Clave).map(limpiar_).reverse()
  };
}
function saldoVacaciones(ses, desde) {
  const p = persona_(ses);
  return resumenVac_(p, leer_('Pedidos'), temporadaDe_(desde), leer_('VacPendientes'));
}

function enviarPedido(ses, d) {
  const out = conLock_(() => {
    const p = persona_(ses), cfg = config_();
    if (!TIPOS[d.tipo]) throw new Error('Tipo de pedido desconocido.');
    const r = { Id: token_(), Clave: p.Clave, Empleado: nombreDe_(p), Tipo: d.tipo, Creado: ahora_(), Estado: 'Pendiente',
                Motivo: String(d.motivo || '').slice(0, 500), Modalidad: d.modalidad || '', Categoria: d.categoria || '',
                Certificado: d.certificado ? 'Sí' : '' };
    if (d.tipo === 'adelanto') {
      const monto = Math.round(Number(d.monto));
      if (!(monto > 0)) throw new Error('Escribí el monto que necesitás.');
      r.Monto = monto; r.Devolucion = d.devolucion || ''; r.Periodo = periodoDe_(r.Creado);
      const tope = topeAdelanto_(p, cfg);
      if (tope !== null) {
        const disp = Math.max(0, tope - adelantosPeriodo_(p.Clave, leer_('Pedidos'), r.Periodo));
        if (monto > disp) throw new Error(disp ? 'Supera tu tope: en este período podés pedir hasta $' + disp.toLocaleString('es-AR') + '.' : 'Ya llegaste al tope de adelantos de este período. Desde el 26 podés pedir para la próxima liquidación.');
      }
    } else {
      if (!/^\d{4}-\d{2}-\d{2}$/.test(d.desde || '')) throw new Error('Elegí la fecha.');
      r.Desde = d.desde;
      if (d.tipo === 'franco') {
        const b = resumenBanco_(p.Clave, leer_('Banco'), leer_('Pedidos'));
        if (b.disponible < 1) throw new Error(b.saldo >= 1 ? 'Ya pediste todos tus días a favor. Esperá la respuesta de los pedidos que tenés.' : 'No tenés días a favor en tu banco. Si trabajaste un feriado o un franco, pedile a tu encargado que lo cargue.');
        r.Hasta = d.desde; r.Dias = 1;
      } else {
        r.Hasta = d.hasta || d.desde;
        if (r.Hasta < r.Desde) throw new Error('La fecha de fin es anterior a la de inicio.');
        r.Dias = diasEntre_(r.Desde, r.Hasta);
        if (r.Dias > 120) throw new Error('Revisá las fechas: son más de 120 días.');
      }
    }
    agregar_('Pedidos', r);
    return { res: { id: r.Id, wa: linkWhatsApp_(p, r) }, r, p };
  });
  const extra = extraPedido_(out.r, out.p);
  let tarjeta = '';
  try {
    tarjeta = crearTarjetaRevision_(out.r, out.p, extra) || '';
    if (tarjeta) conLock_(() => { const f = leer_('Pedidos').find(x => x.Id === out.r.Id); if (f) escribir_('Pedidos', f._fila, { Trello: tarjeta }); });
  } catch (e) { console.error('Trello: ' + e.message); }
  slack_((out.r.Tipo === 'ausencia' ? '🤒 *Aviso de ausencia*' : '📥 *Pedido nuevo*') + ' de *' + nombreDe_(out.p) + '*' + (out.p.Sector ? ' (' + out.p.Sector + ')' : '') + '\n' + textoPedido_(out.r) + extra +
    abrirApp_(out.r.Id, '👉 Revisar y responder') + (tarjeta ? ' · <' + tarjeta + '|Tarjeta en Trello>' : ''));
  return out.res;
}

function cancelarPedido(ses, id) {
  return conLock_(() => {
    const p = persona_(ses);
    const r = leer_('Pedidos').find(x => x.Id === id && x.Clave === p.Clave);
    if (!r) throw new Error('No se encontró el pedido.');
    if (r.Estado !== 'Pendiente') throw new Error('Ese pedido ya fue respondido.');
    escribir_('Pedidos', r._fila, { Estado: 'Cancelado' });
    try { if (r.Trello) cerrarTarjeta_(r.Trello, '🚫 El empleado canceló el pedido.'); } catch (e) { console.error('Trello: ' + e.message); }
    slack_('🚫 *' + nombreDe_(p) + '* canceló su pedido\n' + textoPedido_(r));
    return true;
  });
}

function subirCertificado(ses, id, nombreArchivo, mime, base64) {
  const p = persona_(ses);
  if (!/^(image\/(jpeg|png)|application\/pdf)$/.test(mime)) throw new Error('Subí una foto o un PDF.');
  const r = leer_('Pedidos').find(x => x.Id === id && x.Clave === p.Clave);
  if (!r) throw new Error('No se encontró el pedido.');
  const carpeta = carpetaLegajo_(p);
  const ext = mime === 'application/pdf' ? 'pdf' : mime === 'image/png' ? 'png' : 'jpg';
  const titulo = (r.Desde || r.Creado.slice(0, 10)) + ' Certificado ' + TIPOS[r.Tipo].toLowerCase() + (r.Categoria ? ' - ' + r.Categoria : '') + '.' + ext;
  const blob = Utilities.newBlob(Utilities.base64Decode(base64), mime, titulo);
  const f = carpeta.createFile(blob);
  return conLock_(() => {
    const act = leer_('Pedidos').find(x => x.Id === id);
    escribir_('Pedidos', act._fila, { Archivos: (act.Archivos ? act.Archivos + ' ' : '') + f.getUrl(), Certificado: 'Sí' });
    slack_('📎 *' + nombreDe_(p) + '* subió un certificado: ' + TIPOS[r.Tipo] + (r.Desde ? ' del ' + fechaTxt_(r.Desde) : '') + ' · <' + f.getUrl() + '|Ver archivo>' + abrirApp_(r.Id, 'Ver el pedido'));
    return f.getUrl();
  });
}

function carpetaLegajo_(p) {
  if (p.CarpetaId) { try { return DriveApp.getFolderById(p.CarpetaId); } catch (e) {} }
  const raiz = DriveApp.getFolderById(config_().legajosId);
  const nombre = (p.Apellido + ' ' + p.Nombre).trim();
  const it = raiz.getFoldersByName(nombre);
  const carpeta = it.hasNext() ? it.next() : raiz.createFolder(nombre);
  conLock_(() => escribir_('Nomina', p._fila, { CarpetaId: carpeta.getId() }));
  return carpeta;
}

function linkWhatsApp_(p, r) {
  const e = encargadoDe_(p); if (!e) return null;
  let tel = e.tel.replace(/\D/g, ''); if (!tel.startsWith('54')) tel = '549' + tel.replace(/^0/, '');
  let det = TIPOS[r.Tipo];
  if (r.Tipo === 'adelanto') det += ' por $' + Number(r.Monto).toLocaleString('es-AR');
  else if (r.Tipo === 'franco') det += ' el ' + r.Desde;
  else det += r.Desde === r.Hasta ? ' el ' + r.Desde : ' del ' + r.Desde + ' al ' + r.Hasta;
  const txt = 'Hola ' + e.nombre + ', soy ' + nombreDe_(p) + (p.Sector ? ' (' + p.Sector + ')' : '') + '. Cargué en la app de Personal: ' + det + (r.Motivo ? '. Motivo: ' + r.Motivo : '') + '. ¿Lo podés revisar?';
  return 'https://wa.me/' + tel + '?text=' + encodeURIComponent(txt);
}
function limpiar_(r) { const o = Object.assign({}, r); delete o._fila; return o; }

/* ================= Recursos Humanos ================= */

function datosAdmin(ses) {
  admin_(ses);
  const cfg = config_(), pedidos = leer_('Pedidos'), banco = leer_('Banco'), vacPend = leer_('VacPendientes'), temp = temporadaDe_(hoy_()), url = ScriptApp.getService().getUrl();
  const personas = leer_('Nomina').map(p => ({
    clave: p.Clave, apellido: p.Apellido, nombre: p.Nombre, tarea: p.Tarea, sector: p.Sector, encargado: p.Encargado,
    ingreso: p.Ingreso, salario: p.Salario, vacDias: p.VacDias, telefono: p.Telefono,
    usuario: p.Usuario, temp: p.TempPass, activo: !!p.PassHash, carpeta: p.CarpetaId ? 'https://drive.google.com/drive/folders/' + p.CarpetaId : '',
    vac: resumenVac_(p, pedidos, temp, vacPend), tope: topeAdelanto_(p, cfg), vacAuto: vacCorresponde_(p.Ingreso, temp),
    banco: resumenBanco_(p.Clave, banco, pedidos)
  }));
  return {
    personas, pedidos: pedidos.map(limpiar_), temporada: temp, banco: banco.map(limpiar_),
    motivosCredito: MOTIVOS_CREDITO, motivosUso: MOTIVOS_USO,
    encargados: leer_('Encargados').map(e => ({ nombre: e.Nombre, tel: e.Telefono })),
    topePct: cfg.topePct, planilla: ss_().getUrl(), url,
    admins: leer_('Admins').map(a => ({ usuario: a.Usuario, nombre: a.Nombre, temp: a.TempPass })),
    legajos: 'https://drive.google.com/drive/folders/' + cfg.legajosId,
    integraciones: { slack: !!prop_('SLACK_WEBHOOK'), trello: !!(prop_('TRELLO_KEY') && prop_('TRELLO_TOKEN') && prop_('TRELLO_LIST_ID')),
      calendario: cfg.calendarioId ? 'https://calendar.google.com/calendar/r?cid=' + encodeURIComponent(cfg.calendarioId) : '' }
  };
}

function responderPedido(ses, id, estado, nota) {
  const por = admin_(ses);
  if (['Aprobado', 'Rechazado', 'Pendiente'].indexOf(estado) < 0) throw new Error('Estado inválido.');
  return conLock_(() => {
    const r = leer_('Pedidos').find(x => x.Id === id);
    if (!r) throw new Error('No se encontró el pedido.');
    if (r.Tipo === 'franco' && estado === 'Aprobado' && r.Estado !== 'Aprobado') {
      const b = resumenBanco_(r.Clave, leer_('Banco'), []);
      if (b.saldo < (Number(r.Dias) || 1)) throw new Error('No tiene días a favor en el banco. Cargá primero el día trabajado en la pestaña Banco de días.');
    }
    escribir_('Pedidos', r._fila, { Estado: estado, Nota: estado === 'Pendiente' ? '' : String(nota || '').slice(0, 300), Respondido: estado === 'Pendiente' ? '' : ahora_() });
    if (r.Tipo === 'franco') sincronizarFranco_(r, estado, por);
    const avisos = [];
    // Calendario Vacaciones: vacaciones, pedidos de día y francos aprobados
    if (['vacaciones', 'dia', 'franco'].indexOf(r.Tipo) >= 0) {
      try {
        if (estado === 'Aprobado' && !r.EventoId) { escribir_('Pedidos', r._fila, { EventoId: crearEvento_(r) }); avisos.push('Agregado al calendario Vacaciones'); }
        else if (estado !== 'Aprobado' && r.EventoId) { borrarEvento_(r.EventoId); escribir_('Pedidos', r._fila, { EventoId: '' }); avisos.push('Quitado del calendario Vacaciones'); }
      } catch (e) { avisos.push('No se pudo actualizar el calendario: ' + e.message); }
    }
    // Trello: la tarjeta de revisión se cierra al responder; la de un adelanto aprobado pasa a ser el recordatorio de descuento
    if (trelloOk_()) {
      try {
        const res = { Aprobado: '✅ ' + (r.Tipo === 'ausencia' ? 'Ausencia justificada' : 'Aprobado'), Rechazado: '❌ ' + (r.Tipo === 'ausencia' ? 'Ausencia injustificada' : 'Rechazado') }[estado];
        if (estado === 'Pendiente') {
          if (r.Trello) { reabrirTarjeta_(r.Trello, r); avisos.push('Tarjeta de Trello de nuevo en Pendientes'); }
        } else if (estado === 'Aprobado' && r.Tipo === 'adelanto') {
          if (r.Trello) { tarjetaDescuento_(r.Trello, r, por, nota); }
          else { r.Trello = crearTarjetaDescuento_(r, por, nota); escribir_('Pedidos', r._fila, { Trello: r.Trello }); }
          avisos.push('Tarjeta de Trello lista para descontar en la liquidación');
        } else if (r.Trello) {
          cerrarTarjeta_(r.Trello, res + ' por ' + por + (nota ? '. Nota: ' + nota : '') + '.');
          avisos.push('Tarjeta de Trello cerrada');
        }
      } catch (e) { avisos.push('No se pudo actualizar Trello: ' + e.message); }
    }
    const et = { Aprobado: r.Tipo === 'ausencia' ? '✅ *Ausencia justificada*' : '✅ *Aprobado*', Rechazado: r.Tipo === 'ausencia' ? '❌ *Ausencia injustificada*' : '❌ *Rechazado*', Pendiente: '↩️ *Vuelto a pendiente*' }[estado];
    slack_(et + ' · ' + r.Empleado + ' (por ' + por + ')\n' + textoPedido_(r) + (nota && estado !== 'Pendiente' ? '\n• Nota: ' + nota : '') +
      (avisos.length ? '\n_' + avisos.join(' · ') + '_' : '') + (r.Trello ? '\n<' + r.Trello + '|Ver tarjeta en Trello>' : '') + (estado === 'Pendiente' ? abrirApp_(r.Id, 'Revisar de nuevo') : ''));
    return { avisos };
  });
}

/** Contexto útil para decidir: saldo, tope, quién más falta. */
function extraPedido_(r, p) {
  try {
    const pedidos = leer_('Pedidos');
    let t = '';
    if (r.Tipo === 'vacaciones' || (r.Tipo === 'dia' && r.Modalidad === A_CUENTA_VAC)) {
      const v = resumenVac_(p, pedidos, temporadaDe_(r.Desde), leer_('VacPendientes'));
      if (v.total !== null) t += '\n• Saldo temporada ' + v.temporada + '/' + String(v.temporada + 1).slice(2) + ': ' + (v.total - v.usados - v.pend) + ' días después de este pedido' + (v.anteriorDisponible ? ' + ' + v.anteriorDisponible + ' pendientes anteriores' : '');
    }
    if (r.Tipo === 'franco') { const b = resumenBanco_(p.Clave, leer_('Banco'), pedidos); t += '\n• Banco de días: ' + b.saldo + ' a favor, ' + b.pedidosPend + ' pedidos'; }
    if (r.Tipo === 'adelanto') {
      const tope = topeAdelanto_(p, config_()), tot = adelantosPeriodo_(p.Clave, pedidos, r.Periodo);
      t += '\n• ' + (tope === null ? 'Sin sueldo cargado: no hay tope' : 'Tope del período $' + tope.toLocaleString('es-AR') + ', con este pedido suma $' + tot.toLocaleString('es-AR'));
    }
    if (r.Desde && r.Tipo !== 'adelanto') {
      const hasta = r.Hasta || r.Desde, nom = {};
      leer_('Nomina').forEach(x => nom[x.Clave] = nombreDe_(x));
      const otros = [...new Set(pedidos.filter(x => x.Clave !== r.Clave && x.Desde && x.Tipo !== 'adelanto' && (x.Estado === 'Aprobado' || x.Estado === 'Pendiente') && x.Desde <= hasta && (x.Hasta || x.Desde) >= r.Desde).map(x => nom[x.Clave] || x.Empleado))];
      if (otros.length) t += '\n• ⚠️ Esos días también faltan: ' + otros.join(', ');
    }
    return t;
  } catch (e) { return ''; }
}

/** Resumen diario en Slack de lo que está esperando respuesta. Se activa con activarResumenDiario(). */
function resumenDiario() {
  const pend = leer_('Pedidos').filter(r => r.Estado === 'Pendiente');
  if (!pend.length) return;
  const viejo = pend.reduce((a, r) => a.Creado < r.Creado ? a : r);
  slack_('🗂️ *Pedidos esperando respuesta: ' + pend.length + '*\n' + pend.slice(0, 15).map(r => '• ' + r.Empleado + ': ' + textoPedido_(r).split('\n')[0].replace(/^• /, '') + ' (desde ' + fechaTxt_(r.Creado.slice(0, 10)) + ')').join('\n') +
    (pend.length > 15 ? '\n…y ' + (pend.length - 15) + ' más' : '') + '\nEl más antiguo es de ' + fechaTxt_(viejo.Creado.slice(0, 10)) + '.' + abrirApp_('', '👉 Abrir la app'));
}
/** Ejecutalo una vez: manda el resumen de pendientes todos los días entre las 8 y las 9. */
function activarResumenDiario() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'resumenDiario').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('resumenDiario').timeBased().everyDays(1).atHour(8).inTimezone(TZ).create();
  Logger.log('Listo: todos los días entre las 8 y las 9 llega a #rrhh el resumen de pedidos sin responder.');
}

/* ================= Integraciones: Slack, Google Calendar y Trello ================= */
/*
 * Propiedades del script (Configuración del proyecto > Propiedades del script):
 *   SLACK_WEBHOOK   URL del webhook entrante del canal #rrhh
 *   TRELLO_KEY      API key de Trello
 *   TRELLO_TOKEN    token de Trello
 *   TRELLO_LIST_ID  ID de la lista "Pendientes" (ejecutá listasTrello() para verlo)
 * Si alguna falta, esa integración simplemente no se usa.
 */
function prop_(k) { return String(PropertiesService.getScriptProperties().getProperty(k) || '').trim(); }
function fechaTxt_(iso) { return iso ? iso.slice(8, 10) + '/' + iso.slice(5, 7) + '/' + iso.slice(0, 4) : ''; }
function abrirApp_(id, texto) { const u = ScriptApp.getService().getUrl(); return u ? '\n<' + u + (id ? '?p=' + id : '') + '|' + (texto || 'Abrir la app') + '>' : ''; }
function textoPedido_(r) {
  let t = '• ' + TIPOS[r.Tipo];
  if (r.Tipo === 'adelanto') t += ': $' + Number(r.Monto).toLocaleString('es-AR') + ' · ' + r.Devolucion + ' · liquidación ' + r.Periodo;
  else if (r.Desde === r.Hasta || !r.Hasta) t += ': ' + fechaTxt_(r.Desde);
  else t += ': del ' + fechaTxt_(r.Desde) + ' al ' + fechaTxt_(r.Hasta) + ' (' + r.Dias + ' días)';
  if (r.Modalidad) t += ' · ' + r.Modalidad;
  if (r.Categoria) t += ' · ' + r.Categoria;
  if (r.Motivo) t += '\n• Motivo: ' + r.Motivo;
  return t;
}

function slack_(texto) {
  const url = prop_('SLACK_WEBHOOK'); if (!url) return false;
  try {
    const res = UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json', payload: JSON.stringify({ text: texto }), muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) console.error('Slack respondió ' + res.getResponseCode() + ': ' + res.getContentText());
    return res.getResponseCode() === 200;
  } catch (e) { console.error('Slack: ' + e); return false; }
}

function asegurarCalendario_() {
  const cfg = config_();
  if (cfg.calendarioId) { try { if (CalendarApp.getCalendarById(cfg.calendarioId)) return cfg.calendarioId; } catch (e) {} }
  const cal = CalendarApp.getCalendarsByName('Vacaciones')[0] || CalendarApp.createCalendar('Vacaciones', { color: CalendarApp.Color.TEAL, summary: 'Vacaciones, días y francos aprobados en la app de Personal' });
  const fila = leer_('Config').find(x => x.Clave === 'calendarioId');
  if (fila) escribir_('Config', fila._fila, { Valor: cal.getId() }); else agregar_('Config', { Clave: 'calendarioId', Valor: cal.getId() });
  Logger.log('Calendario Vacaciones listo. Compartilo desde Google Calendar: Configuración del calendario > Compartir con determinadas personas.');
  return cal.getId();
}
function calendario_() {
  const id = config_().calendarioId;
  const cal = id ? CalendarApp.getCalendarById(id) : null;
  if (!cal) throw new Error('no está creado el calendario Vacaciones (ejecutá migrar)');
  return cal;
}
function crearEvento_(r) {
  const d = s => { const [y, m, dd] = s.split('-').map(Number); return new Date(y, m - 1, dd); };
  const ini = d(r.Desde), fin = d(r.Hasta || r.Desde); fin.setDate(fin.getDate() + 1); // fin exclusivo
  const ic = { vacaciones: '🏖️', dia: '📅', franco: '🔁' }[r.Tipo];
  const titulo = ic + ' ' + r.Empleado + ' · ' + (r.Tipo === 'dia' ? 'Día' + (r.Modalidad ? ' (' + r.Modalidad.toLowerCase() + ')' : '') : TIPOS[r.Tipo]);
  const ev = calendario_().createAllDayEvent(titulo, ini, fin, { description: textoPedido_(r).replace(/\*/g, '') + '\n\nAprobado desde la app de Personal.' });
  return ev.getId();
}
function borrarEvento_(id) { const ev = calendario_().getEventById(id); if (ev) ev.deleteEvent(); }

function trello_(metodo, ruta, params) {
  const key = prop_('TRELLO_KEY'), token = prop_('TRELLO_TOKEN');
  const q = Object.entries(Object.assign({ key, token }, params || {})).filter(([, v]) => v !== '' && v !== null && v !== undefined).map(([k, v]) => k + '=' + encodeURIComponent(v)).join('&');
  const res = UrlFetchApp.fetch('https://api.trello.com/1' + ruta + '?' + q, { method: metodo, muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw new Error('Trello respondió ' + res.getResponseCode() + ': ' + res.getContentText().slice(0, 150));
  return JSON.parse(res.getContentText());
}
function trelloOk_() { return !!(prop_('TRELLO_KEY') && prop_('TRELLO_TOKEN') && prop_('TRELLO_LIST_ID')); }
function idTarjeta_(url) { const x = String(url || '').split('/c/')[1]; return x ? x.split('/')[0] : ''; }
const QUE_REVISAR = { vacaciones: 'pedido de vacaciones', dia: 'pedido de día', franco: 'pedido de franco compensatorio',
  adelanto: 'pedido de adelanto', ausencia: 'aviso de ausencia', licencia: 'pedido de licencia especial' };
function linkApp_(id) { const u = ScriptApp.getService().getUrl(); return u ? u + '?p=' + id : ''; }
function descTarjeta_(r, extra) {
  return textoPedido_(r).replace(/\*/g, '') + (extra || '').replace(/\*/g, '') + '\n• Enviado el ' + r.Creado +
    (linkApp_(r.Id) ? '\n\n👉 Revisar y responder en la app: ' + linkApp_(r.Id) : '');
}
function comentar_(id, texto) { trello_('post', '/cards/' + id + '/actions/comments', { text: texto }); }

/** Pedido nuevo: "Revisar pedido de vacaciones de Juan Pérez" en Pendientes. */
function crearTarjetaRevision_(r, p, extra) {
  if (!trelloOk_()) return null;
  const c = trello_('post', '/cards', {
    idList: prop_('TRELLO_LIST_ID'), pos: 'top',
    name: 'Revisar ' + QUE_REVISAR[r.Tipo] + ' de ' + nombreDe_(p),
    desc: descTarjeta_(r, extra),
    due: r.Desde && r.Tipo !== 'ausencia' ? r.Desde + 'T09:00:00-03:00' : ''
  });
  return c.shortUrl;
}
/** Respondido (o cancelado): deja un comentario y la cierra. Si existe TRELLO_LIST_HECHO, la mueve ahí en lugar de archivarla. */
function cerrarTarjeta_(url, comentario) {
  const id = idTarjeta_(url); if (!id) return;
  comentar_(id, comentario);
  if (prop_('TRELLO_LIST_HECHO')) trello_('put', '/cards/' + id, { idList: prop_('TRELLO_LIST_HECHO'), dueComplete: 'true', pos: 'top' });
  else trello_('put', '/cards/' + id, { closed: 'true', dueComplete: 'true' });
}
/** Vuelto a pendiente: la tarjeta vuelve a Pendientes con el título de revisión. */
function reabrirTarjeta_(url, r) {
  const id = idTarjeta_(url); if (!id) return;
  trello_('put', '/cards/' + id, { closed: 'false', idList: prop_('TRELLO_LIST_ID'), pos: 'top', dueComplete: 'false',
    name: 'Revisar ' + QUE_REVISAR[r.Tipo] + ' de ' + r.Empleado,
    due: r.Desde && r.Tipo !== 'ausencia' ? r.Desde + 'T09:00:00-03:00' : '' });
  comentar_(id, '↩️ El pedido volvió a pendiente.');
}
/** Adelanto aprobado: la misma tarjeta queda como recordatorio para descontarlo, con vencimiento el 25. */
function datosDescuento_(r, por, nota) {
  const [y, m] = (r.Periodo || '').split('-');
  return {
    name: '💵 Descontar adelanto $' + Number(r.Monto).toLocaleString('es-AR') + ' de ' + r.Empleado + (m ? ' · liquidación ' + m + '/' + y : ''),
    desc: 'Adelanto aprobado por ' + por + (nota ? ' (' + nota + ')' : '') + '.\n\n' + textoPedido_(r).replace(/\*/g, '') + '\n• Pedido el ' + r.Creado + '\n\nDescontarlo en el recibo de esa liquidación.',
    due: y && m ? y + '-' + m + '-25T12:00:00-03:00' : '', dueComplete: 'false'
  };
}
function tarjetaDescuento_(url, r, por, nota) {
  const id = idTarjeta_(url); if (!id) return;
  trello_('put', '/cards/' + id, Object.assign({ closed: 'false', idList: prop_('TRELLO_LIST_ID'), pos: 'top' }, datosDescuento_(r, por, nota)));
  comentar_(id, '✅ Aprobado por ' + por + (nota ? '. Nota: ' + nota : '') + '.');
}
function crearTarjetaDescuento_(r, por, nota) {
  return trello_('post', '/cards', Object.assign({ idList: prop_('TRELLO_LIST_ID'), pos: 'top' }, datosDescuento_(r, por, nota))).shortUrl;
}

/** Ejecutalo desde el editor para ver los tableros y listas de Trello con sus IDs. */
function listasTrello() {
  if (!prop_('TRELLO_KEY') || !prop_('TRELLO_TOKEN')) throw new Error('Primero cargá TRELLO_KEY y TRELLO_TOKEN en las propiedades del script.');
  trello_('get', '/members/me/boards', { filter: 'open', fields: 'name', lists: 'open' })
    .forEach(b => { Logger.log('Tablero: ' + b.name); (b.lists || []).forEach(l => Logger.log('   Lista "' + l.name + '" → ' + l.id)); });
}

/** Busca la lista "Pendientes" en Trello y guarda su ID en TRELLO_LIST_ID. */
function configurarListaTrello() {
  const TABLERO = ''; // si tenés varias listas "Pendientes", escribí acá el nombre del tablero
  const encontradas = [];
  trello_('get', '/members/me/boards', { filter: 'open', fields: 'name', lists: 'open' }).forEach(b => {
    if (TABLERO && b.name.trim().toLowerCase() !== TABLERO.trim().toLowerCase()) return;
    (b.lists || []).forEach(l => {
      if (l.name.trim().toLowerCase() === 'pendientes') encontradas.push({ tablero: b.name, id: l.id });
    });
  });
  if (!encontradas.length) throw new Error('No encontré ninguna lista llamada "Pendientes".');
  if (encontradas.length > 1) {
    encontradas.forEach(e => Logger.log('Hay una lista Pendientes en el tablero: ' + e.tablero));
    throw new Error('Hay más de una lista "Pendientes". Escribí el nombre del tablero en TABLERO y volvé a ejecutar.');
  }
  const props = PropertiesService.getScriptProperties();
  props.setProperty('TRELLO_LIST_ID', encontradas[0].id);
  Logger.log('Listo: TRELLO_LIST_ID guardado con la lista Pendientes del tablero "' + encontradas[0].tablero + '".');
  // Si en ese tablero hay una lista "Hecho", "Hechos", "Listo", "Terminado" o "Respondidos", los pedidos respondidos se mueven ahí
  const tablero = trello_('get', '/members/me/boards', { filter: 'open', fields: 'name', lists: 'open' }).find(b => (b.lists || []).some(l => l.id === encontradas[0].id));
  const hecho = tablero && tablero.lists.find(l => ['hecho', 'hechos', 'listo', 'listos', 'terminado', 'terminados', 'respondidos', 'done'].indexOf(l.name.trim().toLowerCase()) >= 0);
  if (hecho) { props.setProperty('TRELLO_LIST_HECHO', hecho.id); Logger.log('Los pedidos respondidos se van a mover a la lista "' + hecho.name + '".'); }
  else { props.deleteProperty('TRELLO_LIST_HECHO'); Logger.log('No hay lista "Hecho" en ese tablero: los pedidos respondidos se van a archivar.'); }
}

/** Ejecutalo desde el editor para probar las tres conexiones. */
function probarIntegraciones() {
  Logger.log('Slack: ' + (prop_('SLACK_WEBHOOK') ? (slack_('✅ Prueba: la app de Personal ya puede avisar en este canal.') ? 'mensaje enviado' : 'falló, revisá SLACK_WEBHOOK') : 'falta SLACK_WEBHOOK'));
  try { Logger.log('Calendario: ' + calendario_().getName() + ' (' + config_().calendarioId + ')'); } catch (e) { Logger.log('Calendario: ' + e.message); }
  if (!prop_('TRELLO_KEY') || !prop_('TRELLO_TOKEN') || !prop_('TRELLO_LIST_ID')) Logger.log('Trello: faltan TRELLO_KEY, TRELLO_TOKEN o TRELLO_LIST_ID');
  else { try { const l = trello_('get', '/lists/' + prop_('TRELLO_LIST_ID'), { fields: 'name' }); Logger.log('Trello: lista "' + l.name + '" encontrada'); } catch (e) { Logger.log('Trello: ' + e.message); } }
  if (prop_('TRELLO_LIST_HECHO')) { try { const l = trello_('get', '/lists/' + prop_('TRELLO_LIST_HECHO'), { fields: 'name' }); Logger.log('Trello: los pedidos respondidos van a la lista "' + l.name + '"'); } catch (e) { Logger.log('Trello (lista de respondidos): ' + e.message); } }
  else Logger.log('Trello: los pedidos respondidos se archivan (no hay TRELLO_LIST_HECHO)');
}

function quien_(ses) { try { return sesion_(ses).c; } catch (e) { return 'RRHH'; } }

function guardarPersona(ses, claveEmpleado, d) {
  admin_(ses);
  return conLock_(() => {
    const p = leer_('Nomina').find(x => x.Clave === claveEmpleado);
    if (!p) throw new Error('No se encontró a esa persona.');
    const usuario = String(d.usuario || '').trim();
    if (!usuario) throw new Error('El usuario no puede quedar vacío.');
    const otro = buscarUsuario_(usuario);
    if (otro && !(otro.rol === 'e' && otro.id === p.Clave)) throw new Error('Ese usuario ya lo tiene otra persona.');
    escribir_('Nomina', p._fila, { Ingreso: d.ingreso || '', Salario: d.salario === '' ? '' : Math.round(Number(d.salario)),
      VacDias: d.vacDias === '' ? '' : Number(d.vacDias), Encargado: d.encargado || '', Telefono: d.telefono || '', Usuario: usuario });
    const cambios = [['Ingreso', p.Ingreso, d.ingreso || ''], ['Sueldo', p.Salario, d.salario === '' ? '' : String(Math.round(Number(d.salario)))], ['Días de vacaciones', p.VacDias, d.vacDias === '' ? '' : String(Number(d.vacDias))],
      ['Encargado', p.Encargado, d.encargado || ''], ['Celular', p.Telefono, d.telefono || ''], ['Usuario', p.Usuario, usuario]].filter(c => String(c[1]) !== String(c[2]));
    if (cambios.length) slack_('✏️ *' + quien_(ses) + '* actualizó los datos de *' + nombreDe_(p) + '*: ' + cambios.map(c => c[0] + ' ' + (c[1] || '—') + ' → ' + (c[2] || '—')).join(' · '));
    return true;
  });
}

/* ---- Banco de días: carga de RRHH ---- */

function cargarBanco(ses, mov) {
  const por = admin_(ses);
  const credito = mov.movimiento === 'Crédito';
  const motivos = credito ? MOTIVOS_CREDITO : MOTIVOS_USO;
  if (!/^\d{4}-\d{2}-\d{2}$/.test(mov.fecha || '')) throw new Error('Elegí la fecha.');
  const dias = Number(mov.dias);
  if (!(dias > 0 && dias <= 30) || Math.round(dias * 2) !== dias * 2) throw new Error('Los días tienen que ser un número entre 0,5 y 30 (enteros o medios días).');
  if (motivos.indexOf(mov.motivo) < 0) throw new Error('Elegí el motivo.');
  const claves = [...new Set(mov.claves || [])];
  if (!claves.length) throw new Error('Elegí al menos una persona.');
  return conLock_(() => {
    const nomina = leer_('Nomina'), banco = leer_('Banco');
    const repetidos = [];
    claves.forEach(c => {
      const p = nomina.find(x => x.Clave === c); if (!p) return;
      if (credito && banco.some(m => m.Clave === c && m.Movimiento === 'Crédito' && m.Fecha === mov.fecha)) { repetidos.push(nombreDe_(p)); return; }
      if (!credito && resumenBanco_(c, banco, []).saldo < dias) throw new Error(nombreDe_(p) + ' no tiene ' + dias + (dias === 1 ? ' día' : ' días') + ' a favor en el banco.');
      agregar_('Banco', { Id: token_(), Clave: c, Empleado: nombreDe_(p), Fecha: mov.fecha, Movimiento: credito ? 'Crédito' : 'Uso', Dias: dias,
        Motivo: mov.motivo + (mov.detalle ? ': ' + String(mov.detalle).slice(0, 120) : ''), Pedido: '', Cargado: ahora_(), Por: por });
    });
    const hechos = claves.map(c => nomina.find(x => x.Clave === c)).filter(p => p && repetidos.indexOf(nombreDe_(p)) < 0).map(nombreDe_);
    if (hechos.length) slack_((credito ? '➕ *Banco de días*: se sumó ' : '➖ *Banco de días*: se descontó ') + dias + (dias === 1 ? ' día' : ' días') + ' a ' + hechos.join(', ') +
      ' · ' + mov.motivo + (mov.detalle ? ' (' + mov.detalle + ')' : '') + ' · ' + fechaTxt_(mov.fecha) + ' · cargado por ' + por);
    return { cargados: claves.length - repetidos.length, repetidos };
  });
}

/* ---- Vacaciones pendientes de temporadas anteriores ---- */

function cargarVacPendientes(ses, d) {
  const por = admin_(ses);
  const dias = Number(d.dias);
  if (!(dias > 0 && dias <= 60) || Math.round(dias) !== dias) throw new Error('Los días tienen que ser un número entero entre 1 y 60.');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(d.vence || '')) throw new Error('Elegí la fecha de vencimiento.');
  if (!/^\d{4}\/\d{2}$/.test(d.temporada || '')) throw new Error('Elegí la temporada de origen.');
  return conLock_(() => {
    const p = leer_('Nomina').find(x => x.Clave === d.clave);
    if (!p) throw new Error('Elegí el empleado.');
    if (leer_('VacPendientes').some(x => x.Clave === d.clave && x.Temporada === d.temporada))
      throw new Error(nombreDe_(p) + ' ya tiene pendientes cargados de la temporada ' + d.temporada + '. Borralos primero si querés corregirlos.');
    agregar_('VacPendientes', { Id: token_(), Clave: p.Clave, Empleado: nombreDe_(p), Temporada: d.temporada, Dias: dias, Vence: d.vence,
      Nota: String(d.nota || '').slice(0, 150), Cargado: ahora_(), Por: por });
    slack_('🏖️ *Vacaciones pendientes*: ' + nombreDe_(p) + ' tiene ' + dias + ' días de la temporada ' + d.temporada + ', vencen el ' + fechaTxt_(d.vence) + ' · cargado por ' + por + (d.nota ? ' · ' + d.nota : ''));
    return true;
  });
}

function borrarVacPendiente(ses, id) {
  admin_(ses);
  return conLock_(() => {
    const x = leer_('VacPendientes').find(r => r.Id === id);
    if (!x) throw new Error('No se encontró la carga.');
    borrarFilas_('VacPendientes', [x._fila]);
    slack_('🗑️ *' + quien_(ses) + '* borró los ' + x.Dias + ' días pendientes de ' + x.Empleado + ' (temporada ' + x.Temporada + ')');
    return true;
  });
}

function borrarMovimiento(ses, id) {
  admin_(ses);
  return conLock_(() => {
    const m = leer_('Banco').find(x => x.Id === id);
    if (!m) throw new Error('No se encontró el movimiento.');
    if (m.Pedido) throw new Error('Este día salió de un pedido de franco aprobado. Para devolverlo, volvé el pedido a pendiente o rechazalo.');
    borrarFilas_('Banco', [m._fila]);
    slack_('🗑️ *' + quien_(ses) + '* borró un movimiento del banco de ' + m.Empleado + ': ' + (m.Movimiento === 'Crédito' ? '+' : '−') + m.Dias + ' · ' + m.Motivo + ' · ' + fechaTxt_(m.Fecha));
    return true;
  });
}

/** Genera una contraseña temporal nueva (la anterior deja de funcionar y se cierran sus sesiones). */
function resetearContrasena(ses, tipo, id) {
  admin_(ses);
  return conLock_(() => {
    const hoja = tipo === 'admin' ? 'Admins' : 'Nomina';
    const fila = leer_(hoja).find(x => (tipo === 'admin' ? x.Usuario : x.Clave) === id);
    if (!fila) throw new Error('No se encontró el usuario.');
    const t = passTemporal_();
    escribir_(hoja, fila._fila, { TempPass: t, PassHash: '', Salt: '' });
    cerrarSesionesDe_(tipo === 'admin' ? 'a' : 'e', id);
    slack_('🔑 *' + quien_(ses) + '* generó una contraseña temporal nueva para ' + (tipo === 'admin' ? 'el usuario de RRHH `' + id + '`' : '*' + nombreDe_(fila) + '*'));
    return t;
  });
}

function agregarAdmin(ses, usuario, nombre) {
  admin_(ses);
  usuario = String(usuario || '').trim();
  if (!/^[A-Za-z0-9._-]{3,30}$/.test(usuario)) throw new Error('El usuario tiene que tener entre 3 y 30 letras, números o puntos, sin espacios.');
  return conLock_(() => {
    if (buscarUsuario_(usuario)) throw new Error('Ese usuario ya existe.');
    const t = passTemporal_();
    agregar_('Admins', { Usuario: usuario, Nombre: String(nombre || '').trim(), TempPass: t });
    slack_('👤 *' + quien_(ses) + '* agregó el usuario de RRHH `' + usuario + '`' + (nombre ? ' (' + nombre + ')' : ''));
    return t;
  });
}

function quitarAdmin(ses, usuario) {
  const yo = admin_(ses);
  if (usuario === yo) throw new Error('No podés quitar tu propio usuario.');
  return conLock_(() => {
    const a = leer_('Admins').find(x => x.Usuario === usuario);
    if (!a) throw new Error('No se encontró el usuario.');
    ss_().getSheetByName('Admins').deleteRow(a._fila);
    cerrarSesionesDe_('a', usuario);
    slack_('👤 *' + yo + '* quitó el usuario de RRHH `' + usuario + '`');
    return true;
  });
}

function guardarAjustes(ses, topePct, encargados) {
  admin_(ses);
  const pc = Number(topePct);
  if (!(pc > 0 && pc <= 100)) throw new Error('El tope tiene que ser un porcentaje entre 1 y 100.');
  return conLock_(() => {
    const c = leer_('Config').find(x => x.Clave === 'topePct');
    escribir_('Config', c._fila, { Valor: pc });
    const sh = ss_().getSheetByName('Encargados');
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).clearContent();
    const filas = (encargados || []).filter(e => e.nombre && e.tel).map(e => [String(e.nombre).trim(), String(e.tel).trim()]);
    if (filas.length) sh.getRange(2, 1, filas.length, 2).setValues(filas);
    slack_('⚙️ *' + quien_(ses) + '* guardó los ajustes: tope de adelantos ' + pc + ' % · encargados: ' + (filas.map(f => f[0]).join(', ') || 'ninguno'));
    return true;
  });
}
