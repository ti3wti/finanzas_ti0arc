/**
 * TI0ARC – Sistema automatizado de finanzas y socios
 * Asociación de Radioaficionados de Cartago · cuenta en colones del Banco Nacional
 * creado por TI3WTI - William Marin
 *
 * Qué hace:
 *  1. Busca en Gmail los estados de cuenta del Banco Nacional (formato actual y antiguo),
 *     guarda cada PDF en Drive y extrae sus movimientos, verificando el cuadre de saldos.
 *  2. Mantiene las hojas "Resumen de finanzas", "Movimientos" (con notas editables) e
 *     "Informe" (cualquier período). Detecta meses sin estado de cuenta por diferencia de saldos.
 *  3. Lleva la hoja "Socios": detecta pagos de membresía en el banco y calcula el estado
 *     (Activo, Por confirmar, Vencido) según la configuración de la misma hoja.
 *  4. Atiende un formulario de Google: un socio activo pide el informe de un período y lo
 *     recibe en PDF, con marca de agua, solo en su correo registrado. Todo queda en la
 *     "Bitácora de solicitudes".
 *  5. Envía avisos a los correos de notificaciones de la Junta.
 *
 * Hojas ocultas: Estados, Datos y Notas. No las edite a mano: las demás se reconstruyen desde ellas.
 * Requiere el servicio avanzado Drive API (v3). Ver README.md.
 */

const CFG = {
  QUERY: 'subject:"Estado de cuenta del Banco Nacional de Costa Rica" has:attachment',
  CUENTA: '200-01-097-011427-0',          // se verifica dentro del PDF
  CARPETA: 'Estados de cuenta BN - TI0ARC',
  CORREO_AVISO: '',                        // opcional; los correos se configuran en la hoja Socios
};

const HOJA_RESUMEN = 'Resumen de finanzas';
const MESES = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Setiembre','Octubre','Noviembre','Diciembre'];
const NUM = '#,##0.00;[Red]-#,##0.00;"-"';
const NUM_NETO = '[Blue]#,##0.00;[Red]-#,##0.00;"-"';
const AZUL = '#2E3A87', AMARILLO = '#FFF2CC', TONOS = ['#DDEBF7', '#E2EFDA'], GRIS = '#F2F2F2';

const H_EST = ['Mes','Inicio','Fin','Saldo anterior','Ingresos','N.º ingresos','Egresos','N.º egresos',
               'Saldo final','Cuadre','PDF','MessageId','Procesado'];
const H_DAT = ['Mes','Fecha aplicación','Fecha transacción','Número','Descripción','Ingreso','Egreso',
               'Saldo banco','Fin del estado'];

/* ───────────────────────── Menú ───────────────────────── */

function onOpen() {
  try { colorPestanas_(SpreadsheetApp.getActive()); } catch (e) {}
  SpreadsheetApp.getUi().createMenu('TI0ARC')
    .addItem('Actualizar todo (revisa el correo)', 'actualizarTodo')
    .addItem('Actualizar estado de socios', 'actualizarSocios')
    .addItem('Crear o actualizar formulario de solicitudes', 'crearFormulario')
    .addSeparator()
    .addItem('Activar actualización diaria y envíos desde la hoja', 'instalarDisparador')
    .addToUi();
}

/* ───────────────────────── Proceso principal ───────────────────────── */

// Menú: revisa el correo y además rearma todas las hojas, haya o no estados nuevos
function actualizarTodo() {
  actualizar();
  reconstruir();
  SpreadsheetApp.getActive().toast('Listo: correo revisado y hojas actualizadas.', 'TI0ARC');
}

function actualizar() {
  const ss = SpreadsheetApp.getActive();
  PropertiesService.getScriptProperties().setProperty('ssId', ss.getId());
  const est = hoja_(ss, 'Estados', H_EST);
  const dat = hoja_(ss, 'Datos', H_DAT);
  const props = PropertiesService.getScriptProperties();
  // v2: acepta también los correos antiguos (BUZON_BN_EMAIL, adjunto "Estado de cuenta.pdf").
  // Al cambiar de versión se vuelven a revisar todos los correos; los estados ya importados no se duplican.
  if (props.getProperty('version') !== '2') { props.deleteProperty('vistos'); props.setProperty('version', '2'); }
  const vistos = new Set(JSON.parse(props.getProperty('vistos') || '[]'));
  const t0 = Date.now();
  let pendientes = false;

  // Fechas de estados ya importados (evita duplicados aunque llegue el mismo PDF dos veces)
  const finesYa = new Set();
  if (est.getLastRow() > 1) {
    est.getRange(2, 3, est.getLastRow() - 1, 1).getValues()
       .forEach(r => r[0] instanceof Date && finesYa.add(r[0].getTime()));
  }

  const carpeta = carpeta_();
  const nuevos = [], errores = [];

  GmailApp.search(CFG.QUERY, 0, 500).forEach(hilo => hilo.getMessages().forEach(msg => {
    const id = msg.getId();
    if (vistos.has(id) || pendientes) return;
    if (Date.now() - t0 > 4.5 * 60 * 1000) { pendientes = true; return; }   // límite de 6 min de Apps Script
    let ok = true;

    msg.getAttachments().forEach(att => {
      const nombre = att.getName();
      if (!/\.pdf$/i.test(nombre) || /glosario/i.test(nombre)) return;
      try {
        const r = leerEstado_(att.copyBlob());
        if (!r) return;                                          // PDF de otra cuenta (p. ej. la de dólares)
        if (finesYa.has(r.fin.getTime())) return;               // estado repetido
        const file = carpeta.createFile(att.copyBlob())
                            .setName('Estado BN ' + iso_(r.fin) + '.pdf');
        est.appendRow([r.mes, r.ini || '', r.fin, r.saldoAnt, r.ing, r.nIng, r.egr, r.nEgr,
                       r.saldoFin, r.cuadre ? 'OK' : 'REVISAR', file.getUrl(), id, new Date()]);
        if (r.movs.length) {
          dat.getRange(dat.getLastRow() + 1, 1, r.movs.length, H_DAT.length).setValues(
            r.movs.map(m => [r.mes, m.fApl, m.fTx || '', m.num, m.desc,
                             m.monto > 0 ? m.monto : '', m.monto < 0 ? -m.monto : '', m.saldo, r.fin]));
        }
        finesYa.add(r.fin.getTime());
        nuevos.push(r);
      } catch (e) {
        ok = false;
        errores.push(`${msg.getDate().toLocaleDateString('es-CR')} – ${nombre}: ${e.message}`);
      }
    });
    if (ok) vistos.add(id);   // con error se reintenta en la próxima ejecución
  }));

  props.setProperty('vistos', JSON.stringify([...vistos]));
  if (nuevos.length || !ss.getSheetByName(HOJA_RESUMEN)) reconstruir();
  actualizarSocios_(ss);
  colorPestanas_(ss);
  avisar_(nuevos, errores);
  Logger.log(`Estados nuevos: ${nuevos.length}. Errores: ${errores.length}`);
  if (pendientes) {
    Logger.log('Quedaron correos sin revisar por el límite de tiempo. Ejecute "actualizar" otra vez.');
  }
  errores.forEach(e => Logger.log(e));
}

/* ───────────────────────── Lectura del PDF ───────────────────────── */

function leerEstado_(blob) {
  return parsear_(textoPdf_(blob));
}

// Convierte el PDF a Google Docs para obtener el texto (requiere el servicio avanzado "Drive API")
function textoPdf_(blob) {
  const tmp = Drive.Files.create({ name: 'tmp_estado_bn', mimeType: MimeType.GOOGLE_DOCS }, blob);
  try {
    return DocumentApp.openById(tmp.id).getBody().getText();
  } finally {
    DriveApp.getFileById(tmp.id).setTrashed(true);
  }
}

function parsear_(raw) {
  const txt = raw.replace(/\s+/g, ' ');
  if (txt.indexOf(CFG.CUENTA) < 0 && txt.indexOf(CFG.CUENTA.replace(/-/g, '')) < 0) return null;

  // Formato actual: "Fecha éste estado: 02/01/2026". Formato antiguo: "FECHA ESTE ESTADO ... 03/12/2021"
  const mFin = txt.match(/Fecha [ée]ste estado:?[^\/]{0,80}?(\d{2}\/\d{2}\/\d{4})/i);
  const mIni = txt.match(/Fecha [úu]ltimo estado:?[^\/]{0,80}?(\d{2}\/\d{2}\/\d{4})/i);
  if (!mFin) throw new Error('No se encontró la fecha del estado');
  const fin = fecha_(mFin[1]), ini = mIni ? fecha_(mIni[1]) : null;

  // Saldo anterior y final del cuadro resumen (formato actual)
  let saldoAnt = null, saldoFin = null;
  const s = txt.match(/Cantidad Monto Cantidad Monto ([\d,]+\.\d{2}) (\d+) ([\d,]+\.\d{2}) (\d+) ([\d,]+\.\d{2}) ([\d,]+\.\d{2})/i);
  if (s) { saldoAnt = num_(s[1]); saldoFin = num_(s[6]); }

  // El formato antiguo trae la tabla por columnas (todas las fechas, luego todos los números...)
  const L = raw.split(/\r?\n/).map(x => x.trim());
  const porColumnas = L.some((x, i) => /^\d{2}-\d{2}$/.test(x) && /^\d{2}-\d{2}$/.test(sigNoVacia_(L, i + 1).v));

  let filas, cuadreCol = true;
  if (porColumnas) {
    if (saldoAnt === null) saldoAnt = saldoAnteriorCol_(L);
    if (saldoAnt === null) throw new Error('Formato antiguo: no se encontró el saldo anterior');
    const c = columnas_(L, saldoAnt, txt);
    filas = c.filas; cuadreCol = c.ok;
  } else {
    filas = [];
    const re = /(\d{2})-(\d{2}) (\d{1,10}) (.+?) ([\d,]+\.\d{2})([+-]) ([\d,]+\.\d{2})/g;
    let m;
    while ((m = re.exec(txt)) !== null) {
      filas.push({ dd: +m[1], mm: +m[2], num: m[3], desc: m[4].trim(),
                   monto: (m[6] === '-' ? -1 : 1) * num_(m[5]), saldo: num_(m[7]) });
    }
  }

  const movs = filas.map(f => {
    let anio = fin.getFullYear();
    if (f.mm > fin.getMonth() + 1) anio--;              // diciembre dentro de un estado de enero
    const t = f.desc.match(/^(\d{2})-(\d{2})-(\d{2}) /); // fecha real de la transacción, si viene
    return { fApl: new Date(anio, f.mm - 1, f.dd), fTx: t ? new Date(2000 + +t[3], +t[2] - 1, +t[1]) : null,
             num: f.num, desc: f.desc, monto: f.monto, saldo: f.saldo };
  });

  // Estado sin movimientos (solo trae saldos): se toma el saldo anterior del cuadro resumen
  let sinMovs = false;
  if (saldoAnt === null && !movs.length) {
    saldoAnt = saldoAnteriorCol_(L);
    if (saldoAnt !== null) {
      sinMovs = true;
      // si el PDF parece tener montos con signo pero no se leyó ninguno, se marca para revisar
      if (/[\d,]+\.\d{2}[+-]/.test(txt)) cuadreCol = false;
    }
  }
  if (saldoAnt === null) {
    if (!movs.length) throw new Error('No se pudieron leer saldos ni movimientos');
    saldoAnt = r2_(movs[0].saldo - movs[0].monto);
  }
  if (saldoFin === null) saldoFin = movs.length ? movs[movs.length - 1].saldo : saldoAnt;

  // Verificación: cada saldo diario debe cuadrar con el anterior
  let corrido = saldoAnt, cuadre = cuadreCol;
  movs.forEach(x => { corrido = r2_(corrido + x.monto); if (Math.abs(corrido - x.saldo) > 0.004) cuadre = false; });
  if (Math.abs(corrido - saldoFin) > 0.004) cuadre = false;

  const ing = movs.filter(x => x.monto > 0), egr = movs.filter(x => x.monto < 0);
  return {
    mes: etiqueta_(masDias_(fin, -15)), ini, fin, saldoAnt, saldoFin, movs, cuadre,
    ing: r2_(ing.reduce((a, x) => a + x.monto, 0)), nIng: ing.length,
    egr: r2_(-egr.reduce((a, x) => a + x.monto, 0)), nEgr: egr.length,
  };
}

/* ── Lectura del formato antiguo (por columnas) ── */

const RE_MONTO = /^[\d,]+\.\d{2}$/;
const ENCABEZADOS = /^(FECHA|NUMERO|N[ÚU]MERO|DESCRIPCI[ÓO]N|MONTO|SALDO.*|DEBITOS|D[ÉE]BITOS|CREDITOS|CR[ÉE]DITOS|CANTIDAD|SOBREGIRO.*)$/i;

function sigNoVacia_(L, i) {
  while (i < L.length && !L[i]) i++;
  return { i, v: i < L.length ? L[i] : '' };
}

// Lee N líneas no vacías seguidas que cumplan "re", empezando en la primera coincidencia desde "desde"
function corrida_(L, desde, re, N) {
  for (let k = desde; k < L.length; k++) {
    if (!re.test(L[k])) continue;
    const vals = [];
    let j = k;
    while (j < L.length && vals.length < N) {
      if (!L[j]) { j++; continue; }
      if (!re.test(L[j])) break;
      vals.push(L[j]); j++;
    }
    if (vals.length === N) return { vals, fin: j };
  }
  return null;
}

function saldoAnteriorCol_(L) {
  const i = L.findIndex(x => /^SALDO ANTERIOR$/i.test(x));
  if (i < 0) return null;
  for (let k = i + 1; k < Math.min(L.length, i + 15); k++) if (RE_MONTO.test(L[k])) return num_(L[k]);
  return null;
}

function columnas_(L, saldoAnt, txt) {
  const filas = [];
  let pos = 0, prev = saldoAnt, ok = true;
  while (true) {
    // 1) fechas
    const k0 = L.findIndex((x, i) => i >= pos && /^\d{2}-\d{2}$/.test(x));
    if (k0 < 0) break;
    const fechas = [];
    let j = k0;
    while (j < L.length) {
      if (!L[j]) { j++; continue; }
      if (!/^\d{2}-\d{2}$/.test(L[j])) break;
      fechas.push(L[j]); j++;
    }
    const N = fechas.length;
    // 2) números de transacción
    const nums = corrida_(L, j, /^\d{3,10}$/, N);
    if (!nums) { ok = false; break; }
    // 3) descripciones: líneas que empiezan con fecha o contienen "/" abren una nueva; las demás continúan la anterior
    const desc = [];
    let k = nums.fin;
    while (k < L.length && !(L[k] && (/\//.test(L[k]) || /^\d{2}-\d{2}-\d{2} /.test(L[k])) && !ENCABEZADOS.test(L[k]))) k++;
    for (; k < L.length; k++) {
      const x = L[k];
      if (!x) continue;
      if (ENCABEZADOS.test(x) || RE_MONTO.test(x) || /^\d{2}-\d{2}$/.test(x)) break;
      if (/\//.test(x) || /^\d{2}-\d{2}-\d{2} /.test(x) || !desc.length) desc.push(x);
      else desc[desc.length - 1] += ' ' + x;
    }
    // 4) saldo diario (N montos seguidos)
    const sal = corrida_(L, k, RE_MONTO, N);
    if (!sal) { ok = false; break; }
    if (desc.length !== N) ok = false;
    for (let n = 0; n < N; n++) {
      const [dd, mm] = fechas[n].split('-').map(Number);
      const saldo = num_(sal.vals[n]);
      filas.push({ dd, mm, num: nums.vals[n], desc: desc.length === N ? desc[n] : '(descripción no disponible)',
                   monto: r2_(saldo - prev), saldo });
      prev = saldo;
    }
    pos = sal.fin;
  }
  // Control: el primer monto calculado debe aparecer como tal en el PDF
  if (filas.length && txt.indexOf(fmt_(Math.abs(filas[0].monto))) < 0) ok = false;
  return { filas, ok };
}

/* ───────────────────────── Hojas Resumen y Movimientos ───────────────────────── */

function reconstruir() {
  const ss = SpreadsheetApp.getActive();
  const B = bloques_(ss);
  if (!B.length) return;
  resumen_(ss, B);
  movimientos_(ss, B);
  informe_(ss, B, true);
  actualizarSocios_(ss);
  actualizarMesesFormulario_(ss);
  colorPestanas_(ss);
}

function bloques_(ss) {
  const est = hoja_(ss, 'Estados', H_EST), dat = hoja_(ss, 'Datos', H_DAT);
  if (est.getLastRow() < 2) return [];

  const E = est.getRange(2, 1, est.getLastRow() - 1, H_EST.length).getValues()
               .sort((a, b) => a[2] - b[2]);
  const D = dat.getLastRow() > 1 ? dat.getRange(2, 1, dat.getLastRow() - 1, H_DAT.length).getValues() : [];
  const porFin = {};
  D.forEach(r => (porFin[r[8].getTime()] = porFin[r[8].getTime()] || []).push(r));

  // Bloques por mes, insertando los meses faltantes (cuando el saldo no empata)
  const B = [];
  let prev = null;
  E.forEach(e => {
    const [, ini, fin, sAnt, , , , , sFin] = e;
    const mes = etiqueta_(masDias_(fin, -15));
    if (prev && Math.abs(sAnt - prev.sFin) > 0.004) {
      const a = etiqueta_(masDias_(prev.fin, 15));
      const b = etiqueta_(masDias_(ini instanceof Date ? ini : masDias_(fin, -30), -15));
      B.push({ mes: rango_(a, b), periodo: 'Sin estado de cuenta', sIni: prev.sFin,
               neto: r2_(sAnt - prev.sFin), sBanco: sAnt, movs: null });
    }
    B.push({ mes, periodo: (ini instanceof Date ? dmy_(ini) : '?') + ' – ' + dmy_(fin),
             sIni: sAnt, sBanco: sFin, movs: ordenar_(porFin[fin.getTime()] || []) });
    prev = { fin, sFin };
  });
  return B;
}

function resumen_(ss, B) {
  const vieja = ss.getSheetByName('Resumen');
  if (vieja && !ss.getSheetByName(HOJA_RESUMEN)) vieja.setName(HOJA_RESUMEN);
  const sh = limpia_(ss, HOJA_RESUMEN, 0);
  const primero = B[0].mes, ultimo = B[B.length - 1].mes;
  sh.getRange('A1').setValue(`TI0ARC — Cuenta BN ${CFG.CUENTA} · Resumen mensual ${primero} – ${ultimo} (colones)`)
    .setFontWeight('bold').setFontSize(13);
  const H = ['Mes','Período','Saldo inicial','Ingresos','Egresos','N.º mov.','Neto sin detalle',
             'Saldo calculado','Saldo según banco','Diferencia'];
  encabezado_(sh.getRange(2, 1, 1, H.length).setValues([H]));

  const filas = [], fondos = [];
  B.forEach((b, i) => {
    const r = 3 + i, falta = !b.movs;
    const ing = falta ? '' : r2_(b.movs.reduce((a, x) => a + (+x[5] || 0), 0));
    const egr = falta ? '' : r2_(b.movs.reduce((a, x) => a + (+x[6] || 0), 0));
    filas.push([b.mes, b.periodo, i === 0 ? b.sIni : `=H${r - 1}`, ing, egr, falta ? '' : b.movs.length,
                falta ? b.neto : 0, `=C${r}+D${r}-E${r}+G${r}`, b.sBanco, `=H${r}-I${r}`]);
    fondos.push(Array(H.length).fill(falta ? AMARILLO : null));
  });
  const n = filas.length, t = 3 + n;
  sh.getRange(3, 1, n + 1, 2).setNumberFormat('@');
  sh.getRange(3, 1, n, H.length).setValues(filas).setBackgrounds(fondos);
  sh.getRange(t, 1, 1, H.length).setValues([['TOTAL', '', '', `=SUM(D3:D${t - 1})`, `=SUM(E3:E${t - 1})`,
    `=SUM(F3:F${t - 1})`, `=SUM(G3:G${t - 1})`, `=H${t - 1}`, '', '']])
    .setFontWeight('bold').setBackground(GRIS);
  sh.getRange(3, 3, n + 1, 8).setNumberFormat(NUM);
  sh.getRange(3, 6, n + 1, 1).setNumberFormat('0');
  sh.getRange(3, 7, n + 1, 1).setNumberFormat(NUM_NETO);
  sh.getRange(2, 1, n + 2, H.length).setBorder(true, true, true, true, true, true, '#BBBBBB', null);

  const notas = ['Notas:',
    '• Cada mes corresponde al período de corte del estado de cuenta del BN.',
    '• Filas en amarillo: no hay estado de cuenta. "Neto sin detalle" es lo que cambió el saldo en ese lapso (ingresos menos egresos), calculado por diferencia de saldos: azul = aumentó, rojo = disminuyó.',
    '• Saldo calculado = saldo inicial + ingresos − egresos + neto sin detalle. Saldo según banco = el que reporta el estado de cuenta.',
    '• Diferencia = Saldo calculado − Saldo según banco. Debe ser 0: confirma que no falta ni sobra ningún movimiento.',
    '• Hoja generada automáticamente. Última actualización: ' +
      Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')];
  sh.getRange(t + 2, 1, notas.length, 1).setValues(notas.map(x => [x])).setFontSize(9);
  sh.getRange(t + 2, 1).setFontWeight('bold');
  [90, 190, 115, 115, 115, 70, 115, 115, 125, 85].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.setFrozenRows(2);
}

function movimientos_(ss, B) {
  const sh = limpia_(ss, 'Movimientos');
  const notas = leerNotas_(ss);
  const H = ['Mes','Fecha aplicación','Fecha transacción','Número','Descripción','Ingreso','Egreso','Saldo',
             'Nota ✎','','Resumen del mes','Ingresos','Egresos','Saldo en cuenta'];
  encabezado_(sh.getRange(1, 1, 1, 9).setValues([H.slice(0, 9)]));
  encabezado_(sh.getRange(1, 11, 1, 4).setValues([H.slice(10)]));
  sh.getRange(1, 9).setBackground('#B45F06');

  const filas = [['Saldo inicial', '', '', '', 'Saldo anterior al primer estado', '', '', B[0].sIni, '', '', '', '', '', '']];
  const fondos = [Array(9).fill(GRIS).concat(Array(5).fill(null))];
  const finales = [];
  let r = 3;

  B.forEach((b, k) => {
    const color = b.movs ? TONOS[k % 2] : AMARILLO, a = r;
    const lineas = b.movs
      ? b.movs.map(x => [b.mes, x[1], x[2], String(x[3]), x[4], x[5], x[6], null, notas[clave_(x[3], x[1])] || ''])
      : [[b.mes, '', '', '', `Sin estado de cuenta — neto ₡${fmt_(b.neto)} por diferencia de saldos`, '', '', null, '']];
    const z = a + lineas.length - 1;
    lineas.forEach((l, j) => {
      const fila = r;
      l[7] = b.movs ? `=H${fila - 1}+F${fila}-G${fila}` : b.sBanco;
      const mini = j === 0
        ? [b.mes, b.movs ? `=SUM(F${a}:F${z})` : 'Sin detalle', b.movs ? `=SUM(G${a}:G${z})` : 'Sin detalle', `=H${z}`]
        : ['', '', '', ''];
      filas.push(l.concat([''], mini));
      fondos.push(Array(9).fill(color).concat([null], Array(4).fill(j === 0 ? color : null)));
      r++;
    });
    finales.push(z);
  });

  const n = filas.length, ult = 1 + n;
  sh.getRange(2, 1, n, 1).setNumberFormat('@');
  sh.getRange(2, 4, n, 1).setNumberFormat('@');
  sh.getRange(2, 9, n, 1).setNumberFormat('@');
  sh.getRange(2, 11, n, 1).setNumberFormat('@');
  sh.getRange(2, 1, n, 14).setValues(filas).setBackgrounds(fondos).setFontFamily('Arial');
  sh.getRange(2, 2, n, 2).setNumberFormat('dd/mm/yyyy');
  sh.getRange(2, 6, n + 1, 3).setNumberFormat(NUM);
  sh.getRange(2, 9, n, 1).setFontColor('#7F3F00').setFontStyle('italic').setWrap(true);
  sh.getRange(2, 12, n, 3).setNumberFormat(NUM).setHorizontalAlignment('right');
  sh.getRange(2, 11, n, 1).setFontWeight('bold');
  sh.getRange(2, 14, n, 1).setFontWeight('bold');
  sh.getRange(2, 1, n, 9).setBorder(true, true, true, true, true, true, '#BBBBBB', null);
  finales.forEach(z => sh.getRange(z, 1, 1, 9)
    .setBorder(null, null, true, null, null, null, AZUL, SpreadsheetApp.BorderStyle.SOLID_MEDIUM));

  sh.getRange(ult + 1, 5, 1, 3).setValues([['TOTAL', `=SUM(F2:F${ult})`, `=SUM(G2:G${ult})`]])
    .setFontWeight('bold').setNumberFormat(NUM);
  sh.getRange(ult + 1, 5).setNumberFormat('@');

  [90, 105, 110, 95, 400, 105, 105, 115, 260, 20, 115, 105, 105, 120].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.setFrozenRows(1);
  sh.getRange(1, 1, ult, 9).createFilter();
}

/* ───────────────────────── Notas ───────────────────────── */

function clave_(num, fecha) {
  return String(num).trim() + '|' + (fecha instanceof Date ? iso_(fecha) : String(fecha));
}

function leerNotas_(ss) {
  const sh = hoja_(ss, 'Notas', ['Clave', 'Nota', 'Modificado']);
  const out = {};
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(r => { if (r[0] && r[1] !== '') out[r[0]] = r[1]; });
  }
  return out;
}

// Inserta o actualiza notas en la hoja oculta "Notas". pares = [[clave, nota], ...]
function guardarPares_(ss, pares) {
  if (!pares.length) return;
  const sh = hoja_(ss, 'Notas', ['Clave', 'Nota', 'Modificado']);
  const ult = sh.getLastRow();
  const claves = ult > 1 ? sh.getRange(2, 1, ult - 1, 1).getValues().map(r => r[0]) : [];
  pares.forEach(([k, nota]) => {
    const i = claves.indexOf(k);
    if (i >= 0) sh.getRange(i + 2, 2, 1, 2).setValues([[nota, new Date()]]);
    else if (nota) { sh.appendRow([k, nota, new Date()]); claves.push(k); }
  });
}

// Nota escrita en Movimientos (columna I): se guarda y se refleja en Informe
function notaDesdeMovimientos_(ss, rango) {
  const fila0 = Math.max(rango.getRow(), 2), filas = rango.getLastRow() - fila0 + 1;
  if (filas < 1) return;
  const pares = [];
  rango.getSheet().getRange(fila0, 1, filas, 9).getValues().forEach(v => {
    if (v[3] && v[1] instanceof Date) pares.push([clave_(v[3], v[1]), String(v[8] || '').trim()]);
  });
  if (!pares.length) return;
  guardarPares_(ss, pares);
  const B = bloques_(ss);
  if (B.length) informe_(ss, B, true);     // actualiza Informe sin quitarle el foco a Movimientos
}

// Nota escrita en Informe (columna H de la tabla de movimientos): se guarda y se refleja en Movimientos
function notaDesdeInforme_(ss, rango) {
  const fila0 = rango.getRow(), filas = rango.getLastRow() - fila0 + 1;
  const pares = [];
  rango.getSheet().getRange(fila0, 8, filas, 2).getValues().forEach(v => {
    if (v[1]) pares.push([String(v[1]), String(v[0] || '').trim()]);      // columna I oculta = clave
  });
  if (!pares.length) return;
  guardarPares_(ss, pares);
  const mov = ss.getSheetByName('Movimientos');
  if (!mov || mov.getLastRow() < 2) return;
  const n = mov.getLastRow() - 1;
  const v = mov.getRange(2, 1, n, 9).getValues();
  const mapa = Object.fromEntries(pares);
  v.forEach((x, i) => {
    if (x[3] && x[1] instanceof Date) {
      const k = clave_(x[3], x[1]);
      if (k in mapa) mov.getRange(i + 2, 9).setValue(mapa[k]);
    }
  });
}

/* ───────────────────────── Informe por período ───────────────────────── */

// Se ejecuta solo: al cambiar "Desde"/"Hasta" en Informe, al escribir una Nota, o al editar la hoja Socios
function onEdit(e) {
  const r = e && e.range;
  if (!r) return;
  const hoja = r.getSheet().getName();
  if (hoja === 'Informe' && r.getColumn() === 2 && (r.getRow() === 3 || r.getRow() === 4)) {
    generarInforme();
  } else if (hoja === 'Informe' && r.getColumn() <= 8 && r.getLastColumn() >= 8 && r.getRow() > 4) {
    notaDesdeInforme_(SpreadsheetApp.getActive(), r);
  } else if (hoja === 'Movimientos' && r.getColumn() <= 9 && r.getLastColumn() >= 9) {
    notaDesdeMovimientos_(SpreadsheetApp.getActive(), r);
  } else if (hoja === 'Socios' && r.getRow() > 1 &&
             ([SC.ind + 1, SC.nom + 1, SC.per + 1].includes(r.getColumn()) || r.getColumn() >= COL_CFG)) {
    actualizarSocios_(SpreadsheetApp.getActive(), false);   // sin reordenar mientras se edita
  }
}

function generarInforme() {
  const ss = SpreadsheetApp.getActive();
  const B = bloques_(ss);
  if (B.length) informe_(ss, B, false);
}

function informe_(ss, B, conservar) {
  let sh = ss.getSheetByName('Informe') || ss.getSheetByName('Hoja 1') || ss.getSheetByName('Hoja1');
  if (!sh) sh = ss.insertSheet('Informe', 0);
  if (sh.getName() !== 'Informe') sh.setName('Informe');

  const meses = B.map(b => b.mes);
  let desde = String(sh.getRange('B3').getDisplayValue() || '');
  let hasta = String(sh.getRange('B4').getDisplayValue() || '');
  if (!meses.includes(desde)) desde = meses[Math.max(0, meses.length - 12)];   // por defecto: últimos 12 meses
  if (!meses.includes(hasta)) hasta = meses[meses.length - 1];
  let i0 = meses.indexOf(desde), i1 = meses.indexOf(hasta);
  if (i0 > i1) [i0, i1] = [i1, i0];
  const sel = B.slice(i0, i1 + 1);

  sh.getRange('A6:Z' + Math.max(sh.getMaxRows(), 6)).clear();
  sh.getRange('A1:Z4').clear();
  sh.getRange('A1').setValue('TI0ARC — Informe financiero por período').setFontWeight('bold').setFontSize(14);
  sh.getRange('A2').setValue(`Cuenta BN ${CFG.CUENTA} · montos en colones`).setFontSize(9).setFontColor('#666666');

  const regla = SpreadsheetApp.newDataValidation().requireValueInList(meses, true).setAllowInvalid(false).build();
  sh.getRange('A3:A4').setValues([['Desde:'], ['Hasta:']]).setFontWeight('bold');
  sh.getRange('B3:B4').setNumberFormat('@').setValues([[meses[i0]], [meses[i1]]])
    .setDataValidation(regla).setBackground(AMARILLO).setFontWeight('bold').setBorder(true, true, true, true, true, true);
  sh.getRange('C3').setValue('← elija los meses; el informe se actualiza solo').setFontSize(9).setFontColor('#666666');

  // Totales del período
  const ing = r2_(sel.reduce((a, b) => a + (b.movs ? b.movs.reduce((x, m) => x + (+m[5] || 0), 0) : 0), 0));
  const egr = r2_(sel.reduce((a, b) => a + (b.movs ? b.movs.reduce((x, m) => x + (+m[6] || 0), 0) : 0), 0));
  const neto = r2_(sel.reduce((a, b) => a + (b.movs ? 0 : b.neto), 0));
  const nMov = sel.reduce((a, b) => a + (b.movs ? b.movs.length : 0), 0);
  const sIni = sel[0].sIni, sFin = sel[sel.length - 1].sBanco;
  const faltan = sel.filter(b => !b.movs).map(b => b.mes);

  let r = 6;
  sh.getRange(r, 1).setValue('RESUMEN DEL PERÍODO').setFontWeight('bold').setFontColor(AZUL);
  const tot = [
    ['Saldo al inicio', sIni],
    ['(+) Ingresos', ing],
    ['(−) Egresos', egr],
    ['(±) Neto sin detalle (meses sin estado)', neto],
    ['Saldo al cierre', sFin],
    ['Variación del período', r2_(sFin - sIni)],
    ['Movimientos registrados', nMov],
  ];
  sh.getRange(r + 1, 1, tot.length, 2).setValues(tot).setBorder(true, true, true, true, true, true, '#BBBBBB', null);
  sh.getRange(r + 1, 2, tot.length - 1, 1).setNumberFormat(NUM);
  sh.getRange(r + 4, 2).setNumberFormat(NUM_NETO);
  sh.getRange(r + 6, 2).setNumberFormat(NUM_NETO);
  sh.getRange(r + 7, 2).setNumberFormat('0');
  [r + 1, r + 5].forEach(x => sh.getRange(x, 1, 1, 2).setFontWeight('bold').setBackground(GRIS));
  r += tot.length + 1;
  if (faltan.length) {
    sh.getRange(r, 1).setValue('Sin estado de cuenta en el período: ' + faltan.join(', '))
      .setFontSize(9).setFontColor('#B45F06');
  }
  r += 2;

  // Tabla mensual
  sh.getRange(r, 1).setValue('DETALLE POR MES').setFontWeight('bold').setFontColor(AZUL);
  r++;
  const HM = ['Mes', 'Período', 'Saldo inicial', 'Ingresos', 'Egresos', 'Neto sin detalle', 'Saldo final'];
  encabezado_(sh.getRange(r, 1, 1, HM.length).setValues([HM]));
  const filasM = sel.map(b => [b.mes, b.periodo, b.sIni,
    b.movs ? r2_(b.movs.reduce((x, m) => x + (+m[5] || 0), 0)) : '',
    b.movs ? r2_(b.movs.reduce((x, m) => x + (+m[6] || 0), 0)) : '',
    b.movs ? '' : b.neto, b.sBanco]);
  filasM.push(['TOTAL', '', '', ing, egr, neto, '']);
  const rm = sh.getRange(r + 1, 1, filasM.length, HM.length);
  sh.getRange(r + 1, 1, filasM.length, 2).setNumberFormat('@');
  rm.setValues(filasM).setBorder(true, true, true, true, true, true, '#BBBBBB', null);
  sh.getRange(r + 1, 3, filasM.length, 5).setNumberFormat(NUM);
  sh.getRange(r + 1, 6, filasM.length, 1).setNumberFormat(NUM_NETO);
  sel.forEach((b, k) => sh.getRange(r + 1 + k, 1, 1, HM.length).setBackground(b.movs ? TONOS[k % 2] : AMARILLO));
  sh.getRange(r + filasM.length, 1, 1, HM.length).setFontWeight('bold').setBackground(GRIS);
  r += filasM.length + 3;

  // Movimientos del período
  sh.getRange(r, 1).setValue('MOVIMIENTOS DEL PERÍODO').setFontWeight('bold').setFontColor(AZUL);
  r++;
  const notas = leerNotas_(ss);
  const HD = ['Mes', 'Fecha', 'Número', 'Descripción', 'Ingreso', 'Egreso', 'Saldo', 'Nota ✎', 'Clave'];
  encabezado_(sh.getRange(r, 1, 1, HD.length).setValues([HD]));
  const filasD = [], fondos = [];
  sel.forEach((b, k) => {
    const c = b.movs ? TONOS[k % 2] : AMARILLO;
    let saldo = b.sIni;
    (b.movs ? b.movs.map(m => [b.mes, m[2] || m[1], String(m[3]), m[4], m[5], m[6],
                               (saldo = r2_(saldo + (+m[5] || 0) - (+m[6] || 0))), notas[clave_(m[3], m[1])] || '', clave_(m[3], m[1])])
            : [[b.mes, '', '', `Sin estado de cuenta — neto ₡${fmt_(b.neto)}`, '', '', b.sBanco, '', '']])
      .forEach(f => { filasD.push(f); fondos.push(Array(HD.length).fill(c)); });
  });
  if (filasD.length) {
    sh.getRange(r + 1, 1, filasD.length, 1).setNumberFormat('@');
    sh.getRange(r + 1, 3, filasD.length, 1).setNumberFormat('@');
    sh.getRange(r + 1, 1, filasD.length, HD.length).setValues(filasD).setBackgrounds(fondos)
      .setBorder(true, true, true, true, true, true, '#BBBBBB', null);
    sh.getRange(r + 1, 2, filasD.length, 1).setNumberFormat('dd/mm/yyyy');
    sh.getRange(r + 1, 5, filasD.length, 3).setNumberFormat(NUM);
    sh.getRange(r + 1, 8, filasD.length, 1).setFontColor('#7F3F00').setFontStyle('italic').setWrap(true);
  }
  sh.getRange(r, 8).setBackground('#B45F06');
  sh.hideColumns(9);

  [150, 175, 105, 330, 110, 110, 115, 260].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.setColumnWidth(4, 360);
  if (!conservar) ss.setActiveSheet(sh);
}

/* ───────────────────────── Socios ───────────────────────── */

const H_SOC = ['Indicativo ✎', 'Estado', 'Enviar informe ☐', 'Nombre ✎', 'Observaciones ✎',
  'Último pago detectado', 'Monto', 'Detalle del pago', 'Período pagado ✎',
  'Correo ✎', 'Teléfono ✎', 'Última solicitud', 'N.º solicitudes'];
// Posición (base 0) de cada dato en una fila de Socios
const SC = { ind: 0, est: 1, env: 2, nom: 3, obs: 4, fecha: 5, monto: 6, det: 7, per: 8, mail: 9, tel: 10, sol: 11, nsol: 12 };
const EDITABLES = [SC.ind, SC.env, SC.nom, SC.obs, SC.per, SC.mail, SC.tel].map(i => i + 1);   // columnas editables
// Arma una fila de Socios a partir de un objeto {ind, est, nom, ...}
const filaSocio_ = o => H_SOC.map((_, i) => { const k = Object.keys(SC).find(x => SC[x] === i); return o[k] === undefined ? '' : o[k]; });
const COL_CFG = 15;                                  // columna O: configuración

// Lista inicial de socios: solo se usa si la hoja Socios no existe. Se deja vacía para no publicar
// datos personales en el código; los socios se registran directamente en la hoja.
// Formato de cada fila: ['TI0XXX', 'Nombre completo', 'correo@ejemplo.com', '8888-8888', 'Observaciones'],
const SOCIOS_INICIALES = [
];

const CFG_FILAS = [
  ['CONFIGURACIÓN', ''],
  ['Cuota de membresía (₡)', 10000],
  ['Lista de socios vigente hasta', '31/10/2027'],
  ['Mes de corte (1-12)', 8],
  ['Máximo de solicitudes por día', 1],
  ['WhatsApp TI0ARC', '8307-6432'],
  ['Enlace del formulario', ''],
  ['Informe de la Junta: desde', ''],
  ['Informe de la Junta: hasta', ''],
  ['Correo de notificaciones 1', ''],
  ['Correo de notificaciones 2', ''],
  ['Correo de notificaciones 3', ''],
  ['Correo de notificaciones 4', ''],
];
const CFG_ALIAS = {};

// Libro de la Asociación, también cuando el script corre desde la página web o un disparador
function ssApp_() {
  const id = PropertiesService.getScriptProperties().getProperty('ssId');
  if (id) { try { return SpreadsheetApp.openById(id); } catch (e) {} }
  return SpreadsheetApp.getActive();
}

// Correos de notificaciones, separados por coma; '' si no hay
function avisos_() {
  try {
    const lista = leerConfig_(ssApp_()).notif;
    if (CFG.CORREO_AVISO) lista.push(CFG.CORREO_AVISO);
    return lista.join(',');
  } catch (e) { return CFG.CORREO_AVISO || ''; }
}

function setCfg_(ss, etiqueta, valor) {
  const fila = CFG_FILAS.findIndex(f => f[0] === etiqueta) + 1;
  if (fila > 0) hojaSocios_(ss).getRange(fila, COL_CFG + 1).setValue(valor);
}

function hojaSocios_(ss) {
  let sh = ss.getSheetByName('Socios');
  if (sh) {
    const okEnc = sh.getRange(1, 1, 1, H_SOC.length).getValues()[0].join('|') === H_SOC.join('|');
    const etq = sh.getRange(1, COL_CFG, CFG_FILAS.length, 1).getValues().map(r => r[0]);
    if (okEnc && CFG_FILAS.every((f, i) => etq.at(i) === f[0])) return sh;
  }

  // Crear, o migrar desde una versión anterior de la hoja (se conservan datos y configuración)
  let filas = SOCIOS_INICIALES.map(x => filaSocio_({ ind: x[0], nom: x[1], mail: x[2], tel: x[3], obs: x[4] || '', env: false }));
  const cfgPrevia = {};
  if (sh) {
    const nc = Math.max(sh.getLastColumn(), 1), n = sh.getLastRow() - 1;
    const enc = sh.getRange(1, 1, 1, nc).getValues()[0].map(h => String(h).replace(' ✎', '').trim());
    const col = nombre => enc.indexOf(nombre);
    if (n > 0 && col('Indicativo') >= 0) {
      const v = (r, nombre) => col(nombre) >= 0 ? r[col(nombre)] : '';
      filas = sh.getRange(2, 1, n, nc).getValues().filter(r => String(r[0]).trim()).map(r => filaSocio_({
        ind: v(r, 'Indicativo'), nom: v(r, 'Nombre'), obs: v(r, 'Observaciones'), per: v(r, 'Período pagado'),
        mail: v(r, 'Correo'), tel: v(r, 'Teléfono'), sol: v(r, 'Última solicitud'), nsol: v(r, 'N.º solicitudes'), env: false }));
    }
    const ic = enc.indexOf('CONFIGURACIÓN');
    if (ic >= 0) {
      sh.getRange(1, ic + 1, 40, 2).getDisplayValues().forEach(r => {
        const k = CFG_ALIAS[r[0]] || r[0];
        if (k) cfgPrevia[k] = r[1];
      });
    }
    ss.deleteSheet(sh);
  }
  sh = ss.insertSheet('Socios');
  encabezado_(sh.getRange(1, 1, 1, H_SOC.length).setValues([H_SOC]));
  EDITABLES.forEach(c => sh.getRange(1, c).setBackground('#B45F06'));
  if (filas.length) {
    sh.getRange(2, 1, filas.length, 1).setNumberFormat('@');
    sh.getRange(2, SC.per + 1, filas.length, 1).setNumberFormat('@');
    sh.getRange(2, SC.tel + 1, filas.length, 1).setNumberFormat('@');
    sh.getRange(2, 1, filas.length, H_SOC.length).setValues(filas);
  }

  const cfg = CFG_FILAS.map((f, i) => (i > 0 && cfgPrevia[f[0]] !== undefined && cfgPrevia[f[0]] !== '') ? [f[0], cfgPrevia[f[0]]] : f);
  sh.getRange(1, COL_CFG, CFG_FILAS.length, 2).setNumberFormat('@').setValues(cfg)
    .setBorder(true, true, true, true, true, true, '#BBBBBB', null);
  encabezado_(sh.getRange(1, COL_CFG, 1, 2));
  sh.getRange(2, COL_CFG + 1, CFG_FILAS.length - 1, 1).setBackground(AMARILLO).setHorizontalAlignment('left');
  const fila = e => CFG_FILAS.findIndex(f => f[0] === e) + 1;
  sh.getRange(fila('Lista de socios vigente hasta'), COL_CFG + 1).setNote('Los socios de esta lista se consideran ' +
    'activos hasta esta fecha (fecha de la próxima asamblea). Después, solo quienes tengan un pago que los cubra. ' +
    'Actualícela en cada asamblea.');
  sh.getRange(fila('Mes de corte (1-12)'), COL_CFG + 1).setBackground('#F9CB9C').setNote('Pagos hechos a partir de ' +
    'este mes cuentan para el período siguiente. Valor provisional: definir con la Junta Directiva.');
  sh.getRange(fila('Informe de la Junta: desde'), COL_CFG + 1).setNote('Meses del informe que se envía al marcar ' +
    '"Enviar informe" en la fila de un socio. Se ponen solos en el último mes disponible cuando entra un estado nuevo.');

  [80, 100, 75, 230, 260, 110, 80, 300, 100, 220, 90, 120, 80, 20, 230, 140]
    .forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.setFrozenRows(1);
  sh.setFrozenColumns(1);                 // solo el indicativo: en el celular las columnas fijas se angostan
  return sh;
}

function leerConfig_(ss) {
  const sh = hojaSocios_(ss);
  const v = sh.getRange(1, COL_CFG, CFG_FILAS.length, 2).getDisplayValues();
  const m = {};
  v.forEach(r => m[r[0]] = String(r[1]).trim());
  const f = (m['Lista de socios vigente hasta'] || '31/10/2027').split('/').map(Number);
  const hasta = new Date(f[2] || 2027, (f[1] || 10) - 1, f[0] || 31);
  const correos = pref => Object.keys(m).filter(k => k.startsWith(pref) && /@/.test(m[k])).map(k => m[k].toLowerCase());
  return {
    cuota: num_(String(m['Cuota de membresía (₡)'] || '10000').replace(/[.\s]/g, '')) || 10000,
    hasta, vDia: hasta.getDate(), vMes: hasta.getMonth() + 1,
    corte: Number(m['Mes de corte (1-12)']) || 0,
    porDia: Number(m['Máximo de solicitudes por día']) || 1,
    whatsapp: m['WhatsApp TI0ARC'] || '8307-6432',
    formUrl: m['Enlace del formulario'] || '',
    juntaDesde: m['Informe de la Junta: desde'] || '',
    juntaHasta: m['Informe de la Junta: hasta'] || '',
    notif: correos('Correo de notificaciones'),
  };
}

// Fecha de vencimiento que corresponde a un pago hecho en "f"
function venceDe_(f, c) {
  const y = f.getFullYear(), v = new Date(y, c.vMes - 1, c.vDia);
  if (f > v) return new Date(y + 1, c.vMes - 1, c.vDia);
  if (c.corte && f.getMonth() + 1 >= c.corte) return new Date(y + 1, c.vMes - 1, c.vDia);
  return v;
}

const norm_ = s => String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase();
const KW_MEMB = /MEMB|ANUALIDAD|AFILIA|SOCIO|RENOV|INSCRIP/;
const KW_NO = /RIFA|N.?MERO|NUMS|TAMAL|DONACI|APORTA|CONTRIBU|EVENTO|FIELD|CAMISA|JARRA|STIKER/;
const indic_ = x => norm_(x).replace(/[^A-Z0-9]/g, '');

// Recalcula último pago, vencimiento y estado de cada socio, y ordena (pagos más recientes arriba)
function actualizarSocios_(ss, ordenar = true) {
  const sh = hojaSocios_(ss);
  const c = leerConfig_(ss);
  const n = sh.getLastRow() - 1;
  if (n < 1) return;
  const S = sh.getRange(2, 1, n, H_SOC.length).getValues().filter(r => String(r[0]).trim());
  const dat = ss.getSheetByName('Datos');
  const D = dat && dat.getLastRow() > 1 ? dat.getRange(2, 1, dat.getLastRow() - 1, H_DAT.length).getValues() : [];

  const socios = S.map(r => ({
    ind: indic_(r[0]),
    tokens: norm_(r[SC.nom]).split(/\s+/).filter(t => t.length >= 3),
    pago: null,
  }));

  D.forEach(d => {
    const monto = +d[5] || 0;
    if (monto <= 0) return;
    const desc = norm_(d[4]), compacto = desc.replace(/\s/g, '');
    if (!(KW_MEMB.test(desc) || (monto % c.cuota === 0 && !KW_NO.test(desc)))) return;
    const fecha = d[2] instanceof Date ? d[2] : d[1];
    if (!(fecha instanceof Date)) return;

    let quien = socios.filter(s => s.ind && compacto.indexOf(s.ind) >= 0);
    if (!quien.length) {                                     // por nombre del que paga (antes de "/")
      const pagador = desc.replace(/^\d{2}-\d{2}-\d{2}\s+/, '').split('/')[0].split(/\s+/).filter(t => t.length >= 3);
      quien = socios.filter(s => s.tokens.filter(t => pagador.includes(t)).length >= 2);
      if (quien.length !== 1) return;
    }
    quien.forEach(s => {
      if (!s.pago || fecha > s.pago.fecha) s.pago = { fecha, monto, desc: d[4] };
    });
  });

  const hoy = new Date(); hoy.setHours(0, 0, 0, 0);
  const finActual = c.hasta.getFullYear();               // ej. 2027 → período vigente 2026-2027
  const filas = S.map((r, i) => {
    const p = socios[i].pago;
    // Período: el mayor entre el detectado por el banco y el escrito a mano en la celda
    const finPago = p ? venceDe_(p.fecha, c).getFullYear() : 0;
    const m = String(r[SC.per] || '').match(/(\d{4})\D+(\d{4})/);
    const fin = Math.max(finPago, m ? +m[2] : 0);
    const vence = fin ? new Date(fin, c.vMes - 1, c.vDia) : null;
    // Activo: su período pagado lo cubre hoy. Por confirmar: está en la lista validada en asamblea
    // pero aún no se ve su pago. Vencido: ni lo uno ni lo otro.
    const estado = (vence && vence >= hoy) ? 'Activo' : (hoy <= c.hasta ? 'Por confirmar' : 'Vencido');
    return filaSocio_({ ind: String(r[SC.ind]).trim().toUpperCase(), est: estado, env: false, nom: r[SC.nom], obs: r[SC.obs],
      fecha: p ? p.fecha : '', monto: p ? p.monto : '', det: p ? p.desc : '', per: fin ? `${fin - 1}-${fin}` : '',
      mail: r[SC.mail], tel: r[SC.tel], sol: r[SC.sol], nsol: r[SC.nsol] });
  });
  const t = f => f[SC.fecha] instanceof Date ? f[SC.fecha].getTime() : 0;
  if (ordenar) filas.sort((a, b) => t(b) - t(a) || (a[0] < b[0] ? -1 : 1));

  // Se escriben las columnas de datos sin tocar la de casillas (así su historial queda limpio)
  const m = filas.length, antes = SC.env, despues = H_SOC.length - SC.env - 1;
  sh.getRange(2, 1, n, antes).clearContent();
  sh.getRange(2, SC.env + 2, n, despues).clearContent();
  sh.getRange(2, 1, m, 1).setNumberFormat('@');
  sh.getRange(2, SC.per + 1, m, 1).setNumberFormat('@');
  sh.getRange(2, SC.tel + 1, m, 1).setNumberFormat('@');
  sh.getRange(2, 1, m, antes).setValues(filas.map(f => f.slice(0, antes)));
  sh.getRange(2, SC.env + 2, m, despues).setValues(filas.map(f => f.slice(SC.env + 1)));
  const colorEstado = { 'Activo': '#D9EAD3', 'Por confirmar': '#FCE5CD', 'Vencido': '#F4CCCC' };
  const colores = filas.map(f => [colorEstado[f[SC.est]] || null]);
  sh.getRange(2, SC.est + 1, m, 1).setBackgrounds(colores).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(2, SC.ind + 1, m, 1).setBackgrounds(colores);      // el indicativo también lleva el color del estado
  sh.getRange(2, SC.fecha + 1, m, 1).setNumberFormat('dd/mm/yyyy').setHorizontalAlignment('center');
  sh.getRange(2, SC.monto + 1, m, 1).setNumberFormat(NUM);
  sh.getRange(2, SC.per + 1, m, 1).setHorizontalAlignment('center').setFontWeight('bold')
    .setBackgrounds(filas.map(f => {
      const fin = +String(f[SC.per]).slice(-4) || 0;
      return [fin >= finActual ? '#D9EAD3' : (fin ? '#FCE5CD' : '#F4CCCC')];
    }));
  sh.getRange(2, SC.sol + 1, m, 1).setNumberFormat('dd/mm/yyyy hh:mm');
  // Casillas: solo se agregan donde falten, para no llenar el historial de la celda con cambios del script
  const cas = sh.getRange(2, SC.env + 1, m, 1);
  if (cas.getDataValidations().some(v => !v[0])) cas.insertCheckboxes();
  cas.setHorizontalAlignment('center');
  meses_Junta_(ss, sh, c);
  // Conteo de socios
  const r0 = CFG_FILAS.length + 2;
  sh.getRange(r0, COL_CFG, 24, 2).clear();
  sh.getRange(r0, COL_CFG, 5, 1).setValues([['SOCIOS'], ['Socios registrados'], ['Activos'], ['Por confirmar'], ['Vencidos']]);
  // Conteo calculado por el script (se actualiza cada vez que se edita un socio o corre la actualización)
  const cuenta = e => filas.filter(f => f[SC.est] === e).length;
  sh.getRange(r0 + 1, COL_CFG + 1, 4, 1).setNumberFormat('0')
    .setValues([[filas.length], [cuenta('Activo')], [cuenta('Por confirmar')], [cuenta('Vencido')]]);
  sh.getRange(r0, COL_CFG, 5, 2).setBorder(true, true, true, true, true, true, '#BBBBBB', null);
  encabezado_(sh.getRange(r0, COL_CFG, 1, 2));
  sh.getRange(r0 + 1, COL_CFG + 1, 4, 1).setFontWeight('bold').setHorizontalAlignment('center')
    .setBackgrounds([[null], ['#D9EAD3'], ['#FCE5CD'], ['#F4CCCC']]);
  sh.getRange(r0 + 6, COL_CFG, 11, 1).setValues([
    ['✎ = columna editable. Agregue socios en una fila nueva al final.'],
    ['Estado Activo: su "Período pagado" lo cubre (pago visto en el banco o escrito a mano).'],
    ['Estado Por confirmar: está en la lista validada en asamblea, pero aún no se ve su pago.'],
    ['Estado Vencido: sin pago que lo cubra y pasada la fecha "Lista de socios vigente hasta".'],
    ['"Período pagado" se escribe a mano solo para pagos que el banco no muestra'],
    ['(efectivo, pagado por otra persona, intercambio). Anote el motivo en Observaciones.'],
    ['Período en verde = vigente; naranja = anterior; rojo = sin pago.'],
    ['"Enviar informe": marque la casilla de un socio (también desde el celular) para enviarle'],
    ['el informe de los meses de "Informe de la Junta". La casilla se desmarca sola.'],
    ['Correos de notificaciones: reciben avisos de estados nuevos, errores y envíos.'],
    ['']])
    .setFontSize(9).setFontColor('#666666');
  sh.getRange(1, SC.per + 1).setNote('Se calcula del último pago detectado en el banco. Escríbalo a mano solo para ' +
    'pagos que el banco no muestra (efectivo, pagado por otra persona, intercambio). Un período más antiguo que ' +
    'el detectado se reemplaza automáticamente.');
  sh.getRange(1, 1, m + 1, H_SOC.length).setBorder(true, true, true, true, true, true, '#BBBBBB', null);
}

function buscarSocio_(ss, indicativo) {
  const sh = hojaSocios_(ss);
  const n = sh.getLastRow() - 1;
  if (n < 1) return null;
  const S = sh.getRange(2, 1, n, H_SOC.length).getValues();
  const x = indic_(indicativo);
  for (let i = 0; i < n; i++) {
    const r = S[i];
    if (x && indic_(r[SC.ind]) === x) {
      return { fila: i + 2, ind: String(r[SC.ind]).trim(), nombre: r[SC.nom], correo: String(r[SC.mail]).trim(),
               estado: r[SC.est], activo: r[SC.est] === 'Activo' || r[SC.est] === 'Por confirmar',
               ultimo: r[SC.fecha] instanceof Date ? r[SC.fecha] : null, monto: +r[SC.monto] || 0,
               periodo: String(r[SC.per] || ''), solicitudes: +r[SC.nsol] || 0 };
    }
  }
  return null;
}

// Datos de membresía que se muestran al socio (PDF y correos)
function membresia_(s, c) {
  const fin = +String(s.periodo).slice(-4) || 0;
  const vence = fin ? new Date(fin, c.vMes - 1, c.vDia) : (s.estado === 'Por confirmar' ? c.hasta : null);
  return Object.assign({}, s, { vence });
}

/* ───────────────────────── Envío por la Junta desde la hoja ───────────────────────── */

// Pone listas desplegables de meses en "Informe de la Junta" y, si están vacías o desactualizadas,
// el último mes disponible.
function meses_Junta_(ss, sh, c) {
  const meses = bloques_(ss).map(b => b.mes);
  if (!meses.length) return;
  const regla = SpreadsheetApp.newDataValidation().requireValueInList(meses.slice().reverse(), true).build();
  ['Informe de la Junta: desde', 'Informe de la Junta: hasta'].forEach(et => {
    const fila = CFG_FILAS.findIndex(f => f[0] === et) + 1;
    const celda = sh.getRange(fila, COL_CFG + 1).setDataValidation(regla);
    if (!meses.includes(String(celda.getDisplayValue()))) celda.setValue(meses.at(-1));
  });
}

// Disparador instalable "al editar" (corre con la cuenta del club, también desde el celular):
// al marcar "Enviar informe" en la fila de un socio, le envía el informe y desmarca la casilla.
async function alEditar(e) {
  const r = e && e.range;
  if (!r || r.getSheet().getName() !== 'Socios' || r.getColumn() !== SC.env + 1 || r.getRow() < 2) return;
  if (r.getNumRows() !== 1 || r.getNumColumns() !== 1 || r.getValue() !== true) return;
  const ss = e.source, sh = r.getSheet();
  const ind = String(sh.getRange(r.getRow(), SC.ind + 1).getValue()).trim();
  let quien = '';
  try { quien = (e.user && e.user.getEmail()) || ''; } catch (err) {}
  const c = leerConfig_(ss);
  r.setNote('Enviando…');
  try {
    const v = validarSolicitud_(ss, ind, c.juntaDesde, c.juntaHasta, { canal: 'Hoja (Junta)', por: quien, junta: true });
    if (!v.ok) { r.setValue(false).setNote(`${dmy_(new Date())}: no enviado. ${v.msg}`); return; }
    const s = await enviarInforme_(ss, { ind, desde: v.desde, hasta: v.hasta, canal: 'Hoja (Junta)', por: quien, fila: v.fila });
    const celda = sh.getRange(s.fila, SC.env + 1);
    celda.setValue(false).setNote(`Enviado ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')}` +
      ` · ${v.desde} – ${v.hasta}`);
  } catch (err) {
    r.setValue(false).setNote('Error: ' + err.message);
    avisarError_(ss, { ind, desde: c.juntaDesde, hasta: c.juntaHasta, canal: 'Hoja (Junta)' }, err);
  }
}

/* ───────────────────────── Formulario de solicitud de informe ───────────────────────── */

const H_BIT = ['Fecha y hora', 'Indicativo', 'Nombre', 'Desde', 'Hasta', 'Resultado', 'Enviado a', 'Canal', 'Solicitado por'];

function crearFormulario() {
  const ss = SpreadsheetApp.getActive();
  const props = PropertiesService.getScriptProperties();
  let form;
  try { form = FormApp.openById(props.getProperty('formId')); } catch (e) { form = null; }
  if (!form) {
    form = FormApp.create('Solicitud de informe financiero – TI0ARC');
    props.setProperty('formId', form.getId());
  }
  props.setProperty('ssId', ss.getId());
  const c = leerConfig_(ss);
  form.setDescription('Asociación de Radioaficionados de Cartago (TI0ARC).\n' +
    'Los socios activos pueden solicitar el informe de ingresos y egresos de la cuenta de la Asociación ' +
    'entre dos meses. El informe se envía en PDF únicamente al correo registrado del socio. ' +
    `Límite: ${c.porDia} solicitud(es) por día.`);
  form.setCollectEmail(false);
  form.setAllowResponseEdits(false);
  form.setLimitOneResponsePerUser(false);
  form.setConfirmationMessage('Solicitud recibida. Si usted es socio activo y tiene correo registrado, recibirá ' +
    'el informe en su correo en unos minutos. Si no le llega, o no tiene correo registrado, comuníquese con un ' +
    `miembro de la Junta Directiva o al WhatsApp de TI0ARC: ${c.whatsapp}.`);

  form.getItems().forEach(it => form.deleteItem(it));
  form.addTextItem().setTitle('Indicativo').setHelpText('Ejemplo: TI3ABC').setRequired(true);
  form.addListItem().setTitle('Desde').setRequired(true);
  form.addListItem().setTitle('Hasta').setRequired(true);
  actualizarMesesFormulario_(ss);

  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'alSolicitarInforme')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('alSolicitarInforme').forForm(form).onFormSubmit().create();

  const url = form.getPublishedUrl();
  setCfg_(ss, 'Enlace del formulario', url);
  hojaBitacora_(ss);
  colorPestanas_(ss);
  Logger.log('Formulario listo: ' + url);
  try { SpreadsheetApp.getUi().alert('Formulario listo. Enlace para compartir:\n\n' + url); } catch (e) {}
}

function actualizarMesesFormulario_(ss) {
  const id = PropertiesService.getScriptProperties().getProperty('formId');
  if (!id) return;
  let form;
  try { form = FormApp.openById(id); } catch (e) { return; }
  const B = bloques_(ss);
  if (!B.length) return;
  const meses = B.map(b => b.mes).reverse();             // los más recientes primero
  form.getItems(FormApp.ItemType.LIST).forEach(it => it.asListItem().setChoiceValues(meses));
}

function hojaBitacora_(ss) {
  const sh = hoja_(ss, 'Bitácora de solicitudes', H_BIT);
  if (sh.getRange(1, H_BIT.length).getValue() !== H_BIT.at(-1)) {
    encabezado_(sh.getRange(1, 1, 1, H_BIT.length).setValues([H_BIT]));
  }
  return sh;
}

// Agrega una fila a la bitácora y devuelve su número de fila
function registrar_(ss, fila) {
  const sh = hojaBitacora_(ss);
  const f = fila.slice(0, H_BIT.length);
  while (f.length < H_BIT.length) f.push('');
  const n = sh.getLastRow() + 1;
  sh.getRange(n, 1, 1, H_BIT.length).setValues([f]);
  sh.getRange(n, 1).setNumberFormat('dd/mm/yyyy hh:mm');
  return n;
}

function marcarFila_(ss, fila, resultado, correo) {
  if (!fila) return;
  hojaBitacora_(ss).getRange(fila, 6, 1, 2).setValues([[resultado, correo]]);
}

function avisarError_(ss, job, e) {
  if (avisos_()) MailApp.sendEmail(avisos_(), 'TI0ARC – Error al procesar una solicitud de informe',
    `Indicativo: ${job.ind}\nPeríodo: ${job.desde} – ${job.hasta}\nCanal: ${job.canal}\nError: ${e.message}\n\nHoja: ${ss.getUrl()}`);
}

// Se ejecuta al enviar el formulario de los socios (disparador instalado por crearFormulario)
async function alSolicitarInforme(e) {
  const ss = ssApp_();
  const resp = {};
  e.response.getItemResponses().forEach(r => resp[r.getItem().getTitle()] = String(r.getResponse()).trim());
  const ind = indic_(resp['Indicativo'] || ''), desde = resp['Desde'] || '', hasta = resp['Hasta'] || '';
  const v = validarSolicitud_(ss, ind, desde, hasta, { canal: 'Formulario' });
  if (!v.ok) return;
  try {
    await enviarInforme_(ss, { ind, desde: v.desde, hasta: v.hasta, canal: 'Formulario', fila: v.fila });
  } catch (err) {
    marcarFila_(ss, v.fila, 'Error: ' + err.message, '');
    avisarError_(ss, { ind, desde, hasta, canal: 'Formulario' }, err);
    throw err;
  }
}

// Prueba desde el editor de Apps Script: envía a TI3WTI el informe de los últimos 3 meses (sin PIN)
async function probarSolicitud() {
  const ss = SpreadsheetApp.getActive();
  const meses = bloques_(ss).map(b => b.mes);
  const v = validarSolicitud_(ss, 'TI3WTI', meses[Math.max(0, meses.length - 3)], meses.at(-1), { canal: 'Prueba', por: usuarioActual_() });
  if (!v.ok) return Logger.log('Rechazado: ' + v.msg);
  await enviarInforme_(ss, { ind: 'TI3WTI', desde: v.desde, hasta: v.hasta, canal: 'Prueba', por: usuarioActual_(), fila: v.fila });
  Logger.log('Listo: revise la Bitácora de solicitudes y su correo.');
}

// Verifica socio, estado, correo, límite diario y meses. Si es válida, la anota en la bitácora como
// "En proceso" y devuelve {ok, fila, desde, hasta, correo}; si no, la anota como rechazada y devuelve {ok:false, msg}.
function validarSolicitud_(ss, ind, desde, hasta, o) {
  const ahora = new Date(), c = leerConfig_(ss);
  actualizarSocios_(ss, false);
  const s = buscarSocio_(ss, ind);
  const fila = (res, correo) => registrar_(ss, [ahora, ind, s ? s.nombre : '', desde, hasta, res, correo || '', o.canal, o.por || '']);
  const no = (res, msg) => { fila(res); return { ok: false, msg }; };
  const contacto = `Comuníquese con la Junta Directiva (WhatsApp ${c.whatsapp}).`;

  if (!s) return no('Rechazado: no es socio', 'El indicativo no está registrado como socio. ' + contacto);
  if (!s.activo && !o.junta) return no('Rechazado: membresía vencida', 'Su membresía no está al día. ' + contacto);
  if (!s.correo) {
    if (avisos_()) MailApp.sendEmail(avisos_(), `TI0ARC – ${s.ind} solicitó un informe y no tiene correo registrado`,
      `${s.nombre} (${s.ind}) solicitó el informe ${desde} – ${hasta}, pero no tiene correo en la hoja Socios.\n` +
      'Contáctelo para registrar su correo.\n\nHoja: ' + ss.getUrl());
    return no('Rechazado: sin correo registrado', 'No tiene correo registrado. ' + contacto);
  }
  if (!o.junta) {                                        // límite diario (cuentan los enviados y en proceso)
    const bit = hojaBitacora_(ss), hoyTxt = dmy_(ahora);
    const hoy = bit.getLastRow() > 1 ? bit.getRange(2, 1, bit.getLastRow() - 1, 6).getValues()
      .filter(r => indic_(r[1]) === indic_(s.ind) && r[0] instanceof Date && dmy_(r[0]) === hoyTxt &&
                   /^(Enviado|En proceso)/.test(r[5])).length : 0;
    if (hoy >= c.porDia) return no('Rechazado: límite diario', `Ya alcanzó el máximo de ${c.porDia} informe(s) por día. Intente mañana.`);
  }
  const meses = bloques_(ss).map(b => b.mes);
  let i0 = meses.indexOf(desde), i1 = meses.indexOf(hasta);
  if (i0 < 0 || i1 < 0) return no('Rechazado: meses no válidos', 'Los meses elegidos no son válidos.');
  if (i0 > i1) [i0, i1] = [i1, i0];
  return { ok: true, fila: fila('En proceso', s.correo), desde: meses.at(i0), hasta: meses.at(i1), correo: s.correo };
}

// Genera el PDF con marca de agua, lo envía al correo registrado y actualiza bitácora, avisos y Socios
async function enviarInforme_(ss, job) {
  const c = leerConfig_(ss), ahora = new Date();
  const s = buscarSocio_(ss, job.ind);
  if (!s || !s.correo) throw new Error('El socio ya no está registrado o no tiene correo.');
  const B = bloques_(ss), meses = B.map(b => b.mes);
  const i0 = meses.indexOf(job.desde), i1 = meses.indexOf(job.hasta);
  if (i0 < 0 || i1 < 0) throw new Error('Meses no válidos.');
  const mem = membresia_(s, c);
  const marca = `${s.ind} · ${s.nombre} · ${dmy_(ahora)}`;
  const { pdf, ok: conMarca } = await marcaDeAgua_(informePdf_(ss, B.slice(i0, i1 + 1), mem, c), marca);
  MailApp.sendEmail({
    to: s.correo,
    name: 'TI0ARC – Asociación de Radioaficionados de Cartago',
    subject: `Informe financiero TI0ARC: ${job.desde} – ${job.hasta}`,
    body: `Estimado(a) ${s.nombre}:\n\n` +
      `Adjunto encontrará el informe financiero de la Asociación del período ${job.desde} a ${job.hasta}, ` +
      `${/Junta/.test(job.canal) ? 'enviado por la Junta Directiva' : 'solicitado'} el ${dmy_(ahora)} ` +
      `con el indicativo ${s.ind}. El informe se basa en los estados de cuenta del Banco Nacional.\n\n` +
      `Su membresía: ${mem.estado}` + (mem.vence ? ` · vence el ${dmy_(mem.vence)}` : '') +
      (mem.ultimo ? ` · último pago detectado el ${dmy_(mem.ultimo)}` : '') + '.\n\n' +
      `Si usted no hizo esta solicitud, comuníquese con la Junta Directiva (WhatsApp ${c.whatsapp}).\n\n` +
      'TI0ARC – Asociación de Radioaficionados de Cartago',
    attachments: [pdf],
  });
  const res = conMarca ? 'Enviado' : 'Enviado (sin marca de agua)';
  if (job.fila) marcarFila_(ss, job.fila, res, s.correo);
  else registrar_(ss, [ahora, s.ind, s.nombre, job.desde, job.hasta, res, s.correo, job.canal, job.por || '']);
  if (avisos_()) {
    MailApp.sendEmail(avisos_(), `TI0ARC – Informe enviado a ${s.ind}`,
      `Se envió un informe financiero a un socio.\n\n` +
      `Socio: ${s.nombre} (${s.ind})\nEstado: ${s.estado}\nPeríodo: ${job.desde} – ${job.hasta}\n` +
      `Enviado a: ${s.correo}\nCanal: ${job.canal}${job.por ? ' (' + job.por + ')' : ''}\n` +
      `Fecha: ${Utilities.formatDate(ahora, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')}\n` +
      `Solicitudes de este socio (incluida esta): ${s.solicitudes + 1}\n\nBitácora: ${ss.getUrl()}`);
  }
  hojaSocios_(ss).getRange(s.fila, SC.sol + 1, 1, 2).setValues([[ahora, s.solicitudes + 1]]);
  return s;
}

// Genera el PDF del informe con un documento temporal de Google Docs:
// encabezado y pie en todas las páginas con el nombre e indicativo de quien lo solicitó.
function informePdf_(ss, sel, socio, c, tituloDoc) {
  const sum = (b, k) => b.movs ? r2_(b.movs.reduce((x, m) => x + (+m[k] || 0), 0)) : 0;
  const ing = r2_(sel.reduce((a, b) => a + sum(b, 5), 0));
  const egr = r2_(sel.reduce((a, b) => a + sum(b, 6), 0));
  const neto = r2_(sel.reduce((a, b) => a + (b.movs ? 0 : b.neto), 0));
  const sIni = sel[0].sIni, sFin = sel[sel.length - 1].sBanco;
  const faltan = sel.filter(b => !b.movs).map(b => b.mes);
  const m = x => x === '' || x == null ? '' : (+x < 0 ? '−₡' : '₡') + fmt_(Math.abs(+x));
  const ahora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  const marca = `Generado para ${socio.nombre} (${socio.ind}) el ${ahora} · Documento personal, uso exclusivo del socio`;
  const periodo = `${sel[0].mes} – ${sel[sel.length - 1].mes}`;
  const DA = DocumentApp;

  let doc = DA.create('tmp_informe_' + socio.ind);
  const docId = doc.getId();
  // Google Docs limita los cambios sin guardar: se guarda y reabre el documento por lotes
  let ops = 0, body = doc.getBody();
  const guardarSiHaceFalta = () => {
    if (ops < 800) return;
    doc.saveAndClose();
    doc = DA.openById(docId);
    body = doc.getBody();
    ops = 0;
  };
  try {
    body.setMarginTop(50).setMarginBottom(50).setMarginLeft(40).setMarginRight(40);

    const pie = (sec, txt, color, tam) => {
      const p = sec.getParagraphs()[0] || sec.appendParagraph('');
      p.setText(txt);                                   // setText no devuelve el párrafo
      p.setAlignment(DA.HorizontalAlignment.CENTER);
      p.editAsText().setFontSize(tam).setForegroundColor(color).setBold(false).setFontFamily('Arial');
    };
    pie(doc.addHeader(), marca, '#B45F06', 8);
    pie(doc.addFooter(), `TI0ARC · Cuenta BN ${CFG.CUENTA} · ${marca}`, '#888888', 7);

    const parrafo = (txt, tam, color, negrita) => {
      guardarSiHaceFalta();
      ops += 3;
      const p = body.appendParagraph(txt);
      p.editAsText().setFontFamily('Arial').setFontSize(tam).setForegroundColor(color).setBold(!!negrita);
      p.setSpacingAfter(4);
      return p;
    };
    // Formato por tabla y por fila (no celda por celda) para hacer pocos cambios
    const tabla = (filas, numericas, resaltar, anchos) => {
      guardarSiHaceFalta();
      const t = body.appendTable(filas);
      t.setBorderColor('#BBBBBB');
      if (anchos) anchos.forEach((w, j) => t.setColumnWidth(j, w));
      t.editAsText().setFontFamily('Arial').setFontSize(8).setForegroundColor('#222222').setBold(false);
      const enc = t.getRow(0);
      enc.editAsText().setForegroundColor('#FFFFFF').setBold(true);
      for (let j = 0; j < filas[0].length; j++) enc.getCell(j).setBackgroundColor(AZUL);
      ops += 10 + filas[0].length * 2;
      for (let i = 1; i < filas.length; i++) {
        const fila = t.getRow(i), color = resaltar && resaltar(i);
        for (let j = 0; j < filas[i].length; j++) {
          const celda = fila.getCell(j);
          if (color) { celda.setBackgroundColor(color); ops++; }
          if (numericas.includes(j)) { celda.getChild(0).asParagraph().setAlignment(DA.HorizontalAlignment.RIGHT); ops++; }
          if (!filas[i][j]) {                   // marca invisible: indicativo del solicitante
            celda.setText(socio.ind);
            celda.editAsText().setFontSize(6).setForegroundColor(color || '#FFFFFF');
            ops += 3;
          }
        }
      }
      return t;
    };

    const titulo = body.getParagraphs()[0];
    titulo.setText('TI0ARC — ' + (tituloDoc || 'Informe financiero'));
    titulo.editAsText().setFontFamily('Arial').setFontSize(16).setBold(true).setForegroundColor(AZUL);
    parrafo(`Asociación de Radioaficionados de Cartago · Período ${periodo}`, 10, '#444444');
    parrafo(marca, 8, '#B45F06', true);

    // Aviso destacado: origen de la información
    guardarSiHaceFalta();
    const aviso = body.appendTable([['Este informe se basa exclusivamente en los estados de cuenta del Banco Nacional ' +
      `(cuenta ${CFG.CUENTA}). No incluye dinero en efectivo ni movimientos de otras cuentas.`]]);
    aviso.setBorderColor('#B45F06');
    aviso.getCell(0, 0).setBackgroundColor(AMARILLO).setPaddingTop(6).setPaddingBottom(6).setPaddingLeft(8).setPaddingRight(8);
    aviso.editAsText().setFontFamily('Arial').setFontSize(9).setBold(true).setForegroundColor('#7F3F00');
    ops += 8;

    // Membresía del socio
    parrafo('Su membresía', 11, AZUL, true).setSpacingBefore(10);
    const est = socio.estado || '';
    const colorEst = est === 'Activo' ? '#D9EAD3' : (est === 'Por confirmar' ? '#FCE5CD' : '#F4CCCC');
    const memb = [['Dato', 'Detalle'], ['Socio', `${socio.nombre} (${socio.ind})`], ['Estado', est],
      ['Último pago detectado', socio.ultimo ? `${dmy_(socio.ultimo)} (${m(socio.monto)})` : 'No detectado en los estados de cuenta'],
      ['Período pagado', socio.periodo || '—'],
      ['Su membresía vence el', socio.vence ? dmy_(socio.vence) : '—']];
    tabla(memb, [], i => (i === 2 || i === 5) ? colorEst : null, [180, 240]);
    if (socio.vence) parrafo(`Recuerde renovar su membresía antes del ${dmy_(socio.vence)}.`, 9, '#38761D', true);

    parrafo('Resumen del período', 11, AZUL, true).setSpacingBefore(10);
    const res = [['Concepto', 'Monto'], ['Saldo al inicio', m(sIni)], ['(+) Ingresos', m(ing)], ['(−) Egresos', m(egr)]];
    if (neto) res.push(['(±) Neto de meses sin estado de cuenta', m(neto)]);
    res.push(['Saldo al cierre', m(sFin)], ['Variación del período', m(r2_(sFin - sIni))]);
    tabla(res, [1], i => res[i][0] === 'Saldo al cierre' ? GRIS : null, [300, 120]);
    if (faltan.length) parrafo('Sin estado de cuenta en el período: ' + faltan.join(', '), 8, '#B45F06');

    parrafo('Detalle por mes', 11, AZUL, true).setSpacingBefore(10);
    const meses = [['Mes', 'Período', 'Saldo inicial', 'Ingresos', 'Egresos', 'Neto sin detalle', 'Saldo final']];
    sel.forEach(b => meses.push([b.mes, b.periodo, m(b.sIni), b.movs ? m(sum(b, 5)) : '', b.movs ? m(sum(b, 6)) : '',
                                 b.movs ? '' : m(b.neto), m(b.sBanco)]));
    meses.push(['TOTAL', '', '', m(ing), m(egr), neto ? m(neto) : '', '']);
    tabla(meses, [2, 3, 4, 5, 6], i => i === meses.length - 1 ? GRIS : (!sel[i - 1].movs ? AMARILLO : null),
          [70, 100, 74, 72, 72, 72, 72]);

    parrafo('Movimientos', 11, AZUL, true).setSpacingBefore(10);
    const notas = leerNotas_(ss);
    const encMov = ['Fecha', 'Número', 'Descripción', 'Ingreso', 'Egreso', 'Saldo', 'Nota'];
    const anchosMov = [52, 50, 168, 62, 62, 66, 72];
    sel.forEach(b => {
      parrafo(b.mes, 9, AZUL, true).setSpacingBefore(6);
      if (!b.movs) {
        tabla([encMov, ['', '', `Sin estado de cuenta (neto ${m(b.neto)})`, '', '', m(b.sBanco), '']],
              [3, 4, 5], () => AMARILLO, anchosMov);
        return;
      }
      const filas = [encMov];
      let saldo = b.sIni;
      b.movs.forEach(x => {
        saldo = r2_(saldo + (+x[5] || 0) - (+x[6] || 0));
        filas.push([dmy_(x[2] instanceof Date ? x[2] : x[1]), String(x[3]), String(x[4]),
                    x[5] ? m(x[5]) : '', x[6] ? m(x[6]) : '', m(saldo), notas[clave_(x[3], x[1])] || '']);
      });
      // tablas muy largas se parten para no superar el límite de cambios
      for (let k = 1; k < filas.length; k += 120) tabla([encMov].concat(filas.slice(k, k + 120)), [3, 4, 5], null, anchosMov);
    });

    parrafo('Documento generado automáticamente a partir de los estados de cuenta del Banco Nacional. ' +
      `Consultas: Junta Directiva o WhatsApp ${c.whatsapp}.`, 7, '#888888').setSpacingBefore(10);

    doc.saveAndClose();
    return DriveApp.getFileById(docId).getAs(MimeType.PDF)
      .setName(`Informe TI0ARC ${periodo} - ${socio.ind}.pdf`);
  } finally {
    DriveApp.getFileById(docId).setTrashed(true);
  }
}

function actualizarSocios() {
  actualizarSocios_(SpreadsheetApp.getActive());
  colorPestanas_(SpreadsheetApp.getActive());
}

// Colores y orden de las pestañas: Resumen, Movimientos, Informe, Socios, Bitácora de solicitudes
function colorPestanas_(ss) {
  const vieja = ss.getSheetByName('Resumen');
  if (vieja && !ss.getSheetByName(HOJA_RESUMEN)) vieja.setName(HOJA_RESUMEN);
  const orden = [[HOJA_RESUMEN, '#2E3A87'], ['Movimientos', '#38761D'], ['Informe', '#E69138'],
                 ['Socios', '#674EA7'], ['Bitácora de solicitudes', '#666666']];
  const activa = ss.getActiveSheet();
  let pos = 1;
  orden.forEach(([n, color]) => {
    const sh = ss.getSheetByName(n);
    if (!sh) return;
    sh.setTabColor(color);
    if (sh.getIndex() !== pos) { ss.setActiveSheet(sh); ss.moveActiveSheet(pos); }
    pos++;
  });
  if (activa) ss.setActiveSheet(activa);
}

/* ───────────────────────── Marca de agua del PDF ───────────────────────── */

// Carga la biblioteca pdf-lib (https://pdf-lib.js.org) desde un CDN, una vez por ejecución
function cargarPdfLib_() {
  if (globalThis.PDFLib) return globalThis.PDFLib;
  // pdf-lib usa setTimeout, que no existe en Apps Script
  if (!globalThis.setTimeout) globalThis.setTimeout = (f, t) => { if (t) Utilities.sleep(Math.min(t, 500)); f(); return 0; };
  const urls = ['https://cdnjs.cloudflare.com/ajax/libs/pdf-lib/1.17.1/pdf-lib.min.js',
                'https://cdn.jsdelivr.net/npm/pdf-lib@1.17.1/dist/pdf-lib.min.js'];
  for (const u of urls) {
    try {
      const src = UrlFetchApp.fetch(u, { muteHttpExceptions: true }).getContentText();
      (0, eval)(src);                                      // eval global: define globalThis.PDFLib
      if (globalThis.PDFLib) return globalThis.PDFLib;
    } catch (e) { Logger.log('No se pudo cargar pdf-lib desde ' + u + ': ' + e.message); }
  }
  throw new Error('No se pudo cargar la biblioteca de marca de agua');
}

// Estampa en todas las páginas un texto diagonal repetido en varias filas.
// Devuelve {pdf, ok}. Si algo falla, devuelve el PDF original (el informe se envía igual, con su encabezado y pie).
async function marcaDeAgua_(blob, texto) {
  try {
    const { PDFDocument, StandardFonts, rgb, degrees } = cargarPdfLib_();
    const doc = await PDFDocument.load(new Uint8Array(blob.getBytes()));
    const fuente = await doc.embedFont(StandardFonts.HelveticaBold);
    const tam = 14, ancho = fuente.widthOfTextAtSize(texto, tam);
    const pasoX = ancho + 50, pasoY = 62, angulo = 30;
    doc.getPages().forEach(pag => {
      const { width, height } = pag.getSize();
      for (let y = -height; y < height * 1.6; y += pasoY) {
        for (let x = -width, k = 0; x < width * 1.6; x += pasoX, k++) {
          pag.drawText(texto, {
            x: x + ((y / pasoY) % 2 ? pasoX / 2 : 0), y, size: tam, font: fuente,
            color: rgb(0.45, 0.45, 0.45), opacity: 0.25, rotate: degrees(angulo),
          });
        }
      }
    });
    const bytes = await doc.save();
    return { pdf: Utilities.newBlob(Array.from(bytes, b => (b > 127 ? b - 256 : b)), MimeType.PDF, blob.getName()), ok: true };
  } catch (e) {
    Logger.log('Marca de agua no aplicada: ' + e.message);
    return { pdf: blob, ok: false };
  }
}

/* ───────────────────────── Diagnóstico ───────────────────────── */

// Ejecútela desde Apps Script si algún estado sale con error o "REVISAR":
// muestra en el registro el texto que Google extrae de los PDF antiguos y lo que el script entiende.
function diagnosticarFormatoAntiguo() {
  const hilos = GmailApp.search('from:BUZON_BN_EMAIL@bncr.fi.cr ' + CFG.QUERY, 0, 5);
  for (const h of hilos) for (const msg of h.getMessages()) for (const att of msg.getAttachments()) {
    if (!/\.pdf$/i.test(att.getName()) || /glosario/i.test(att.getName())) continue;
    const raw = textoPdf_(att.copyBlob());
    if (raw.indexOf(CFG.CUENTA) < 0 && raw.indexOf(CFG.CUENTA.replace(/-/g, '')) < 0) continue;
    Logger.log('TEXTO EXTRAÍDO (primeros 3000 caracteres):\n' + raw.slice(0, 3000));
    try {
      const r = parsear_(raw);
      Logger.log(`RESULTADO: ${r.mes} · ${r.movs.length} movimientos · saldo anterior ${r.saldoAnt} · ` +
                 `saldo final ${r.saldoFin} · cuadre ${r.cuadre ? 'OK' : 'REVISAR'}`);
      r.movs.slice(0, 5).forEach(m => Logger.log(`  ${dmy_(m.fApl)} ${m.num} ${m.desc} ${m.monto}`));
    } catch (e) { Logger.log('ERROR: ' + e.message); }
    return;
  }
  Logger.log('No se encontró ningún PDF antiguo de la cuenta en colones.');
}

/* ───────────────────────── Utilidades ───────────────────────── */

// Ordena los movimientos de un estado por fecha de transacción (si no viene, fecha de aplicación).
// Los del mismo día conservan el orden del estado de cuenta.
function ordenar_(movs) {
  const f = m => ((m[2] instanceof Date) ? m[2] : m[1]).getTime();
  return movs.map((m, i) => [m, i]).sort((a, b) => f(a[0]) - f(b[0]) || a[1] - b[1]).map(x => x[0]);
}

function instalarDisparador() {
  const ss = SpreadsheetApp.getActive();
  PropertiesService.getScriptProperties().setProperty('ssId', ss.getId());
  ScriptApp.getProjectTriggers()
    .filter(t => ['actualizar', 'alEditar'].includes(t.getHandlerFunction()))
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('actualizar').timeBased().everyDays(1).atHour(7).create();
  ScriptApp.newTrigger('alEditar').forSpreadsheet(ss).onEdit().create();
  ss.toast('Actualización diaria (7:00 a. m.) y envíos desde la hoja activados.', 'TI0ARC');
}

function reprocesarTodo() {
  const ss = SpreadsheetApp.getActive();
  ['Estados', 'Datos', HOJA_RESUMEN, 'Movimientos'].forEach(n => {
    const sh = ss.getSheetByName(n);
    if (sh) ss.deleteSheet(sh);
  });
  PropertiesService.getScriptProperties().deleteProperty('vistos');
  actualizar();
}

function avisar_(nuevos, errores) {
  if (!avisos_() || (!nuevos.length && !errores.length)) return;
  const url = SpreadsheetApp.getActive().getUrl();
  let cuerpo = '';
  nuevos.forEach(r => {
    cuerpo += `${r.mes} (${r.ini ? dmy_(r.ini) : '?'} – ${dmy_(r.fin)})\n` +
              `  Ingresos: ₡${fmt_(r.ing)} (${r.nIng})\n  Egresos: ₡${fmt_(r.egr)} (${r.nEgr})\n` +
              `  Saldo final: ₡${fmt_(r.saldoFin)}\n  Cuadre: ${r.cuadre ? 'OK' : 'REVISAR'}\n\n`;
  });
  if (errores.length) cuerpo += 'Errores:\n' + errores.join('\n') + '\n\n';
  MailApp.sendEmail(avisos_(),
    `TI0ARC – ${nuevos.length ? 'Nuevo estado de cuenta BN' : 'Error al procesar estado BN'}`,
    cuerpo + 'Hoja: ' + url);
}

function hoja_(ss, nombre, encab) {
  let sh = ss.getSheetByName(nombre);
  if (!sh) {
    sh = ss.insertSheet(nombre);
    encabezado_(sh.getRange(1, 1, 1, encab.length).setValues([encab]));
    sh.setFrozenRows(1);
    if (['Estados', 'Datos', 'Notas'].includes(nombre)) sh.hideSheet();
  }
  return sh;
}

function limpia_(ss, nombre) {
  let sh = ss.getSheetByName(nombre);
  if (!sh) sh = ss.insertSheet(nombre, nombre === HOJA_RESUMEN ? 0 : 1);
  if (sh.getFilter()) sh.getFilter().remove();
  sh.clear();
  return sh;
}

function encabezado_(rng) {
  return rng.setBackground(AZUL).setFontColor('#FFFFFF').setFontWeight('bold')
            .setHorizontalAlignment('center').setWrap(true);
}

function carpeta_() {
  const it = DriveApp.getFoldersByName(CFG.CARPETA);
  return it.hasNext() ? it.next() : DriveApp.createFolder(CFG.CARPETA);
}

function fecha_(s) { const [d, m, y] = s.split('/').map(Number); return new Date(y, m - 1, d); }
function masDias_(d, n) { const x = new Date(d); x.setDate(x.getDate() + n); return x; }
function etiqueta_(d) { return MESES[d.getMonth()] + ' ' + d.getFullYear(); }
function rango_(a, b) {
  if (a === b) return a;
  const [ma, ya] = a.split(' '), [mb, yb] = b.split(' ');
  return ya === yb ? `${ma}–${mb} ${ya}` : `${a} – ${b}`;
}
function num_(s) { return parseFloat(s.replace(/,/g, '')); }
function r2_(x) { return Math.round(x * 100) / 100; }
function dmy_(d) { return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy'); }
function iso_(d) { return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function fmt_(x) { return x.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }
