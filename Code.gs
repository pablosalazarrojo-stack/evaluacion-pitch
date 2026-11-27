/**
 * Evaluación de Pitch (solo jueces, presentaciones individuales) — el "cerebro" que vive en Google (Apps Script).
 * Guarda las calificaciones en una hoja privada. La lista de alumnos y la contraseña
 * del profe viven aquí, NO en la página pública.
 *
 * Instalación (una sola vez):
 *  1. script.google.com → Proyecto nuevo → pegar este archivo.
 *  2. Llenar ALUMNOS y ADMIN_INICIAL (solo en el editor; en GitHub van vacíos).
 *  3. Ejecutar setup (crea la hoja «Evaluación de Pitch - respuestas»).
 *  4. Implementar → Aplicación web · Ejecutar como: yo · Acceso: cualquier persona.
 *  5. Pegar la URL /exec en CONFIG.API de index.html.
 */

/* Rubros y puntos máximos (suman 100). */
const RUBROS = [
  ["Gancho de entrada", 10],
  ["Problema", 20],
  ["Validación primaria del problema (entrevistas)", 15],
  ["Validación secundaria del problema (fuentes)", 15],
  ["Idea emprendedora", 10],
  ["Diseño del experimento (hipótesis crucial y métrica)", 15],
  ["Tiempo (dos a dos minutos y medio como máximo)", 5],
  ["Forma", 10]
];
const MAX = RUBROS.map(function (r) { return r[1]; });
const ADMIN_INICIAL = "CAMBIAR";
/* Nombres en el orden en que presentan — solo en el editor de Apps Script */
const ALUMNOS = [];

const ND = MAX.length;
const COLS = ["juez_id", "juez", "alumno"]
  .concat(RUBROS.map(function (r, i) { return "r" + (i + 1) + " " + r[0] + " (" + r[1] + ")"; }), ["total", "estado", "actualizado"]);
const I0 = 3;
const I_TOT = I0 + ND, I_EST = I_TOT + 1, I_ACT = I_TOT + 2;
const col = function (i) { let s = ""; i++; while (i > 0) { const r = (i - 1) % 26; s = String.fromCharCode(65 + r) + s; i = Math.floor((i - 1) / 26); } return s; };

function setup() {
  const p = props();
  const id = p.getProperty("SHEET_ID");
  const ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.create("Evaluación de Pitch - respuestas");
  p.setProperty("SHEET_ID", ss.getId());
  if (!p.getProperty("ADMIN_PASS")) p.setProperty("ADMIN_PASS", ADMIN_INICIAL);

  const a = ss.getSheetByName("Alumnos") || ss.insertSheet("Alumnos");
  if (ALUMNOS.length) {
    a.clear();
    a.getRange(1, 1, 1, 2).setValues([["orden", "nombre"]]).setFontWeight("bold");
    a.getRange(2, 1, ALUMNOS.length, 2).setValues(ALUMNOS.map(function (n, i) { return [i + 1, n]; }));
    a.setFrozenRows(1);
    a.autoResizeColumns(1, 2);
  }

  const e = ss.getSheetByName("Evaluaciones") || ss.insertSheet("Evaluaciones", 0);
  e.getRange(1, 1, 1, e.getMaxColumns()).clearContent();
  e.getRange(1, 1, 1, COLS.length).setValues([COLS]).setFontWeight("bold");
  e.setFrozenRows(1);

  resumen(ss);
  ["Sheet1", "Hoja 1", "Hoja1"].forEach(function (n) {
    const s = ss.getSheetByName(n);
    if (s && ss.getSheets().length > 1) ss.deleteSheet(s);
  });
  CacheService.getScriptCache().remove("alumnos");
  Logger.log(ss.getUrl());
}

/* Pestaña Resumen: promedio de jueces por alumno (solo lo enviado). */
function resumen(ss) {
  const r = ss.getSheetByName("Resumen") || ss.insertSheet("Resumen", 0);
  r.clear();
  const T = col(I_TOT), EST = col(I_EST);
  r.getRange(1, 1, 1, 3).setValues([["Alumno", "Calificación (prom. jueces)", "Jueces que lo calificaron"]]).setFontWeight("bold");
  r.getRange("A2").setFormula("=Alumnos!B2:B");
  const n = Math.max(ALUMNOS.length, 40);
  for (let i = 0; i < n; i++) {
    const row = i + 2;
    const cond = 'Evaluaciones!$C:$C,$A' + row + ',Evaluaciones!$' + EST + ':$' + EST + ',"enviado"';
    r.getRange(row, 2).setFormula('=IF($A' + row + '="","",IFERROR(ROUND(AVERAGEIFS(Evaluaciones!$' + T + ':$' + T + ',' + cond + '),1),""))');
    r.getRange(row, 3).setFormula('=IF($A' + row + '="","",COUNTIFS(' + cond + '))');
  }
  r.setFrozenRows(1);
  r.getRange("B:B").setFontWeight("bold");
  r.autoResizeColumns(1, 3);
}

function props() { return PropertiesService.getScriptProperties(); }
function libro() { return SpreadsheetApp.openById(props().getProperty("SHEET_ID")); }

function doGet() { return salida({ ok: true, msg: "Evaluación activa" }); }

function doPost(e) {
  let out;
  try {
    const d = JSON.parse(e.postData.contents);
    switch (d.action) {
      case "login": out = login(quien(d)); break;
      case "draft": out = guardar(quien(d), d, false); break;
      case "submit": out = guardar(quien(d), d, true); break;
      case "profe": admin(d); out = tablero(); break;
      case "reset": admin(d); reiniciar(); out = tablero(); break;
      default: throw new Error("Acción desconocida.");
    }
  } catch (err) {
    out = { ok: false, error: String(err.message || err) };
  }
  return salida(out);
}

function salida(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function admin(d) {
  if (String(d.pass || "") !== props().getProperty("ADMIN_PASS")) {
    Utilities.sleep(1500);
    throw new Error("Contraseña incorrecta.");
  }
}

function ronda() { return props().getProperty("RONDA") || "1"; }

function reiniciar() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sh = libro().getSheetByName("Evaluaciones");
    if (sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1);
    props().setProperty("RONDA", String(Date.now()));
  } finally {
    lock.releaseLock();
  }
}

/* El nombre del juez es su llave: sin acentos, sin mayúsculas, espacios simples. */
function llaveJuez(n) {
  return "JUEZ:" + String(n || "").normalize("NFD").replace(/[̀-ͯ]/g, "").toLowerCase().replace(/[^a-z0-9ñ]+/g, " ").trim();
}

function alumnos() {
  const c = CacheService.getScriptCache();
  const hit = c.get("alumnos");
  if (hit) return JSON.parse(hit);
  const sh = libro().getSheetByName("Alumnos");
  const v = sh.getLastRow() > 1 ? sh.getRange(2, 2, sh.getLastRow() - 1, 1).getValues() : [];
  const out = v.map(function (r) { return String(r[0]).trim(); }).filter(String);
  c.put("alumnos", JSON.stringify(out), 600);
  return out;
}

function quien(d) {
  const nombre = String(d.nombre || "").replace(/\s+/g, " ").trim().slice(0, 80);
  if (nombre.length < 2) throw new Error("Escribe tu nombre.");
  return { id: llaveJuez(nombre), nombre: nombre, saludo: nombre.split(" ")[0] };
}

function filas() {
  const sh = libro().getSheetByName("Evaluaciones");
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, COLS.length).getValues().filter(function (r) { return r[0]; });
}

function de(id, rows) {
  const ev = {};
  let enviado = false;
  rows.forEach(function (r) {
    if (String(r[0]) !== id) return;
    const t = r[I_ACT] instanceof Date ? r[I_ACT].getTime() : (Date.parse(r[I_ACT]) || 0);
    ev[r[2]] = { d: r.slice(I0, I0 + ND).map(Number), t: t };
    if (r[I_EST] === "enviado") enviado = true;
  });
  return { evals: ev, enviado: enviado };
}

function login(q) {
  const yo = de(q.id, filas());
  return { ok: true, id: q.id, nombre: q.saludo, nombreCompleto: q.nombre, alumnos: alumnos(), rubros: RUBROS,
    evals: yo.evals, enviado: yo.enviado, ronda: ronda() };
}

function puntos(x) {
  if (!x || x.length !== ND) throw new Error("Faltan calificaciones.");
  return x.map(function (v, i) {
    v = Number(v);
    if (!(v >= 0 && v <= MAX[i]) || Math.round(v) !== v) throw new Error("Calificación fuera de rango.");
    return v;
  });
}

/* draft: un alumno, "borrador".  submit: todos los que calificó, "enviado" (puede no ser la lista completa). */
function guardar(q, d, final) {
  if (d.ronda && String(d.ronda) !== ronda()) throw new Error("RONDA: El profe reinició la evaluación. Entra de nuevo.");
  const lista = alumnos();
  const evals = {};
  if (final) {
    const ks = Object.keys(d.evals || {});
    if (!ks.length) throw new Error("Faltan calificaciones.");
    ks.forEach(function (al) {
      if (lista.indexOf(al) < 0) throw new Error("No encontramos a " + al + ".");
      evals[al] = puntos(d.evals[al]);
    });
  } else {
    if (lista.indexOf(d.alumno) < 0) throw new Error("No encontramos a ese alumno.");
    evals[d.alumno] = puntos(d.d);
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sh = libro().getSheetByName("Evaluaciones");
    const rows = filas();
    if (de(q.id, rows).enviado) {
      if (final) return { ok: true, ya: true };
      throw new Error("Ya enviaste tus calificaciones.");
    }
    const idx = {};
    rows.forEach(function (r, i) { idx[String(r[0]) + "|" + r[2]] = i + 2; });
    const ahora = new Date();
    const nuevas = [];
    Object.keys(evals).forEach(function (al) {
      const v = evals[al];
      const tot = v.reduce(function (s, x) { return s + x; }, 0);
      const fila = [q.id, q.nombre, al].concat(v, [tot, final ? "enviado" : "borrador", ahora]);
      const at = idx[q.id + "|" + al];
      if (at) sh.getRange(at, 1, 1, fila.length).setValues([fila]);
      else nuevas.push(fila);
    });
    /* al enviar, los borradores de alumnos que ya no mandó también quedan como enviados de su lista final */
    if (final) {
      rows.forEach(function (r, i) {
        if (String(r[0]) === q.id && !evals[r[2]]) sh.getRange(i + 2, 1, 1, COLS.length).clearContent();
      });
    }
    if (nuevas.length) sh.getRange(sh.getLastRow() + 1, 1, nuevas.length, COLS.length).setValues(nuevas);
    return { ok: true, t: ahora.getTime() };
  } finally {
    lock.releaseLock();
  }
}

/* Promedios exactos por alumno; se redondea solo al final. */
function tablero() {
  const rows = filas();
  const lista = alumnos();
  const acc = {};
  lista.forEach(function (n) { acc[n] = { n: 0, b: 0, s: MAX.map(function () { return 0; }) }; });
  const jueces = {};
  rows.forEach(function (r) {
    const id = String(r[0]), env = r[I_EST] === "enviado";
    jueces[id] = jueces[id] || { nombre: String(r[1]), k: 0, env: false };
    jueces[id].k++;
    if (env) jueces[id].env = true;
    const t = acc[r[2]];
    if (!t) return;
    if (env) { t.n++; for (let i = 0; i < ND; i++) t.s[i] += Number(r[I0 + i]) || 0; }
    else t.b++;
  });
  const r1 = function (x) { return Math.round(x * 10) / 10; };
  return {
    ok: true,
    ronda: ronda(),
    rubros: RUBROS,
    alumnos: lista.map(function (n) {
      const t = acc[n];
      if (!t.n) return { nombre: n, n: 0, borradores: t.b, prom: null, total: null };
      const p = t.s.map(function (s) { return s / t.n; });
      return { nombre: n, n: t.n, borradores: t.b, prom: p.map(r1), total: r1(p.reduce(function (s, x) { return s + x; }, 0)) };
    }),
    jueces: Object.keys(jueces).map(function (id) {
      const j = jueces[id];
      return { nombre: j.nombre, estado: j.env ? "enviado" : "en curso", k: j.k };
    })
  };
}
