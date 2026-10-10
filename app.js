/* Seguimiento de práctica clínica · app local (sin servidor). Los datos viven solo en este dispositivo (localStorage). */
'use strict';
const LS_KEY = 'seguimientoPC_v1';
const PRACTICAS = [
  { id: 'FUN', nombre: 'Fundamentos de Enfermería I', corto: 'Fundamentos', re: /fundamento/ },
  { id: 'MQ',  nombre: 'Enfermería Médico Quirúrgica', corto: 'Médico Quirúrgica', re: /medico\s*-?\s*quir/ },
  { id: 'QX',  nombre: 'Enfermería Quirúrgica', corto: 'Quirúrgica', re: /quir/ },
  { id: 'PSQ', nombre: 'Enfermería Psiquiátrica', corto: 'Psiquiátrica', re: /psiq/ },
  { id: 'SP',  nombre: 'Salud Pública', corto: 'Salud Pública', re: /salud\s*publica/ },
];
const PMAP = Object.fromEntries(PRACTICAS.map(p => [p.id, p]));
const MOTIVOS = ['Comisión', 'Servicio', 'Cita programada', 'Exceptuado', 'Encamado', 'Juzgado', 'Otros'];
const MCOLOR = { 'Comisión': '#4b5320', 'Servicio': '#c9a227', 'Cita programada': '#2f6c8f', 'Exceptuado': '#8a6d3b', 'Encamado': '#b23a2e', 'Juzgado': '#6b4c8a', 'Otros': '#9aa08a' };
const SEDES = [
  [/central militar|\bhcm\b/, 'HCM'], [/mujer y neonat|hmemyn|hmeyn/, 'HMEMYN'], [/especialidades medicas|\buem\b/, 'UEM'],
  [/mexico\s*-?\s*espana/, 'CST III México-España'], [/vasconcelos/, 'CST III Vasconcelos'], [/manuel gonzalez/, 'CST III Manuel González Rivera'],
  [/cardiel/, 'CST III Lago Cardiel'], [/\binp\b|instituto nacional de psiq/, 'INP'], [/fray|bernardino/, 'Hosp. Fray Bernardino Álvarez'],
  [/navarro/, 'Hosp. Juan N. Navarro'], [/naval/, 'Centro Médico Naval'],
];
const DEF_CFG = { umbral: 8, adv: 20, alerta: 30, horasFalta: 5, pInst: 70, pInv: 30,
  horas: { FUN: 306, MQ: 224, QX: 185, PSQ: 115, SP: 75 } };

/* ---------- utilidades ---------- */
const $ = (s, el = document) => el.querySelector(s);
const $$ = (s, el = document) => [...el.querySelectorAll(s)];
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const norm = s => String(s ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[_]+/g, ' ').replace(/\s+/g, ' ').trim();
const STOP = new Set(['de', 'del', 'la', 'las', 'los', 'y', 'jefe', 'subjefe']);
const tokens = n => norm(String(n).replace(/\(.*?\)/g, ' ')).replace(/[^a-zñ ]/g, ' ').split(' ').filter(t => t.length > 1 && !STOP.has(t));
const keyOf = n => tokens(n).sort().join(' ');
const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
const fmt = (n, d = 1) => (n == null || isNaN(n)) ? '—' : Number(n).toLocaleString('es-MX', { minimumFractionDigits: d, maximumFractionDigits: d });
const fmt2 = n => fmt(n, 2);
const hoyISO = () => { const d = new Date(); return new Date(d - d.getTimezoneOffset() * 6e4).toISOString().slice(0, 10); };
const fFecha = iso => { if (!iso) return '—'; const [y, m, d] = iso.split('-'); const M = ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic']; return `${+d} ${M[+m - 1]} ${y}`; };
const cleanName = n => String(n ?? '').replace(/\(.*?\)/g, ' ').replace(/[.]+$/, '').replace(/\s+/g, ' ').trim();
const iniciales = n => tokens(n).length ? cleanName(n).split(' ').filter(w => w.length > 1).slice(0, 2).map(w => w[0].toUpperCase()).join('') : '?';
function detectPractica(t) { const s = norm(t); if (!s) return null; for (const p of PRACTICAS) if (p.re.test(s)) return p.id; return null; }
function sedeCorta(t) { const s = norm(t); for (const [re, n] of SEDES) if (re.test(s)) return n; return String(t ?? '').replace(/\s+/g, ' ').trim(); }
function toNum(v) { if (v == null || v === '') return null; if (typeof v === 'number') return isFinite(v) ? v : null; const s = String(v).replace(',', '.').trim(); if (!/^-?\d+(\.\d+)?$/.test(s)) return null; return parseFloat(s); }
function parseFecha(v) {
  if (v == null || v === '') return null;
  if (v instanceof Date && !isNaN(v)) return new Date(v - v.getTimezoneOffset() * 6e4).toISOString().slice(0, 10);
  if (typeof v === 'number' && window.XLSX) { const d = XLSX.SSF.parse_date_code(v); if (d) return `${d.y}-${String(d.m).padStart(2, '0')}-${String(d.d).padStart(2, '0')}`; }
  const s = String(v).trim(); let m;
  if ((m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/))) return `${m[1]}-${m[2].padStart(2, '0')}-${m[3].padStart(2, '0')}`;
  if ((m = s.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})/))) { const y = m[3].length === 2 ? '20' + m[3] : m[3]; return `${y}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`; }
  return null;
}
function normMotivo(m, orig) {
  const a = norm(m), b = norm(orig);
  if (/retardo/.test(a) || (!a && /retardo/.test(b))) return 'RETARDO';
  for (const k of MOTIVOS) if (a === norm(k)) return k;
  for (const t of [a, b]) {
    if (!t) continue;
    if (/juzgad|audiencia|declarar/.test(t)) return 'Juzgado';
    if (/encamad/.test(t)) return 'Encamado';
    if (/exceptuad/.test(t)) return 'Exceptuado';
    if (/cita|consulta|peloton de sanidad|laboratorio|revaloracion|rehabilitacion/.test(t)) return 'Cita programada';
    if (/comision/.test(t)) return 'Comisión';
    if (/servicio|cuartelera|guardia|\bsv\b|cabo de turno|imaginaria/.test(t)) return 'Servicio';
  }
  return 'Otros';
}
function simil(a, b) { const A = new Set(a), B = new Set(b); let s = 0; A.forEach(t => { if (B.has(t)) s++; }); return { s, max: Math.max(A.size, B.size) }; }
function toast(msg, err) { const t = document.createElement('div'); t.className = 'toast' + (err ? ' e' : ''); t.textContent = msg; $('#toasts').appendChild(t); setTimeout(() => t.remove(), 3800); }

/* ---------- estado ---------- */
let S = load();
function vacio() { return { v: 1, cfg: JSON.parse(JSON.stringify(DEF_CFG)), cadetes: {}, insc: {}, faltas: {}, acciones: {}, padron: {}, imports: [] }; }
function load() {
  try { const s = JSON.parse(localStorage.getItem(LS_KEY)); if (s && s.v) { s.cfg = Object.assign(JSON.parse(JSON.stringify(DEF_CFG)), s.cfg || {}); s.cfg.horas = Object.assign({}, DEF_CFG.horas, s.cfg.horas || {}); return s; } } catch (e) { }
  return vacio();
}
function save() {
  try { localStorage.setItem(LS_KEY, JSON.stringify(S)); } catch (e) { toast('No se pudo guardar en este navegador (espacio insuficiente). Exporta un respaldo.', true); }
  D = null;
}

/* resolución de cadetes por nombre (tolerante a orden y acentos) */
// modo 'falta': solo se empata (por parecido) con cadetes de la misma práctica.
// modo 'lista': solo con cadetes de la práctica que aún no vienen de una lista (evita juntar a dos personas distintas de la misma lista).
function buscarCadete(nombre, practica, modo) {
  const k = keyOf(nombre); if (!k) return null;
  if (S.cadetes[k]) return k;
  if (S.alias && S.alias[k]) return S.alias[k];
  if (!practica) return null;
  const tk = k.split(' '); let best = null, bs = 0, tie = false;
  for (const ck in S.cadetes) {
    const ins = S.insc[ck + '|' + practica];
    if (!ins || (modo === 'lista' && ins.origen === 'lista')) continue;
    const { s, max } = simil(tk, ck.split(' '));
    if (s >= 3 && s >= max - 1) {
      let sc = s * 10 - max; if (practica && S.insc[ck + '|' + practica]) sc += 5;
      if (sc > bs) { best = ck; bs = sc; tie = false; } else if (sc === bs) tie = true;
    }
  }
  return tie ? null : best;
}
function asegurarCadete(nombre, practica, modo) {
  let k = buscarCadete(nombre, practica, modo);
  if (!k) { k = keyOf(nombre); S.cadetes[k] = { key: k, nombre: cleanName(nombre), grupo: '', matricula: '' }; }
  else if (k !== keyOf(nombre)) { S.alias = S.alias || {}; S.alias[keyOf(nombre)] = k; }
  return k;
}
function asignarGrupos() {
  const pk = Object.keys(S.padron); if (!pk.length) return;
  for (const c of Object.values(S.cadetes)) {
    if (c.grupoManual) continue;
    let p = S.padron[c.key];
    if (!p) {
      const tk = c.key.split(' '); let best = null, bs = 0, tie = false;
      for (const k of pk) { const { s, max } = simil(tk, k.split(' ')); if (s >= 3 && s >= max - 1) { const sc = s * 10 - max; if (sc > bs) { best = k; bs = sc; tie = false; } else if (sc === bs) tie = true; } }
      if (best && !tie) p = S.padron[best];
    }
    if (p) { c.grupo = p.grupo; if (p.matricula) c.matricula = p.matricula; }
  }
}

/* ---------- cálculos derivados ---------- */
let D = null;
function finalDe(i) {
  if (i.finalManual != null) return i.finalManual;
  const c = S.cfg;
  if (i.inst != null && i.inv != null) return Math.round((i.inst * c.pInst / 100 + i.inv * c.pInv / 100) * 100) / 100;
  if (i.inst == null && i.inv == null && i.finalLista != null) return i.finalLista;
  return null;
}
function estadoDe(p) { return p >= S.cfg.alerta ? 'alerta' : p >= S.cfg.adv ? 'advertencia' : 'regular'; }
function derive() {
  if (D) return D;
  const c = S.cfg, byIns = {}, byCad = {};
  for (const f of Object.values(S.faltas)) { (byIns[f.cadete + '|' + f.practica] ||= []).push(f); }
  const rows = [];
  for (const i of Object.values(S.insc)) {
    const cad = S.cadetes[i.cadete]; if (!cad) continue;
    const fl = (byIns[i.id] || []).sort((a, b) => a.fecha.localeCompare(b.fecha));
    const horas = c.horas[i.practica] || 0;
    const pct = horas ? fl.length * c.horasFalta * 100 / horas : 0;
    const est = estadoDe(pct + 1e-9);
    const limite = horas ? Math.ceil((c.alerta / 100 * horas / c.horasFalta) - 1e-9) : 0;
    const fin = finalDe(i);
    const mot = {}; fl.forEach(f => mot[f.motivo] = (mot[f.motivo] || 0) + 1);
    const acc = Object.values(S.acciones).filter(a => a.cadete === i.cadete).sort((a, b) => (b.fecha || '').localeCompare(a.fecha || ''));
    const r = { ...i, cad, nombre: cad.nombre, grupo: cad.grupo || '', faltas: fl, nf: fl.length, pct: Math.round(pct * 100) / 100, estado: est,
      paraAlerta: Math.max(0, limite - fl.length), final: fin, riesgo: fin != null && fin < c.umbral, mot, ultAccion: acc[0] || null, nAcc: acc.length };
    rows.push(r); (byCad[i.cadete] ||= []).push(r);
  }
  rows.sort((a, b) => b.pct - a.pct || a.nombre.localeCompare(b.nombre));
  D = { rows, byCad };
  return D;
}
const ORD = { alerta: 3, advertencia: 2, regular: 1 };

/* ---------- importación ---------- */
const G_RE = { inst: /instrumento/, inv: /trabajo|investigaci/, fin: /evaluaci|final|calificaci|promedio/ };
function sheetAOA(ws) {
  const ref = ws['!ref']; if (!ref) return { aoa: [], hidden: new Set(), merges: [], off: { r: 0, c: 0 } };
  const rg = XLSX.utils.decode_range(ref);
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true });
  const hidden = new Set(); (ws['!rows'] || []).forEach((r, i) => { if (r && r.hidden) hidden.add(i - rg.s.r); });
  const merges = (ws['!merges'] || []).map(m => ({ r0: m.s.r - rg.s.r, c0: m.s.c - rg.s.c, r1: m.e.r - rg.s.r, c1: m.e.c - rg.s.c }));
  return { aoa, hidden, merges, off: rg.s, ws };
}
const cellTxt = v => v == null ? '' : String(v);
function findHeader(aoa, test, maxRows = 30) {
  for (let r = 0; r < Math.min(aoa.length, maxRows); r++) { const row = (aoa[r] || []).map(v => norm(cellTxt(v))); if (test(row)) return r; }
  return -1;
}
function analizarHoja(name, sh, fileName) {
  const { aoa } = sh;
  let h = findHeader(aoa, row => row.some(x => /^fecha/.test(x)) && row.some(x => /cadete|discente|nombre|alumn/.test(x)) && row.some(x => /motivo/.test(x)));
  if (h >= 0) return { tipo: 'faltas', h, map: mapFaltas(aoa[h]) };
  const titulo = aoa.slice(0, 12).flat().map(cellTxt).find(t => /lista de asistencia de/i.test(t));
  h = findHeader(aoa, row => row.some(x => /nombre del discente|^nombre/.test(x)) && !row.some(x => G_RE.inst.test(x) || G_RE.inv.test(x) || /evaluaci|calificaci/.test(x)));
  if (titulo && h >= 0) {
    const g = (titulo.match(/lista de asistencia de\s+(.+?)\s+correspondiente/i) || [])[1] || name;
    const row = aoa[h].map(x => norm(cellTxt(x)));
    return { tipo: 'padron', h, grupo: g.replace(/\s+/g, ' ').trim(), map: { nombre: row.findIndex(x => /nombre/.test(x)), matricula: row.findIndex(x => /matricula/.test(x)) } };
  }
  h = findHeader(aoa, row => row.some(x => /^nombre|discente|cadete|alumn/.test(x)) && row.some(x => G_RE.inst.test(x) || G_RE.inv.test(x) || G_RE.fin.test(x)));
  if (h >= 0) return { tipo: 'calif', h, ...mapCalif(sh, h, name, fileName) };
  return { tipo: 'desconocida', h: -1 };
}
function mapFaltas(hdr) {
  const row = hdr.map(x => norm(cellTxt(x))); const f = re => row.findIndex(x => re.test(x));
  const motivo = row.findIndex(x => /motivo/.test(x) && !/escrito|bitacora|original/.test(x));
  return { fecha: f(/^fecha/), nombre: f(/cadete|discente|nombre|alumn/), practica: f(/practica|materia|asignatura/), grupoTurno: f(/grupo|turno/),
    sede: f(/hospital|sede|unidad|centro/), obs: f(/servicio|rotaci|observ/), motivo, motivoOrig: f(/motivo.*(escrito|bitacora|original)/), instructora: f(/instructor|docente|profesor/) };
}
function mapCalif(sh, h, name, fileName) {
  const { aoa, merges } = sh; const row = (aoa[h] || []).map(x => norm(cellTxt(x)));
  const f = re => row.findIndex(x => re.test(x));
  const nombre = f(/^nombre|discente|cadete|alumn/);
  const inst = f(G_RE.inst), inv = f(/trabajo|investigaci/);
  const fin = row.findIndex((x, i) => G_RE.fin.test(x) && i !== inst && i !== inv && !/firma/.test(x));
  const mw = c => { const m = merges.find(m => m.r0 <= h && m.r1 >= h && m.c0 === c); return m ? m.c1 - m.c0 + 1 : 1; };
  // Trabajo de investigación en una sola columna y sumado directo en la final: viene ya ponderado (30 %)
  let invPond = false;
  if (inv >= 0 && mw(inv) === 1 && fin >= 0 && sh.ws) {
    for (let r = h + 1; r < Math.min(aoa.length, h + 8); r++) {
      const cell = sh.ws[XLSX.utils.encode_cell({ r: r + sh.off.r, c: fin + sh.off.c })];
      if (cell && cell.f) { const col = XLSX.utils.encode_col(inv + sh.off.c); invPond = new RegExp('(^|[^A-Z])' + col + '\\d').test(cell.f.toUpperCase()); break; }
    }
  }
  const meta = { materia: '', docente: '', periodo: '', horario: '', sede: '' };
  const textos = aoa.slice(0, h).flat().filter(v => typeof v === 'string').map(v => v.trim()).filter(Boolean);
  for (const t of textos) {
    const n = norm(t);
    if (!meta.materia && /materia|practica clinica/.test(n)) { meta.materia = t; continue; }
    if (!meta.docente && /^(docente|instructor|profesor)/.test(n)) { meta.docente = t.replace(/^(docente|instructora?|profesora?)\s*:?\s*/i, '').trim(); continue; }
    if (!meta.periodo && /(enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|octubre|noviembre|diciembre)/.test(n) && /\d/.test(n) && !/ciclo/.test(n)) { meta.periodo = t.replace(/\s+/g, ' '); continue; }
    if (!meta.horario && !/ciclo/.test(n) && /\b\d{1,2}:?\d{2}\s*-\s*\d{1,2}:?\d{2}\b/.test(n) && /(lunes|martes|miercoles|jueves|viernes|sabado|\d:\d)/.test(n)) { meta.horario = t.replace(/\s+/g, ' '); continue; }
  }
  // la sede se toma primero del nombre de la hoja (lo actualiza abril); si no la indica, del encabezado
  meta.sedeEncabezado = '';
  for (const t of textos) { const n = norm(t); if (/ejercito|dir\.|rect|esc\. mil|secc|subsec|^docente/.test(n)) continue; if (SEDES.some(([re]) => re.test(n))) { meta.sedeEncabezado = sedeCorta(t); break; } }
  meta.sede = SEDES.some(([re]) => re.test(norm(name))) ? sedeCorta(name) : (meta.sedeEncabezado || sedeCorta(name));
  const practica = detectPractica(meta.materia) || detectPractica(name) || detectPractica(fileName);
  const all = norm(textos.join(' '));
  const turno = /segundo turno|14:?00/.test(all) ? 'Vespertino' : (/07:?30|primer turno/.test(all) ? 'Matutino' : '');
  return { map: { nombre, inst, inv, fin, practica: f(/^(practica|materia|asignatura)/), sede: f(/^(sede|hospital|unidad)/), docente: f(/^(docente|instructor|profesor)/), grupo: f(/^grupo/) }, invPond, meta: { ...meta, practica, turno } };
}
const NO_NOMBRE = /promedio|firma|desviacion|mediana|moda|reprob|^nombre|reviso|certifico|total|profesor/;
function filasDatos(sh, h, colNombre) {
  const out = []; let vacias = 0;
  for (let r = h + 1; r < sh.aoa.length; r++) {
    const row = sh.aoa[r] || []; if (sh.hidden.has(r)) continue;
    const n = cellTxt(row[colNombre]).trim();
    if (!n) { if (++vacias > 25) break; continue; }
    if (NO_NOMBRE.test(norm(n))) { if (/promedio|firma|reviso/.test(norm(n))) break; continue; }
    if (tokens(n).length < 2) continue;
    vacias = 0; out.push({ r, row, nombre: n });
  }
  return out;
}
async function leerArchivo(file) {
  if (/\.(csv|txt)$/i.test(file.name)) return XLSX.read(await file.text(), { type: 'string', raw: false });
  return XLSX.read(new Uint8Array(await file.arrayBuffer()), { type: 'array', cellStyles: true, cellFormula: true, cellDates: false });
}
async function importarArchivos(files, revisar) {
  const log = $('#implog'); if (log) log.innerHTML = '';
  const add = (cls, html) => { if (!log) return; const d = document.createElement('div'); d.className = 'li ' + cls; d.innerHTML = '<div>' + html + '</div>'; log.appendChild(d); };
  const total = { faltas: 0, calif: 0, padron: 0 };
  const trabajos = [];
  for (const file of files) {
    let wb; try { wb = await leerArchivo(file); } catch (e) { add('e', `<b>${esc(file.name)}</b>: no se pudo leer (${esc(e.message)}).`); continue; }
    for (const name of wb.SheetNames) {
      const ws = wb.Sheets[name];
      if (wb.Workbook && wb.Workbook.Sheets) { const sm = wb.Workbook.Sheets.find(s => s.name === name); if (sm && sm.Hidden) continue; }
      const sh = sheetAOA(ws); if (!sh.aoa.some(r => r && r.some(v => v != null && v !== ''))) continue;
      let an = analizarHoja(name, sh, file.name);
      const faltan = needMap(an);
      if ((an.tipo === 'desconocida' || faltan.length || revisar) && !(an.tipo === 'padron' && !revisar)) {
        const res = await dialogoMapeo(file.name, name, sh, an, faltan);
        if (!res) { add('w', `<b>${esc(file.name)} · ${esc(name.trim())}</b>: hoja omitida.`); continue; }
        an = res;
      }
      trabajos.push({ file, name, sh, an });
    }
    S.imports.unshift({ archivo: file.name, fecha: new Date().toISOString() }); S.imports = S.imports.slice(0, 40);
  }
  // primero listas de grupo y de práctica, al final el registro de faltas (para cruzarlo con las listas ya cargadas)
  const ordenT = { padron: 0, calif: 1, faltas: 2 };
  trabajos.sort((a, b) => ordenT[a.an.tipo] - ordenT[b.an.tipo]);
  for (const { file, name, sh, an } of trabajos) {
    const tag = `<b>${esc(file.name)} · ${esc(name.trim())}</b>`;
    const avisos = w => w.length ? '<br><span class="small">' + w.slice(0, 8).map(esc).join('<br>') + (w.length > 8 ? `<br>… y ${w.length - 8} avisos más` : '') + '</span>' : '';
    try {
      if (an.tipo === 'faltas') { const r = importarFaltas(sh, an, file.name); total.faltas += r.n; add(r.warn.length ? 'w' : 's', `${tag}: registro de faltas, ${r.n} faltas importadas${r.ret ? `, ${r.ret} retardos omitidos (no cuentan)` : ''}${r.dup ? `, ${r.dup} ya estaban y se actualizaron` : ''}.${avisos(r.warn)}`); }
      else if (an.tipo === 'padron') { const n = importarPadron(sh, an); total.padron += n; add('s', `${tag}: lista del grupo ${esc(an.grupo)}, ${n} discentes.`); }
      else if (an.tipo === 'calif') { const r = importarCalif(sh, an, file.name, name); total.calif += r.n; const difSede = an.meta.sedeEncabezado && an.meta.sedeEncabezado !== an.meta.sede ? [`El encabezado dice «${an.meta.sedeEncabezado}» y la hoja «${name.trim()}»; se tomó ${an.meta.sede}.`] : []; const w = [...difSede, ...r.warn]; add(w.length ? 'w' : 's', `${tag}: ${esc(PMAP[an.meta.practica]?.corto || '¿práctica?')} · ${esc(an.meta.sede)} · ${r.n} cadetes, ${r.conCalif} con calificación.${an.meta.docente ? ' Docente: ' + esc(an.meta.docente) + '.' : ' Sin docente en la hoja.'}${r.sinMapa.length ? '<br><span class="small">Sin columna para: ' + r.sinMapa.join(', ') + '.</span>' : ''}${avisos(w)}`); }
    } catch (e) { console.error(e); add('e', `${tag}: error al importar (${esc(e.message)}).`); }
  }
  asignarGrupos(); save();
  add('s', `<b>Listo.</b> ${total.calif} registros de listas, ${total.faltas} faltas y ${total.padron} discentes de listas de grupo. <a href="#/tablero">Ver tablero →</a>`);
  toast('Importación lista');
  return total;
}
function needMap(an) {
  if (an.tipo === 'faltas') { const m = an.map, f = []; if (m.fecha < 0) f.push('fecha'); if (m.nombre < 0) f.push('nombre'); if (m.practica < 0 && !an.fijo?.practica) f.push('práctica'); if (m.motivo < 0 && m.motivoOrig < 0) f.push('motivo'); return f; }
  if (an.tipo === 'calif') { const m = an.map, f = []; if (m.nombre < 0) f.push('nombre'); if (!an.meta.practica) f.push('práctica'); if (m.inst < 0 && m.inv < 0 && m.fin < 0) f.push('calificaciones'); return f; }
  return [];
}
function importarFaltas(sh, an, archivo) {
  const m = an.map, out = { n: 0, ret: 0, dup: 0, warn: [] };
  const v = (row, c) => c >= 0 ? row[c] : null;
  for (let r = an.h + 1; r < sh.aoa.length; r++) {
    const row = sh.aoa[r] || []; if (sh.hidden.has(r)) continue;
    const nombre = cellTxt(v(row, m.nombre)).trim(); if (!nombre) continue;
    const fecha = parseFecha(v(row, m.fecha)); if (!fecha) { out.warn.push(`Fila ${r + 1}: fecha no reconocida (${cellTxt(v(row, m.fecha))}).`); continue; }
    const practica = detectPractica(cellTxt(v(row, m.practica))) || an.fijo?.practica;
    if (!practica) { out.warn.push(`Fila ${r + 1}: práctica no reconocida (${cellTxt(v(row, m.practica))}).`); continue; }
    const motivoOrig = cellTxt(v(row, m.motivoOrig)).trim(), motivoTx = cellTxt(v(row, m.motivo)).trim();
    const motivo = normMotivo(motivoTx, motivoOrig);
    if (motivo === 'RETARDO') { out.ret++; continue; }
    const k = asegurarCadete(nombre, practica, 'falta');
    const sedeTx = cellTxt(v(row, m.sede)).trim() || an.fijo?.sede || '';
    const iid = k + '|' + practica;
    if (!S.insc[iid]) { S.insc[iid] = { id: iid, cadete: k, practica, sede: sedeCorta(sedeTx), hoja: cellTxt(v(row, m.grupoTurno)).trim(), docente: cellTxt(v(row, m.instructora)).trim(), periodo: '', horario: '', inst: null, inv: null, finalLista: null, finalManual: null, origen: 'registro de faltas' }; out.warn.push(`${cleanName(nombre)} no aparece en las listas de ${PMAP[practica].corto}; se agregó desde el registro de faltas.`); }
    const id = fecha + '|' + iid;
    if (S.faltas[id]) out.dup++;
    S.faltas[id] = { id, cadete: k, practica, fecha, motivo, motivoOrig: motivoOrig || motivoTx, sede: sedeCorta(sedeTx), grupoTurno: cellTxt(v(row, m.grupoTurno)).trim(), instructora: cellTxt(v(row, m.instructora)).trim(), obs: cellTxt(v(row, m.obs)).trim(), archivo };
    out.n++;
  }
  return out;
}
function importarPadron(sh, an) {
  let n = 0;
  for (const { row, nombre } of filasDatos(sh, an.h, an.map.nombre)) {
    const k = keyOf(nombre); if (!k) continue;
    S.padron[k] = { nombre: cleanName(nombre), grupo: an.grupo, matricula: an.map.matricula >= 0 ? cellTxt(row[an.map.matricula]).trim() : '' }; n++;
  }
  return n;
}
function importarCalif(sh, an, archivo, hoja) {
  const m = an.map, out = { n: 0, conCalif: 0, warn: [], sinMapa: [] };
  if (m.inst < 0) out.sinMapa.push('instrumento'); if (m.inv < 0) out.sinMapa.push('trabajo de investigación'); if (m.fin < 0) out.sinMapa.push('evaluación parcial final');
  const vistos = new Set();
  for (const { row, nombre } of filasDatos(sh, an.h, m.nombre)) {
    const practica = (m.practica >= 0 && detectPractica(cellTxt(row[m.practica]))) || an.meta.practica; if (!practica) continue;
    const k = asegurarCadete(nombre, practica, 'lista'); const id = k + '|' + practica;
    const sede = (m.sede >= 0 && cellTxt(row[m.sede]).trim()) ? sedeCorta(row[m.sede]) : an.meta.sede;
    const docente = (m.docente >= 0 && cellTxt(row[m.docente]).trim()) || an.meta.docente;
    if (vistos.has(id)) { out.warn.push(`${cleanName(nombre)} aparece dos veces en esta hoja; se tomó la primera.`); continue; }
    vistos.add(id);
    const prev = S.insc[id];
    if (prev && prev.origen === 'lista' && prev.archivo === archivo && prev.hoja && prev.hoja !== hoja.trim()) out.warn.push(`${cleanName(nombre)} también aparece en la hoja «${prev.hoja}»; queda con la sede de «${hoja.trim()}».`);
    let inst = m.inst >= 0 ? toNum(row[m.inst]) : null, inv = m.inv >= 0 ? toNum(row[m.inv]) : null, fin = m.fin >= 0 ? toNum(row[m.fin]) : null;
    if (inv != null && an.invPond && inv <= S.cfg.pInv / 10 + 1e-9) inv = Math.round(inv / (S.cfg.pInv / 100) * 100) / 100;
    if (inst == null && inv == null && fin === 0) fin = null;
    const ok = x => x == null || (x >= 0 && x <= 10);
    if (!ok(inst) || !ok(inv) || !ok(fin)) out.warn.push(`${cleanName(nombre)}: calificación fuera de 0 a 10, revísala.`);
    const i = prev || { id, cadete: k, practica, inst: null, inv: null, finalLista: null, finalManual: null };
    Object.assign(i, { sede: sede || i.sede || '', hoja: hoja.trim(), docente: docente || i.docente || '', periodo: an.meta.periodo || i.periodo || '', horario: an.meta.horario || i.horario || '', turno: an.meta.turno || i.turno || '', archivo, origen: 'lista' });
    if (inst != null) i.inst = inst; if (inv != null) i.inv = inv; if (fin != null) i.finalLista = fin;
    if (inst != null || inv != null) i.finalManual = null;
    if (m.grupo >= 0 && cellTxt(row[m.grupo]).trim()) { S.cadetes[k].grupo = cellTxt(row[m.grupo]).trim(); S.cadetes[k].grupoManual = true; }
    S.insc[id] = i; out.n++; if (inst != null || inv != null || fin != null) out.conCalif++;
  }
  return out;
}

/* diálogo de mapeo manual de columnas */
function dialogoMapeo(archivo, hoja, sh, an, faltan) {
  return new Promise(resolve => {
    let tipo = an.tipo === 'faltas' ? 'faltas' : 'calif';
    let h = an.h >= 0 ? an.h : Math.max(0, findHeader(sh.aoa, row => row.filter(Boolean).length >= 2));
    const draw = () => {
      const hdr = sh.aoa[h] || []; const ncol = Math.max(1, hdr.length, ...sh.aoa.slice(h, h + 6).map(r => (r || []).length));
      const opts = sel => `<option value="-1">— sin columna —</option>` + Array.from({ length: ncol }, (_, i) => `<option value="${i}" ${sel === i ? 'selected' : ''}>${XLSX.utils.encode_col(i)} · ${esc(cellTxt(hdr[i]).slice(0, 32) || '(vacía)')}</option>`).join('');
      const auto = (tipo === an.tipo && h === an.h) ? an : (tipo === 'faltas' ? { map: mapFaltas(hdr) } : mapCalif(sh, h, hoja, archivo));
      const m = auto.map || {}; const meta = auto.meta || {};
      const campos = tipo === 'faltas'
        ? [['fecha', 'Fecha *'], ['nombre', 'Nombre del cadete *'], ['practica', 'Práctica'], ['motivo', 'Motivo'], ['motivoOrig', 'Motivo escrito en bitácora'], ['sede', 'Sede / hospital'], ['grupoTurno', 'Grupo o turno'], ['instructora', 'Instructora'], ['obs', 'Observaciones']]
        : [['nombre', 'Nombre del cadete *'], ['inst', 'Instrumento de evaluación'], ['inv', 'Trabajo de investigación'], ['fin', 'Evaluación parcial final'], ['practica', 'Práctica (columna)'], ['sede', 'Sede (columna)'], ['docente', 'Docente (columna)'], ['grupo', 'Grupo (columna)']];
      const prev = sh.aoa.slice(h, h + 6).map((r, ri) => `<tr>${Array.from({ length: Math.min(ncol, 12) }, (_, i) => ri ? `<td>${esc(cellTxt((r || [])[i]).slice(0, 26))}</td>` : `<th class="nos">${XLSX.utils.encode_col(i)} ${esc(cellTxt((r || [])[i]).slice(0, 22))}</th>`).join('')}</tr>`).join('');
      $('#modal').innerHTML = `<h2>Revisa las columnas</h2><div class="muted">${esc(archivo)} · hoja «${esc(hoja.trim())}»${faltan.length ? ` · No se reconoció: <b>${faltan.join(', ')}</b>` : ''}</div>
      <div class="mapgrid">
        <div class="field"><label>Tipo de hoja</label><select class="inp" id="mpTipo"><option value="calif" ${tipo === 'calif' ? 'selected' : ''}>Lista de calificaciones</option><option value="faltas" ${tipo === 'faltas' ? 'selected' : ''}>Registro de faltas</option></select></div>
        <div class="field"><label>Fila de encabezados</label><input class="inp" id="mpH" type="number" min="1" value="${h + 1}"></div>
        ${campos.map(([k, l]) => `<div class="field"><label>${l}</label><select class="inp" data-mp="${k}">${opts(m[k] ?? -1)}</select></div>`).join('')}
        <div class="field"><label>Práctica para toda la hoja</label><select class="inp" id="mpPr"><option value="">— tomar de la columna —</option>${PRACTICAS.map(p => `<option value="${p.id}" ${meta.practica === p.id ? 'selected' : ''}>${p.nombre}</option>`).join('')}</select></div>
        <div class="field"><label>Sede para toda la hoja</label><input class="inp" id="mpSede" value="${esc(meta.sede || '')}"></div>
        ${tipo === 'calif' ? `<div class="field"><label>Docente</label><input class="inp" id="mpDoc" value="${esc(meta.docente || '')}"></div>
        <div class="field"><label>Trabajo de investigación viene</label><select class="inp" id="mpPond"><option value="0" ${!auto.invPond ? 'selected' : ''}>En escala de 0 a 10</option><option value="1" ${auto.invPond ? 'selected' : ''}>Ya ponderado (0 a ${S.cfg.pInv / 10})</option></select></div>` : ''}
      </div>
      <div class="preview tbl-wrap"><table class="tbl">${prev}</table></div>
      <div class="mfoot"><button class="btn ghost" id="mpSkip">Omitir hoja</button><button class="btn" id="mpOk">Importar hoja</button></div>`;
      $('#mpTipo').onchange = e => { tipo = e.target.value; draw(); };
      $('#mpH').onchange = e => { h = Math.max(0, (+e.target.value || 1) - 1); draw(); };
      $('#mpSkip').onclick = () => { closeModal(); resolve(null); };
      $('#mpOk').onclick = () => {
        const map = {}; $$('[data-mp]').forEach(s => map[s.dataset.mp] = +s.value);
        const pr = $('#mpPr').value, sede = $('#mpSede').value.trim();
        if (map.nombre < 0) return toast('Elige la columna del nombre.', true);
        if (tipo === 'faltas') {
          if (map.fecha < 0) return toast('Elige la columna de la fecha.', true);
          if (map.practica < 0 && !pr) return toast('Elige la práctica (columna o para toda la hoja).', true);
          closeModal(); resolve({ tipo, h, map, fijo: { practica: pr || null, sede } });
        } else {
          if (!pr && map.practica < 0) return toast('Elige la práctica.', true);
          const res = { tipo, h, map, invPond: $('#mpPond').value === '1', meta: { ...meta, practica: pr || meta.practica, sede: sede || meta.sede || sedeCorta(hoja), docente: $('#mpDoc').value.trim() } }; closeModal(); resolve(res);
        }
      };
    };
    draw(); openModal();
  });
}
function openModal() { $('#modalBg').classList.add('on'); }
function closeModal() { $('#modalBg').classList.remove('on'); $('#modal').innerHTML = ''; }

/* ---------- exportación a Excel (Noto Sans 10) ---------- */
const XF = { name: 'Noto Sans', sz: 10 };
const LB = { style: 'thin', color: { rgb: 'D5D9C8' } };
const BORDER = { top: LB, bottom: LB, left: LB, right: LB };
function hojaEstilada(titulo, sub, head, data, widths) {
  const ws = XLSX.utils.aoa_to_sheet([[titulo], [sub], [], head, ...data]);
  const rg = XLSX.utils.decode_range(ws['!ref']);
  for (let r = rg.s.r; r <= rg.e.r; r++) for (let c = rg.s.c; c <= rg.e.c; c++) {
    const a = XLSX.utils.encode_cell({ r, c }); if (!ws[a]) { if (r < 3) continue; ws[a] = { t: 's', v: '' }; }
    const s = { font: { ...XF }, alignment: { vertical: 'center', wrapText: r > 3 } };
    if (r === 0) s.font = { ...XF, bold: true, color: { rgb: '3B4425' } };
    else if (r === 1) s.font = { ...XF, color: { rgb: '6B7160' } };
    else if (r === 3) { s.font = { ...XF, bold: true, color: { rgb: 'FFFFFF' } }; s.fill = { patternType: 'solid', fgColor: { rgb: '4B5320' } }; s.border = BORDER; s.alignment = { vertical: 'center', horizontal: 'center', wrapText: true }; }
    else if (r > 3) { s.border = BORDER; const v = ws[a].v; if (v === 'Alerta') s.fill = { patternType: 'solid', fgColor: { rgb: 'FBE3DF' } }; else if (v === 'Advertencia') s.fill = { patternType: 'solid', fgColor: { rgb: 'FDF1D8' } }; else if (v === 'Sí') s.fill = { patternType: 'solid', fgColor: { rgb: 'F3E3F0' } }; if (typeof v === 'number') s.alignment.horizontal = 'center'; }
    ws[a].s = s;
  }
  ws['!cols'] = widths.map(w => ({ wch: w }));
  ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: Math.max(0, head.length - 1) } }, { s: { r: 1, c: 0 }, e: { r: 1, c: Math.max(0, head.length - 1) } }];
  if (data.length) ws['!autofilter'] = { ref: XLSX.utils.encode_range({ r: 3, c: 0 }, { r: 3 + data.length, c: head.length - 1 }) };
  return ws;
}
const ESTN = { alerta: 'Alerta', advertencia: 'Advertencia', regular: 'Regular' };
const n2 = v => v == null ? '' : Math.round(v * 100) / 100;
function filaResumen(r) {
  return [r.nombre, r.grupo, PMAP[r.practica].nombre, r.sede, r.hoja || '', r.docente, r.nf, n2(r.pct), ESTN[r.estado], r.estado === 'alerta' ? 0 : r.paraAlerta, ...MOTIVOS.map(m => r.mot[m] || 0), n2(r.inst), n2(r.inv), n2(r.final), r.final == null ? 'Sin calificación' : (r.riesgo ? 'Sí' : 'No'), r.ultAccion ? `${fFecha(r.ultAccion.fecha)} · ${r.ultAccion.accion}` : '', r.ultAccion?.responsable || ''];
}
const HEAD_RES = ['Cadete', 'Grupo', 'Práctica', 'Sede', 'Hoja / grupo de práctica', 'Docente', 'Faltas', '% faltas', 'Estado', 'Faltas para alerta', ...MOTIVOS, 'Instrumento', 'Trabajo de investigación', 'Evaluación parcial final', 'Riesgo académico', 'Última acción', 'Responsable'];
const W_RES = [30, 8, 26, 22, 16, 30, 7, 8, 12, 9, 9, 9, 10, 10, 9, 8, 7, 10, 11, 11, 13, 34, 20];
function libroReporte(tipo, rowsIn) {
  const d = derive(); const c = S.cfg; const fecha = hoyISO();
  let rows = rowsIn || d.rows;
  const wb = XLSX.utils.book_new();
  const sub = `Generado el ${fFecha(fecha)} · % = faltas × ${c.horasFalta} h × 100 / horas de la práctica · advertencia desde ${c.adv} %, alerta desde ${c.alerta} % · riesgo académico: calificación menor a ${c.umbral}`;
  if (tipo === 'riesgo') rows = rows.filter(r => r.estado !== 'regular' || r.riesgo);
  const titulo = tipo === 'riesgo' ? 'Cadetes en riesgo · práctica clínica' : 'Reporte de seguimiento · práctica clínica';
  XLSX.utils.book_append_sheet(wb, hojaEstilada(titulo, sub, HEAD_RES, rows.map(filaResumen), W_RES), tipo === 'riesgo' ? 'En riesgo' : 'Resumen');
  const setIns = new Set(rows.map(r => r.id)); const setCad = new Set(rows.map(r => r.cadete));
  const nom = k => S.cadetes[k]?.nombre || '';
  const fl = Object.values(S.faltas).filter(f => setIns.has(f.cadete + '|' + f.practica)).sort((a, b) => a.fecha.localeCompare(b.fecha) || nom(a.cadete).localeCompare(nom(b.cadete)));
  XLSX.utils.book_append_sheet(wb, hojaEstilada('Detalle de faltas', sub, ['Fecha', 'Cadete', 'Práctica', 'Sede', 'Grupo o turno', 'Motivo', 'Motivo escrito', 'Instructora', 'Observaciones'],
    fl.map(f => [f.fecha, nom(f.cadete), PMAP[f.practica].nombre, f.sede, f.grupoTurno, f.motivo, f.motivoOrig, f.instructora, f.obs]), [11, 30, 26, 22, 16, 14, 24, 30, 40]), 'Faltas');
  const ac = Object.values(S.acciones).filter(a => setCad.has(a.cadete)).sort((a, b) => (a.fecha || '').localeCompare(b.fecha || ''));
  XLSX.utils.book_append_sheet(wb, hojaEstilada('Acciones de seguimiento', sub, ['Fecha', 'Cadete', 'Grupo', 'Acción', 'Responsable', 'Estado', 'Observaciones'],
    ac.map(a => [a.fecha, nom(a.cadete), S.cadetes[a.cadete]?.grupo || '', a.accion, a.responsable, a.estado || '', a.obs]), [11, 30, 8, 40, 24, 12, 44]), 'Acciones');
  if (tipo !== 'riesgo') {
    const ag = agrupar(rows, r => PMAP[r.practica].nombre);
    XLSX.utils.book_append_sheet(wb, hojaEstilada('Resumen por práctica', sub, ['Práctica', 'Registros', 'Alerta', 'Advertencia', 'Con calificación', 'Promedio', 'Riesgo académico', 'Faltas'],
      ag.map(g => [g.k, g.n, g.alerta, g.adv, g.conCal, n2(g.prom), g.riesgo, g.faltas]), [30, 9, 9, 11, 13, 10, 13, 9]), 'Por práctica');
  }
  return { wb, nombre: `${tipo === 'riesgo' ? 'Riesgo' : 'Seguimiento'}_practica_clinica_${fecha}.xlsx` };
}
function exportar(tipo, rowsIn) {
  if (!window.XLSX) return toast('No se cargó el módulo de Excel.', true);
  if (!derive().rows.length) return toast('Aún no hay datos para exportar.', true);
  const { wb, nombre } = libroReporte(tipo, rowsIn);
  XLSX.writeFile(wb, nombre);
  toast('Excel generado');
}
function agrupar(rows, fk) {
  const m = new Map();
  for (const r of rows) { const k = fk(r) || 'Sin dato'; const g = m.get(k) || { k, n: 0, alerta: 0, adv: 0, conCal: 0, suma: 0, riesgo: 0, faltas: 0, pct: 0 }; g.n++; g.pct += r.pct; if (r.estado === 'alerta') g.alerta++; if (r.estado === 'advertencia') g.adv++; if (r.final != null) { g.conCal++; g.suma += r.final; } if (r.riesgo) g.riesgo++; g.faltas += r.nf; m.set(k, g); }
  return [...m.values()].map(g => ({ ...g, prom: g.conCal ? g.suma / g.conCal : null, pctProm: g.pct / g.n })).sort((a, b) => a.k.localeCompare(b.k, 'es', { numeric: true }));
}

/* ---------- gráficas ---------- */
function hbars(items, { max, fmtV = v => fmt(v) } = {}) {
  if (!items.length) return '<div class="muted small">Sin datos.</div>';
  const mx = max ?? Math.max(...items.map(i => i.v || 0), 1);
  return items.map(i => `<div class="hbar" title="${esc(i.l)}"><span class="lb">${esc(i.l)}</span><span class="tr"><span class="fl" data-w="${Math.max(0, Math.min(100, (i.v || 0) / mx * 100))}" ${i.c ? `style="background:${i.c}"` : ''}></span></span><span class="vl">${i.v == null ? '—' : fmtV(i.v)}</span></div>`).join('');
}
function donut(items) {
  const tot = items.reduce((s, i) => s + i.v, 0); if (!tot) return '<div class="muted small">Sin faltas registradas.</div>';
  const R = 52, C = 2 * Math.PI * R; let acc = 0;
  const arcs = items.filter(i => i.v).map(i => { const len = i.v / tot * C; const s = `<circle r="${R}" cx="70" cy="70" fill="none" stroke="${i.c}" stroke-width="22" stroke-dasharray="0 ${C}" data-da="${len} ${C - len}" stroke-dashoffset="${-acc}" transform="rotate(-90 70 70)"><title>${esc(i.l)}: ${i.v}</title></circle>`; acc += len; return s; }).join('');
  return `<div class="donut-wrap"><svg class="donut" viewBox="0 0 140 140">${arcs}<text x="70" y="70" text-anchor="middle" font-size="22" font-weight="800" fill="#1f2418">${tot}</text><text x="70" y="86" text-anchor="middle" font-size="10" fill="#6b7160">faltas</text></svg>
  <div class="dl">${items.filter(i => i.v).map(i => `<div><span><i style="background:${i.c}"></i>${esc(i.l)}</span><b>${i.v} <span class="muted small">(${fmt(i.v / tot * 100, 0)} %)</span></b></div>`).join('')}</div></div>`;
}
function vbars(items) {
  const mx = Math.max(...items.map(i => i.v), 1);
  return `<div class="vbars">${items.map(i => `<div class="vb"><b>${i.v}</b><div data-h="${i.v / mx * 100}" style="background:${i.c}"></div></div>`).join('')}</div><div class="vlabels">${items.map(i => `<span>${esc(i.l)}</span>`).join('')}</div>`;
}
function lineaSemanal(faltas) {
  if (!faltas.length) return '<div class="muted small">Sin faltas registradas.</div>';
  const wk = iso => { const d = new Date(iso + 'T12:00:00'); d.setDate(d.getDate() - (d.getDay() + 6) % 7); return d.toISOString().slice(0, 10); };
  const m = new Map(); faltas.forEach(f => { const k = wk(f.fecha); m.set(k, (m.get(k) || 0) + 1); });
  const ks = [...m.keys()].sort(); const last = new Date(ks.at(-1) + 'T12:00:00');
  const pts = []; for (let d = new Date(ks[0] + 'T12:00:00'); d <= last; d.setDate(d.getDate() + 7)) { const k = d.toISOString().slice(0, 10); pts.push({ k, v: m.get(k) || 0 }); }
  const W = 600, H = 150, P = 30, mx = Math.max(...pts.map(p => p.v), 1);
  const x = i => P + (pts.length === 1 ? (W - 2 * P) / 2 : i * (W - 2 * P) / (pts.length - 1)), y = v => H - 24 - v / mx * (H - 46);
  const path = pts.map((p, i) => `${i ? 'L' : 'M'}${x(i)},${y(p.v)}`).join(' ');
  return `<svg class="line-svg" viewBox="0 0 ${W} ${H}" preserveAspectRatio="xMidYMid meet"><defs><linearGradient id="lg" x1="0" x2="0" y1="0" y2="1"><stop offset="0" stop-color="#6b7a3a" stop-opacity=".35"/><stop offset="1" stop-color="#6b7a3a" stop-opacity="0"/></linearGradient></defs>
  <path d="${path} L${x(pts.length - 1)},${H - 24} L${x(0)},${H - 24} Z" fill="url(#lg)"/><path d="${path}" fill="none" stroke="#4b5320" stroke-width="2.5" stroke-linejoin="round"/>
  ${pts.map((p, i) => `<circle cx="${x(i)}" cy="${y(p.v)}" r="4" fill="#c9a227"><title>Semana del ${fFecha(p.k)}: ${p.v}</title></circle><text x="${x(i)}" y="${y(p.v) - 9}" font-size="11" text-anchor="middle" fill="#1f2418" font-weight="700">${p.v}</text><text x="${x(i)}" y="${H - 6}" font-size="10" text-anchor="middle" fill="#6b7160">sem. ${fFecha(p.k).slice(0, -5)}</text>`).join('')}</svg>`;
}
function animar(root = document) {
  requestAnimationFrame(() => requestAnimationFrame(() => {
    $$('[data-w]', root).forEach(e => e.style.width = e.dataset.w + '%');
    $$('[data-h]', root).forEach(e => e.style.height = e.dataset.h + '%');
    $$('[data-da]', root).forEach(e => e.setAttribute('stroke-dasharray', e.dataset.da));
    $$('[data-count]', root).forEach(e => { const to = +e.dataset.count, dec = +(e.dataset.dec || 0), t0 = performance.now(); const step = t => { const k = Math.min(1, (t - t0) / 700); e.textContent = fmt(to * (1 - Math.pow(1 - k, 3)), dec); if (k < 1) requestAnimationFrame(step); }; requestAnimationFrame(step); });
  }));
}
const badge = e => `<span class="badge b-${e}">${ESTN[e]}</span>`;
const colorPct = p => p >= S.cfg.alerta ? 'var(--bad)' : p >= S.cfg.adv ? 'var(--warn)' : 'var(--ok)';

/* ---------- vistas ---------- */
function bienvenida() {
  return `<div class="hero"><h2>Seguimiento de práctica clínica</h2><p>Carga tus listas de práctica clínica (con las calificaciones de instrumento, trabajo de investigación y evaluación parcial) y el registro de inasistencias. La app cruza faltas y calificaciones, calcula porcentajes y te avisa quién está en advertencia o alerta.</p>
  <a class="btn gold" href="#/importar">Importar archivos</a></div>
  <div class="steps"><div class="step"><b>1</b><div><strong>Importa</strong> las listas por sede (.xlsx) y el registro de faltas.</div></div><div class="step"><b>2</b><div><strong>Revisa</strong> el tablero: alertas, promedios y distribución.</div></div><div class="step"><b>3</b><div><strong>Abre la ficha</strong> de cada cadete y anota las acciones de seguimiento.</div></div><div class="step"><b>4</b><div><strong>Exporta</strong> el reporte en Excel y guarda un respaldo.</div></div></div>`;
}
function vTablero() {
  const d = derive(); const rows = d.rows;
  if (!rows.length) return bienvenida();
  const cads = Object.keys(d.byCad);
  const peor = k => d.byCad[k].reduce((m, r) => Math.max(m, ORD[r.estado]), 0);
  const nAl = cads.filter(k => peor(k) === 3).length, nAd = cads.filter(k => peor(k) === 2).length;
  const conCal = rows.filter(r => r.final != null), prom = conCal.length ? conCal.reduce((s, r) => s + r.final, 0) / conCal.length : null;
  const nRies = new Set(rows.filter(r => r.riesgo).map(r => r.cadete)).size;
  const faltas = Object.values(S.faltas);
  const porPr = PRACTICAS.map(p => { const rs = rows.filter(r => r.practica === p.id); return { p, n: rs.length, al: rs.filter(r => r.estado === 'alerta').length, ad: rs.filter(r => r.estado === 'advertencia').length }; }).filter(x => x.n);
  const mot = MOTIVOS.map(m => ({ l: m, v: faltas.filter(f => f.motivo === m).length, c: MCOLOR[m] }));
  const bins = [['< 6', 0, 6, '#b23a2e'], ['6 a 6.9', 6, 7, '#d0662a'], ['7 a 7.9', 7, 8, '#c9a227'], ['8 a 8.9', 8, 9, '#6b7a3a'], ['9 a 10', 9, 10.01, '#3f7d4e']].map(([l, a, b, c]) => ({ l, v: conCal.filter(r => r.final >= a && r.final < b).length, c }));
  const top = rows.filter(r => r.estado !== 'regular').slice(0, 10);
  const ries = rows.filter(r => r.riesgo).sort((a, b) => a.final - b.final).slice(0, 10);
  return `<div class="grid kpis">
    ${kpi('Cadetes', cads.length, `${rows.length} registros en ${porPr.length} prácticas`, '', '#/listado?estado=')}
    ${kpi('En alerta', nAl, `desde ${S.cfg.alerta} % de faltas`, 'bad', '#/listado?estado=alerta')}
    ${kpi('En advertencia', nAd, `desde ${S.cfg.adv} % de faltas`, 'warn', '#/listado?estado=advertencia')}
    ${kpi('Faltas registradas', faltas.length, 'los retardos no cuentan', 'gold')}
    ${kpi('Promedio general', prom, `${conCal.length} con calificación final`, 'ok', null, 2)}
    ${kpi('Riesgo académico', nRies, `calificación menor a ${S.cfg.umbral}`, nRies ? 'bad' : 'ok', '#/listado?estado=riesgo')}
  </div>
  <div class="grid cols3" style="margin-top:16px">
    <div class="card span2"><h3>Estado de faltas por práctica <span class="sub">cadetes por práctica</span></h3>
      ${porPr.map(x => `<div style="margin:10px 0"><div class="row" style="justify-content:space-between"><b>${esc(x.p.nombre)}</b><span class="small muted">${x.al} alerta · ${x.ad} advertencia · ${x.n - x.al - x.ad} regular</span></div>
      <div class="stack" style="margin-top:5px"><span data-w="${x.al / x.n * 100}" style="background:var(--bad)"></span><span data-w="${x.ad / x.n * 100}" style="background:var(--warn)"></span><span data-w="${(x.n - x.al - x.ad) / x.n * 100}" style="background:#9fb58a"></span></div></div>`).join('')}
      <div class="legend"><span><i style="background:var(--bad)"></i>Alerta</span><span><i style="background:var(--warn)"></i>Advertencia</span><span><i style="background:#9fb58a"></i>Regular</span></div></div>
    <div class="card"><h3>Faltas por motivo</h3>${donut(mot)}</div>
    <div class="card span2"><h3>Promedios de calificación <span class="tabs" id="promTabs"><button data-g="practica" class="on">Práctica</button><button data-g="sede">Sede</button><button data-g="grupo">Grupo</button><button data-g="docente">Docente</button></span></h3><div id="promBox"></div></div>
    <div class="card"><h3>Distribución de calificaciones</h3>${conCal.length ? vbars(bins) : '<div class="empty small">Aún no hay calificaciones finales.<br>Captúralas o importa las listas calificadas.</div>'}</div>
    <div class="card span2"><h3>Faltas por semana</h3>${lineaSemanal(faltas)}</div>
    <div class="card"><h3>% de faltas por sede <span class="sub">promedio</span></h3>${hbars(agrupar(rows, r => r.sede).map(g => ({ l: g.k, v: g.pctProm, c: g.pctProm >= S.cfg.adv ? 'linear-gradient(90deg,#e0b45a,#c98a12)' : null })).sort((a, b) => b.v - a.v).slice(0, 10), { max: Math.max(S.cfg.alerta, 1), fmtV: v => fmt(v) + ' %' })}</div>
    <div class="card"><h3>Alerta y advertencia <a class="btn ghost sm" href="#/listado?estado=riesgoFaltas">Ver todos</a></h3><div class="list-mini">${top.length ? top.map(r => miniRow(r, `${fmt(r.pct, 2)} %`)).join('') : '<div class="muted small">Nadie en advertencia ni en alerta.</div>'}</div></div>
    <div class="card"><h3>Riesgo académico <a class="btn ghost sm" href="#/listado?estado=riesgo">Ver todos</a></h3><div class="list-mini">${ries.length ? ries.map(r => miniRow(r, fmt2(r.final), 'riesgo')).join('') : `<div class="muted small">${conCal.length ? `Nadie con calificación menor a ${S.cfg.umbral}.` : 'Sin calificaciones todavía.'}</div>`}</div></div>
    <div class="card"><h3>Exportar</h3><p class="muted small">Reportes en Excel con letra Noto Sans 10.</p><div class="row"><button class="btn" data-act="export-seg">Reporte de seguimiento</button><button class="btn gold" data-act="export-riesgo">Reporte de riesgo</button></div></div>
  </div>`;
}
function kpi(l, v, s, cls, href, dec = 0) {
  return `<div class="card kpi ${cls} ${href ? 'click' : ''}" ${href ? `data-href="${href}"` : ''}><div class="k-l">${l}</div><div class="k-v" ${v != null ? `data-count="${v}" data-dec="${dec}"` : ''}>${v == null ? '—' : 0}</div><div class="k-s">${s}</div></div>`;
}
function miniRow(r, val, est) {
  return `<a href="#/cadete/${encodeURIComponent(r.cadete)}"><span><b>${esc(r.nombre)}</b><br><span class="small muted">${esc(PMAP[r.practica].corto)} · ${esc(r.sede)}${r.grupo ? ' · ' + esc(r.grupo) : ''}</span></span><span class="badge b-${est === 'riesgo' ? 'riesgo' : r.estado}">${val}</span></a>`;
}
function promBox(g) {
  const d = derive(); const fk = { practica: r => PMAP[r.practica].corto, sede: r => r.sede, grupo: r => r.grupo || 'Sin grupo', docente: r => r.docente || 'Sin docente' }[g];
  const items = agrupar(d.rows, fk).filter(x => x.conCal).map(x => ({ l: `${x.k} (${x.conCal})`, v: x.prom, c: x.prom < S.cfg.umbral ? 'linear-gradient(90deg,#d07a6f,#b23a2e)' : null })).sort((a, b) => b.v - a.v);
  $('#promBox').innerHTML = items.length ? hbars(items.slice(0, 14), { max: 10, fmtV: fmt2 }) : '<div class="empty small">Aún no hay calificaciones finales. Importa las listas calificadas o captúralas en <a href="#/captura">Calificaciones</a>.</div>';
  animar($('#promBox'));
}

/* listado */
let LF = { q: '', practica: '', sede: '', grupo: '', docente: '', estado: '', sort: 'pct', dir: -1 };
function opciones(vals, sel, todo) { return `<option value="">${todo}</option>` + [...new Set(vals.filter(Boolean))].sort((a, b) => a.localeCompare(b, 'es', { numeric: true })).map(v => `<option ${v === sel ? 'selected' : ''} value="${esc(v)}">${esc(v)}</option>`).join(''); }
function filtrar(rows) {
  const q = norm(LF.q);
  return rows.filter(r => (!q || norm(r.nombre + ' ' + r.sede + ' ' + r.docente + ' ' + r.grupo).includes(q)) && (!LF.practica || r.practica === LF.practica) && (!LF.sede || r.sede === LF.sede) && (!LF.grupo || r.grupo === LF.grupo) && (!LF.docente || r.docente === LF.docente) &&
    (!LF.estado || (LF.estado === 'riesgo' ? r.riesgo : LF.estado === 'sincal' ? r.final == null : LF.estado === 'riesgoFaltas' ? r.estado !== 'regular' : r.estado === LF.estado)));
}
function vListado(params) {
  const d = derive(); if (!d.rows.length) return bienvenida();
  if (params.estado != null) LF.estado = params.estado;
  const sel = (id, lab, html) => `<div class="field"><label>${lab}</label><select class="inp" id="${id}">${html}</select></div>`;
  return `<div class="filters">
    <div class="field fsearch"><label>Buscar</label><input class="inp" id="fQ" type="search" placeholder="Nombre, sede, docente…" value="${esc(LF.q)}"></div>
    ${sel('fPr', 'Práctica', `<option value="">Todas</option>${PRACTICAS.map(p => `<option value="${p.id}" ${LF.practica === p.id ? 'selected' : ''}>${p.corto}</option>`).join('')}`)}
    ${sel('fSe', 'Sede', opciones(d.rows.map(r => r.sede), LF.sede, 'Todas'))}
    ${sel('fGr', 'Grupo', opciones(d.rows.map(r => r.grupo), LF.grupo, 'Todos'))}
    ${sel('fDo', 'Docente', opciones(d.rows.map(r => r.docente), LF.docente, 'Todas'))}
    ${sel('fEs', 'Estado', [['', 'Todos'], ['alerta', 'Alerta'], ['advertencia', 'Advertencia'], ['riesgoFaltas', 'Alerta o advertencia'], ['regular', 'Regular'], ['riesgo', `Riesgo académico (menor a ${S.cfg.umbral})`], ['sincal', 'Sin calificación']].map(([v, l]) => `<option value="${v}" ${LF.estado === v ? 'selected' : ''}>${l}</option>`).join(''))}
  </div>
  <div class="count-line"><span id="lCount" class="muted"></span><span class="row"><button class="btn ghost sm" id="lClear">Quitar filtros</button><button class="btn sm" id="lExp">Exportar esta vista</button><button class="btn gold sm" data-act="export-riesgo">Reporte de riesgo</button></span></div>
  <div class="tbl-wrap resp"><table class="tbl" id="lTbl"><thead><tr>
    ${[['nombre', 'Cadete'], ['grupo', 'Grupo'], ['practica', 'Práctica'], ['sede', 'Sede'], ['docente', 'Docente'], ['nf', 'Faltas'], ['pct', '% faltas'], ['estado', 'Estado'], ['final', 'Calificación'], ['nAcc', 'Acciones']].map(([k, l]) => `<th data-s="${k}" class="${['nf', 'pct', 'final', 'nAcc'].includes(k) ? 'num' : ''}">${l}${LF.sort === k ? (LF.dir > 0 ? ' ▲' : ' ▼') : ''}</th>`).join('')}
  </tr></thead><tbody id="lBody"></tbody></table></div><div class="mcards" id="lCards"></div>`;
}
let ultimaVista = null;
function pintarListado() {
  const d = derive(); let rs = filtrar(d.rows);
  const k = LF.sort, dir = LF.dir;
  rs = rs.slice().sort((a, b) => { let x = a[k], y = b[k]; if (k === 'estado') { x = ORD[x]; y = ORD[y]; } if (k === 'practica') { x = PMAP[x].corto; y = PMAP[y].corto; } if (x == null) x = -1; if (y == null) y = -1; return (typeof x === 'string' ? x.localeCompare(String(y), 'es', { numeric: true }) : x - y) * dir || a.nombre.localeCompare(b.nombre); });
  $('#lCount').textContent = `${rs.length} de ${d.rows.length} registros (cadete por práctica)`;
  const fin = r => r.final == null ? '<span class="muted">—</span>' : (r.riesgo ? `<span class="badge b-riesgo">${fmt2(r.final)}</span>` : `<b>${fmt2(r.final)}</b>`);
  $('#lBody').innerHTML = rs.slice(0, 800).map(r => `<tr class="clk" data-c="${esc(r.cadete)}"><td><b>${esc(r.nombre)}</b></td><td>${esc(r.grupo) || '<span class="muted">—</span>'}</td><td>${esc(PMAP[r.practica].corto)}</td><td>${esc(r.sede)}</td><td class="small">${esc(r.docente)}</td><td class="num">${r.nf}</td><td class="num"><span class="pbar"><i style="width:${Math.min(100, r.pct / Math.max(S.cfg.alerta, 1) * 100)}%;background:${colorPct(r.pct)}"></i></span>${fmt(r.pct, 2)} %</td><td>${badge(r.estado)}</td><td class="num">${fin(r)}</td><td class="num">${r.nAcc || ''}</td></tr>`).join('') || `<tr><td colspan="10" class="empty">Sin resultados con estos filtros.</td></tr>`;
  $('#lCards').innerHTML = rs.slice(0, 300).map(r => `<div class="mcard" data-c="${esc(r.cadete)}"><div class="row" style="justify-content:space-between;align-items:center"><b>${esc(r.nombre)}</b>${badge(r.estado)}</div><div class="small muted">${esc(PMAP[r.practica].corto)} · ${esc(r.sede)}${r.grupo ? ' · ' + esc(r.grupo) : ''}</div><div class="row small" style="margin-top:6px;justify-content:space-between"><span>${r.nf} faltas · <b>${fmt(r.pct, 2)} %</b></span><span>Calif.: ${fin(r)}</span></div></div>`).join('');
  $$('[data-c]', $('#view')).forEach(e => e.onclick = () => location.hash = '#/cadete/' + encodeURIComponent(e.dataset.c));
  ultimaVista = rs;
}
function bindListado() {
  const upd = () => { LF.q = $('#fQ').value; LF.practica = $('#fPr').value; LF.sede = $('#fSe').value; LF.grupo = $('#fGr').value; LF.docente = $('#fDo').value; LF.estado = $('#fEs').value; pintarListado(); };
  ['#fQ', '#fPr', '#fSe', '#fGr', '#fDo', '#fEs'].forEach(s => $(s).addEventListener(s === '#fQ' ? 'input' : 'change', upd));
  $('#lClear').onclick = () => { LF = { ...LF, q: '', practica: '', sede: '', grupo: '', docente: '', estado: '' }; location.hash = '#/listado'; render(); };
  $('#lExp').onclick = () => exportar('seguimiento', ultimaVista);
  $$('#lTbl th[data-s]').forEach(th => th.onclick = () => { const k = th.dataset.s; if (LF.sort === k) LF.dir *= -1; else { LF.sort = k; LF.dir = ['nombre', 'grupo', 'practica', 'sede', 'docente'].includes(k) ? 1 : -1; } render(true); });
  pintarListado();
}

/* ficha del cadete */
function vCadete(key) {
  const c = S.cadetes[key]; if (!c) return `<div class="empty"><h2>No se encontró el cadete</h2><a class="btn" href="#/listado">Volver al listado</a></div>`;
  const d = derive(); const rs = (d.byCad[key] || []).slice().sort((a, b) => b.pct - a.pct);
  const peor = rs.reduce((m, r) => ORD[r.estado] > ORD[m] ? r.estado : m, 'regular');
  const acc = Object.values(S.acciones).filter(a => a.cadete === key).sort((a, b) => (b.fecha || '').localeCompare(a.fecha || ''));
  const todas = rs.flatMap(r => r.faltas);
  const mot = MOTIVOS.map(m => ({ l: m, v: todas.filter(f => f.motivo === m).length, c: MCOLOR[m] }));
  const tl = [...todas.map(f => ({ t: 'f', fecha: f.fecha, f })), ...acc.map(a => ({ t: 'a', fecha: a.fecha || '', a }))].sort((a, b) => b.fecha.localeCompare(a.fecha));
  const c30 = S.cfg.alerta, c20 = S.cfg.adv, esc15 = c30 * 1.5;
  const inp = (r, k, v, lab, extra = '') => `<div class="field"><label>${lab}</label><input class="inp" type="number" step="0.01" min="0" max="10" data-g="${k}" data-i="${esc(r.id)}" value="${v ?? ''}" ${extra}></div>`;
  return `<div class="row" style="margin-bottom:12px"><a class="btn ghost sm" href="#/listado">← Volver al listado</a></div>
  <div class="card"><div class="fhead"><div class="avatar">${esc(iniciales(c.nombre))}</div><div class="grow"><h2 class="fname">${esc(c.nombre)}</h2>
    <div class="muted">${c.grupo ? 'Grupo ' + esc(c.grupo) : 'Sin grupo asignado'}${c.matricula ? ' · Matrícula ' + esc(c.matricula) : ''} · ${rs.length} práctica${rs.length === 1 ? '' : 's'} · ${todas.length} falta${todas.length === 1 ? '' : 's'}</div></div>
    <div class="row" style="align-items:center">${badge(peor)}${rs.some(r => r.riesgo) ? '<span class="badge b-riesgo">Riesgo académico</span>' : ''}<button class="btn ghost sm" id="edGrupo">Editar datos</button></div></div></div>
  <div class="grid cols2" style="margin-top:16px">
    <div class="grid" style="align-content:start">
    ${rs.map(r => `<div class="card enr ${r.estado}"><h3><span>${esc(PMAP[r.practica].nombre)}</span>${badge(r.estado)}</h3>
      <div class="small muted">${esc(r.sede)}${r.hoja ? ' · lista «' + esc(r.hoja) + '»' : ''}${r.turno ? ' · ' + esc(r.turno) : ''}${r.docente ? '<br>Docente: ' + esc(r.docente) : ''}${r.periodo ? '<br>' + esc(r.periodo) : ''}${r.horario ? ' · ' + esc(r.horario) : ''}</div>
      <div class="meter"><div class="mf" data-w="${Math.min(100, r.pct / esc15 * 100)}" style="background:${colorPct(r.pct)}"></div><span class="mk" style="left:${c20 / esc15 * 100}%"><span>${c20} %</span></span><span class="mk" style="left:${c30 / esc15 * 100}%"><span>${c30} %</span></span></div>
      <div class="gr3"><div class="gbox"><div class="gl">Faltas</div><div class="gv">${r.nf}</div></div><div class="gbox"><div class="gl">% de faltas</div><div class="gv" style="color:${colorPct(r.pct)}">${fmt(r.pct, 2)}</div></div><div class="gbox"><div class="gl">Faltas para alerta</div><div class="gv">${r.estado === 'alerta' ? '—' : r.paraAlerta}</div></div></div>
      <div class="small muted" style="margin:6px 0 10px">${S.cfg.horas[r.practica]} h de práctica · cada falta suma ${fmt(S.cfg.horasFalta * 100 / S.cfg.horas[r.practica], 2)} %</div>
      <div class="gr3">${inp(r, 'inst', r.inst, 'Instrumento')}${inp(r, 'inv', r.inv, 'Trabajo de invest.')}${inp(r, 'fin', r.final, 'Evaluación final', `placeholder="auto" style="${r.riesgo ? 'border-color:var(--bad);color:var(--bad);font-weight:700' : ''}"`)}</div>
      <div class="small muted" style="margin-top:4px">Final = instrumento × ${S.cfg.pInst} % + trabajo × ${S.cfg.pInv} %.${r.finalManual != null ? ' <b>La final está escrita a mano.</b>' : ' Si escribes la final a mano, queda fija.'}</div>
      ${r.nf ? `<div style="margin-top:10px">${Object.entries(r.mot).map(([m, n]) => `<span class="chip" style="background:${MCOLOR[m]}22">${esc(m)}: <b>${n}</b></span>`).join('')}</div>` : ''}
    </div>`).join('')}
    <div class="card"><h3>Faltas por motivo</h3>${donut(mot)}</div>
    </div>
    <div class="grid" style="align-content:start">
      <div class="card"><h3>Bitácora de seguimiento <button class="btn sm" id="addAcc">+ Acción</button></h3>
        ${acc.length ? `<div class="tbl-wrap"><table class="tbl" style="min-width:0"><thead><tr><th class="nos">Fecha</th><th class="nos">Acción</th><th class="nos">Responsable</th><th class="nos">Estado</th><th class="nos"></th></tr></thead><tbody>${acc.map(a => `<tr><td class="small" style="white-space:nowrap">${fFecha(a.fecha)}</td><td><b>${esc(a.accion)}</b>${a.obs ? `<div class="small muted">${esc(a.obs)}</div>` : ''}</td><td class="small">${esc(a.responsable)}</td><td><span class="badge ${a.estado === 'Cerrada' ? 'b-regular' : a.estado === 'En curso' ? 'b-info' : 'b-advertencia'}">${esc(a.estado || 'Pendiente')}</span></td><td><button class="btn ghost icon" data-ea="${esc(a.id)}" title="Editar">✎</button></td></tr>`).join('')}</tbody></table></div>` : '<div class="muted small">Sin acciones registradas. Agrega la primera: entrevista, aviso a la docente, tarjeta informativa, plan de recuperación…</div>'}
      </div>
      <div class="card"><h3>Línea de tiempo <button class="btn ghost sm" id="addFal">+ Falta</button></h3>
        ${tl.length ? `<div class="timeline">${tl.map(x => x.t === 'f' ? `<div class="tl-i" style="--c:${MCOLOR[x.f.motivo]}"><div class="tl-d">${fFecha(x.fecha)} · ${esc(PMAP[x.f.practica].corto)}</div><div class="tl-x"><div><b>Falta · ${esc(x.f.motivo)}</b>${x.f.motivoOrig && norm(x.f.motivoOrig) !== norm(x.f.motivo) ? ` <span class="muted small">(${esc(x.f.motivoOrig)})</span>` : ''}<div class="small muted">${esc([x.f.sede, x.f.grupoTurno, x.f.obs].filter(Boolean).join(' · '))}</div></div><button class="btn ghost icon" data-df="${esc(x.f.id)}" title="Quitar falta">✕</button></div></div>`
          : `<div class="tl-i" style="--c:var(--gold)"><div class="tl-d">${fFecha(x.fecha)} · Seguimiento</div><b>${esc(x.a.accion)}</b><div class="small muted">${esc([x.a.responsable, x.a.estado].filter(Boolean).join(' · '))}</div></div>`).join('')}</div>` : '<div class="muted small">Sin faltas ni acciones.</div>'}
      </div>
    </div>
  </div>`;
}
function bindCadete(key) {
  $$('input[data-g]').forEach(inp => inp.addEventListener('change', () => {
    const i = S.insc[inp.dataset.i]; const v = inp.value === '' ? null : toNum(inp.value);
    if (v != null && (v < 0 || v > 10)) { toast('La calificación va de 0 a 10.', true); inp.value = ''; return; }
    if (inp.dataset.g === 'fin') i.finalManual = v; else { i[inp.dataset.g] = v; i.finalManual = null; }
    save(); render(true); toast('Calificación guardada');
  }));
  $$('[data-df]').forEach(b => b.onclick = () => { if (confirm('¿Quitar esta falta del registro?')) { delete S.faltas[b.dataset.df]; save(); render(true); toast('Falta quitada'); } });
  $$('[data-ea]').forEach(b => b.onclick = () => formAccion(key, S.acciones[b.dataset.ea]));
  $('#addAcc').onclick = () => formAccion(key);
  $('#addFal').onclick = () => formFalta(key);
  $('#edGrupo').onclick = () => {
    const c = S.cadetes[key];
    $('#modal').innerHTML = `<h2>Datos del cadete</h2><div class="mapgrid"><div class="field"><label>Nombre</label><input class="inp" id="cN" value="${esc(c.nombre)}"></div><div class="field"><label>Grupo</label><input class="inp" id="cG" value="${esc(c.grupo)}" placeholder='Ej. 3 "C"'></div><div class="field"><label>Matrícula</label><input class="inp" id="cM" value="${esc(c.matricula)}"></div></div>
    <div class="mfoot"><button class="btn danger" id="cDel" style="margin-right:auto">Eliminar cadete</button><button class="btn ghost" data-close>Cancelar</button><button class="btn" id="cOk">Guardar</button></div>`;
    openModal();
    $('#cOk').onclick = () => { c.nombre = $('#cN').value.trim() || c.nombre; c.grupo = $('#cG').value.trim(); c.matricula = $('#cM').value.trim(); c.grupoManual = true; save(); closeModal(); render(true); toast('Datos guardados'); };
    $('#cDel').onclick = () => { if (!confirm('¿Eliminar al cadete con sus calificaciones, faltas y acciones?')) return; for (const t of ['insc', 'faltas', 'acciones']) for (const [k, v] of Object.entries(S[t])) if (v.cadete === key) delete S[t][k]; delete S.cadetes[key]; save(); closeModal(); location.hash = '#/listado'; };
  };
}
function formAccion(key, a) {
  const nuevo = !a; a = a || { id: uid(), cadete: key, fecha: hoyISO(), accion: '', responsable: S.ultResp || '', obs: '', estado: 'Pendiente' };
  const sug = ['Entrevista con el cadete', 'Aviso a la docente de práctica', 'Tarjeta informativa', 'Plan de recuperación de horas', 'Canalización a Sección Médica', 'Plática con el jefe de grupo', 'Revisión de calificación'];
  $('#modal').innerHTML = `<h2>${nuevo ? 'Nueva acción de seguimiento' : 'Editar acción'}</h2><div class="muted">${esc(S.cadetes[key].nombre)}</div>
  <div class="mapgrid"><div class="field" style="grid-column:1/-1"><label>Acción *</label><input class="inp" id="aA" list="aSug" value="${esc(a.accion)}"><datalist id="aSug">${sug.map(s => `<option value="${s}">`).join('')}</datalist></div>
  <div class="field"><label>Responsable</label><input class="inp" id="aR" value="${esc(a.responsable)}"></div><div class="field"><label>Fecha</label><input class="inp" id="aF" type="date" value="${esc(a.fecha)}"></div>
  <div class="field"><label>Estado</label><select class="inp" id="aE">${['Pendiente', 'En curso', 'Cerrada'].map(s => `<option ${a.estado === s ? 'selected' : ''}>${s}</option>`).join('')}</select></div>
  <div class="field" style="grid-column:1/-1"><label>Observaciones</label><textarea class="inp" id="aO">${esc(a.obs)}</textarea></div></div>
  <div class="mfoot">${nuevo ? '' : '<button class="btn danger" id="aDel" style="margin-right:auto">Eliminar</button>'}<button class="btn ghost" data-close>Cancelar</button><button class="btn" id="aOk">Guardar</button></div>`;
  openModal(); $('#aA').focus();
  $('#aOk').onclick = () => { const t = $('#aA').value.trim(); if (!t) return toast('Escribe la acción.', true); Object.assign(a, { accion: t, responsable: $('#aR').value.trim(), fecha: $('#aF').value, estado: $('#aE').value, obs: $('#aO').value.trim() }); S.ultResp = a.responsable; S.acciones[a.id] = a; save(); closeModal(); render(true); toast('Acción guardada'); };
  if (!nuevo) $('#aDel').onclick = () => { if (confirm('¿Eliminar esta acción?')) { delete S.acciones[a.id]; save(); closeModal(); render(true); } };
}
function formFalta(key) {
  const ins = Object.values(S.insc).filter(i => i.cadete === key);
  $('#modal').innerHTML = `<h2>Registrar falta</h2><div class="muted">${esc(S.cadetes[key].nombre)}</div>
  <div class="mapgrid"><div class="field"><label>Práctica</label><select class="inp" id="fP">${ins.map(i => `<option value="${i.practica}">${PMAP[i.practica].nombre}</option>`).join('')}</select></div>
  <div class="field"><label>Fecha</label><input class="inp" id="fF" type="date" value="${hoyISO()}"></div>
  <div class="field"><label>Motivo</label><select class="inp" id="fM">${MOTIVOS.map(m => `<option>${m}</option>`).join('')}</select></div>
  <div class="field" style="grid-column:1/-1"><label>Observaciones</label><input class="inp" id="fO"></div></div>
  <div class="mfoot"><button class="btn ghost" data-close>Cancelar</button><button class="btn" id="fOk">Guardar</button></div>`;
  openModal();
  $('#fOk').onclick = () => { const p = $('#fP').value, f = $('#fF').value; if (!p || !f) return toast('Falta la práctica o la fecha.', true); const iid = key + '|' + p, id = f + '|' + iid; if (S.faltas[id]) return toast('Ya hay una falta ese día en esa práctica.', true); const i = S.insc[iid]; S.faltas[id] = { id, cadete: key, practica: p, fecha: f, motivo: $('#fM').value, motivoOrig: '', sede: i?.sede || '', grupoTurno: i?.hoja || '', instructora: i?.docente || '', obs: $('#fO').value.trim(), archivo: 'captura manual' }; save(); closeModal(); render(true); toast('Falta registrada'); };
}

/* captura de calificaciones */
let CF = { practica: '', hoja: '' };
const grupoCap = i => `${i.sede || 'Sin sede'}${i.hoja && norm(i.hoja) !== norm(i.sede) ? ' · ' + i.hoja : ''}`;
function vCaptura() {
  const ins = Object.values(S.insc);
  if (!ins.length) return bienvenida();
  const prs = PRACTICAS.filter(p => ins.some(i => i.practica === p.id));
  if (!prs.some(p => p.id === CF.practica)) CF.practica = prs[0].id;
  const grupos = [...new Set(ins.filter(i => i.practica === CF.practica).map(grupoCap))].sort((a, b) => a.localeCompare(b, 'es', { numeric: true }));
  if (!grupos.includes(CF.hoja)) CF.hoja = grupos[0];
  return `<div class="card"><div class="row"><div class="field" style="min-width:220px"><label>Práctica</label><select class="inp" id="cPr">${prs.map(p => `<option value="${p.id}" ${p.id === CF.practica ? 'selected' : ''}>${p.nombre}</option>`).join('')}</select></div>
  <div class="field" style="min-width:240px"><label>Sede / lista</label><select class="inp" id="cHo">${grupos.map(g => `<option ${g === CF.hoja ? 'selected' : ''}>${esc(g)}</option>`).join('')}</select></div>
  <div class="grow"></div><button class="btn ghost" id="cAdd">+ Agregar cadete</button><button class="btn" id="cSave">Guardar cambios</button></div>
  <p class="small muted" style="margin:10px 0 0">Escribe instrumento y trabajo de investigación (0 a 10); la evaluación final se calcula sola (${S.cfg.pInst} % + ${S.cfg.pInv} %). Si escribes la final directo, queda fija. En rojo, las finales menores a ${S.cfg.umbral}.</p></div>
  <div class="tbl-wrap" style="margin-top:14px"><table class="tbl" style="min-width:640px"><thead><tr><th class="nos">#</th><th class="nos">Cadete</th><th class="nos">Grupo</th><th class="nos num">% faltas</th><th class="nos num">Instrumento</th><th class="nos num">Trabajo</th><th class="nos num">Final</th></tr></thead><tbody id="cBody"></tbody></table></div>`;
}
function bindCaptura() {
  if (!$('#cPr')) return;
  const lista = derive().rows.filter(r => r.practica === CF.practica && grupoCap(r) === CF.hoja).sort((a, b) => a.nombre.localeCompare(b.nombre));
  const pinta = tr => { const a = toNum(tr.querySelector('[data-k=inst]').value), b = toNum(tr.querySelector('[data-k=inv]').value), f = tr.querySelector('[data-k=fin]'); if (!f.dataset.manual && a != null && b != null) f.value = Math.round((a * S.cfg.pInst / 100 + b * S.cfg.pInv / 100) * 100) / 100; const fv = toNum(f.value); const low = fv != null && fv < S.cfg.umbral; f.style.color = low ? 'var(--bad)' : ''; f.style.fontWeight = low ? '700' : ''; };
  const ninp = (k, v, extra = '') => `<td class="num"><input class="inp" type="number" step="0.01" min="0" max="10" data-k="${k}" value="${v ?? ''}" ${extra}></td>`;
  $('#cBody').innerHTML = lista.map((r, n) => `<tr data-i="${esc(r.id)}"><td class="muted">${n + 1}</td><td><a href="#/cadete/${encodeURIComponent(r.cadete)}"><b>${esc(r.nombre)}</b></a></td><td>${esc(r.grupo)}</td><td class="num"><span class="badge b-${r.estado}">${fmt(r.pct, 2)} %</span></td>${ninp('inst', r.inst)}${ninp('inv', r.inv)}${ninp('fin', r.final, r.finalManual != null ? 'data-manual="1"' : '')}</tr>`).join('') || '<tr><td colspan="7" class="empty">Sin cadetes en esta lista.</td></tr>';
  $$('#cBody tr[data-i]').forEach(tr => { pinta(tr); tr.querySelectorAll('input').forEach(inp => inp.addEventListener('input', () => { if (inp.dataset.k === 'fin') inp.dataset.manual = inp.value ? '1' : ''; pinta(tr); })); });
  $('#cPr').onchange = e => { CF.practica = e.target.value; CF.hoja = ''; render(true); };
  $('#cHo').onchange = e => { CF.hoja = e.target.value; render(true); };
  $('#cSave').onclick = () => {
    let n = 0, bad = 0;
    $$('#cBody tr[data-i]').forEach(tr => {
      const i = S.insc[tr.dataset.i]; const g = k => { const v = tr.querySelector(`[data-k=${k}]`).value; return v === '' ? null : toNum(v); };
      const inst = g('inst'), inv = g('inv'), fin = g('fin'); if ([inst, inv, fin].some(v => v != null && (v < 0 || v > 10))) { bad++; return; }
      const manual = tr.querySelector('[data-k=fin]').dataset.manual === '1';
      const antes = JSON.stringify([i.inst, i.inv, i.finalManual]);
      i.inst = inst; i.inv = inv; i.finalManual = manual ? fin : null;
      if (antes !== JSON.stringify([i.inst, i.inv, i.finalManual])) n++;
    });
    save(); render(true); toast(bad ? `${bad} filas con valores fuera de 0 a 10 no se guardaron.` : `${n} cadetes actualizados`, !!bad);
  };
  $('#cAdd').onclick = () => {
    const ej = lista[0];
    $('#modal').innerHTML = `<h2>Agregar cadete a la lista</h2><div class="muted">${esc(PMAP[CF.practica].nombre)} · ${esc(CF.hoja)}</div><div class="mapgrid"><div class="field" style="grid-column:1/-1"><label>Nombre *</label><input class="inp" id="nN"></div><div class="field"><label>Grupo</label><input class="inp" id="nG"></div></div><div class="mfoot"><button class="btn ghost" data-close>Cancelar</button><button class="btn" id="nOk">Agregar</button></div>`;
    openModal();
    $('#nOk').onclick = () => { const n = $('#nN').value.trim(); if (tokens(n).length < 2) return toast('Escribe nombre y apellidos.', true); const k = asegurarCadete(n, CF.practica, 'lista'); if ($('#nG').value.trim()) { S.cadetes[k].grupo = $('#nG').value.trim(); S.cadetes[k].grupoManual = true; } const id = k + '|' + CF.practica; S.insc[id] = S.insc[id] || { id, cadete: k, practica: CF.practica, sede: ej?.sede || CF.hoja, hoja: ej?.hoja || '', docente: ej?.docente || '', periodo: ej?.periodo || '', horario: ej?.horario || '', inst: null, inv: null, finalLista: null, finalManual: null, origen: 'captura manual' }; save(); closeModal(); render(true); toast('Cadete agregado'); };
  };
}

/* importar */
function vImportar() {
  return `<div class="grid cols2"><div class="card"><h3>Importar archivos</h3>
    <label class="drop" id="drop"><input type="file" id="fileIn" multiple accept=".xlsx,.xls,.xlsm,.csv" hidden>
      <svg viewBox="0 0 24 24"><path d="M19.35 10.04A7.49 7.49 0 0 0 12 4C9.11 4 6.6 5.64 5.35 8.04A5.994 5.994 0 0 0 0 14c0 3.31 2.69 6 6 6h13c2.76 0 5-2.24 5-5 0-2.64-2.05-4.78-4.65-4.96M14 13v4h-4v-4H7l5-5 5 5z"/></svg>
      <div style="font-weight:700;font-size:1.1rem;margin-top:6px">Arrastra aquí tus archivos o toca para elegirlos</div><div class="muted small">Excel (.xlsx) o CSV · puedes elegir varios a la vez</div></label>
    <label class="row small" style="margin-top:10px;gap:6px;align-items:center"><input type="checkbox" id="revisar"> Revisar el mapeo de columnas de cada hoja antes de importar</label>
    <div class="log" id="implog" style="margin-top:14px"></div></div>
    <div class="card"><h3>¿Qué puedo importar?</h3><div class="small" style="line-height:1.6">
      <p><b>Listas de práctica clínica</b> (una hoja por sede o grupo): se leen materia, sede, periodo, horario, docente y la tabla <i>No. · Nombre · Instrumento · Trabajo de investigación · Evaluación parcial final</i>. Las filas ocultas y los renglones de promedio, mediana o firmas se ignoran.</p>
      <p><b>Registro de inasistencias</b> (una fila por falta): Fecha, Cadete, Práctica, Grupo/Turno, Hospital, Servicio u observación, Motivo, Motivo escrito en bitácora, Instructora. Los retardos no cuentan. Si vuelves a importar un registro actualizado, las faltas repetidas (igual cadete, práctica y día) se actualizan y no se duplican.</p>
      <p><b>Listas de asistencia por grupo</b> (1 "A" a 4 "D", TSU): sirven para asignar el grupo a cada cadete.</p>
      <p>Si alguna hoja trae otro formato, la app te pide elegir las columnas.</p></div>
      ${S.imports.length ? `<h3 style="margin-top:14px">Últimas importaciones</h3><div class="small">${S.imports.slice(0, 10).map(i => `<div class="row" style="justify-content:space-between;border-bottom:1px solid var(--line);padding:5px 0"><span>${esc(i.archivo)}</span><span class="muted">${new Date(i.fecha).toLocaleString('es-MX', { dateStyle: 'short', timeStyle: 'short' })}</span></div>`).join('')}</div>` : ''}
    </div></div>`;
}
function bindImportar() {
  const drop = $('#drop'), inp = $('#fileIn');
  const go = files => { if (!files.length) return; if (!window.XLSX) return toast('No se cargó el módulo de Excel; revisa tu conexión.', true); importarArchivos([...files], $('#revisar').checked); };
  inp.onchange = () => { go(inp.files); inp.value = ''; };
  ['dragenter', 'dragover'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('over'); }));
  ['dragleave', 'drop'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove('over'); }));
  drop.addEventListener('drop', e => go(e.dataTransfer.files));
}

/* ajustes */
function vAjustes() {
  const c = S.cfg; const n = Object.keys(S.cadetes).length;
  const fld = (id, lab, v, step = '1') => `<div class="field"><label>${lab}</label><input class="inp" type="number" step="${step}" id="${id}" value="${v}"></div>`;
  return `<div class="grid cols2"><div class="card"><h3>Reglas de cálculo</h3>
    <div class="mapgrid">${fld('sU', 'Calificación mínima (riesgo académico si es menor)', c.umbral, '0.1')}${fld('sAd', 'Advertencia desde (% faltas)', c.adv, '0.1')}${fld('sAl', 'Alerta desde (% faltas)', c.alerta, '0.1')}${fld('sHf', 'Horas por falta', c.horasFalta, '0.5')}${fld('sPi', 'Peso del instrumento (%)', c.pInst)}${fld('sPv', 'Peso del trabajo de investigación (%)', c.pInv)}</div>
    <h3 style="margin-top:16px">Horas por práctica</h3><div class="mapgrid">${PRACTICAS.map(p => `<div class="field"><label>${p.nombre}</label><input class="inp" type="number" data-hr="${p.id}" value="${c.horas[p.id]}"></div>`).join('')}</div>
    <div class="mfoot"><button class="btn ghost" id="sDef">Valores originales</button><button class="btn" id="sOk">Guardar reglas</button></div>
    <p class="small muted">Motivos de falta: ${MOTIVOS.join(', ')}. Cita programada incluye consulta, pelotón de sanidad y estudios de laboratorio; Servicio incluye cuartelera y guardia; Juzgado incluye audiencia y a declarar. Los retardos no cuentan como falta.</p></div>
  <div class="grid" style="align-content:start">
    <div class="card"><h3>Reportes en Excel</h3><p class="small muted">Letra Noto Sans 10. El de seguimiento trae a todos los cadetes; el de riesgo, solo advertencia, alerta y calificación menor a ${c.umbral}.</p><div class="row"><button class="btn" data-act="export-seg">Reporte de seguimiento</button><button class="btn gold" data-act="export-riesgo">Reporte de riesgo</button></div></div>
    <div class="card"><h3>Respaldo</h3><p class="small muted">Tus datos (${n} cadetes, ${Object.keys(S.faltas).length} faltas, ${Object.keys(S.acciones).length} acciones) viven solo en este navegador. Descarga un respaldo para guardarlo o pasarlo a otro equipo.</p>
      <div class="row"><button class="btn" id="bExp">Descargar respaldo (.json)</button><label class="btn ghost">Restaurar respaldo<input type="file" id="bImp" accept=".json,application/json" hidden></label></div></div>
    <div class="card" style="border:1px solid var(--bad-l)"><h3>Borrar datos</h3><p class="small muted">Quita de este dispositivo a todos los cadetes, calificaciones, faltas y acciones. No se puede deshacer: descarga antes un respaldo.</p><button class="btn danger" id="bDel">Borrar todos los datos</button></div>
  </div></div>`;
}
function bindAjustes() {
  $('#sOk').onclick = () => {
    const g = id => toNum($(id).value);
    const vals = { umbral: g('#sU'), adv: g('#sAd'), alerta: g('#sAl'), horasFalta: g('#sHf'), pInst: g('#sPi'), pInv: g('#sPv') };
    if (Object.values(vals).some(v => v == null || v < 0)) return toast('Revisa los valores.', true);
    if (vals.adv >= vals.alerta) return toast('La advertencia va por debajo de la alerta.', true);
    if (Math.abs(vals.pInst + vals.pInv - 100) > 0.01) return toast('Los pesos suman 100 %.', true);
    Object.assign(S.cfg, vals); $$('[data-hr]').forEach(i => { const v = toNum(i.value); if (v > 0) S.cfg.horas[i.dataset.hr] = v; });
    save(); toast('Reglas guardadas'); render(true);
  };
  $('#sDef').onclick = () => { S.cfg = JSON.parse(JSON.stringify(DEF_CFG)); save(); render(true); toast('Se restauraron los valores originales'); };
  $('#bExp').onclick = () => { const b = new Blob([JSON.stringify(S)], { type: 'application/json' }); const a = document.createElement('a'); a.href = URL.createObjectURL(b); a.download = `respaldo_seguimiento_pc_${hoyISO()}.json`; document.body.appendChild(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 2000); };
  $('#bImp').onchange = async e => { const f = e.target.files[0]; if (!f) return; try { const s = JSON.parse(await f.text()); if (!s || !s.v || !s.cadetes) throw new Error('formato'); if (!confirm('Esto reemplaza los datos actuales de este dispositivo. ¿Continuar?')) return; s.cfg = Object.assign(JSON.parse(JSON.stringify(DEF_CFG)), s.cfg || {}); S = s; save(); toast('Respaldo restaurado'); render(true); } catch (err) { toast('El archivo no es un respaldo válido.', true); } };
  $('#bDel').onclick = () => { const t = prompt('Para borrar todo escribe BORRAR'); if (t && t.trim().toUpperCase() === 'BORRAR') { const cfg = S.cfg; S = vacio(); S.cfg = cfg; save(); toast('Datos borrados'); location.hash = '#/tablero'; render(); } };
}

/* ---------- ruteo ---------- */
const TIT = { tablero: ['Tablero', 'Faltas y calificaciones de práctica clínica'], listado: ['Cadetes', 'Busca, filtra y abre la ficha de cada cadete'], captura: ['Calificaciones', 'Captura y corrección por lista'], importar: ['Importar', 'Listas de práctica clínica, registro de faltas y listas de grupo'], ajustes: ['Ajustes y respaldo', 'Reglas, reportes y respaldo'], cadete: ['Ficha del cadete', 'Calificaciones, faltas y bitácora de seguimiento'] };
function parseHash() { const h = location.hash.replace(/^#\/?/, ''); const [path, qs] = h.split('?'); const parts = path.split('/'); return { r: parts[0] || 'tablero', arg: parts.slice(1).join('/'), params: Object.fromEntries(new URLSearchParams(qs || '')) }; }
function render(quieto) {
  const { r, arg, params } = parseHash(); const v = $('#view');
  const [t, s] = TIT[r] || TIT.tablero; $('#viewTitle').textContent = t; $('#viewSub').textContent = s;
  $$('.nav a, .bottomnav a').forEach(a => a.classList.toggle('on', a.dataset.r === r || (r === 'cadete' && a.dataset.r === 'listado')));
  if (!quieto) { v.style.animation = 'none'; void v.offsetHeight; v.style.animation = ''; }
  if (r === 'listado') { v.innerHTML = vListado(params); if ($('#lBody')) bindListado(); }
  else if (r === 'cadete') { const k = decodeURIComponent(arg); v.innerHTML = vCadete(k); if (S.cadetes[k]) bindCadete(k); }
  else if (r === 'captura') { v.innerHTML = vCaptura(); bindCaptura(); }
  else if (r === 'importar') { v.innerHTML = vImportar(); bindImportar(); }
  else if (r === 'ajustes') { v.innerHTML = vAjustes(); bindAjustes(); }
  else { v.innerHTML = vTablero(); if ($('#promTabs')) { $$('#promTabs button').forEach(b => b.onclick = () => { $$('#promTabs button').forEach(x => x.classList.toggle('on', x === b)); promBox(b.dataset.g); }); promBox('practica'); } }
  animar(v);
}
window.addEventListener('hashchange', () => { render(); window.scrollTo({ top: 0 }); });
document.addEventListener('click', e => {
  const b = e.target.closest('[data-act]'); if (b) { if (b.dataset.act === 'export-seg') exportar('seguimiento'); else if (b.dataset.act === 'export-riesgo') exportar('riesgo'); return; }
  const k = e.target.closest('[data-href]'); if (k) { location.hash = k.dataset.href; return; }
  if (e.target.closest('[data-close]')) closeModal();
});
$('#modalBg').addEventListener('click', e => { if (e.target.id === 'modalBg' && !$('#mpOk')) closeModal(); });
document.addEventListener('keydown', e => { if (e.key === 'Escape' && !$('#mpOk')) closeModal(); });
/* búsqueda rápida */
const qsI = $('#qsearch'), qr = $('#qres');
qsI.addEventListener('input', () => {
  const q = norm(qsI.value); if (q.length < 2) { qr.classList.remove('on'); return; }
  const d = derive(); const hits = Object.values(S.cadetes).filter(c => norm(c.nombre).includes(q)).slice(0, 12);
  qr.innerHTML = hits.map(c => `<a href="#/cadete/${encodeURIComponent(c.key)}"><b>${esc(c.nombre)}</b><small>${(d.byCad[c.key] || []).map(r => `${PMAP[r.practica].corto} ${fmt(r.pct, 1)} %`).join(' · ')}${c.grupo ? ' · ' + esc(c.grupo) : ''}</small></a>`).join('') || '<div class="muted small" style="padding:10px">Sin coincidencias</div>';
  qr.classList.add('on');
});
qr.addEventListener('click', () => { qr.classList.remove('on'); qsI.value = ''; });
document.addEventListener('click', e => { if (!e.target.closest('.qsearch')) qr.classList.remove('on'); });
/* para pruebas automatizadas */
window.__app = { importarArchivos, libroReporte, derive: () => derive(), estado: () => S };
render();
