// ============================================================
//  FP&A Planner — Google Apps Script Backend
//  Pega este código en Extensiones → Apps Script de tu Google Sheet
//  Luego: Implementar → Nueva implementación → Aplicación web
//  Ejecutar como: Yo | Acceso: Cualquiera
// ============================================================

const SS = SpreadsheetApp.getActiveSpreadsheet();

function getSheet(name) {
  let sh = SS.getSheetByName(name);
  if (!sh) sh = SS.insertSheet(name);
  return sh;
}

// ── Cabeceras esperadas por hoja ──────────────────────────────
const HEADERS = {
  tasks:      ['task_id','name','team','phase','assign','assignName','dstart','dend','horas'],
  people:     ['key','name','team'],
  deps:       ['task_id','dep_id'],
  statuses:   ['task_id','status','month'],
  timestamps: ['task_id','ts_inicio','ts_fin','month'],
};

function ensureHeaders() {
  Object.entries(HEADERS).forEach(([name, cols]) => {
    const sh = getSheet(name);
    if (sh.getLastRow() === 0) {
      sh.appendRow(cols);
      sh.getRange(1,1,1,cols.length).setFontWeight('bold');
    }
  });
}

// ── CORS helper ───────────────────────────────────────────────
function ok(data) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, data: data }))
    .setMimeType(ContentService.MimeType.JSON);
}

function err(msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: false, error: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Entry points ──────────────────────────────────────────────
function doGet(e)  { return route(e); }
function doPost(e) { return route(e); }

function route(e) {
  try {
    ensureHeaders();
    const action  = (e.parameter && e.parameter.action) || '';
    let   payload = {};
    if (e.postData && e.postData.contents) {
      try { payload = JSON.parse(e.postData.contents); } catch(_) {}
    }

    switch (action) {
      case 'getAll':       return actionGetAll();
      case 'addTask':      return actionAddTask(payload);
      case 'updateTask':   return actionUpdateTask(payload);
      case 'deleteTask':   return actionDeleteTask(payload);
      case 'setStatus':    return actionSetStatus(payload);
      case 'setDeps':      return actionSetDeps(payload);
      case 'upsertPerson': return actionUpsertPerson(payload);
      case 'deletePerson': return actionDeletePerson(payload);
      case 'swapAssign':   return actionSwapAssign(payload);
      case 'seed':         return actionSeed(payload);
      default:             return err('Acción desconocida: ' + action);
    }
  } catch(ex) {
    return err(ex.message);
  }
}

// ── Helpers de lectura ────────────────────────────────────────
function sheetToObjects(name) {
  const sh = getSheet(name);
  if (sh.getLastRow() < 2) return [];
  const [headers, ...rows] = sh.getDataRange().getValues();
  return rows.map(r => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = r[i] === '' ? null : r[i]);
    return obj;
  });
}

// ── getAll ────────────────────────────────────────────────────
function actionGetAll() {
  const tasks      = sheetToObjects('tasks');
  const people     = sheetToObjects('people');   // array [{key,name,team}]
  const deps       = sheetToObjects('deps');      // array [{task_id,dep_id}]
  const statuses   = sheetToObjects('statuses');
  const timestamps = sheetToObjects('timestamps');

  // Normalizar IDs a Number para comparaciones en el frontend
  tasks.forEach(t => { t.id = Number(t.task_id || t.id); t.task_id = t.id; });
  deps.forEach(d => { d.task_id = Number(d.task_id); d.dep_id = Number(d.dep_id); });
  statuses.forEach(s => { s.task_id = Number(s.task_id); });
  timestamps.forEach(t => { t.task_id = Number(t.task_id); });

  return ok({ tasks, people, deps, statuses, timestamps });
}

// ── addTask ───────────────────────────────────────────────────
function actionAddTask(p) {
  const sh = getSheet('tasks');
  const newId = nextId('tasks', 'task_id');
  sh.appendRow([
    newId,
    p.name || '',
    p.team || '',
    p.phase || 'General',
    p.assign || '',
    p.assignName || '',
    Number(p.dstart) || 1,
    Number(p.dend)   || 1,
    p.horas !== undefined ? p.horas : '',
  ]);
  return ok({ id: newId });
}

function nextId(sheetName, col) {
  const objs = sheetToObjects(sheetName);
  if (!objs.length) return 1;
  return Math.max(...objs.map(o => Number(o[col]) || 0)) + 1;
}

// ── updateTask ────────────────────────────────────────────────
function actionUpdateTask(p) {
  const sh  = getSheet('tasks');
  const row = findRow(sh, 'task_id', p.id);
  if (!row) return err('Tarea no encontrada: ' + p.id);

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const map = { name:1, team:1, phase:1, assign:1, assignName:1, dstart:1, dend:1, horas:1 };

  Object.entries(p).forEach(([key, val]) => {
    if (!map[key]) return;
    const col = headers.indexOf(key) + 1;
    if (col > 0) sh.getRange(row, col).setValue(val === undefined ? '' : val);
  });
  return ok({});
}

// ── deleteTask ────────────────────────────────────────────────
function actionDeleteTask(p) {
  deleteRowsWhere('tasks',      'task_id', p.id);
  deleteRowsWhere('statuses',   'task_id', p.id);
  deleteRowsWhere('timestamps', 'task_id', p.id);
  deleteRowsWhere('deps',       'task_id', p.id);
  deleteRowsWhere('deps',       'dep_id',  p.id);
  return ok({});
}

// ── setStatus ─────────────────────────────────────────────────
function actionSetStatus(p) {
  const sh      = getSheet('statuses');
  const tsSheet = getSheet('timestamps');
  const month   = p.month || '';
  const tid     = p.task_id;

  // Borrar registro previo del mismo mes
  deleteRowsWhere('statuses', 'task_id', tid, r => r.month === month || !r.month);
  if (p.status) sh.appendRow([tid, p.status, month]);

  // Timestamps
  deleteRowsWhere('timestamps', 'task_id', tid, r => r.month === month || !r.month);
  if (p.ts_inicio || p.ts_fin) {
    tsSheet.appendRow([tid, p.ts_inicio || '', p.ts_fin || '', month]);
  }
  return ok({});
}

// ── setDeps ───────────────────────────────────────────────────
function actionSetDeps(p) {
  deleteRowsWhere('deps', 'task_id', p.task_id);
  const sh = getSheet('deps');
  (p.deps || []).forEach(depId => sh.appendRow([p.task_id, depId]));
  return ok({});
}

// ── upsertPerson ──────────────────────────────────────────────
function actionUpsertPerson(p) {
  const sh  = getSheet('people');
  const row = findRow(sh, 'key', p.key);
  if (row) {
    sh.getRange(row, 2).setValue(p.name);
    sh.getRange(row, 3).setValue(p.team);
  } else {
    sh.appendRow([p.key, p.name, p.team]);
  }
  return ok({});
}

// ── deletePerson ──────────────────────────────────────────────
function actionDeletePerson(p) {
  deleteRowsWhere('people', 'key', p.key);
  return ok({});
}

// ── swapAssign ────────────────────────────────────────────────
function actionSwapAssign(p) {
  const sh      = getSheet('tasks');
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const assignCol     = headers.indexOf('assign') + 1;
  const assignNameCol = headers.indexOf('assignName') + 1;
  if (sh.getLastRow() < 2) return ok({});

  // Buscar nombres actuales en la hoja de personas
  const allPeople = sheetToObjects('people');
  const fromPerson = allPeople.find(x => x.key === p.from);
  const toPerson   = allPeople.find(x => x.key === p.to);
  const fromName   = fromPerson ? fromPerson.name : p.from;
  const toName     = toPerson   ? toPerson.name   : p.to;

  const data = sh.getRange(2, 1, sh.getLastRow()-1, sh.getLastColumn()).getValues();
  data.forEach((row, i) => {
    const assign = String(row[assignCol-1]);
    if (assign === String(p.from)) {
      sh.getRange(i+2, assignCol).setValue(p.to);
      if (assignNameCol > 0) sh.getRange(i+2, assignNameCol).setValue(toName);
    } else if (assign === String(p.to)) {
      sh.getRange(i+2, assignCol).setValue(p.from);
      if (assignNameCol > 0) sh.getRange(i+2, assignNameCol).setValue(fromName);
    }
  });
  return ok({});
}

// ── seed ──────────────────────────────────────────────────────
function actionSeed(p) {
  // Solo inserta si las hojas están vacías
  const tasks = sheetToObjects('tasks');
  if (tasks.length === 0 && p.tasks) {
    const sh = getSheet('tasks');
    p.tasks.forEach(t => sh.appendRow([
      t.id, t.name, t.team, t.phase||'General',
      t.assign, t.assignName, t.dstart, t.dend, t.horas||''
    ]));
  }

  const people = sheetToObjects('people');
  if (people.length === 0 && p.people) {
    const sh = getSheet('people');
    p.people.forEach(p2 => sh.appendRow([p2.key, p2.name, p2.team]));
  }

  const deps = sheetToObjects('deps');
  if (deps.length === 0 && p.deps) {
    const sh = getSheet('deps');
    Object.entries(p.deps).forEach(([tid, depList]) => {
      (depList||[]).forEach(did => sh.appendRow([tid, did]));
    });
  }

  return ok({ seeded: true });
}

// ── Utilidades ────────────────────────────────────────────────
/** Devuelve el número de fila (1-based) donde col === val, o null */
function findRow(sh, colName, val) {
  if (sh.getLastRow() < 2) return null;
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const col = headers.indexOf(colName);
  if (col < 0) return null;
  const data = sh.getRange(2, col+1, sh.getLastRow()-1, 1).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]) === String(val)) return i + 2;
  }
  return null;
}

/** Elimina todas las filas donde colName === val (y el filtro opcional pasa) */
function deleteRowsWhere(sheetName, colName, val, filterFn) {
  const sh = getSheet(sheetName);
  if (sh.getLastRow() < 2) return;
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const col = headers.indexOf(colName);
  if (col < 0) return;

  // Iterar de abajo hacia arriba para no desplazar índices
  for (let r = sh.getLastRow(); r >= 2; r--) {
    const rowVals = sh.getRange(r, 1, 1, sh.getLastColumn()).getValues()[0];
    if (String(rowVals[col]) === String(val)) {
      if (!filterFn) {
        sh.deleteRow(r);
      } else {
        const obj = {};
        headers.forEach((h,i) => obj[h] = rowVals[i]);
        if (filterFn(obj)) sh.deleteRow(r);
      }
    }
  }
}
