// === Plateforme d’évaluation Zone / DDS 2026 — serveur Google Apps Script ===
// Collez ce code dans Extensions > Apps Script du Google Sheet, changez le mot de passe ci-dessous,
// exécutez setup() une fois, puis Déployer > Application Web (Exécuter en tant que : moi ; Accès : tout le monde).
const ADMIN_USER = 'admin';
const ADMIN_PASS = 'ChangezMoi2026';   // <- à changer avant d’exécuter setup()

const sh = n => SpreadsheetApp.getActive().getSheetByName(n) || SpreadsheetApp.getActive().insertSheet(n);
const hash = p => Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(p)));
const out = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);

function setup() {
  const c = sh('Comptes');
  if (c.getLastRow() < 1) { c.appendRow(['user', 'hash', 'role', 'label', 'token']); c.appendRow([ADMIN_USER, hash(ADMIN_PASS), 'admin', 'Administrateur', '']); }
  const e = sh('Evaluations');
  if (e.getLastRow() < 1) e.appendRow(['id', 'owner', 'niveau', 'structure', 'date', 'maj', 'json']);
}

function doPost(e) {
  const L = LockService.getScriptLock(); L.waitLock(20000);
  try { return out(run(JSON.parse(e.postData.contents))); }
  catch (err) { return out({ ok: 0, err: String(err) }); }
  finally { L.releaseLock(); }
}

function who(d) {
  const v = sh('Comptes').getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (v[i][0] == d.user && v[i][4] && v[i][4] == d.token) return { user: v[i][0], role: v[i][2] };
  return null;
}

function run(d) {
  setup();
  const C = sh('Comptes'), E = sh('Evaluations');
  if (d.a == 'login') {
    const v = C.getDataRange().getValues();
    for (let i = 1; i < v.length; i++) if (v[i][0] == d.user && v[i][1] == hash(d.pass)) {
      const t = Utilities.getUuid(); C.getRange(i + 1, 5).setValue(t);
      return { ok: 1, token: t, role: v[i][2], label: v[i][3] };
    }
    return { ok: 0, err: 'Identifiant ou mot de passe incorrect' };
  }
  const u = who(d);
  if (!u) return { ok: 0, err: 'Session expirée', auth: 1 };
  if (d.a == 'push') {
    const ev = d.ev, v = E.getDataRange().getValues(); let r = 0;
    for (let i = 1; i < v.length; i++) if (v[i][0] == ev.id) { r = i + 1; break; }
    if (r && u.role != 'admin' && v[r - 1][1] != u.user) return { ok: 0, err: 'Interdit' };
    const own = r ? v[r - 1][1] : (u.role == 'admin' && ev.o ? ev.o : u.user); ev.o = own;
    const row = [ev.id, own, ev.n, ev.n == 'ZS' ? ev.zone : ev.dep, ev.date, ev.u, JSON.stringify(ev)];
    r ? E.getRange(r, 1, 1, 7).setValues([row]) : E.appendRow(row);
    return { ok: 1 };
  }
  if (d.a == 'pull') {
    const v = E.getDataRange().getValues().slice(1).filter(x => u.role == 'admin' || x[1] == u.user);
    return { ok: 1, ev: v.map(x => JSON.parse(x[6])) };
  }
  if (u.role != 'admin') return { ok: 0, err: 'Réservé à l’administrateur' };
  if (d.a == 'users') return { ok: 1, u: C.getDataRange().getValues().slice(1).map(x => [x[0], x[2], x[3]]) };
  if (d.a == 'addUser') {
    if (!d.nu || !d.np) return { ok: 0, err: 'Identifiant et mot de passe requis' };
    const v = C.getDataRange().getValues();
    for (let i = 1; i < v.length; i++) if (v[i][0] == d.nu) { C.getRange(i + 1, 2, 1, 4).setValues([[hash(d.np), d.nr || 'acteur', d.nl || d.nu, '']]); return { ok: 1 }; }
    C.appendRow([d.nu, hash(d.np), d.nr || 'acteur', d.nl || d.nu, '']);
    return { ok: 1 };
  }
  return { ok: 0, err: 'Action inconnue' };
}
