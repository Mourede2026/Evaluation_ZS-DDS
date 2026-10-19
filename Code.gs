// === Plateforme d’évaluation Zone / DDS 2026 — serveur Google Apps Script ===
// Collez ce code dans Extensions > Apps Script du Google Sheet, puis Déployer > Application Web
// (Exécuter en tant que : moi ; Accès : tout le monde).
// Aucun mot de passe n'est écrit dans ce fichier : à la première ouverture de la plateforme,
// l'écran de connexion propose de CRÉER le mot de passe administrateur.
// Les comptes sont visibles et modifiables directement dans l'onglet « Comptes » du Google Sheet
// (identifiant, mot de passe en clair, rôle admin/acteur, nom affiché). Gardez ce Sheet privé.
const ADMIN_USER = 'admin';

const sh = n => SpreadsheetApp.getActive().getSheetByName(n) || SpreadsheetApp.getActive().insertSheet(n);
const same = (stored, p) => String(stored).trim() !== '' && (String(stored).trim() === String(p).trim() || String(stored) === hash(p));
const role = x => String(x).trim().toLowerCase();
const hash = p => Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(p)));
const out = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);

function setup() {
  const c = sh('Comptes');
  if (c.getRange(1, 2).getValue() !== 'Mot de passe') {
  c.getRange(1, 1, 1, 5).setValues([['Identifiant', 'Mot de passe', 'Rôle (admin ou acteur)', 'Nom affiché', 'Session (ne pas modifier)']]).setFontWeight('bold');
  c.getRange('A:B').setNumberFormat('@');
  c.getRange(2, 3, 500, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['admin', 'acteur'], true).setAllowInvalid(false).build());
  c.setFrozenRows(1);
  }
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
  for (let i = 1; i < v.length; i++) if (v[i][0] == d.user && v[i][4] && v[i][4] == d.token) return { user: v[i][0], role: role(v[i][2]) };
  return null;
}

function run(d) {
  setup();
  const C = sh('Comptes'), E = sh('Evaluations');
  const hasAdmin = () => C.getDataRange().getValues().slice(1).some(x => role(x[2]) == 'admin' && String(x[1]).trim() !== '');
  if (d.a == 'status') return { ok: 1, init: hasAdmin() ? 0 : 1 };
  if (d.a == 'init') {
    if (hasAdmin()) return { ok: 0, err: 'Un administrateur existe déjà' };
    if (!d.np || String(d.np).length < 6) return { ok: 0, err: 'Mot de passe : 6 caractères minimum' };
    const t = Utilities.getUuid();
    const v = C.getDataRange().getValues(); let r = 0;
    for (let i = 1; i < v.length; i++) if (v[i][0] == ADMIN_USER) { r = i + 1; break; }
    const row = [ADMIN_USER, String(d.np), 'admin', 'Administrateur', t];
    r ? C.getRange(r, 1, 1, 5).setValues([row]) : C.appendRow(row);
    return { ok: 1, token: t, role: 'admin', label: 'Administrateur', user: ADMIN_USER };
  }
  if (d.a == 'login') {
    const v = C.getDataRange().getValues();
    for (let i = 1; i < v.length; i++) if (v[i][0] == d.user && same(v[i][1], d.pass)) {
      const t = Utilities.getUuid(); C.getRange(i + 1, 5).setValue(t);
      return { ok: 1, token: t, role: role(v[i][2]), label: v[i][3] };
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
  if (d.a == 'chpass') {
    if (!d.np || String(d.np).length < 6) return { ok: 0, err: 'Mot de passe : 6 caractères minimum' };
    const v = C.getDataRange().getValues();
    for (let i = 1; i < v.length; i++) if (v[i][0] == u.user) {
      if (!same(v[i][1], d.op)) return { ok: 0, err: 'Ancien mot de passe incorrect' };
      C.getRange(i + 1, 2).setValue(String(d.np)); return { ok: 1 };
    }
    return { ok: 0, err: 'Compte introuvable' };
  }
  if (u.role != 'admin') return { ok: 0, err: 'Réservé à l’administrateur' };
  if (d.a == 'delUser') {
    if (d.nu == u.user) return { ok: 0, err: 'Vous ne pouvez pas supprimer votre propre compte' };
    const v = C.getDataRange().getValues();
    for (let i = 1; i < v.length; i++) if (v[i][0] == d.nu) { C.deleteRow(i + 1); return { ok: 1 }; }
    return { ok: 0, err: 'Compte introuvable' };
  }
  if (d.a == 'users') return { ok: 1, u: C.getDataRange().getValues().slice(1).map(x => [x[0], role(x[2]), x[3]]) };
  if (d.a == 'addUser') {
    if (!d.nu || !d.np) return { ok: 0, err: 'Identifiant et mot de passe requis' };
    if (String(d.np).length < 4) return { ok: 0, err: 'Mot de passe : 4 caractères minimum' };
    const v = C.getDataRange().getValues();
    for (let i = 1; i < v.length; i++) if (v[i][0] == d.nu) { C.getRange(i + 1, 2, 1, 4).setValues([[String(d.np), d.nr || 'acteur', d.nl || d.nu, '']]); return { ok: 1 }; }
    C.appendRow([d.nu, String(d.np), d.nr || 'acteur', d.nl || d.nu, '']);
    return { ok: 1 };
  }
  return { ok: 0, err: 'Action inconnue' };
}
