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
  if (c.getRange(1, 3).getValue() !== 'Rôle (admin, acteur, intrants ou formation)') {
  c.getRange(1, 1, 1, 5).setValues([['Identifiant', 'Mot de passe', 'Rôle (admin, acteur, intrants ou formation)', 'Nom affiché', 'Session (ne pas modifier)']]).setFontWeight('bold');
  c.getRange('A:B').setNumberFormat('@');
  c.getRange(2, 3, 500, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['admin', 'acteur', 'intrants', 'formation'], true).setAllowInvalid(false).build());
  c.setFrozenRows(1);
  }
  if (c.getRange(1, 6).getValue() === '') c.getRange(1, 6).setValue('Structure rattachée (zone ou DDS, pour les comptes intrants)').setFontWeight('bold');
  const sts = sh('Structures');
  if (sts.getLastRow() < 1) { sts.getRange(1, 1, 1, 2).setValues([['Niveau (ZS ou DDS)', 'Nom de la structure']]).setFontWeight('bold'); sts.getRange('B:B').setNumberFormat('@'); sts.getRange(2, 1, 500, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['ZS', 'DDS'], true).setAllowInvalid(false).build()); sts.setFrozenRows(1); }
  if (c.getRange(2, 6).getDataValidation() == null) { c.getRange(2, 6, 500, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(sts.getRange('B2:B1000'), true).setAllowInvalid(true).build()); c.getRange(1, 6).setBackground('#fff2a8'); }
  const fm_ = sh('Formations');
  if (fm_.getLastRow() < 1) { fm_.getRange(1, 1, 1, 4).setValues([['Formation sanitaire ou hôpital', 'Type (CS, PRIVE ou HOP)', 'Zone sanitaire rattachée', 'Identifiant du compte']]).setFontWeight('bold'); fm_.setFrozenRows(1); }
  const fs_ = sh('Fiches');
  if (fs_.getLastRow() < 1) fs_.appendRow(['Zone sanitaire', 'Type (CS, PRIVE, HOP)', 'Formation sanitaire', 'Date', 'Envoyé par', 'Horodatage', 'Données (ne pas modifier)', 'Statut (envoyée ou brouillon)']);
  if (fs_.getRange(1, 8).getValue() === '') fs_.getRange(1, 8).setValue('Statut (envoyée ou brouillon)');
  const it = sh('Intrants');
  if (it.getLastRow() < 1) it.appendRow(['Structure (zone ou DDS)', 'Date d’envoi', 'Envoyé par', 'Horodatage', 'Données (ne pas modifier)', 'Niveau']);
  if (it.getRange(1, 6).getValue() === '') it.getRange(1, 6).setValue('Niveau');
  const e = sh('Evaluations');
  if (e.getLastRow() < 1) e.appendRow(['id', 'owner', 'niveau', 'structure', 'date', 'maj', 'json']);
}

const norm = s => String(s || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();

function formReg() {
  return sh('Formations').getDataRange().getValues().slice(1).filter(x => String(x[0]).trim() && String(x[3]).trim()).map(x => ({ name: String(x[0]).trim(), t: String(x[1]).trim().toUpperCase(), zone: String(x[2]).trim(), user: String(x[3]).trim() }));
}

function stList() {
  const L = []; sh('Structures').getDataRange().getValues().slice(1).forEach(x => { const nm = String(x[1]).trim(); if (nm) L.push({ n: String(x[0]).trim().toUpperCase() == 'DDS' ? 'DDS' : 'ZS', name: nm }); });
  return L;
}

function doGet() { return out({ ok: 1, msg: 'Serveur actif : le déploiement fonctionne.' }); }

function doPost(e) {
  const L = LockService.getScriptLock(); L.waitLock(20000);
  try { return out(run(JSON.parse(e.postData.contents))); }
  catch (err) { return out({ ok: 0, err: String(err) }); }
  finally { L.releaseLock(); }
}

function who(d) {
  const v = sh('Comptes').getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (v[i][0] == d.user && v[i][4] && v[i][4] == d.token) return { user: v[i][0], role: role(v[i][2]), struct: String(v[i][5] || '').trim() };
  return null;
}

function run(d) {
  setup();
  const C = sh('Comptes'), E = sh('Evaluations');
  const hasAdmin = () => C.getDataRange().getValues().slice(1).some(x => role(x[2]) == 'admin' && String(x[1]).trim() !== '');
  if (d.a == 'status') return { ok: 1, init: hasAdmin() ? 0 : 1, n: C.getDataRange().getValues().slice(1).filter(x => String(x[0]).trim() !== '').length };
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
      return { ok: 1, token: t, role: role(v[i][2]), label: v[i][3], struct: String(v[i][5] || '').trim() };
    }
    const known = v.slice(1).some(x => String(x[0]).trim() == String(d.user).trim());
    return { ok: 0, err: known ? 'Mot de passe incorrect pour cet identifiant' : 'Identifiant inconnu (' + (v.length - 1) + ' compte(s) lu(s) dans le Sheet)' };
  }
  const u = who(d);
  if (!u) return { ok: 0, err: 'Session expirée', auth: 1 };
  if (u.role == 'formation' && ['ficheSubmit', 'ficheList', 'ficheDel', 'chpass'].indexOf(d.a) < 0) return { ok: 0, err: 'Ce compte sert uniquement à la saisie des fiches' };
  if (u.role == 'intrants' && ['zones', 'intrSubmit', 'chpass'].indexOf(d.a) < 0) return { ok: 0, err: 'Ce compte sert uniquement au formulaire intrants' };
  if (d.a == 'ficheSubmit' || d.a == 'ficheList' || d.a == 'ficheDel') {
    let reg = null, zone = u.role != 'admin' && u.struct ? u.struct : String(d.zone || '').trim();
    if (u.role == 'formation') {
      reg = formReg().filter(x => x.user == u.user)[0];
      if (!reg) return { ok: 0, err: 'Formation sanitaire non enregistrée : demandez à l’administrateur de l’ajouter' };
      zone = reg.zone; d.t = reg.t; d.name = reg.name;
    }
    const zl = stList().filter(x => x.n == 'ZS' && norm(x.name) == norm(zone))[0];
    if (!zl) return { ok: 0, err: 'Zone sanitaire inconnue : ce compte doit être rattaché à une zone sanitaire de la liste' };
    zone = zl.name; const F = sh('Fiches'), v = F.getDataRange().getValues();
    if (d.a == 'ficheList') return { ok: 1, zone: zone, reg: reg, fiches: v.slice(1).filter(x => norm(x[0]) == norm(zone) && (!reg || (x[1] == reg.t && norm(x[2]) == norm(reg.name)))).map(x => { let o = {}; try { o = JSON.parse(x[6]); } catch (e) {} return { t: x[1], name: x[2], date: x[3], by: x[4], ts: Number(x[5]), v: o, st: String(x[7] || 'envoyée') }; }) };
    if (['CS', 'PRIVE', 'HOP'].indexOf(d.t) < 0 || !String(d.name || '').trim()) return { ok: 0, err: 'Type et nom de la formation sanitaire requis' };
    let row = 0, rowD = 0; for (let i = 1; i < v.length; i++) if (norm(v[i][0]) == norm(zone) && v[i][1] == d.t && norm(v[i][2]) == norm(d.name)) { if (String(v[i][7]) == 'brouillon') rowD = i + 1; else row = i + 1; }
    if (d.a == 'ficheDel') { const r = d.draft ? rowD : row; if (!r) return { ok: 0, err: 'Fiche introuvable' }; F.deleteRow(r); return { ok: 1 }; }
    const js = JSON.stringify(d.v || {}); if (js.length > 45000) return { ok: 0, err: 'Données trop volumineuses' };
    const rw = [zone, d.t, String(d.name).trim().slice(0, 120), String(d.date || '').slice(0, 20), u.user, Date.now(), js, d.draft ? 'brouillon' : 'envoyée'], tr = d.draft ? rowD : row;
    tr ? F.getRange(tr, 1, 1, 8).setValues([rw]) : F.appendRow(rw);
    if (!d.draft && rowD) F.deleteRow(rowD);
    return { ok: 1 };
  }
  if (d.a == 'zones') {
    let L = stList();
    if (!L.length) { const z = {}; E.getDataRange().getValues().slice(1).forEach(x => { const n = x[2] == 'DDS' ? 'DDS' : 'ZS', nm = String(x[3]).trim(); if (nm) z[n + '|' + norm(nm)] = { n: n, name: nm }; }); L = Object.keys(z).map(k => z[k]); }
    return { ok: 1, zones: L.sort((a, b) => a.name.localeCompare(b.name)) };
  }
  if (d.a == 'intrSubmit') {
    const crow = C.getDataRange().getValues().filter(x => x[0] == u.user)[0], st = crow ? String(crow[5] || '').trim() : '';
    if (st) { d.zone = st; d.n = '*'; }
    if (!d.zone || !String(d.zone).trim() || !d.items || typeof d.items != 'object') return { ok: 0, err: 'Zone sanitaire et données requises' };
    const js = JSON.stringify(d.items); if (js.length > 45000) return { ok: 0, err: 'Données trop volumineuses' };
    sh('Intrants').appendRow([String(d.zone).trim().slice(0, 120), new Date(), (u.user + (d.by ? ' · ' + String(d.by) : '')).slice(0, 120), Date.now(), js, d.n == '*' ? '*' : d.n == 'DDS' ? 'DDS' : 'ZS']);
    return { ok: 1 };
  }
  if (d.a == 'push') {
    const ev = d.ev, v = E.getDataRange().getValues(); let r = 0;
    for (let i = 1; i < v.length; i++) if (v[i][0] == ev.id) { r = i + 1; break; }
    if (r && u.role != 'admin' && v[r - 1][1] != u.user && !(u.struct && norm(v[r - 1][3]) == norm(u.struct))) return { ok: 0, err: 'Interdit' };
    const sn = ev.n == 'ZS' ? ev.zone : ev.dep;
    if (String(sn || '').trim()) {
      if (u.role != 'admin' && u.struct && norm(sn) != norm(u.struct)) return { ok: 0, err: 'Ce compte est rattaché à la structure « ' + u.struct + ' »' };
      const L = stList(); if (L.length && !L.some(x => x.n == ev.n && norm(x.name) == norm(sn))) return { ok: 0, err: 'Structure absente de la liste : ' + sn };
    }
    const own = r ? v[r - 1][1] : (u.role == 'admin' && ev.o ? ev.o : u.user); ev.o = own;
    const row = [ev.id, own, ev.n, ev.n == 'ZS' ? ev.zone : ev.dep, ev.date, ev.u, JSON.stringify(ev)];
    r ? E.getRange(r, 1, 1, 7).setValues([row]) : E.appendRow(row);
    return { ok: 1 };
  }
  if (d.a == 'pull') {
    const v = E.getDataRange().getValues().slice(1).filter(x => u.role == 'admin' || x[1] == u.user || (u.struct && norm(x[3]) == norm(u.struct)));
    const mine = {}; v.forEach(x => { mine[(x[2] == 'DDS' ? 'DDS' : 'ZS') + '|' + norm(x[3])] = 1; });
    const last = {}; sh('Intrants').getDataRange().getValues().slice(1).forEach(x => {
      const n = x[5] == '*' ? '*' : x[5] == 'DDS' ? 'DDS' : 'ZS', nm = norm(x[0]), p = String(x[2]).split(' · ');
      if (!nm || (u.role != 'admin' && !(n == '*' ? (mine['ZS|' + nm] || mine['DDS|' + nm]) : mine[n + '|' + nm]))) return;
      const k = n + '|' + nm;
      if (!last[k] || Number(x[3]) > last[k].ts) { try { last[k] = { n: n, z: x[0], ts: Number(x[3]), by: p.join(' · '), items: JSON.parse(x[4]) }; } catch (e) {} }
    });
    const fi = []; sh('Fiches').getDataRange().getValues().slice(1).forEach(x => { const nm = norm(x[0]); if (!nm || String(x[7]) == 'brouillon' || (u.role != 'admin' && !mine['ZS|' + nm] && !(u.struct && norm(u.struct) == nm))) return; try { fi.push({ z: x[0], t: x[1], name: x[2], ts: Number(x[5]), v: JSON.parse(x[6]) }); } catch (e) {} });
    return { ok: 1, ev: v.map(x => JSON.parse(x[6])), intr: Object.keys(last).map(k => last[k]), st: stList(), fiches: fi };
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
  if (d.a == 'addStruct') {
    const n = d.n == 'DDS' ? 'DDS' : 'ZS', have = {}; stList().forEach(x => { have[x.n + '|' + norm(x.name)] = 1; });
    const add = []; (d.names || []).forEach(nm => { nm = String(nm).trim(); if (nm && !have[n + '|' + norm(nm)]) { have[n + '|' + norm(nm)] = 1; add.push([n, nm]); } });
    if (add.length) { const S_ = sh('Structures'); S_.getRange(S_.getLastRow() + 1, 1, add.length, 2).setValues(add); }
    return { ok: 1, added: add.length };
  }
  if (d.a == 'formList') {
    const pw = {}; C.getDataRange().getValues().slice(1).forEach(x => { pw[x[0]] = String(x[1]); });
    return { ok: 1, list: formReg().map(x => ({ name: x.name, t: x.t, zone: x.zone, user: x.user, pass: pw[x.user] || '' })) };
  }
  if (d.a == 'addForm') {
    const t = ['CS', 'PRIVE', 'HOP'].indexOf(d.t) >= 0 ? d.t : '', zl = stList().filter(x => x.n == 'ZS' && norm(x.name) == norm(d.zone))[0];
    if (!t || !zl) return { ok: 0, err: 'Choisissez le type et une zone sanitaire de la liste des structures' };
    const have = {}; formReg().forEach(x => { have[norm(x.name) + '|' + norm(x.zone) + '|' + x.t] = 1; });
    const users = {}; C.getDataRange().getValues().slice(1).forEach(x => { users[String(x[0]).toLowerCase()] = 1; });
    const reg = [], acc = [], out = [];
    (d.names || []).forEach(nm => {
      nm = String(nm).trim(); if (!nm || have[norm(nm) + '|' + norm(zl.name) + '|' + t]) return; have[norm(nm) + '|' + norm(zl.name) + '|' + t] = 1;
      const base = 'fs-' + norm(zl.name).replace(/ /g, '').slice(0, 8) + '-' + norm(nm).replace(/ /g, '').slice(0, 10); let user = base, k = 2;
      while (users[user]) user = base + (k++); users[user] = 1;
      const pw = Utilities.getUuid().replace(/-/g, '').slice(0, 6);
      reg.push([nm, t, zl.name, user]); acc.push([user, pw, 'formation', nm, '', nm]); out.push({ name: nm, user: user, pass: pw });
    });
    if (reg.length) { const F_ = sh('Formations'); F_.getRange(F_.getLastRow() + 1, 1, reg.length, 4).setValues(reg); C.getRange(C.getLastRow() + 1, 1, acc.length, 6).setValues(acc); }
    return { ok: 1, added: out.length };
  }
  if (d.a == 'delForm') {
    const F_ = sh('Formations'), fv = F_.getDataRange().getValues();
    for (let i = 1; i < fv.length; i++) if (fv[i][3] == d.user) { F_.deleteRow(i + 1); break; }
    const cv = C.getDataRange().getValues();
    for (let i = 1; i < cv.length; i++) if (cv[i][0] == d.user && role(cv[i][2]) == 'formation') { C.deleteRow(i + 1); break; }
    return { ok: 1 };
  }
  if (d.a == 'delStruct') {
    const v = sh('Structures').getDataRange().getValues();
    for (let i = 1; i < v.length; i++) if (norm(v[i][1]) == norm(d.name) && (String(v[i][0]).trim().toUpperCase() == 'DDS' ? 'DDS' : 'ZS') == d.n) { sh('Structures').deleteRow(i + 1); return { ok: 1 }; }
    return { ok: 0, err: 'Structure introuvable' };
  }
  if (d.a == 'delUser') {
    if (d.nu == u.user) return { ok: 0, err: 'Vous ne pouvez pas supprimer votre propre compte' };
    const v = C.getDataRange().getValues();
    for (let i = 1; i < v.length; i++) if (v[i][0] == d.nu) { C.deleteRow(i + 1); return { ok: 1 }; }
    return { ok: 0, err: 'Compte introuvable' };
  }
  if (d.a == 'users') return { ok: 1, u: C.getDataRange().getValues().slice(1).map(x => [x[0], role(x[2]), x[3], x[5] || '']) };
  if (d.a == 'addUser') {
    if (!d.nu || !d.np) return { ok: 0, err: 'Identifiant et mot de passe requis' };
    if (String(d.np).length < 4) return { ok: 0, err: 'Mot de passe : 4 caractères minimum' };
    if (String(d.ns || '').trim()) { const f = stList().filter(x => norm(x.name) == norm(d.ns))[0]; if (!f) return { ok: 0, err: 'Structure inconnue : ajoutez-la d’abord à la liste des structures' }; d.ns = f.name; }
    const v = C.getDataRange().getValues();
    for (let i = 1; i < v.length; i++) if (v[i][0] == d.nu) { C.getRange(i + 1, 2, 1, 4).setValues([[String(d.np), d.nr || 'acteur', d.nl || d.nu, '']]); C.getRange(i + 1, 6).setValue(String(d.ns || '').trim()); return { ok: 1 }; }
    C.appendRow([d.nu, String(d.np), d.nr || 'acteur', d.nl || d.nu, '', String(d.ns || '').trim()]);
    return { ok: 1 };
  }
  return { ok: 0, err: 'Action inconnue' };
}
