# Backend: reviews, usage analytics & editable cards (Google Apps Script)

The site stores everything in a Google Sheet you own, through a single Apps Script Web App. It powers
three things, all free and unlimited:

1. **Reviews** — each card's review form logs a row (rating + comment). A honeypot + 60-second
   de-dupe drop obvious spam silently.
2. **Usage analytics** — every card view / share / QR / save-contact / review-open is counted, so the
   dashboard shows how often each card is actually used (plus a 30-day trend).
3. **Editable cards** — an admin can edit each person's content from the dashboard (name, role,
   department, office, map link, phone, email, WhatsApp, LinkedIn, avatar initials), **upload a photo**
   for each card, and **retire / reactivate** a card (hides it from the directory without ever changing
   its link). The cards render that content live. Writes are protected by a server-checked admin
   password, exchanged at login for a short-lived session token.

There are **15 fixed cards** (`employe-01` … `employe-15`), one per NFC chip. Their links never change —
the editor only changes the *content* shown on a card, never its address — so the programmed chips keep
working. When someone leaves, you reassign their card (edit name/role/photo); the link stays identical.

Until you connect the backend, everything **degrades gracefully**: the review form falls back to the
visitor's mail app, analytics simply read 0, and cards show their built-in HTML defaults.

The Sheet uses these tabs (created automatically): `Avis` (reviews), `Dashboard` (live review
summary), `Vues` (per-card counters), `Evenements` (raw event log), `Employes` (editable card content),
`Photos` (uploaded avatars, one small image per card).

---

## 1. Create the spreadsheet
Go to <https://sheets.new> and create a sheet (e.g. **Euro Campus**).

## 2. Add the script
In the sheet: **Extensions → Apps Script**, delete the sample code, and paste **all** of this:

```javascript
/* ============================================================
   EURO CAMPUS — reviews + analytics + editable cards
   ============================================================ */

// Columns are matched by HEADER NAME, never by position, so the sheet can gain
// new columns (or you can reorder them) without breaking reads/writes.
var EMP_HEADERS = ['Slug','Nom','Poste','Département','Bureau','MapURL','Téléphone',
                   'Email','WhatsApp','LinkedIn','Initiales','Mis à jour','Actif'];
var HEADER_KEY  = { 'Slug':'slug','Nom':'nom','Poste':'poste','Département':'departement',
                    'Bureau':'bureau','MapURL':'mapurl','Téléphone':'telephone','Email':'email',
                    'WhatsApp':'whatsapp','LinkedIn':'linkedin','Initiales':'initiales',
                    'Mis à jour':'maj','Actif':'actif' };
// Fields the editor is allowed to write (param name → column header).
var EDITABLE = { nom:'Nom', poste:'Poste', departement:'Département', bureau:'Bureau',
                 mapurl:'MapURL', telephone:'Téléphone', email:'Email', whatsapp:'WhatsApp',
                 linkedin:'LinkedIn', initiales:'Initiales' };

/* ---------- WRITE endpoint ---------- */
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (data.type === 'event')     return recordEvent_(ss, data);
    if (data.type === 'savephoto') return json_(savephoto_(data));

    // default: a review
    // Spam guard 1 — honeypot: real users never fill the hidden field.
    if (String(data.hp || '').trim() !== '') return json_({ ok:true });
    // Spam guard 2 — drop a near-identical resubmit within 60s (double-tap / bot loop).
    var sig = 'rv_' + shortHash_(String(data.employee||'') + '|' + String(data.rating||'') + '|' + String(data.comment||''));
    var cache = CacheService.getScriptCache();
    if (cache.get(sig)) return json_({ ok:true });
    cache.put(sig, '1', 60);

    var sheet = ss.getSheetByName('Avis') || ss.insertSheet('Avis');
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Date','Employé','Poste','Bureau','Note','Commentaire','Client','Page']);
    }
    sheet.appendRow([new Date(), data.employee||'', data.role||'', data.office||'',
                     data.rating||'', data.comment||'', data.reviewer||'', data.url||'']);
    ensureDashboard_(ss);
    return json_({ ok:true });
  } catch (err) {
    return json_({ ok:false, error:String(err) });
  }
}

function shortHash_(s) {
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, s, Utilities.Charset.UTF_8);
  return raw.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

function recordEvent_(ss, data) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (e) { return json_({ ok:false, error:'busy' }); }
  try {
    var slug = String(data.slug||'').trim();
    if (!slug) return json_({ ok:false, error:'no slug' });
    var action = String(data.action||'view');
    var now = new Date();

    var log = ss.getSheetByName('Evenements') || ss.insertSheet('Evenements');
    if (log.getLastRow() === 0) log.appendRow(['Date','Slug','Employé','Action','Page']);
    log.appendRow([now, slug, data.employee||'', action, data.url||'']);

    var vues = ss.getSheetByName('Vues') || ss.insertSheet('Vues');
    if (vues.getLastRow() === 0) {
      vues.appendRow(['Slug','Employé','Vues','Partages','QR','Contacts','ClicsAvis','Dernière vue']);
    }
    var col = { view:3, share:4, qr:5, save:6, review_open:7 }[action] || 3;
    var values = vues.getDataRange().getValues();
    var rowIndex = -1;
    for (var i=1;i<values.length;i++){ if (String(values[i][0])===slug){ rowIndex=i+1; break; } }
    if (rowIndex === -1) {
      var row = [slug, data.employee||'', 0,0,0,0,0, now];
      row[col-1] = 1;
      vues.appendRow(row);
    } else {
      var cur = Number(vues.getRange(rowIndex, col).getValue())||0;
      vues.getRange(rowIndex, col).setValue(cur+1);
      vues.getRange(rowIndex, 8).setValue(now);
      if (data.employee) vues.getRange(rowIndex, 2).setValue(data.employee);
    }
    return json_({ ok:true });
  } finally { lock.releaseLock(); }
}

/* ---------- READ endpoint (JSON / JSONP) ---------- */
function doGet(e) {
  var p = (e && e.parameter) || {};
  var action = p.action || 'stats';
  var cb = p.callback || '';
  var out;
  if      (action === 'employees') out = employees_();
  else if (action === 'employee')  out = employee_(p.slug);
  else if (action === 'photo')     out = photo_(p.slug);
  else if (action === 'photos')    out = photos_();
  else if (action === 'views')     out = views_();
  else if (action === 'dashboard') out = dashboard_(isAuthed_(p));
  else if (action === 'login')     out = login_(p);
  else if (action === 'save')      out = save_(p);
  else if (action === 'setactive') out = setactive_(p);
  else                             out = stats_(isAuthed_(p));
  var payload = JSON.stringify(out);
  var body = cb ? cb + '(' + payload + ')' : payload;
  return ContentService.createTextOutput(body)
    .setMimeType(cb ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- reviews aggregate ---------- */
// includePII=true returns each review's comment + reviewer name (customer PII).
// Anonymous callers (the public directory) get aggregates only — counts/averages
// — never the raw comments or who wrote them.
function stats_(includePII) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Avis');
  var stats = [], total = 0;
  if (sheet && sheet.getLastRow() > 1) {
    var rows = sheet.getRange(2,1,sheet.getLastRow()-1,8).getValues();
    var map = {}, order = [];
    rows.forEach(function (r) {
      var name = r[1]; if (!name) return;
      if (!map[name]) { map[name] = { employee:name, role:r[2], office:r[3], count:0, sum:0, reviews:[] }; order.push(name); }
      var note = Number(r[4])||0, m = map[name];
      m.count++; m.sum += note; total++;
      m.reviews.push({ date:formatDate_(r[0]), rating:note,
        comment:  includePII ? String(r[5]||'') : '',
        reviewer: includePII ? String(r[6]||'') : '' });
    });
    order.forEach(function (k) {
      var m = map[k];
      stats.push({ employee:m.employee, role:m.role, office:m.office, count:m.count, avg:m.count?m.sum/m.count:0, reviews:m.reviews });
    });
    stats.sort(function (a, b) { return b.avg - a.avg; });
  }
  return { ok:true, total:total, stats:stats };
}

/* ---------- per-card view counters ---------- */
function views_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Vues');
  var cards = {}, total = 0;
  if (sh && sh.getLastRow() > 1) {
    sh.getRange(2,1,sh.getLastRow()-1,8).getValues().forEach(function (r) {
      var slug = String(r[0]).trim(); if (!slug) return;
      var views = Number(r[2])||0;
      cards[slug] = { views:views, share:Number(r[3])||0, qr:Number(r[4])||0,
                      save:Number(r[5])||0, reviewOpen:Number(r[6])||0,
                      last: r[7] ? new Date(r[7]).toISOString() : '' };
      total += views;
    });
  }
  return { ok:true, total:total, cards:cards };
}

/* ---------- combined dashboard payload ---------- */
function dashboard_(includePII) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = stats_(includePII);
  var byName = {}; s.stats.forEach(function (x) { byName[x.employee] = x; });

  var emps = readEmployees_();
  var vMap = {};
  var vsh = ss.getSheetByName('Vues');
  if (vsh && vsh.getLastRow() > 1) {
    vsh.getRange(2,1,vsh.getLastRow()-1,8).getValues().forEach(function (r) {
      var slug = String(r[0]).trim(); if (!slug) return;
      vMap[slug] = { employee:String(r[1]||''), views:Number(r[2])||0, share:Number(r[3])||0,
                     qr:Number(r[4])||0, save:Number(r[5])||0, reviewOpen:Number(r[6])||0,
                     last: r[7] ? new Date(r[7]).toISOString() : '' };
    });
  }

  var rowsBySlug = {}, order = [];
  function ensure(slug, name, role, office) {
    if (!rowsBySlug[slug]) {
      rowsBySlug[slug] = { slug:slug, employee:name||'', role:role||'', office:office||'',
        views:0, actions:{share:0,qr:0,save:0,reviewOpen:0}, last:'', reviewCount:0, avg:null, reviews:[] };
      order.push(slug);
    }
    return rowsBySlug[slug];
  }
  emps.forEach(function (e) { ensure(e.slug, e.nom, e.poste, e.bureau); });
  Object.keys(vMap).forEach(function (slug) {
    var v = vMap[slug], row = ensure(slug, v.employee, '', '');
    row.views = v.views; row.actions = { share:v.share, qr:v.qr, save:v.save, reviewOpen:v.reviewOpen };
    row.last = v.last; if (!row.employee) row.employee = v.employee;
  });
  order.forEach(function (slug) {
    var row = rowsBySlug[slug], st = byName[row.employee];
    if (st) { row.reviewCount = st.count; row.avg = st.avg; row.reviews = st.reviews;
              if (!row.role) row.role = st.role; if (!row.office) row.office = st.office; }
  });
  var rows = order.map(function (slug) { return rowsBySlug[slug]; });

  var totalViews = rows.reduce(function (a, r) { return a + (r.views||0); }, 0);
  var totals = { views:totalViews, reviews:s.total, rated:s.stats.length,
    avgOverall: s.total ? (s.stats.reduce(function (a,x) { return a + x.avg*x.count; },0)/s.total) : null };

  return { ok:true, rows:rows, totals:totals, trend:trend_() };
}

function trend_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var log = ss.getSheetByName('Evenements');
  var days = 30, labels = [], counts = [], idx = {};
  var tz = Session.getScriptTimeZone();
  var today = new Date(); today.setHours(0,0,0,0);
  for (var i=days-1;i>=0;i--) {
    var d = new Date(today); d.setDate(d.getDate()-i);
    idx[Utilities.formatDate(d, tz, 'yyyy-MM-dd')] = labels.length;
    labels.push(Utilities.formatDate(d, tz, 'dd/MM')); counts.push(0);
  }
  if (log && log.getLastRow() > 1) {
    log.getRange(2,1,log.getLastRow()-1,4).getValues().forEach(function (r) {
      if (String(r[3]) !== 'view') return;
      var d = r[0]; if (Object.prototype.toString.call(d) !== '[object Date]') return;
      var key = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
      if (key in idx) counts[idx[key]]++;
    });
  }
  return { labels:labels, views:counts };
}

/* ---------- editable card content (Employes) ---------- */
// Opens the sheet and BACK-FILLS any missing column headers (e.g. the new
// "Actif") without ever moving or clearing existing data.
function empSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Employes');
  if (!sh) { sh = ss.insertSheet('Employes'); sh.appendRow(EMP_HEADERS); return sh; }
  var lastCol = sh.getLastColumn();
  var have = lastCol ? sh.getRange(1,1,1,lastCol).getValues()[0].map(function (h) { return String(h).trim(); }) : [];
  EMP_HEADERS.forEach(function (h) {
    if (have.indexOf(h) === -1) { sh.getRange(1, sh.getLastColumn()+1).setValue(h); have.push(h); }
  });
  return sh;
}
// header name → 1-based column index, read from row 1.
function headerMap_(sh) {
  var lastCol = sh.getLastColumn();
  var heads = sh.getRange(1,1,1,lastCol).getValues()[0], map = {};
  for (var i=0;i<heads.length;i++) map[String(heads[i]).trim()] = i+1;
  return map;
}
// blank cell = active (so the original rows count as active without editing).
function activeBool_(v) {
  if (v === '' || v == null || v === true) return true;
  if (v === false) return false;
  var s = String(v).trim().toLowerCase();
  return !(s === 'false' || s === 'non' || s === '0' || s === 'no');
}
function readEmployees_() {
  var sh = empSheet_(), out = [];
  if (sh.getLastRow() < 2) return out;
  var hm = headerMap_(sh), slugCol = hm['Slug'];
  var rows = sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).getValues();
  rows.forEach(function (r) {
    if (slugCol && !String(r[slugCol-1]).trim()) return;
    var o = {};
    for (var h in HEADER_KEY) {
      var col = hm[h], v = (col && r[col-1] != null) ? r[col-1] : '';
      o[HEADER_KEY[h]] = (h === 'Actif') ? activeBool_(v) : String(v);
    }
    out.push(o);
  });
  return out;
}
function employees_() { return { ok:true, employees: readEmployees_() }; }
function employee_(slug) {
  slug = String(slug||'').trim();
  var list = readEmployees_();
  for (var i=0;i<list.length;i++) if (list[i].slug === slug) return { ok:true, employee:list[i] };
  return { ok:true, employee:null };
}

function save_(p) {
  var auth = checkAuth_(p);
  if (!auth.ok) return auth;
  var slug = String(p.slug||'').trim();
  if (!slug) return { ok:false, error:'no slug' };
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (e) { return { ok:false, error:'busy' }; }
  try {
    var sh = empSheet_(), hm = headerMap_(sh), slugCol = hm['Slug'];
    var values = sh.getDataRange().getValues();
    var rowIndex = -1;
    for (var i=1;i<values.length;i++){ if (String(values[i][slugCol-1]).trim() === slug){ rowIndex=i+1; break; } }
    if (rowIndex === -1) {
      var row = []; for (var c=0;c<sh.getLastColumn();c++) row.push('');
      row[slugCol-1] = slug;
      for (var k in EDITABLE) if (p[k] != null && hm[EDITABLE[k]]) row[hm[EDITABLE[k]]-1] = p[k];
      if (hm['Mis à jour']) row[hm['Mis à jour']-1] = new Date();
      sh.appendRow(row);
    } else {
      for (var k2 in EDITABLE) if (p[k2] != null && hm[EDITABLE[k2]]) sh.getRange(rowIndex, hm[EDITABLE[k2]]).setValue(p[k2]);
      if (hm['Mis à jour']) sh.getRange(rowIndex, hm['Mis à jour']).setValue(new Date());
    }
    return { ok:true };
  } finally { lock.releaseLock(); }
}

// Soft retire / reactivate — flips the "Actif" flag. The card's page + link
// are never touched, so a retired card still works; it just drops out of the
// internal directory until you reactivate (or reassign) it.
function setactive_(p) {
  var auth = checkAuth_(p);
  if (!auth.ok) return auth;
  var slug = String(p.slug||'').trim();
  if (!slug) return { ok:false, error:'no slug' };
  var active = activeBool_(p.active);
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (e) { return { ok:false, error:'busy' }; }
  try {
    var sh = empSheet_(), hm = headerMap_(sh), slugCol = hm['Slug'], actifCol = hm['Actif'];
    var values = sh.getDataRange().getValues();
    for (var i=1;i<values.length;i++){
      if (String(values[i][slugCol-1]).trim() === slug){
        sh.getRange(i+1, actifCol).setValue(active);
        if (hm['Mis à jour']) sh.getRange(i+1, hm['Mis à jour']).setValue(new Date());
        return { ok:true };
      }
    }
    return { ok:false, error:'not found' };
  } finally { lock.releaseLock(); }
}

/* ---------- uploaded photos (Photos tab) ---------- */
// Photos are kept in their own tab (one small data-URI per slug) so the
// employee payload stays light. The dashboard shrinks each image to a small
// square JPEG before upload, so a cell stays well under the size limit.
function photoSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Photos');
  if (!sh) { sh = ss.insertSheet('Photos'); sh.appendRow(['Slug','DataURI','Mis à jour']); }
  return sh;
}
function photo_(slug) {
  slug = String(slug||'').trim();
  if (!slug) return { ok:true, photo:'' };
  var sh = photoSheet_();
  if (sh.getLastRow() < 2) return { ok:true, photo:'' };
  var values = sh.getRange(2,1,sh.getLastRow()-1,2).getValues();
  for (var i=0;i<values.length;i++) if (String(values[i][0]).trim() === slug) return { ok:true, photo:String(values[i][1]||'') };
  return { ok:true, photo:'' };
}
function photos_() {
  var sh = photoSheet_(), map = {};
  if (sh.getLastRow() >= 2) {
    sh.getRange(2,1,sh.getLastRow()-1,2).getValues().forEach(function (r) {
      var slug = String(r[0]).trim(); if (slug) map[slug] = String(r[1]||'');
    });
  }
  return { ok:true, photos:map };
}
function savephoto_(p) {
  var auth = checkAuth_(p);
  if (!auth.ok) return auth;
  var slug = String(p.slug||'').trim();
  if (!slug) return { ok:false, error:'no slug' };
  var datauri = String(p.datauri||'');
  // Strict allow-list: only a base64 JPEG (what the dashboard produces). The
  // base64 charset can't contain ", <, > or spaces, so a stored value can never
  // break out of the <img src="…"> it's later rendered into.
  if (datauri !== '' && !/^data:image\/jpeg;base64,[A-Za-z0-9+/=]+$/.test(datauri)) return { ok:false, error:'bad image' };
  if (datauri.length > 200000) return { ok:false, error:'too large' };
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (e) { return { ok:false, error:'busy' }; }
  try {
    var sh = photoSheet_();
    var values = sh.getLastRow() >= 2 ? sh.getRange(2,1,sh.getLastRow()-1,1).getValues() : [];
    var rowIndex = -1;
    for (var i=0;i<values.length;i++) if (String(values[i][0]).trim() === slug) { rowIndex = i+2; break; }
    if (datauri === '') {                       // empty = remove the photo
      if (rowIndex !== -1) sh.deleteRow(rowIndex);
      return { ok:true };
    }
    if (rowIndex === -1) sh.appendRow([slug, datauri, new Date()]);
    else { sh.getRange(rowIndex,2).setValue(datauri); sh.getRange(rowIndex,3).setValue(new Date()); }
    return { ok:true };
  } finally { lock.releaseLock(); }
}

/* ---------- admin auth (session tokens) ---------- */
// The long-lived admin hash is sent at most once (login) and exchanged for a
// short-lived session id. Writes then carry the session id, not the password,
// so the secret no longer rides in every save URL / log line.
function login_(p) {
  var cache = CacheService.getScriptCache();
  // Brute-force brake. Apps Script can't see the client IP, so this is a GLOBAL
  // counter: 5 wrong tries lock all logins for 60s. Crude, but combined with a
  // strong passphrase it makes online guessing infeasible.
  var FAILS = 'login_fails';
  if (Number(cache.get(FAILS) || 0) >= 5) return { ok:false, error:'throttled' };
  var stored = PropertiesService.getScriptProperties().getProperty('ADMIN_HASH') || '';
  if (!stored || String(p.token||'') !== stored) {
    cache.put(FAILS, String(Number(cache.get(FAILS) || 0) + 1), 60);
    return { ok:false, error:'auth' };
  }
  cache.remove(FAILS);
  var sid = Utilities.getUuid().replace(/-/g,'') + Math.random().toString(36).slice(2);
  cache.put('sess_' + sid, '1', 1800);   // 30-min TTL
  return { ok:true, session: sid };
}
// Session-only. The raw admin hash is no longer accepted as a bearer token, so
// it can't be replayed against the API — it's used solely inside login_.
function checkAuth_(p) {
  if (p.session) {
    var cache = CacheService.getScriptCache(), key = 'sess_' + String(p.session);
    if (cache.get(key)) { cache.put(key, '1', 1800); return { ok:true }; }   // sliding window
  }
  return { ok:false, error:'auth' };
}
function isAuthed_(p) { return checkAuth_(p).ok; }

/* ---------- review summary tab (unchanged) ---------- */
function ensureDashboard_(ss) {
  var dash = ss.getSheetByName('Dashboard');
  if (!dash) dash = ss.insertSheet('Dashboard', 0);
  if (dash.getRange('A1').getFormula()) return;
  dash.getRange('A1').setFormula(
    '=QUERY(Avis!B2:E, "select B, C, D, count(E), avg(E) ' +
    'where B is not null group by B, C, D order by avg(E) desc ' +
    'label B \'Employé\', C \'Poste\', D \'Bureau\', ' +
    'count(E) \'Avis reçus\', avg(E) \'Note moyenne\'")'
  );
  dash.getRange('E:E').setNumberFormat('0.0');
  dash.getRange('1:1').setFontWeight('bold');
  dash.setFrozenRows(1);
}
function setupDashboard() { ensureDashboard_(SpreadsheetApp.getActiveSpreadsheet()); }

function formatDate_(d) {
  if (Object.prototype.toString.call(d) === '[object Date]' && !isNaN(d)) {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  }
  return String(d || '');
}

/* ---------- one-time setup helpers ---------- */
// Run once to set the admin password, then DELETE the password from the code.
function setAdminPassword() {
  var password = 'change-me';   // ← put your password here, run this function, then clear it
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password, Utilities.Charset.UTF_8);
  var hex = raw.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
  PropertiesService.getScriptProperties().setProperty('ADMIN_HASH', hex);
}

// Populates the Employes tab with the real team (cards 01–13). Re-running it
// REPLACES the Employes rows (photos in the Photos tab are kept). Edit any
// field afterwards from the dashboard.
function seedEmployees() {
  var sh = empSheet_();
  if (sh.getLastRow() > 1) sh.getRange(2,1,sh.getLastRow()-1,EMP_HEADERS.length).clearContent();
  var now = new Date();
  // slug, nom, poste, departement(spécialisation), bureau, mapurl, telephone, email, whatsapp, linkedin, initiales
  var MAP = 'https://maps.app.goo.gl/svpwEEanJJWajUJe9';       // siège Tizi Ouzou
  var MAIL = 'contact@euro-campus.com';
  var LI = 'https://www.linkedin.com/company/euro-campus-dz/';
  var B = 'Tizi Ouzou — Siège';
  [
    ['employe-01','Hakim Bennini','Co-Fondateur / Co-Directeur','Direction',B,MAP,'+213550631388',MAIL,'213550631388',LI,'HB'],
    ['employe-02','Said Bezzaou','Co-Fondateur / Co-Directeur','Direction',B,MAP,'',MAIL,'',LI,'SB'],
    ['employe-03','Sarah Dehmas','Responsable Communication','Communication',B,MAP,'',MAIL,'',LI,'SD'],
    ['employe-04','Mohamed Amine Bouaiche','Conseiller · Relations internationales','Écoles privées France, Canada, USA, Italie, Espagne, Roumanie, Turquie',B,MAP,'+213560981136',MAIL,'213560981136',LI,'MB'],
    ['employe-05','Dahia Slimani','Conseillère en orientation','Écoles privées France, Canada, USA, Italie, Espagne, Roumanie, Turquie',B,MAP,'+213560980962',MAIL,'213560980962',LI,'DS'],
    ['employe-06','Sadia Ouerdi','Conseillère en orientation','Canada, USA, Roumanie, Turquie',B,MAP,'+213550233166',MAIL,'213550233166',LI,'SO'],
    ['employe-07','Yasmine Mazed','Conseillère en orientation','Canada, USA, Roumanie, Turquie',B,MAP,'+213550233710',MAIL,'213550233710',LI,'YM'],
    ['employe-08','Ania Saad','Conseillère · Chef de service','Canada, USA, Roumanie, Turquie',B,MAP,'+213550231459',MAIL,'213550231459',LI,'AS'],
    ['employe-09','Dehbia Halem','Conseillère en orientation','Écoles privées France, Italie, Espagne',B,MAP,'+213560981531',MAIL,'213560981531',LI,'DH'],
    ['employe-10','Manel Aitout','Conseillère en orientation','Écoles privées France, Canada, USA, Italie, Espagne, Roumanie, Turquie',B,MAP,'+213560981575',MAIL,'213560981575',LI,'MA'],
    ['employe-11','Melinda Belkadi','Conseillère en orientation','Écoles privées France, Italie, Espagne',B,MAP,'+213550233711',MAIL,'213550233711',LI,'MB'],
    ['employe-12','Salim Yahiaoui','Conseiller en orientation','Écoles privées France, Italie, Espagne',B,MAP,'+213550235278',MAIL,'213550235278',LI,'SY'],
    ['employe-13','Amel Sefsaf','Conseillère · Chef de service','Écoles privées France, Italie, Espagne',B,MAP,'+213550235419',MAIL,'213550235419',LI,'AS']
  ].forEach(function (d) { sh.appendRow(d.concat([now])); });
}
```

Click **Save** (💾).

## 3. One-time setup (run from the editor)
In the Apps Script editor's function dropdown:
1. Put your password in `setAdminPassword` (the `password` variable), pick it in the dropdown, click
   **Run ▶**, approve the permission prompt — then **delete the password from the code** and Save (only
   the hash is kept, in Script Properties).
2. Run **`seedEmployees`** once to create the 15 card rows. (Run **`setupDashboard`** too if you want
   the review-summary tab before the first review.)

## 4. Deploy as a Web App
1. **Deploy → New deployment** → gear → **Web app**.
2. **Execute as:** Me · **Who has access:** Anyone · **Deploy**, then **Authorize access**.
3. Copy the **Web app URL** (ends in `/exec`).

## 5. Connect the site
The same `/exec` URL is referenced in three places — set them to your URL:
- `assets/card.js` → `REVIEW_ENDPOINT`
- `index.html` → `ENDPOINT`
- `dashboard.html` → `ENDPOINT`

Commit and deploy (merge to `main`).

## Editing cards from the dashboard
Open `dashboard.html` → **✏️ Gérer les cartes** → enter the admin password → **Modifier** on any
person. Saving writes to the `Employes` tab; the profile pages and directory pick up the new content
on their next load (rendered live from the Sheet). The avatar **initials**, **name/role/department**,
**office + map link**, **phone/WhatsApp/email/LinkedIn**, and the generated **vCard** all update.

- **Photo:** in **Modifier**, click **Téléverser une photo** — the image is shrunk to a small square in
  your browser and stored (in the `Photos` tab) for that card; the profile page and directory show it
  in place of the initials. **Retirer la photo** clears it (back to initials).
- **Retire / reactivate:** **Retirer** hides a card from the internal directory (its page + NFC link
  still work); **Réactiver** brings it back. Handy while a card is between employees.

## Notes
- **Updating the script later:** **Deploy → Manage deployments → Edit → Version: New version** so the
  `/exec` URL stays the same (a brand-new deployment makes a new URL). The new `Actif` column and
  `Photos` tab are created/back-filled automatically on first use — **existing data is never cleared
  or moved** (columns are matched by header name, so don't re-run `seedEmployees`, which overwrites).
- **Auth:** the password never leaves the browser in plaintext — only its SHA-256 hash is sent, once,
  at login; the server returns a short-lived **session token** (30 min) that every write carries. The
  raw hash is **no longer accepted as an API credential** (session-only), and login is **rate-limited**
  (5 wrong tries lock all logins for ~60s) to blunt brute-force. Use a long, random admin passphrase —
  it's the one real security boundary. Keep `ADMIN_HASH` out of the repo.
- **Privacy (review PII):** raw review **comments** and **reviewer names** are returned **only** to a
  logged-in dashboard (valid session). Anonymous callers of `?action=stats` / `?action=dashboard` get
  aggregates only (counts / averages / view stats). So open the dashboard's review drill-down while
  logged in (**Gérer les cartes**) to read comments.
- **Read access:** the `assets/gate.js` passcode is a soft UI curtain only — the directory/dashboard
  HTML and the employee contact fields are public by nature (static site + NFC cards). Don't rely on it
  for confidentiality; the protections above guard *edits* and *customer PII*, which is what matters.
- **Photo uploads** are validated server-side as base64 JPEG only (the format the dashboard produces),
  so a stored value can't smuggle markup into the page.
- **Clearing a field:** because cards keep their built-in HTML when a Sheet field is blank, emptying a
  field in the editor reverts that field to its static default rather than showing nothing.
- **Social-share previews:** WhatsApp/Facebook link-preview text comes from each page's static OG meta
  tags (read by crawlers before JS), so those previews keep the original text until the HTML is
  regenerated — the visible page itself always reflects the latest Sheet content.
- **Privacy / CORS:** reviews, events and photo uploads submit fire-and-forget (`mode:no-cors`); reads,
  saves and retire use JSONP so the browser can see the result. After a photo upload the dashboard
  re-reads it to confirm it stuck. If a review submission ever fails, the form falls back to the mail
  app so nothing is lost.
- **Spam:** the review endpoint has a hidden honeypot field and drops near-duplicate resubmits within
  60 seconds; the form also ignores submissions made in under a second. No CAPTCHA needed.
```
