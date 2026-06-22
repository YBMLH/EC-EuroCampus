# Backend: reviews, usage analytics & editable cards (Google Apps Script)

The site stores everything in a Google Sheet you own, through a single Apps Script Web App. It powers
three things, all free and unlimited:

1. **Reviews** — each card's review form logs a row (rating + comment).
2. **Usage analytics** — every card view / share / QR / save-contact / review-open is counted, so the
   dashboard shows how often each card is actually used (plus a 30-day trend).
3. **Editable cards** — an admin can edit each person's content from the dashboard (name, role,
   department, office, map link, phone, email, WhatsApp, LinkedIn, avatar initials); the cards render
   that content live. Writes are protected by a server-checked admin password.

Until you connect the backend, everything **degrades gracefully**: the review form falls back to the
visitor's mail app, analytics simply read 0, and cards show their built-in HTML defaults.

The Sheet uses these tabs (created automatically): `Avis` (reviews), `Dashboard` (live review
summary), `Vues` (per-card counters), `Evenements` (raw event log), `Employes` (editable card content).

---

## 1. Create the spreadsheet
Go to <https://sheets.new> and create a sheet (e.g. **Euro Campus**).

## 2. Add the script
In the sheet: **Extensions → Apps Script**, delete the sample code, and paste **all** of this:

```javascript
/* ============================================================
   EURO CAMPUS — reviews + analytics + editable cards
   ============================================================ */

var EMP_HEADERS = ['Slug','Nom','Poste','Département','Bureau','MapURL','Téléphone',
                   'Email','WhatsApp','LinkedIn','Initiales','Mis à jour'];
var EMP_KEYS    = ['slug','nom','poste','departement','bureau','mapurl','telephone',
                   'email','whatsapp','linkedin','initiales','maj'];

/* ---------- WRITE endpoint ---------- */
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (data.type === 'event') return recordEvent_(ss, data);

    // default: a review
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
  else if (action === 'views')     out = views_();
  else if (action === 'dashboard') out = dashboard_();
  else if (action === 'save')      out = save_(p);
  else                             out = stats_();
  var payload = JSON.stringify(out);
  var body = cb ? cb + '(' + payload + ')' : payload;
  return ContentService.createTextOutput(body)
    .setMimeType(cb ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- reviews aggregate ---------- */
function stats_() {
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
      m.reviews.push({ date:formatDate_(r[0]), rating:note, comment:String(r[5]||''), reviewer:String(r[6]||'') });
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
function dashboard_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = stats_();
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
function empSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Employes');
  if (!sh) { sh = ss.insertSheet('Employes'); sh.appendRow(EMP_HEADERS); }
  return sh;
}
function readEmployees_() {
  var sh = empSheet_(), out = [];
  if (sh.getLastRow() < 2) return out;
  sh.getRange(2,1,sh.getLastRow()-1,EMP_HEADERS.length).getValues().forEach(function (r) {
    if (!String(r[0]).trim()) return;
    var o = {};
    for (var i=0;i<EMP_KEYS.length;i++) o[EMP_KEYS[i]] = (r[i]==null ? '' : String(r[i]));
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
  var stored = PropertiesService.getScriptProperties().getProperty('ADMIN_HASH') || '';
  if (!stored || String(p.token||'') !== stored) return { ok:false, error:'auth' };
  var slug = String(p.slug||'').trim();
  if (!slug) return { ok:false, error:'no slug' };
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (e) { return { ok:false, error:'busy' }; }
  try {
    var sh = empSheet_();
    var values = sh.getDataRange().getValues();
    var rowIndex = -1;
    for (var i=1;i<values.length;i++){ if (String(values[i][0]).trim() === slug){ rowIndex=i+1; break; } }
    var paramToCol = { nom:2, poste:3, departement:4, bureau:5, mapurl:6, telephone:7,
                       email:8, whatsapp:9, linkedin:10, initiales:11 };
    if (rowIndex === -1) {
      var row = []; for (var c=0;c<EMP_HEADERS.length;c++) row.push('');
      row[0] = slug;
      for (var k in paramToCol) if (p[k] != null) row[paramToCol[k]-1] = p[k];
      row[11] = new Date();
      sh.appendRow(row);
    } else {
      for (var k2 in paramToCol) if (p[k2] != null) sh.getRange(rowIndex, paramToCol[k2]).setValue(p[k2]);
      sh.getRange(rowIndex, 12).setValue(new Date());
    }
    return { ok:true };
  } finally { lock.releaseLock(); }
}

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

// Run once to populate the Employes tab with the 15 cards. Edit the rest from the dashboard afterwards.
function seedEmployees_() {
  var sh = empSheet_();
  if (sh.getLastRow() > 1) sh.getRange(2,1,sh.getLastRow()-1,EMP_HEADERS.length).clearContent();
  var now = new Date();
  // slug, nom, poste, departement, bureau, mapurl, telephone, email, whatsapp, linkedin, initiales
  [
    ['employe-01','Hakim Bennini','Directeur Général','Direction','Tizi Ouzou — Siège','https://maps.app.goo.gl/svpwEEanJJWajUJe9','+213550631388','contact@euro-campus.com','213550631388','https://www.linkedin.com/company/euro-campus-dz/','HB'],
    ['employe-02','Prénom Nom','Directrice des Opérations','','Tizi Ouzou','','','contact@euro-campus.com','','','PN'],
    ['employe-03','Prénom Nom','Conseiller en Orientation','','Tizi Ouzou','','','contact@euro-campus.com','','','PN'],
    ['employe-04','Prénom Nom','Conseillère en Orientation','','Tizi Ouzou','','','contact@euro-campus.com','','','PN'],
    ['employe-05','Prénom Nom','Responsable Admissions','','Tizi Ouzou','','','contact@euro-campus.com','','','PN'],
    ['employe-06','Prénom Nom','Responsable Béjaïa','','Béjaïa','','','contact@euro-campus.com','','','PN'],
    ['employe-07','Prénom Nom','Conseiller en Orientation','','Béjaïa','','','contact@euro-campus.com','','','PN'],
    ['employe-08','Prénom Nom','Responsable Alger','','Alger','','','contact@euro-campus.com','','','PN'],
    ['employe-09','Prénom Nom','Conseillère en Orientation','','Alger','','','contact@euro-campus.com','','','PN'],
    ['employe-10','Prénom Nom','Responsable Oran','','Oran','','','contact@euro-campus.com','','','PN'],
    ['employe-11','Prénom Nom','Conseiller en Orientation','','Oran','','','contact@euro-campus.com','','','PN'],
    ['employe-12','Prénom Nom','Responsable Constantine','','Constantine','','','contact@euro-campus.com','','','PN'],
    ['employe-13','Prénom Nom','Responsable Annaba','','Annaba','','','contact@euro-campus.com','','','PN'],
    ['employe-14','Prénom Nom','Responsable Boumerdès','','Boumerdès','','','contact@euro-campus.com','','','PN'],
    ['employe-15','Prénom Nom','Chargée Communication','','Tizi Ouzou','','','contact@euro-campus.com','','','PN']
  ].forEach(function (d) { sh.appendRow(d.concat([now])); });
}
```

Click **Save** (💾).

## 3. One-time setup (run from the editor)
In the Apps Script editor's function dropdown:
1. Put your password in `setAdminPassword` (the `password` variable), pick it in the dropdown, click
   **Run ▶**, approve the permission prompt — then **delete the password from the code** and Save (only
   the hash is kept, in Script Properties).
2. Run **`seedEmployees_`** once to create the 15 card rows. (Run **`setupDashboard`** too if you want
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

## Notes
- **Updating the script later:** **Deploy → Manage deployments → Edit → Version: New version** so the
  `/exec` URL stays the same (a brand-new deployment makes a new URL).
- **Auth:** the password never leaves the browser in plaintext — only its SHA-256 hash is sent and
  compared to `ADMIN_HASH` in Script Properties. Keep that hash out of the repo. The public passcode
  gate (`assets/gate.js`) still controls read access; `ADMIN_HASH` is what guards edits.
- **Clearing a field:** because cards keep their built-in HTML when a Sheet field is blank, emptying a
  field in the editor reverts that field to its static default rather than showing nothing.
- **Social-share previews:** WhatsApp/Facebook link-preview text comes from each page's static OG meta
  tags (read by crawlers before JS), so those previews keep the original text until the HTML is
  regenerated — the visible page itself always reflects the latest Sheet content.
- **Privacy / CORS:** reviews and events submit fire-and-forget (`mode:no-cors`); reads and saves use
  JSONP so the browser can see the result. If submissions ever fail, the review form falls back to the
  mail app so nothing is lost.
- **Spam:** the endpoint is public by nature. If you get spam, ask and I'll add a honeypot + rate check.
```
