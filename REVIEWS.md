# Collecting reviews in a Google Sheet

The review form on each card can log every submission as a row in a Google
Sheet you own — free, unlimited, and easy to track (sort, filter, average
rating per employee). Until you connect it, the form falls back to opening the
visitor's mail app with the review pre-filled, so nothing breaks in the meantime.

Setup takes about 5 minutes.

## 1. Create the spreadsheet
1. Go to <https://sheets.new> and create a sheet (name it e.g. **Avis Euro Campus**).

## 2. Add the script
1. In the sheet: **Extensions → Apps Script**.
2. Delete any sample code and paste this:

```javascript
function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Avis') || ss.insertSheet('Avis');
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Date', 'Employé', 'Poste', 'Bureau', 'Note', 'Commentaire', 'Client', 'Page']);
    }
    sheet.appendRow([
      new Date(),
      data.employee || '',
      data.role || '',
      data.office || '',
      data.rating || '',
      data.comment || '',
      data.reviewer || '',
      data.url || ''
    ]);
    ensureDashboard_(ss);
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Creates/refreshes a "Dashboard" tab: one row per employee with their
// review count and average rating, sorted best-first. It uses a live
// QUERY formula, so it updates automatically as new reviews arrive.
function ensureDashboard_(ss) {
  var dash = ss.getSheetByName('Dashboard');
  if (!dash) dash = ss.insertSheet('Dashboard', 0); // put it as the first tab
  if (dash.getRange('A1').getFormula()) return;      // already set up

  dash.getRange('A1').setFormula(
    '=QUERY(Avis!B2:E, "select B, C, D, count(E), avg(E) ' +
    'where B is not null group by B, C, D order by avg(E) desc ' +
    'label B \'Employé\', C \'Poste\', D \'Bureau\', ' +
    'count(E) \'Avis reçus\', avg(E) \'Note moyenne\'")'
  );
  dash.getRange('E:E').setNumberFormat('0.0');        // average to 1 decimal
  dash.getRange('1:1').setFontWeight('bold');
  dash.setFrozenRows(1);
}

// Optional: run this once manually to build the Dashboard before any reviews.
function setupDashboard() {
  ensureDashboard_(SpreadsheetApp.getActiveSpreadsheet());
}
```

3. Click **Save** (💾).

## 3. Deploy as a Web App
1. **Deploy → New deployment**.
2. Click the gear next to "Select type" → **Web app**.
3. Set:
   - **Description:** Reviews
   - **Execute as:** **Me**
   - **Who has access:** **Anyone**
4. **Deploy**, then **Authorize access** (approve the Google permission prompt —
   it's your own script writing to your own sheet).
5. Copy the **Web app URL**. It ends in `/exec`, e.g.
   `https://script.google.com/macros/s/AKfy...../exec`.

## 4. Connect the site
Open `assets/card.js` and paste the URL into the `REVIEW_ENDPOINT` constant near
the top:

```javascript
var REVIEW_ENDPOINT = "https://script.google.com/macros/s/AKfy...../exec";
```

Commit and deploy (merge to `main`). Done — new reviews now append to the sheet.

## The Dashboard tab
The script automatically maintains a **Dashboard** tab (the first tab) with one
row per employee:

| Employé | Poste | Bureau | Nombre d'avis | Note moyenne |
|---------|-------|--------|---------------|--------------|

It's a live formula sorted by highest average rating first, so it updates on its
own as new reviews come in — no maintenance needed. It appears automatically
with the first review; to build it right away (before any reviews), run the
`setupDashboard` function once from the Apps Script editor (**Run** ▶).

## Notes
- **Raw data:** the `Avis` tab keeps every review (date, employee, role, office,
  rating, comment, client, page) so you can filter, export, or build your own
  pivots beyond the Dashboard.
- **Privacy / CORS:** the form submits "fire-and-forget" (`mode: no-cors`), so the
  browser never needs to read the response. If a submission ever fails (e.g. no
  network), the form automatically falls back to the mail-app method so the
  review is never lost.
- **Updating the script later:** after editing the Apps Script, use
  **Deploy → Manage deployments → Edit → Version: New version** so the URL stays
  the same. (A brand-new deployment creates a new URL.)
- **Spam:** submissions are public by nature. If you start getting spam, tell me
  and I'll add a hidden honeypot field + basic rate check.
