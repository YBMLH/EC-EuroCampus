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
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
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

## Notes
- **Per-employee tracking:** each row records the employee, role, and office, so
  you can filter or build a pivot table for each person's average rating.
- **Privacy / CORS:** the form submits "fire-and-forget" (`mode: no-cors`), so the
  browser never needs to read the response. If a submission ever fails (e.g. no
  network), the form automatically falls back to the mail-app method so the
  review is never lost.
- **Updating the script later:** after editing the Apps Script, use
  **Deploy → Manage deployments → Edit → Version: New version** so the URL stays
  the same. (A brand-new deployment creates a new URL.)
- **Spam:** submissions are public by nature. If you start getting spam, tell me
  and I'll add a hidden honeypot field + basic rate check.
