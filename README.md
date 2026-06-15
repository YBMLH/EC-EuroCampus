# EC-EuroCampus

Digital NFC business cards for the Euro Campus team (study-abroad & immigration
agency, Algeria). Each `employe-NN.html` is a self-contained mobile profile page
whose URL is programmed onto a physical NFC card; `index.html` is the internal
directory of all 15 cards.

## Features

Every card includes:

- **Contact actions** — call, email, WhatsApp, LinkedIn, and a one-tap
  *vCard* download to save the contact.
- **QR code** — a scannable code (top-right toolbar) that encodes the same
  vCard, so anyone can save the contact by scanning even without an NFC reader.
- **Share** — native share sheet via the Web Share API, with a copy-link
  fallback on desktop.
- **Dark mode** — follows the system preference and can be toggled manually;
  the choice is remembered across visits.
- **Installable (PWA)** — a web manifest + icon let the directory be added to
  the home screen as an app.

## Structure

```
index.html              Directory of all cards
employe-01..15.html     Individual NFC profile pages
manifest.webmanifest    PWA manifest
assets/
  card.css              Shared toolbar / QR modal styles + dark-mode theme
  card.js               Shared behaviour (theme, share, QR) — injected, no-JS safe
  qrcode.js             Vendored QR generator (MIT, Kazuhiko Arase)
  icon.svg              App / favicon icon
```

The shared behaviour lives in `assets/` and is loaded by every page, so a single
edit there updates all cards at once. `card.js` reads each page's existing data
and injects its own UI, so the pages still work if JavaScript is disabled.

## Usage

It's a static site — open `index.html` in a browser, or serve the folder:

```bash
python3 -m http.server 8000
```

Then program each card's URL (e.g. `.../employe-01.html`) onto its NFC chip.
