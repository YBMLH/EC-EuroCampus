/* =========================================================
   EURO CAMPUS — shared card behaviour
   Adds three features to every NFC card / the directory:
     • Dark mode  (system default + manual toggle, remembered)
     • Share      (native share sheet, copy-link fallback)
     • QR code    (scan to save the contact / open the card)
   The script is purely additive: it reads data already in the
   page and injects its own UI, so the existing markup is
   untouched and the page still works if JS is disabled.
   ========================================================= */
(function () {
  "use strict";

  var doc = document;
  var root = doc.documentElement;
  var STORE_KEY = "ec-theme";

  /* ---------- icons ---------- */
  var ICONS = {
    sun:  '<svg viewBox="0 0 24 24"><path d="M12 7a5 5 0 100 10 5 5 0 000-10zm0-5a1 1 0 011 1v2a1 1 0 11-2 0V3a1 1 0 011-1zm0 16a1 1 0 011 1v2a1 1 0 11-2 0v-2a1 1 0 011-1zM4.22 4.22a1 1 0 011.42 0l1.41 1.41A1 1 0 115.64 7.05L4.22 5.64a1 1 0 010-1.42zm12.71 12.72a1 1 0 011.42 0l1.41 1.41a1 1 0 01-1.42 1.42l-1.41-1.41a1 1 0 010-1.42zM2 12a1 1 0 011-1h2a1 1 0 110 2H3a1 1 0 01-1-1zm17 0a1 1 0 011-1h2a1 1 0 110 2h-2a1 1 0 01-1-1zM4.22 19.78a1 1 0 010-1.42l1.41-1.41a1 1 0 011.42 1.42l-1.41 1.41a1 1 0 01-1.42 0zM16.93 7.05a1 1 0 010-1.42l1.41-1.41a1 1 0 011.42 1.42l-1.41 1.41a1 1 0 01-1.42 0z"/></svg>',
    moon: '<svg viewBox="0 0 24 24"><path d="M21 12.79A9 9 0 1111.21 3 7 7 0 0021 12.79z"/></svg>',
    share:'<svg viewBox="0 0 24 24"><path d="M18 16.08c-.76 0-1.44.3-1.96.77L8.91 12.7c.05-.23.09-.46.09-.7s-.04-.47-.09-.7l7.05-4.11c.54.5 1.25.81 2.04.81a3 3 0 100-6 3 3 0 00-3 3c0 .24.04.47.09.7L8.04 9.81A3 3 0 003 12a3 3 0 005.04 2.19l7.12 4.16c-.05.21-.08.43-.08.65a2.92 2.92 0 102.92-2.92z"/></svg>',
    qr:   '<svg viewBox="0 0 24 24"><path d="M3 11h8V3H3v8zm2-6h4v4H5V5zm-2 16h8v-8H3v8zm2-6h4v4H5v-4zM13 3v8h8V3h-8zm6 6h-4V5h4v4zm-6 4h2v2h-2v-2zm2 2h2v2h-2v-2zm-2 2h2v2h-2v-2zm4 0h2v2h-2v-2zm0-4h2v2h-2v-2zm-2 0h-2 4-2zm2 0h2v2h-2v-2z"/></svg>'
  };

  /* ---------- theme ---------- */
  function systemPrefersDark() {
    return window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches;
  }
  function currentTheme() {
    var saved = null;
    try { saved = localStorage.getItem(STORE_KEY); } catch (e) {}
    return saved || (systemPrefersDark() ? "dark" : "light");
  }
  function applyTheme(theme) {
    root.setAttribute("data-theme", theme);
    var meta = doc.querySelector('meta[name="theme-color"]');
    if (meta) meta.setAttribute("content", theme === "dark" ? "#081417" : "#1EA9BA");
  }
  applyTheme(currentTheme()); // run ASAP to limit flash

  /* ---------- small helpers ---------- */
  function el(tag, cls, html) {
    var n = doc.createElement(tag);
    if (cls) n.className = cls;
    if (html != null) n.innerHTML = html;
    return n;
  }
  function toast(msg) {
    var t = doc.querySelector(".ec-toast");
    if (!t) { t = el("div", "ec-toast"); doc.body.appendChild(t); }
    t.textContent = msg;
    requestAnimationFrame(function () { t.classList.add("show"); });
    clearTimeout(t._timer);
    t._timer = setTimeout(function () { t.classList.remove("show"); }, 2200);
  }

  /* ---------- read the card's own data ---------- */
  function readCard() {
    var nameEl = doc.querySelector(".person-name");
    var roleEl = doc.querySelector(".person-role");
    var saveEl = doc.querySelector("a.save[href^='data:text/vcard']");
    var vcard = null;
    if (saveEl) {
      try { vcard = decodeURIComponent(saveEl.getAttribute("href").split(",")[1]); } catch (e) {}
    }
    return {
      isCard: !!saveEl,
      name: nameEl ? nameEl.textContent.trim() : (doc.title || "Euro Campus"),
      role: roleEl ? roleEl.textContent.trim() : "",
      vcard: vcard
    };
  }

  /* ---------- share ---------- */
  function doShare(info) {
    var url = location.href;
    var shareData = {
      title: info.name + " — Euro Campus",
      text: info.role ? info.name + " · " + info.role : info.name,
      url: url
    };
    if (navigator.share) {
      navigator.share(shareData).catch(function () {});
      return;
    }
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(url).then(
        function () { toast("Lien copié ✓"); },
        function () { toast(url); }
      );
    } else {
      toast(url);
    }
  }

  /* ---------- QR modal ---------- */
  function buildQrModal(info) {
    var overlay = el("div", "ec-modal");
    var card = el("div", "ec-modal-card");
    var encodingContact = !!info.vcard;
    card.appendChild(el("div", "ec-modal-title", encodingContact ? "Carte de contact" : "Partager cette page"));
    card.appendChild(el("div", "ec-modal-sub",
      encodingContact ? "Scannez pour enregistrer le contact directement." : "Scannez pour ouvrir cette page."));

    var qrBox = el("div", "ec-qr");
    card.appendChild(qrBox);

    if (info.name) card.appendChild(el("div", "ec-modal-name", info.name));
    if (info.role) card.appendChild(el("div", "ec-modal-role", info.role));

    var close = el("button", "ec-close", "Fermer");
    card.appendChild(close);
    overlay.appendChild(card);
    doc.body.appendChild(overlay);

    var rendered = false;
    function render() {
      if (rendered || typeof qrcode === "undefined") return;
      var data = info.vcard || location.href;
      var qr = qrcode(0, "M");        // type 0 = auto-fit, error level M
      qr.addData(data);
      qr.make();
      qrBox.innerHTML = qr.createSvgTag({ cellSize: 6, margin: 0, scalable: true });
      rendered = true;
    }
    function open()  { render(); overlay.classList.add("open"); }
    function hide()  { overlay.classList.remove("open"); }

    close.addEventListener("click", hide);
    overlay.addEventListener("click", function (e) { if (e.target === overlay) hide(); });
    doc.addEventListener("keydown", function (e) { if (e.key === "Escape") hide(); });
    return open;
  }

  /* ---------- toolbar ---------- */
  function build() {
    var info = readCard();
    var bar = el("div", "ec-toolbar");

    // theme toggle
    var themeBtn = el("button", "ec-tool");
    themeBtn.type = "button";
    themeBtn.setAttribute("aria-label", "Changer de thème");
    function syncThemeIcon() {
      themeBtn.innerHTML = root.getAttribute("data-theme") === "dark" ? ICONS.sun : ICONS.moon;
    }
    syncThemeIcon();
    themeBtn.addEventListener("click", function () {
      var next = root.getAttribute("data-theme") === "dark" ? "light" : "dark";
      applyTheme(next);
      try { localStorage.setItem(STORE_KEY, next); } catch (e) {}
      syncThemeIcon();
    });
    bar.appendChild(themeBtn);

    // share + QR only on actual cards
    if (info.isCard) {
      var shareBtn = el("button", "ec-tool", ICONS.share);
      shareBtn.type = "button";
      shareBtn.setAttribute("aria-label", "Partager");
      shareBtn.addEventListener("click", function () { doShare(info); });
      bar.appendChild(shareBtn);

      var openQr = buildQrModal(info);
      var qrBtn = el("button", "ec-tool", ICONS.qr);
      qrBtn.type = "button";
      qrBtn.setAttribute("aria-label", "Afficher le QR code");
      qrBtn.addEventListener("click", openQr);
      bar.appendChild(qrBtn);
    }

    doc.body.appendChild(bar);
  }

  if (doc.readyState === "loading") {
    doc.addEventListener("DOMContentLoaded", build);
  } else {
    build();
  }
})();
