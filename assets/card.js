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

  // ── Reviews ──────────────────────────────────────────────────────────────
  // Paste your Google Apps Script Web-App URL between the quotes to log reviews
  // to your spreadsheet. While empty, the review form falls back to opening the
  // visitor's mail app instead. Deployment steps are in REVIEWS.md.
  var REVIEW_ENDPOINT = "https://script.google.com/macros/s/AKfycbxVrPv_7MOGORDnFOqsMUdyNVTzBb9ZRjLPsplg79prvc6qd6zBzwGEP01pf0olQ4v_/exec";

  /* ---------- icons ---------- */
  var ICONS = {
    sun:  '<svg viewBox="0 0 24 24"><path d="M12 7a5 5 0 100 10 5 5 0 000-10zm0-5a1 1 0 011 1v2a1 1 0 11-2 0V3a1 1 0 011-1zm0 16a1 1 0 011 1v2a1 1 0 11-2 0v-2a1 1 0 011-1zM4.22 4.22a1 1 0 011.42 0l1.41 1.41A1 1 0 115.64 7.05L4.22 5.64a1 1 0 010-1.42zm12.71 12.72a1 1 0 011.42 0l1.41 1.41a1 1 0 01-1.42 1.42l-1.41-1.41a1 1 0 010-1.42zM2 12a1 1 0 011-1h2a1 1 0 110 2H3a1 1 0 01-1-1zm17 0a1 1 0 011-1h2a1 1 0 110 2h-2a1 1 0 01-1-1zM4.22 19.78a1 1 0 010-1.42l1.41-1.41a1 1 0 011.42 1.42l-1.41 1.41a1 1 0 01-1.42 0zM16.93 7.05a1 1 0 010-1.42l1.41-1.41a1 1 0 011.42 1.42l-1.41 1.41a1 1 0 01-1.42 0z"/></svg>',
    moon: '<svg viewBox="0 0 24 24"><path d="M21 12.79A9 9 0 1111.21 3 7 7 0 0021 12.79z"/></svg>',
    share:'<svg viewBox="0 0 24 24"><path d="M18 16.08c-.76 0-1.44.3-1.96.77L8.91 12.7c.05-.23.09-.46.09-.7s-.04-.47-.09-.7l7.05-4.11c.54.5 1.25.81 2.04.81a3 3 0 100-6 3 3 0 00-3 3c0 .24.04.47.09.7L8.04 9.81A3 3 0 003 12a3 3 0 005.04 2.19l7.12 4.16c-.05.21-.08.43-.08.65a2.92 2.92 0 102.92-2.92z"/></svg>',
    qr:   '<svg viewBox="0 0 24 24"><path d="M3 11h8V3H3v8zm2-6h4v4H5V5zm-2 16h8v-8H3v8zm2-6h4v4H5v-4zM13 3v8h8V3h-8zm6 6h-4V5h4v4zm-6 4h2v2h-2v-2zm2 2h2v2h-2v-2zm-2 2h2v2h-2v-2zm4 0h2v2h-2v-2zm0-4h2v2h-2v-2zm-2 0h-2 4-2zm2 0h2v2h-2v-2z"/></svg>',
    star: '<svg viewBox="0 0 24 24"><path d="M12 17.27l5.18 3.12-1.37-5.9 4.58-3.97-6.03-.52L12 4.5 9.64 10l-6.03.52 4.58 3.97-1.37 5.9z"/></svg>'
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

  /* ---------- slug + JSONP helpers ---------- */
  function slugFromPath() {
    var f = (location.pathname.split("/").pop() || "").replace(/\.html?$/i, "");
    return /^employe-/.test(f) ? f : "";
  }
  function jsonp(url, cb) {
    var name = "_ecjp_" + Math.random().toString(36).slice(2);
    var s = doc.createElement("script");
    var done = false;
    function finish(data) { if (done) return; done = true; try { cb(data); } catch (e) {} try { delete window[name]; } catch (e) { window[name] = undefined; } if (s.parentNode) s.parentNode.removeChild(s); }
    window[name] = function (data) { finish(data); };
    s.src = url + (url.indexOf("?") < 0 ? "?" : "&") + "callback=" + name + "&_=" + Date.now();
    s.onerror = function () { finish(null); };
    doc.body.appendChild(s);
  }

  /* ---------- usage analytics (fire-and-forget) ---------- */
  function trackEvent(action, info) {
    try {
      if (!REVIEW_ENDPOINT) return;
      var slug = slugFromPath();
      if (!slug) return;
      if (action === "view") {
        var k = "ec-viewed-" + slug;
        try { if (sessionStorage.getItem(k)) return; sessionStorage.setItem(k, "1"); } catch (e) {}
      }
      fetch(REVIEW_ENDPOINT, {
        method: "POST",
        mode: "no-cors",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify({ type: "event", action: action, slug: slug, employee: (info && info.name) || "", url: location.href })
      });
    } catch (e) {}
  }

  /* ---------- live content from the Sheet (progressive enhancement) ---------- */
  function buildVCard(r) {
    var name = (r.nom || "").trim();
    var parts = name.split(/\s+/);
    var last = parts.length > 1 ? parts[parts.length - 1] : name;
    var first = parts.length > 1 ? parts.slice(0, -1).join(" ") : "";
    var tel = (r.telephone || "").replace(/\s+/g, "");
    var lines = ["BEGIN:VCARD", "VERSION:3.0", "FN:" + name, "N:" + last + ";" + first + ";;;",
      "TITLE:" + (r.poste || ""), "ORG:Euro Campus;" + (r.departement || "")];
    if (tel) lines.push("TEL;TYPE=WORK,VOICE:" + tel);
    if (r.email) lines.push("EMAIL;TYPE=WORK:" + r.email);
    if (r.bureau) lines.push("ADR;TYPE=WORK:;;" + r.bureau + ";;;;Algeria");
    lines.push("URL:https://eurocampuss.com", "END:VCARD");
    return lines.join("\r\n");
  }
  function setText(sel, val) { var e = doc.querySelector(sel); if (e && val != null && val !== "") e.textContent = val; }
  function setHref(sel, val) { var e = doc.querySelector(sel); if (e && val) e.setAttribute("href", val); }
  function applyEmployee(info, r) {
    setText(".person-name", r.nom);
    setText(".person-role", r.poste);
    setText(".person-dept", r.departement);
    setText(".loc-v", r.bureau);
    setHref("a.loc", r.mapurl);
    if (r.telephone) setHref("a.b-call", "tel:" + r.telephone.replace(/\s+/g, ""));
    if (r.email)     setHref("a.b-mail", "mailto:" + r.email);
    if (r.whatsapp)  setHref("a.b-wa", "https://wa.me/" + r.whatsapp.replace(/[^\d]/g, ""));
    if (r.linkedin)  setHref("a.b-li", r.linkedin);
    if (r.nom)    info.name = r.nom;
    if (r.poste)  info.role = r.poste;
    if (r.bureau) info.office = r.bureau;
    if (r.email)  info.email = r.email;
    var saveEl = doc.querySelector("a.save[href^='data:text/vcard']");
    if (saveEl) {
      var vcf = buildVCard(r);
      saveEl.setAttribute("href", "data:text/vcard;charset=utf-8," + encodeURIComponent(vcf));
      info.vcard = vcf;
    }
    if (r.nom) doc.title = r.nom + " — Euro Campus";
  }
  function hydrateCard(info) {
    if (!REVIEW_ENDPOINT || !info.isCard) return;
    var slug = slugFromPath();
    if (!slug) return;
    jsonp(REVIEW_ENDPOINT + "?action=employee&slug=" + encodeURIComponent(slug), function (data) {
      if (data && data.ok && data.employee) applyEmployee(info, data.employee);
    });
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
    var locEl = doc.querySelector(".loc-v");
    var mailEl = doc.querySelector("a[href^='mailto:']");
    var email = mailEl
      ? mailEl.getAttribute("href").replace(/^mailto:/, "").split("?")[0]
      : "contact@euro-campus.com";
    return {
      isCard: !!saveEl,
      name: nameEl ? nameEl.textContent.trim() : (doc.title || "Euro Campus"),
      role: roleEl ? roleEl.textContent.trim() : "",
      office: locEl ? locEl.textContent.trim() : "",
      email: email,
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

  /* ---------- review (mailto) ---------- */
  function buildReviewModal(info) {
    var overlay = el("div", "ec-modal ec-review");
    var card = el("div", "ec-modal-card");

    card.appendChild(el("div", "ec-modal-title", "Laisser un avis"));
    card.appendChild(el("div", "ec-modal-sub",
      "Évaluez votre expérience avec " + info.name + (info.role ? " · " + info.role : "") + "."));

    // star rating
    var rating = 0;
    var stars = el("div", "ec-stars");
    var starEls = [];
    function paint(n) {
      for (var i = 0; i < 5; i++) starEls[i].classList.toggle("on", i < n);
    }
    for (var i = 1; i <= 5; i++) {
      (function (val) {
        var s = el("button", "ec-star", ICONS.star);
        s.type = "button";
        s.setAttribute("aria-label", val + " étoile" + (val > 1 ? "s" : ""));
        s.addEventListener("mouseenter", function () { paint(val); });
        s.addEventListener("focus", function () { paint(val); });
        s.addEventListener("click", function () { rating = val; paint(val); });
        starEls.push(s);
        stars.appendChild(s);
      })(i);
    }
    stars.addEventListener("mouseleave", function () { paint(rating); });
    card.appendChild(stars);

    // fields
    var nameField = el("input", "ec-field");
    nameField.type = "text";
    nameField.placeholder = "Votre nom (optionnel)";
    nameField.setAttribute("aria-label", "Votre nom");
    card.appendChild(nameField);

    var comment = el("textarea", "ec-field ec-textarea");
    comment.placeholder = "Votre commentaire…";
    comment.setAttribute("aria-label", "Votre commentaire");
    card.appendChild(comment);

    var hint = el("div", "ec-hint");
    card.appendChild(hint);

    var send = el("button", "ec-close", "Envoyer l'avis");
    card.appendChild(send);

    var cancel = el("button", "ec-cancel", "Annuler");
    card.appendChild(cancel);

    overlay.appendChild(card);
    doc.body.appendChild(overlay);

    function hide() { overlay.classList.remove("open"); }

    send.addEventListener("click", function () {
      if (!rating) {
        hint.textContent = "Merci de choisir une note ★";
        hint.classList.add("show");
        return;
      }
      var fullStars = "";
      for (var k = 0; k < 5; k++) fullStars += (k < rating ? "★" : "☆");

      var who = nameField.value.trim();

      // build the body as blank-line-separated sections (kept tidy in any mail app)
      var head = "Avis pour : " + info.name + (info.role ? " (" + info.role + ")" : "");
      if (info.office) head += "\r\nBureau : " + info.office;
      var note = "Note : " + fullStars + "  (" + rating + "/5)";
      var body = "Commentaire :\r\n" + (comment.value.trim() || "(aucun)");
      var foot = (who ? "De la part de : " + who + "\r\n" : "") + "— Envoyé depuis la carte Euro Campus";

      var subject = "Avis client — " + info.name + (info.role ? " (" + info.role + ")" : "");

      // Fallback: open the visitor's mail app with the formatted review.
      function mailtoFallback() {
        var href = "mailto:" + info.email +
          "?subject=" + encodeURIComponent(subject) +
          "&body=" + encodeURIComponent([head, note, body, foot].join("\r\n\r\n"));
        window.location.href = href;
        toast("Ouverture de votre messagerie…");
        hide();
      }

      // No backend configured → use the mail-app fallback.
      if (!REVIEW_ENDPOINT) { mailtoFallback(); return; }

      // Backend configured → submit silently to the Google Sheet.
      var payload = {
        employee: info.name,
        role: info.role,
        office: info.office,
        rating: rating,
        comment: comment.value.trim(),
        reviewer: who,
        url: location.href
      };
      send.disabled = true;
      send.textContent = "Envoi…";
      fetch(REVIEW_ENDPOINT, {
        method: "POST",
        mode: "no-cors", // Apps Script doesn't send CORS headers; fire-and-forget
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify(payload)
      }).then(function () {
        toast("Merci pour votre avis ✓");
        hide();
      }).catch(function () {
        // network failure — don't lose the review, hand off to the mail app
        mailtoFallback();
      });
    });

    cancel.addEventListener("click", hide);
    overlay.addEventListener("click", function (e) { if (e.target === overlay) hide(); });
    doc.addEventListener("keydown", function (e) { if (e.key === "Escape") hide(); });

    return function open() { overlay.classList.add("open"); };
  }

  /* ---------- toolbar ---------- */
  function build() {
    var info = readCard();
    if (info.isCard) { hydrateCard(info); trackEvent("view", info); }
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
      shareBtn.addEventListener("click", function () { trackEvent("share", info); doShare(info); });
      bar.appendChild(shareBtn);

      var openQr = buildQrModal(info);
      var qrBtn = el("button", "ec-tool", ICONS.qr);
      qrBtn.type = "button";
      qrBtn.setAttribute("aria-label", "Afficher le QR code");
      qrBtn.addEventListener("click", function () { trackEvent("qr", info); openQr(); });
      bar.appendChild(qrBtn);

      // count "save contact" taps on the vCard download link
      var saveLink = doc.querySelector("a.save[href^='data:text/vcard']");
      if (saveLink) saveLink.addEventListener("click", function () { trackEvent("save", info); });
    }

    // Cards have a header bar — drop the buttons in there (in normal flow,
    // so they never cover the logo / NFC badge). The directory has no header
    // bar, so the toolbar floats in the corner instead.
    var brand = doc.querySelector(".brand-bar");
    if (brand) {
      var group = el("div", "ec-bar-right");
      var badge = brand.querySelector(".nfc-badge");
      if (badge) group.appendChild(badge); // keep the NFC badge, grouped right
      group.appendChild(bar);
      brand.appendChild(group);
    } else {
      bar.classList.add("ec-floating");
      doc.body.appendChild(bar);
    }

    // Review CTA at the bottom of the card
    if (info.isCard) {
      var page = doc.querySelector(".page");
      if (page) {
        var openReview = buildReviewModal(info);
        var reviewBtn = el("button", "ec-review-btn", ICONS.star + "<span>Laisser un avis</span>");
        reviewBtn.type = "button";
        reviewBtn.addEventListener("click", function () { trackEvent("review_open", info); openReview(); });
        var divider = page.querySelector(".divider");
        page.insertBefore(reviewBtn, divider || null);
      }
    }
  }

  if (doc.readyState === "loading") {
    doc.addEventListener("DOMContentLoaded", build);
  } else {
    build();
  }
})();
