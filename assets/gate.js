/* =========================================================
   EURO CAMPUS — passcode gate
   Soft client-side lock for the directory + dashboard. Hides
   the page until the correct passcode is entered; the choice
   is remembered per device. The passcode is never stored in
   plain text — only the SHA-256 hash below is compared.
   NOTE: this keeps casual visitors out; it is not strong
   security (the individual cards stay public by design).
   ========================================================= */
(function () {
  "use strict";

  // SHA-256 of the passcode.
  var PASS_HASH = "3de8623c168895576f471bcb4491c37d9082bb712cf3c0a68796fd6e66f937d4";
  var KEY = "ec-gate-ok";

  // Already unlocked on this device → do nothing.
  try { if (localStorage.getItem(KEY) === PASS_HASH) return; } catch (e) {}

  var doc = document;
  var root = doc.documentElement;

  // Hide page content immediately (head runs before body is parsed).
  root.classList.add("ec-locked");
  var style = doc.createElement("style");
  style.textContent =
    "html.ec-locked .wrap{display:none!important}" +
    ".ec-gate{position:fixed;inset:0;z-index:9999;display:flex;align-items:center;justify-content:center;padding:22px;" +
      "background:linear-gradient(150deg,#1EA9BA 0%,#157f8c 60%,#0e5c66 100%);font-family:'Manrope',sans-serif}" +
    ".ec-gate-card{width:100%;max-width:340px;background:#fff;border-radius:22px;padding:30px 26px;text-align:center;" +
      "box-shadow:0 30px 70px rgba(8,32,36,.35)}" +
    ".ec-gate-mark{width:54px;height:54px;margin:0 auto 14px}" +
    ".ec-gate-title{font-family:'Poppins',sans-serif;font-weight:700;font-size:1.2rem;color:#0d2c30}" +
    ".ec-gate-sub{font-size:.84rem;color:#5d7a7e;margin-top:5px;line-height:1.4}" +
    ".ec-gate-in{width:100%;margin-top:18px;padding:13px 15px;border:1px solid rgba(30,169,186,.3);border-radius:12px;" +
      "font-size:1rem;font-family:inherit;text-align:center;letter-spacing:.04em}" +
    ".ec-gate-in:focus{outline:none;border-color:#1EA9BA;box-shadow:0 0 0 3px rgba(30,169,186,.15)}" +
    ".ec-gate-btn{width:100%;margin-top:12px;padding:14px;border:none;border-radius:12px;cursor:pointer;" +
      "background:#1EA9BA;color:#fff;font-family:inherit;font-weight:700;font-size:.96rem;transition:transform .15s}" +
    ".ec-gate-btn:hover{transform:translateY(-2px)}.ec-gate-btn:active{transform:scale(.97)}" +
    ".ec-gate-err{color:#d9534f;font-size:.82rem;font-weight:600;min-height:18px;margin-top:10px}" +
    ".ec-shake{animation:ecShake .4s}" +
    "@keyframes ecShake{0%,100%{transform:translateX(0)}20%,60%{transform:translateX(-7px)}40%,80%{transform:translateX(7px)}}";
  (doc.head || root).appendChild(style);

  function sha256Hex(str) {
    var enc = new TextEncoder().encode(str);
    return crypto.subtle.digest("SHA-256", enc).then(function (buf) {
      return Array.prototype.map.call(new Uint8Array(buf), function (b) {
        return ("0" + b.toString(16)).slice(-2);
      }).join("");
    });
  }

  var built = false;
  function build() {
    if (built) return;            // never create more than one overlay
    built = true;
    var gate = doc.createElement("div");
    gate.className = "ec-gate";
    gate.innerHTML =
      '<div class="ec-gate-card">' +
        '<svg class="ec-gate-mark" viewBox="0 0 64 64" fill="none">' +
          '<rect x="4" y="14" width="26" height="6" rx="3" fill="#1EA9BA"/>' +
          '<rect x="4" y="29" width="26" height="6" rx="3" fill="#1EA9BA"/>' +
          '<rect x="4" y="44" width="26" height="6" rx="3" fill="#1EA9BA"/>' +
          '<path d="M58 20a18 18 0 1 0 0 24" stroke="#1EA9BA" stroke-width="6" stroke-linecap="round" fill="none"/>' +
        '</svg>' +
        '<div class="ec-gate-title">Accès réservé</div>' +
        '<div class="ec-gate-sub">Entrez le code d\'accès pour continuer.</div>' +
        '<input class="ec-gate-in" type="password" inputmode="text" autocomplete="off" ' +
          'placeholder="Code d\'accès" aria-label="Code d\'accès"/>' +
        '<button class="ec-gate-btn" type="button">Entrer</button>' +
        '<div class="ec-gate-err" role="alert"></div>' +
      '</div>';
    doc.body.appendChild(gate);

    var input = gate.querySelector(".ec-gate-in");
    var btn = gate.querySelector(".ec-gate-btn");
    var err = gate.querySelector(".ec-gate-err");
    var card = gate.querySelector(".ec-gate-card");
    input.focus();

    function fail(msg) {
      err.textContent = msg || "Code incorrect.";
      card.classList.remove("ec-shake");
      void card.offsetWidth;            // restart animation
      card.classList.add("ec-shake");
      input.select();
    }

    function submit() {
      var val = input.value.trim();
      if (!val) { fail("Veuillez saisir le code."); return; }
      sha256Hex(val).then(function (hex) {
        if (hex === PASS_HASH) {
          try { localStorage.setItem(KEY, PASS_HASH); } catch (e) {}
          root.classList.remove("ec-locked");
          gate.remove();
        } else {
          fail("Code incorrect.");
        }
      }).catch(function () { fail("Erreur, réessayez."); });
    }

    btn.addEventListener("click", submit);
    input.addEventListener("keydown", function (e) { if (e.key === "Enter") submit(); });
  }

  if (doc.readyState === "loading") {
    doc.addEventListener("DOMContentLoaded", build);
  } else {
    build();
  }
})();
