/* blockwall shared JS — loader, balance hide, refresh, modal, coin picker.
 * Vanilla ES5+, no dependencies. Safe to include on every page.
 */
(function () {
  "use strict";

  // ------------------------------------------------------------------
  // 1. Loader — show / hide full-page overlay
  // ------------------------------------------------------------------
  var LOADER_MIN_MS = 120;       // never flash the loader for shorter than this
  var LOADER_STARTED_AT = 0;
  var LOADER_PENDING = 0;

  function startLoader() {
    LOADER_PENDING++;
    document.body.setAttribute("data-bc-loading", "1");
    LOADER_STARTED_AT = Date.now();
  }

  function stopLoader() {
    LOADER_PENDING = Math.max(0, LOADER_PENDING - 1);
    if (LOADER_PENDING > 0) return;
    var elapsed = Date.now() - LOADER_STARTED_AT;
    var wait = Math.max(0, LOADER_MIN_MS - elapsed);
    setTimeout(function () {
      // re-check (a new loader may have started in the meantime)
      if (LOADER_PENDING === 0) {
        document.body.removeAttribute("data-bc-loading");
      }
    }, wait);
  }

  window.startLoader = startLoader;
  window.stopLoader  = stopLoader;

  // Show loader on initial page load for a brief moment, then hide.
  document.addEventListener("DOMContentLoaded", function () {
    startLoader();
    // hide on next tick (DOMContentLoaded fires before paint, so this is fast)
    setTimeout(stopLoader, 30);
  });

  // Show loader between navigations — click any internal link
  document.addEventListener("click", function (e) {
    var a = e.target.closest && e.target.closest("a[href]");
    if (!a) return;
    var href = a.getAttribute("href");
    if (!href || href.charAt(0) === "#") return;
    if (a.target === "_blank") return;
    if (a.hasAttribute("data-bc-no-loader")) return;
    if (e.metaKey || e.ctrlKey || e.shiftKey || e.altKey) return;
    if (a.origin && a.origin !== window.location.origin) return;
    // Treat as a navigation — show loader until the new page paints
    startLoader();
  });

  // ------------------------------------------------------------------
  // 2. Balance hide — toggle "Ihre Guthaben ausblenden"
  // ------------------------------------------------------------------
  var HIDDEN_PLACEHOLDER = "€******";

  function applyBalanceHide() {
    var hidden = document.body.getAttribute("data-bc-balance-hidden") === "1";
    document.querySelectorAll(".bc-amount").forEach(function (el) {
      if (!el.dataset.bcBalance) {
        el.dataset.bcBalance = el.textContent;
      }
      el.textContent = hidden ? HIDDEN_PLACEHOLDER : el.dataset.bcBalance;
    });
  }

  window.addEventListener("bc:balances-updated", applyBalanceHide);

  // Eye-icon button
  document.addEventListener("DOMContentLoaded", function () {
    var eyes = document.querySelectorAll('[title="Ein-/Ausblenden"]');
    eyes.forEach(function (btn) {
      btn.addEventListener("click", function (e) {
        e.preventDefault();
        var hidden = document.body.getAttribute("data-bc-balance-hidden") === "1";
        if (hidden) {
          document.body.removeAttribute("data-bc-balance-hidden");
        } else {
          document.body.setAttribute("data-bc-balance-hidden", "1");
        }
        applyBalanceHide();
      });
    });

    // ------------------------------------------------------------------
    // 3. Refresh — fetch /api/wallet + /api/activity, update DOM
    // ------------------------------------------------------------------
    var r = document.getElementById("bc-refresh");
    if (r) {
      r.addEventListener("click", function (e) {
        e.preventDefault();
        r.disabled = true;
        startLoader();
        var done = 0;
        function maybeStop() {
          done++;
          if (done >= 2) {
            r.disabled = false;
            stopLoader();
          }
        }
        fetch("/api/wallet", { credentials: "same-origin" })
          .then(function (resp) { return resp.ok ? resp.json() : null; })
          .then(function (data) {
            if (data && data.wallets) updateWalletDom(data);
            maybeStop();
          })
          .catch(maybeStop);
        fetch("/api/activity", { credentials: "same-origin" })
          .then(function (resp) { return resp.ok ? resp.json() : null; })
          .then(function (data) {
            if (data && data.activity) updateActivityDom(data);
            maybeStop();
          })
          .catch(maybeStop);
      });
    }
  });

  function updateWalletDom(data) {
    // Update every element with data-bc-wallet-symbol="<SYM>" with new balance text
    var bySym = {};
    (data.wallets || []).forEach(function (w) {
      bySym[String(w.symbol).toUpperCase()] = w;
    });
    document.querySelectorAll("[data-bc-wallet-symbol]").forEach(function (el) {
      var sym = String(el.getAttribute("data-bc-wallet-symbol")).toUpperCase();
      var w = bySym[sym];
      if (!w) return;
      var eur = Number(w.balance_eur || 0);
      var qty = Number(w.balance_qty || 0);
      var qtyUnit = w.qty_unit || sym;
      var eurTxt = "€" + eur.toFixed(2);
      var qtyTxt = "(" + Math.round(qty * 10000) / 10000 + " " + qtyUnit + ")";
      if (el.dataset.bcBalance !== undefined || el.classList.contains("bc-amount")) {
        el.dataset.bcBalance = eurTxt;
        el.textContent = eurTxt;
      }
      var qtyEl = el.parentNode && el.parentNode.querySelector("[data-bc-wallet-qty]");
      if (qtyEl) qtyEl.textContent = qtyTxt;
    });

    // Update net-worth element if present
    var net = document.querySelector("[data-bc-networth]");
    if (net && typeof data.networth !== "undefined") {
      var t = "€" + Number(data.networth).toFixed(2);
      net.dataset.bcBalance = t;
      net.textContent = t;
    }
    window.dispatchEvent(new CustomEvent("bc:balances-updated"));
  }

  function updateActivityDom(data) {
    // Light touch — only update rows that already exist
    var list = document.querySelector("[data-bc-activity-list]");
    if (!list) return;
    var rows = data.activity || [];
    list.innerHTML = rows.slice(0, 5).map(function (r) {
      return '<div class="bc-row"><div class="bc-row-main"><div class="bc-row-main__icon">'
        + '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 5v14M5 12h14"/></svg>'
        + '</div><div class="bc-row-main__text"><div class="bc-row-main__label">'
        + escapeHtml(String(r.label || ""))
        + '</div><div class="bc-row-main__sub">'
        + escapeHtml(String(r.ts || ""))
        + '</div></div></div>'
        + '<div class="bc-row-aside__amount">€' + Number(r.amount_eur || 0).toFixed(2) + '</div>'
        + '</div>';
    }).join("");
  }

  function escapeHtml(s) {
    return s.replace(/[&<>"']/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c];
    });
  }

  // ------------------------------------------------------------------
  // 4. Einzahlung modal — open / close / search / pick
  // ------------------------------------------------------------------
  var modal = null;
  var coinSearch = null;
  var coinList = null;

  function openEinzahlungModal() {
    if (!modal) modal = document.getElementById("bc-einzahlung-modal");
    if (!modal) return;
    modal.classList.add("open");
    modal.setAttribute("aria-hidden", "false");
    document.body.style.overflow = "hidden";
    setTimeout(function () {
      if (coinSearch) coinSearch.focus();
    }, 50);
  }
  function closeEinzahlungModal() {
    if (!modal) modal = document.getElementById("bc-einzahlung-modal");
    if (!modal) return;
    modal.classList.remove("open");
    modal.setAttribute("aria-hidden", "true");
    document.body.style.overflow = "";
    if (coinSearch) {
      coinSearch.value = "";
      filterCoins("");
    }
  }
  window.openEinzahlungModal  = openEinzahlungModal;
  window.closeEinzahlungModal = closeEinzahlungModal;

  function filterCoins(q) {
    if (!coinList) return;
    q = (q || "").trim().toLowerCase();
    coinList.querySelectorAll(".bc-modal__coin").forEach(function (row) {
      var name = (row.dataset.bcCoinName || "").toLowerCase();
      var sym  = (row.dataset.bcCoin || "").toLowerCase();
      var hay  = name + " " + sym;
      row.style.display = (!q || hay.indexOf(q) !== -1) ? "" : "none";
    });
  }

  document.addEventListener("DOMContentLoaded", function () {
    modal = document.getElementById("bc-einzahlung-modal");
    coinSearch = document.getElementById("bc-coin-search");
    coinList = document.getElementById("bc-coin-list");
    if (!modal) return;

    // Click outside / on close button
    modal.addEventListener("click", function (e) {
      var t = e.target;
      if (t && t.getAttribute && t.getAttribute("data-bc-modal-close") === "1") {
        closeEinzahlungModal();
      }
    });
    document.addEventListener("keydown", function (e) {
      if (e.key === "Escape" && modal.classList.contains("open")) {
        closeEinzahlungModal();
      }
    });

    // Search filter
    if (coinSearch) {
      coinSearch.addEventListener("input", function () {
        filterCoins(coinSearch.value);
      });
    }

    // Pick coin
    if (coinList) {
      coinList.addEventListener("click", function (e) {
        var row = e.target.closest(".bc-modal__coin");
        if (!row) return;
        var sym = row.dataset.bcCoin;
        var name = row.dataset.bcCoinName;
        var color = row.dataset.bcCoinColor;
        window.dispatchEvent(new CustomEvent("bc:coin-picked", {
          detail: { symbol: sym, name: name, color: color }
        }));
        closeEinzahlungModal();
      });
    }

    // Generic "open einzahlung" triggers
    document.querySelectorAll("[data-bc-open-modal='einzahlung']").forEach(function (btn) {
      btn.addEventListener("click", function (e) {
        e.preventDefault();
        openEinzahlungModal();
      });
    });
  });
})();