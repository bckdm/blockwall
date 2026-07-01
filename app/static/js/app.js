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

  // ------------------------------------------------------------------
  // 5. Senden modal — pick asset, enter address + amount, submit
  // ------------------------------------------------------------------
  var sendModal       = null;
  var sendForm        = null;
  var sendPicker      = null;
  var sendCoinSearch  = null;
  var sendCoinList    = null;
  var sendSymbolInput = null;
  var sendCoinName    = null;
  var sendCoinSub     = null;
  var sendCoinIcon    = null;
  var sendAmountInput = null;
  var sendAmountUnit  = null;
  var sendAmountEquiv = null;
  var sendMaxBtn      = null;

  // Per-coin data: price (EUR) + available balance (qty) + decimals
  var SEND_COIN_DATA = (function () {
    var map = {};
    document.querySelectorAll("#bc-send-coin-list .bc-modal__coin").forEach(function (row) {
      var sym = String(row.dataset.bcSendCoin || "").toUpperCase();
      if (!sym) return;
      map[sym] = {
        name: row.dataset.bcSendCoinName || sym,
        color: row.dataset.bcSendCoinColor || "#888",
        balanceEur: parseFloat(row.dataset.bcSendCoinBalance || "0") || 0,
        balanceQty: parseFloat(row.dataset.bcSendCoinQty || "0") || 0,
      };
    });
    return map;
  })();

  // Pull COIN_PRICE from /api/wallet? No — embed prices in the data attributes
  // of each pick row to avoid an extra round trip. (Simpler: prices are static.)
  var SEND_COIN_PRICES = {};
  document.querySelectorAll("#bc-send-coin-list .bc-modal__coin").forEach(function (row) {
    // If the row has data-bc-send-coin-price, use it. Otherwise compute from balance.
    // We didn't put price on the data attrs; derive from EUR / QTY when both > 0.
    var sym = String(row.dataset.bcSendCoin || "").toUpperCase();
    var eur = parseFloat(row.dataset.bcSendCoinBalance || "0") || 0;
    var qty = parseFloat(row.dataset.bcSendCoinQty || "0") || 0;
    if (qty > 0) SEND_COIN_PRICES[sym] = eur / qty;
  });
  // Fallback: hardcoded prices (must match COINS in app.py).
  // Used only when balance is 0 and we can't derive price from balance.
  if (!Object.keys(SEND_COIN_PRICES).length) {
    SEND_COIN_PRICES = {
      BTC: 87234.50, ETH: 3208.72, USDT: 0.92, BNB: 584.30, SOL: 152.18,
      USDC: 0.92, XRP: 0.48, ADA: 0.36, DOGE: 0.12, TRX: 0.28,
      MATIC: 0.42, DOT: 5.83, BCH: 384.50,
    };
  }

  function fmtEur(n) {
    return "€" + (Math.round((n || 0) * 100) / 100)
      .toFixed(2).replace(".", ",");
  }
  function fmtQty(n, decimals) {
    if (typeof decimals !== "number") decimals = 8;
    return (Math.round((n || 0) * Math.pow(10, decimals)) / Math.pow(10, decimals))
      .toFixed(decimals).replace(/\.?0+$/, "");
  }

  function updateSendDisplay() {
    var sym = (sendSymbolInput && sendSymbolInput.value || "BTC").toUpperCase();
    var d = SEND_COIN_DATA[sym];
    if (!d) return;
    if (sendCoinIcon) {
      sendCoinIcon.textContent = sym.charAt(0);
      sendCoinIcon.style.background = d.color;
    }
    if (sendCoinName) sendCoinName.textContent = d.name;
    if (sendCoinSub) {
      sendCoinSub.textContent = fmtQty(d.balanceQty, 8) + " " + sym + " · ≈ " + fmtEur(d.balanceEur);
    }
    if (sendAmountUnit) sendAmountUnit.textContent = sym;
    updateSendEquiv();
  }

  function updateSendEquiv() {
    var sym = (sendSymbolInput && sendSymbolInput.value || "BTC").toUpperCase();
    var price = SEND_COIN_PRICES[sym] || 0;
    var qty = parseFloat((sendAmountInput && sendAmountInput.value) || "0") || 0;
    if (sendAmountEquiv) sendAmountEquiv.textContent = "≈ " + fmtEur(qty * price);
  }

  function openSendModal() {
    if (!sendModal) sendModal = document.getElementById("bc-send-modal");
    if (!sendModal) return;
    sendModal.classList.add("open");
    sendModal.setAttribute("aria-hidden", "false");
    document.body.style.overflow = "hidden";
    updateSendDisplay();
    setTimeout(function () {
      var addr = document.getElementById("bc-send-address");
      if (addr) addr.focus();
    }, 50);
  }
  function closeSendModal() {
    if (!sendModal) sendModal = document.getElementById("bc-send-modal");
    if (!sendModal) return;
    sendModal.classList.remove("open");
    sendModal.setAttribute("aria-hidden", "true");
    document.body.style.overflow = "";
    if (sendForm) sendForm.reset();
    showSendPicker(false);
  }
  window.openSendModal  = openSendModal;
  window.closeSendModal = closeSendModal;

  function showSendPicker(show) {
    if (!sendPicker) sendPicker = document.getElementById("bc-send-picker");
    if (!sendForm)    sendForm    = document.getElementById("bc-send-form");
    if (!sendPicker || !sendForm) return;
    if (show) {
      sendPicker.hidden = false;
      sendForm.style.display = "none";
      if (sendCoinSearch) {
        sendCoinSearch.value = "";
        filterSendCoins("");
        setTimeout(function () { sendCoinSearch.focus(); }, 30);
      }
    } else {
      sendPicker.hidden = true;
      sendForm.style.display = "";
    }
  }

  function filterSendCoins(q) {
    if (!sendCoinList) return;
    q = (q || "").trim().toLowerCase();
    sendCoinList.querySelectorAll(".bc-modal__coin").forEach(function (row) {
      var name = (row.dataset.bcSendCoinName || "").toLowerCase();
      var sym  = (row.dataset.bcSendCoin || "").toLowerCase();
      row.style.display = (!q || (name + " " + sym).indexOf(q) !== -1) ? "" : "none";
    });
  }

  document.addEventListener("DOMContentLoaded", function () {
    sendModal       = document.getElementById("bc-send-modal");
    sendForm        = document.getElementById("bc-send-form");
    sendPicker      = document.getElementById("bc-send-picker");
    sendCoinSearch  = document.getElementById("bc-send-coin-search");
    sendCoinList    = document.getElementById("bc-send-coin-list");
    sendSymbolInput = document.getElementById("bc-send-symbol");
    sendCoinName    = document.getElementById("bc-send-coin-name");
    sendCoinSub     = document.getElementById("bc-send-coin-sub");
    sendCoinIcon    = document.getElementById("bc-send-coin-icon");
    sendAmountInput = document.getElementById("bc-send-amount");
    sendAmountUnit  = document.getElementById("bc-send-amount-unit");
    sendAmountEquiv = document.getElementById("bc-send-amount-equiv");
    sendMaxBtn      = document.getElementById("bc-send-max");
    if (!sendModal) return;

    // Click outside / on close button
    sendModal.addEventListener("click", function (e) {
      var t = e.target;
      if (t && t.getAttribute && t.getAttribute("data-bc-modal-close") === "1") {
        closeSendModal();
      }
    });
    document.addEventListener("keydown", function (e) {
      if (e.key === "Escape" && sendModal.classList.contains("open")) {
        // Close the picker first if it's open, then the modal
        if (sendPicker && !sendPicker.hidden) {
          showSendPicker(false);
        } else {
          closeSendModal();
        }
      }
    });

    // "Ändern" opens the picker
    var changeBtn = document.getElementById("bc-send-coin-change");
    if (changeBtn) changeBtn.addEventListener("click", function () { showSendPicker(true); });

    // Picker search
    if (sendCoinSearch) {
      sendCoinSearch.addEventListener("input", function () { filterSendCoins(sendCoinSearch.value); });
    }

    // Pick a coin in the picker
    if (sendCoinList) {
      sendCoinList.addEventListener("click", function (e) {
        var row = e.target.closest(".bc-modal__coin");
        if (!row) return;
        var sym = String(row.dataset.bcSendCoin || "").toUpperCase();
        if (sendSymbolInput) sendSymbolInput.value = sym;
        showSendPicker(false);
        updateSendDisplay();
        // Reset amount + focus
        if (sendAmountInput) {
          sendAmountInput.value = "";
          updateSendEquiv();
          setTimeout(function () { sendAmountInput.focus(); }, 30);
        }
      });
    }

    // Recompute EUR equivalent on amount input
    if (sendAmountInput) {
      sendAmountInput.addEventListener("input", updateSendEquiv);
    }

    // Max button — fill available balance
    if (sendMaxBtn) {
      sendMaxBtn.addEventListener("click", function () {
        var sym = (sendSymbolInput.value || "BTC").toUpperCase();
        var d = SEND_COIN_DATA[sym];
        if (!d || !sendAmountInput) return;
        sendAmountInput.value = String(d.balanceQty);
        updateSendEquiv();
      });
    }

    // Generic "open send" triggers
    document.querySelectorAll("[data-bc-open-modal='send']").forEach(function (btn) {
      btn.addEventListener("click", function (e) {
        e.preventDefault();
        openSendModal();
      });
    });

    // Initial render
    updateSendDisplay();
  });
})();