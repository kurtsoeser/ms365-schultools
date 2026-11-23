/**
 * Angebots-Konfigurator: Bundles → optional Bausteine → Kontakt.
 * Details in Dialog (kein Scroll-Sprung). Katalog: offer-catalog.js
 */
(() => {
  const catalog = window.MS365_OFFER_CATALOG;
  if (!catalog || !Array.isArray(catalog.modules)) {
    console.warn("[offer] MS365_OFFER_CATALOG fehlt");
    return;
  }

  const CONTACT_MAIL = catalog.contactMail || "kontakt@ms365.schule";
  const BUNDLES = catalog.bundles || [];
  const MODULES = catalog.modules || [];
  const SETUP =
    (catalog.addons || []).find((a) => a.id === "addon-setup") || {
      id: "addon-setup",
      label: "Setup / Onboarding",
      priceLabel: "590–790 € einmalig",
    };

  const root = document.querySelector("[data-offer-builder]");
  if (!root) return;

  const state = {
    selected: new Set(),
    setupOn: false,
    activeBundle: null,
    search: "",
    configOpen: false,
  };

  const els = {
    config: root.querySelector("[data-offer-config]"),
    moduleGrid: root.querySelector("[data-offer-modules]"),
    tip: root.querySelector("[data-offer-tip]"),
    price: root.querySelector("[data-offer-price]"),
    choice: root.querySelector("[data-offer-choice]"),
    listSum: root.querySelector("[data-offer-listsum]"),
    form: root.querySelector("[data-offer-form]"),
    status: root.querySelector("[data-offer-status]"),
    search: root.querySelector("[data-offer-search]"),
    searchHint: root.querySelector("[data-offer-search-hint]"),
    dialog: document.querySelector("[data-offer-dialog]"),
    detail: document.querySelector("[data-offer-detail]"),
    kontakt: document.getElementById("kontakt"),
    toggleConfig: root.querySelector("[data-offer-toggle-config]"),
    setup: root.querySelector("[data-offer-setup]"),
    setupCheck: root.querySelector("[data-offer-setup-check]"),
  };

  function escapeHtml(s) {
    return String(s == null ? "" : s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function moduleById(id) {
    return MODULES.find((m) => m.id === id) || null;
  }

  function expandDeps(ids) {
    const set = new Set(ids);
    let changed = true;
    while (changed) {
      changed = false;
      for (const id of [...set]) {
        const m = moduleById(id);
        (m && m.requires ? m.requires : []).forEach((r) => {
          if (!set.has(r)) {
            set.add(r);
            changed = true;
          }
        });
      }
    }
    return MODULES.map((m) => m.id).filter((id) => set.has(id));
  }

  function pruneMissingDeps(ids) {
    const set = new Set(ids);
    let changed = true;
    while (changed) {
      changed = false;
      for (const m of MODULES) {
        if (!set.has(m.id)) continue;
        const req = m.requires || [];
        if (req.some((r) => !set.has(r))) {
          set.delete(m.id);
          changed = true;
        }
      }
    }
    return MODULES.map((m) => m.id).filter((id) => set.has(id));
  }

  function listSum(ids) {
    return ids.reduce((sum, id) => {
      const m = moduleById(id);
      return sum + (m && m.priceEuro ? m.priceEuro : 0);
    }, 0);
  }

  function sameSet(a, b) {
    if (a.length !== b.length) return false;
    const s = new Set(a);
    return b.every((x) => s.has(x));
  }

  function matchExactBundle(ids) {
    return BUNDLES.find((b) => sameSet(ids, b.packageIds)) || null;
  }

  function nearestBundle(ids) {
    let best = null;
    let bestScore = Infinity;
    let bestSize = -1;
    const sel = new Set(ids);
    for (const b of BUNDLES) {
      const missing = b.packageIds.filter((id) => !sel.has(id)).length;
      const extra = ids.filter((id) => b.packageIds.indexOf(id) < 0).length;
      const score = missing * 2 + extra;
      const size = b.packageIds.length;
      if (score < bestScore || (score === bestScore && size > bestSize)) {
        bestScore = score;
        bestSize = size;
        best = b;
      }
    }
    return best;
  }

  function recommend(ids) {
    const selected = expandDeps(ids);
    const sum = listSum(selected);
    if (!selected.length) {
      return {
        mode: "custom",
        bundle: null,
        listSum: 0,
        offerEuro: null,
        tip: "Wählen Sie oben ein Bundle – oder öffnen Sie Bausteine zum Feintuning.",
      };
    }
    const exact = matchExactBundle(selected);
    if (exact) {
      const save = Math.max(0, sum - (exact.priceEuro || 0));
      return {
        mode: "bundle",
        bundle: exact,
        listSum: sum,
        offerEuro: exact.priceEuro,
        tip: save
          ? "Gegenüber Einzelmodulen ca. " + formatEuro(save) + " günstiger."
          : "Schulpauschale statt Einzelmodule.",
        saveEuro: save,
      };
    }
    const near = nearestBundle(selected);
    if (!near) {
      return {
        mode: "custom",
        bundle: null,
        listSum: sum,
        offerEuro: sum,
        tip: "Individuelle Zusammenstellung – wir klären die Schulpauschale per E-Mail.",
      };
    }
    const missing = near.packageIds.filter((id) => selected.indexOf(id) < 0);
    const extra = selected.filter((id) => near.packageIds.indexOf(id) < 0);
    if (missing.length && missing.length <= 3) {
      return {
        mode: "custom",
        bundle: null,
        listSum: sum,
        offerEuro: sum,
        tip:
          "Oft sinnvoll: „" +
          near.label +
          "“ (" +
          near.priceLabel +
          "). Es fehlt noch: " +
          missing.map((id) => (moduleById(id) || {}).label || id).join(", ") +
          ".",
      };
    }
    if (!missing.length && extra.length) {
      const extraSum = listSum(extra);
      return {
        mode: "custom",
        bundle: null,
        listSum: sum,
        offerEuro: sum,
        tip:
          "Kern passt zu „" +
          near.label +
          "“. Extra: " +
          extra.map((id) => (moduleById(id) || {}).label || id).join(", ") +
          (extraSum ? " (ca. +" + formatEuro(extraSum) + ")." : "."),
      };
    }
    return {
      mode: "custom",
      bundle: null,
      listSum: sum,
      offerEuro: sum,
      tip:
        "Individuell (ca. " +
        formatEuro(sum) +
        "). Nächstes Bundle: „" +
        near.label +
        "“.",
    };
  }

  function formatEuro(n) {
    if (n == null || !Number.isFinite(n)) return "–";
    return new Intl.NumberFormat("de-AT").format(n) + " €";
  }

  function normalizeSearch(q) {
    return String(q || "")
      .trim()
      .toLowerCase()
      .replace(/\s+/g, " ");
  }

  function moduleMatchesSearch(m, q) {
    if (!q) return true;
    const hay = String(m.searchText || m.label || "").toLowerCase();
    if (hay.indexOf(q) >= 0) return true;
    return q.split(" ").filter(Boolean).every((p) => hay.indexOf(p) >= 0);
  }

  function findModulesForQuery(q) {
    const nq = normalizeSearch(q);
    if (!nq) return [];
    return MODULES.filter((m) => moduleMatchesSearch(m, nq));
  }

  function setConfigOpen(open) {
    state.configOpen = !!open;
    if (els.config) {
      els.config.hidden = !state.configOpen;
    }
    if (els.toggleConfig) {
      els.toggleConfig.textContent = state.configOpen
        ? "Bausteine ausblenden"
        : "Bausteine anzeigen";
    }
    document.querySelectorAll('[data-offer-pick="custom"]').forEach((btn) => {
      btn.classList.toggle("is-active", state.configOpen);
      btn.setAttribute("aria-pressed", state.configOpen ? "true" : "false");
    });
  }

  function scrollToKontakt() {
    if (els.kontakt) {
      els.kontakt.scrollIntoView({ behavior: "smooth", block: "start" });
    }
  }

  function setSelected(ids, bundleId, opts) {
    state.selected = new Set(expandDeps(ids));
    state.activeBundle = bundleId || null;
    syncUi();
    if (opts && opts.scrollKontakt) scrollToKontakt();
  }

  function syncCards() {
    document.querySelectorAll("[data-offer-pick]").forEach((btn) => {
      const id = btn.getAttribute("data-offer-pick");
      if (id === "custom") return;
      const on = state.activeBundle === id;
      btn.classList.toggle("is-active", Boolean(on));
      btn.setAttribute("aria-pressed", on ? "true" : "false");
      const card = btn.closest(".pricing-card");
      if (card) card.classList.toggle("is-selected", Boolean(on));
    });
    const bar = document.querySelector(".offer-custom-bar");
    if (bar) {
      bar.classList.toggle(
        "is-active",
        state.configOpen || state.activeBundle === "custom"
      );
    }
  }

  function closeDialog() {
    if (els.dialog && typeof els.dialog.close === "function") {
      els.dialog.close();
    } else if (els.dialog) {
      els.dialog.removeAttribute("open");
    }
  }

  function openDetail(m) {
    if (!m || !els.detail) return;
    const primary = (m.tools || []).filter((t) => !t.setup);
    const setup = (m.tools || []).filter((t) => t.setup);
    const li = (tools) =>
      tools
        .map(
          (t) =>
            "<li><strong>" +
            escapeHtml(t.label) +
            "</strong>" +
            (t.blurb ? "<span>" + escapeHtml(t.blurb) + "</span>" : "") +
            "</li>"
        )
        .join("");

    els.detail.innerHTML =
      '<div class="offer-dialog__head">' +
      "<div>" +
      '<p class="offer-dialog__eyebrow">Baustein</p>' +
      '<h4 id="offer-dialog-title">' +
      escapeHtml(m.label) +
      "</h4>" +
      '<p class="offer-dialog__blurb">' +
      escapeHtml(m.blurb || "") +
      "</p>" +
      "</div>" +
      '<button type="button" class="offer-dialog__close" data-offer-dialog-close aria-label="Schließen">×</button>' +
      "</div>" +
      (primary.length
        ? '<p class="offer-dialog__section">Werkzeuge für den Alltag</p><ul class="offer-dialog__list">' +
          li(primary) +
          "</ul>"
        : "") +
      (setup.length
        ? '<p class="offer-dialog__section">IT-Einrichtung</p><ul class="offer-dialog__list offer-dialog__list--muted">' +
          li(setup) +
          "</ul>"
        : "") +
      '<div class="offer-dialog__actions">' +
      '<p class="offer-dialog__foot">' +
      escapeHtml(formatEuro(m.priceEuro)) +
      " / Jahr</p>" +
      '<button type="button" class="btn btn-primary btn-sm" data-offer-dialog-add="' +
      escapeHtml(m.id) +
      '">' +
      (state.selected.has(m.id) ? "Bereits gewählt" : "Baustein wählen") +
      "</button>" +
      "</div>";

    els.detail.querySelector("[data-offer-dialog-close]")?.addEventListener("click", closeDialog);
    const addBtn = els.detail.querySelector("[data-offer-dialog-add]");
    if (addBtn && !state.selected.has(m.id)) {
      addBtn.addEventListener("click", () => {
        const next = expandDeps([...state.selected, m.id]);
        const exact = matchExactBundle(next);
        state.selected = new Set(next);
        state.activeBundle = exact ? exact.id : "custom";
        closeDialog();
        syncUi();
      });
    }

    if (els.dialog) {
      if (typeof els.dialog.showModal === "function") els.dialog.showModal();
      else els.dialog.setAttribute("open", "");
    }
  }

  function renderModules() {
    if (!els.moduleGrid) return;
    const q = normalizeSearch(state.search);
    const matched = q ? findModulesForQuery(q) : MODULES;
    const matchedIds = new Set(matched.map((m) => m.id));

    if (els.searchHint) {
      if (!q) {
        els.searchHint.textContent = "Tipp: „Schulwahl“ → Schul-Apps. Details öffnen im Fenster.";
      } else if (!matched.length) {
        els.searchHint.textContent = "Kein Treffer zu „" + q + "“.";
      } else {
        els.searchHint.textContent =
          matched.length +
          " Treffer" +
          (matched.length === 1 ? ": " + matched[0].label : "") +
          ".";
      }
    }

    els.moduleGrid.innerHTML = MODULES.map((m) => {
      const checked = state.selected.has(m.id);
      const hit = !q || matchedIds.has(m.id);
      const chips = (m.highlights || [])
        .slice(0, 2)
        .map((h) => '<span class="offer-module__chip">' + escapeHtml(h) + "</span>")
        .join("");
      const n = (m.tools || []).filter((t) => !t.setup).length || (m.tools || []).length;
      const toolLabel =
        n === 1 ? "1 Werkzeug · Details" : n + " Werkzeuge · Details";
      return (
        '<div class="offer-module' +
        (checked ? " is-on" : "") +
        (hit ? "" : " is-dim") +
        (q && hit ? " is-hit" : "") +
        '">' +
        '<label class="offer-module__check">' +
        '<input type="checkbox" value="' +
        escapeHtml(m.id) +
        '" ' +
        (checked ? "checked " : "") +
        ' aria-label="' +
        escapeHtml(m.label) +
        '"/>' +
        "</label>" +
        '<div class="offer-module__body">' +
        '<span class="offer-module__title">' +
        escapeHtml(m.label) +
        "</span>" +
        '<span class="offer-module__blurb">' +
        escapeHtml(m.blurb || "") +
        "</span>" +
        (chips ? '<span class="offer-module__chips">' + chips + "</span>" : "") +
        '<span class="offer-module__meta">' +
        '<span class="offer-module__price">' +
        escapeHtml(formatEuro(m.priceEuro)) +
        "</span>" +
        '<button type="button" class="offer-module__more" data-offer-detail="' +
        escapeHtml(m.id) +
        '">' +
        toolLabel +
        "</button>" +
        "</span></div></div>"
      );
    }).join("");

    els.moduleGrid.querySelectorAll('input[type="checkbox"]').forEach((input) => {
      input.addEventListener("change", () => {
        const explicit = new Set();
        els.moduleGrid.querySelectorAll('input[type="checkbox"]').forEach((el) => {
          if (el.checked) explicit.add(el.value);
        });
        const finalIds = input.checked
          ? expandDeps([...explicit])
          : pruneMissingDeps([...explicit]);
        const exact = matchExactBundle(finalIds);
        state.selected = new Set(finalIds);
        state.activeBundle = exact ? exact.id : finalIds.length ? "custom" : null;
        syncUi();
      });
    });

    els.moduleGrid.querySelectorAll("[data-offer-detail]").forEach((btn) => {
      btn.addEventListener("click", (ev) => {
        ev.preventDefault();
        ev.stopPropagation();
        openDetail(moduleById(btn.getAttribute("data-offer-detail")));
      });
    });

    if (q && matched.length === 1) {
      // auto-open only on Enter handled elsewhere; keep quiet on typing
    }
  }

  function renderTip(rec) {
    if (!els.tip) return;
    if (rec.mode === "bundle" && rec.bundle) {
      const chips = (rec.bundle.packageIds || [])
        .map((id) => moduleById(id))
        .filter(Boolean)
        .map(
          (m) =>
            '<span class="offer-summary__chip">' + escapeHtml(m.label) + "</span>"
        )
        .join("");
      const save =
        rec.saveEuro > 0
          ? '<span class="offer-summary__save">−' +
            escapeHtml(formatEuro(rec.saveEuro)) +
            "</span>"
          : "";
      els.tip.innerHTML =
        '<div class="offer-summary__visual">' +
        save +
        '<span class="offer-summary__visual-label">gegenüber Einzelmodulen</span>' +
        "</div>" +
        (chips ? '<div class="offer-summary__chips">' + chips + "</div>" : "");
      return;
    }
    els.tip.textContent = rec.tip || "";
  }

  function syncSummary() {
    const ids = [...state.selected];
    const rec = recommend(ids);
    renderTip(rec);
    if (els.choice) {
      els.choice.textContent =
        rec.mode === "bundle" && rec.bundle
          ? "Bundle · " + rec.bundle.label
          : ids.length
            ? "Individuell"
            : "Noch keine Auswahl";
    }
    if (els.price) {
      els.price.textContent =
        rec.mode === "bundle" && rec.bundle
          ? rec.bundle.priceLabel
          : rec.offerEuro
            ? formatEuro(rec.offerEuro) + " / Schuljahr (Listenpreis)"
            : "–";
    }
    if (els.listSum) {
      els.listSum.textContent = ids.length
        ? "Einzelmodule: " + formatEuro(rec.listSum)
        : "Einzelmodule: –";
    }
    if (els.setup) els.setup.classList.toggle("is-on", state.setupOn);
    if (els.setupCheck) els.setupCheck.checked = state.setupOn;
    syncCards();
  }

  function syncUi() {
    if (state.configOpen) renderModules();
    syncSummary();
  }

  function buildMailto(form) {
    const fd = new FormData(form);
    const name = String(fd.get("name") || "").trim();
    const school = String(fd.get("school") || "").trim();
    const email = String(fd.get("email") || "").trim();
    const message = String(fd.get("message") || "").trim();
    const ids = [...state.selected];
    const rec = recommend(ids);
    const choiceLabel =
      rec.mode === "bundle" && rec.bundle
        ? "Bundle: " + rec.bundle.label + " (" + rec.bundle.priceLabel + ")"
        : "Individuell (Listenpreis " + (rec.listSum || 0) + " € / Schuljahr)";

    const moduleLines = ids.length
      ? ids
          .map((id) => {
            const m = moduleById(id);
            return (
              "- " +
              (m ? m.label : id) +
              " (" +
              (m ? m.priceEuro + " €" : "–") +
              ")"
            );
          })
          .join("\n")
      : "- (keine)";

    const extraLine = state.setupOn
      ? "- " + SETUP.label + " (" + SETUP.priceLabel + ")"
      : "- keines";

    const subject =
      "Angebotsanfrage MS365-Schul-Tools – " +
      (rec.mode === "bundle" && rec.bundle ? rec.bundle.label : "Individuell");

    const body = [
      "Guten Tag,",
      "",
      "ich interessiere mich für MS365-Schul-Tools.",
      "",
      "Schule: " + (school || "–"),
      "Name: " + (name || "–"),
      "E-Mail: " + (email || "–"),
      "",
      "Auswahl: " + choiceLabel,
      "",
      "Bausteine:",
      moduleLines,
      "",
      "Extra:",
      extraLine,
      "",
      "Hinweis Homepage:",
      rec.tip || "–",
      "",
      "Nachricht:",
      message || "–",
      "",
      "Mit freundlichen Grüßen",
      name || "",
    ].join("\n");

    return (
      "mailto:" +
      CONTACT_MAIL +
      "?subject=" +
      encodeURIComponent(subject) +
      "&body=" +
      encodeURIComponent(body)
    );
  }

  document.querySelectorAll("[data-offer-pick]").forEach((btn) => {
    btn.addEventListener("click", () => {
      const id = btn.getAttribute("data-offer-pick");
      if (id === "custom") {
        setConfigOpen(true);
        state.activeBundle = state.activeBundle || "custom";
        syncUi();
        els.config?.scrollIntoView({ behavior: "smooth", block: "start" });
        els.search?.focus();
        return;
      }
      const bundle = BUNDLES.find((b) => b.id === id);
      if (!bundle) return;
      setConfigOpen(false);
      setSelected(bundle.packageIds, bundle.id, { scrollKontakt: true });
    });
  });

  if (els.toggleConfig) {
    els.toggleConfig.addEventListener("click", () => {
      setConfigOpen(!state.configOpen);
      if (state.configOpen) {
        syncUi();
        els.config?.scrollIntoView({ behavior: "smooth", block: "start" });
      }
    });
  }

  if (els.setupCheck) {
    els.setupCheck.addEventListener("change", () => {
      state.setupOn = !!els.setupCheck.checked;
      syncSummary();
    });
  }

  if (els.dialog) {
    els.dialog.addEventListener("click", (ev) => {
      if (ev.target === els.dialog) closeDialog();
    });
    els.dialog.addEventListener("cancel", (ev) => {
      ev.preventDefault();
      closeDialog();
    });
  }

  if (els.search) {
    els.search.addEventListener("input", () => {
      state.search = els.search.value;
      renderModules();
    });
    els.search.addEventListener("keydown", (ev) => {
      if (ev.key !== "Enter") return;
      ev.preventDefault();
      const hits = findModulesForQuery(state.search);
      if (hits.length === 1) openDetail(hits[0]);
    });
  }

  if (els.form) {
    els.form.addEventListener("submit", (ev) => {
      ev.preventDefault();
      const href = buildMailto(els.form);
      if (els.status) {
        els.status.textContent =
          "E-Mail-Programm öffnet sich mit Anfrage an " + CONTACT_MAIL + ".";
      }
      window.location.href = href;
    });
  }

  setConfigOpen(false);
  const featured = BUNDLES.find((b) => b.featured) || BUNDLES[0];
  if (featured) setSelected(featured.packageIds, featured.id);
  else syncUi();
})();
