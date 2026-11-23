(() => {
  const reduceMotion = window.matchMedia("(prefers-reduced-motion: reduce)").matches;
  const nav = document.querySelector("[data-nav]");
  const year = document.querySelector("[data-year]");
  if (year) year.textContent = String(new Date().getFullYear());

  const heroVisual = document.querySelector(".hero-visual");
  const heroShot = document.querySelector(".hero-visual .device-shot");
  const heroFallback = document.querySelector(".hero-visual .device-fallback");
  if (heroShot && heroVisual) {
    const showFallback = () => {
      heroVisual.classList.add("is-fallback");
      if (heroFallback) heroFallback.hidden = false;
      heroShot.remove();
    };
    heroShot.addEventListener("error", showFallback);
    if (heroShot.complete && heroShot.naturalWidth === 0) showFallback();
  }

  const onScroll = () => {
    if (!nav) return;
    nav.classList.toggle("is-scrolled", window.scrollY > 40);
  };
  onScroll();
  window.addEventListener("scroll", onScroll, { passive: true });

  const atmosphere = document.querySelector(".hero-atmosphere");
  if (atmosphere && !reduceMotion) {
    let ticking = false;
    window.addEventListener(
      "scroll",
      () => {
        if (ticking) return;
        ticking = true;
        requestAnimationFrame(() => {
          const y = Math.min(window.scrollY, 480);
          atmosphere.style.transform = `translate3d(0, ${y * 0.18}px, 0) scale(${1 + y * 0.00015})`;
          ticking = false;
        });
      },
      { passive: true }
    );
  }

  const reveals = document.querySelectorAll(".reveal");
  if (reveals.length) {
    if (reduceMotion || !("IntersectionObserver" in window)) {
      reveals.forEach((el) => el.classList.add("is-visible"));
    } else {
      const io = new IntersectionObserver(
        (entries) => {
          entries.forEach((entry) => {
            if (!entry.isIntersecting) return;
            entry.target.classList.add("is-visible");
            io.unobserve(entry.target);
          });
        },
        { rootMargin: "0px 0px -8% 0px", threshold: 0.12 }
      );
      reveals.forEach((el) => io.observe(el));
    }
  }

  /* Persona-Showcase */
  const showcase = document.querySelector("[data-showcase]");
  const catalog = window.MS365_SCREENS_CATALOG;
  if (showcase && catalog && Array.isArray(catalog.personas) && catalog.personas.length) {
    const frame = showcase.querySelector(".screen-main .device-frame");
    const img = showcase.querySelector("#screen-main-img");
    const cap = showcase.querySelector("#screen-main-cap");
    const urlEl = showcase.querySelector("#screen-main-url");
    const personaHost = showcase.querySelector("[data-screen-personas]");
    const thumbsHost = showcase.querySelector("[data-screen-thumbs]");
    const leadEl = showcase.querySelector("[data-screen-persona-lead]");
    const cacheV =
      (img && (img.getAttribute("src") || "").match(/[?&]v=([^&]+)/) || [])[1] ||
      "20261121";

    let personaIdx = Math.max(
      0,
      catalog.personas.findIndex((p) => p.id === "lehrer")
    );
    if (personaIdx < 0) personaIdx = 0;
    let screenIdx = 0;

    function screenSrc(file) {
      return "assets/screens/" + file + "?v=" + cacheV;
    }

    function showScreen(persona, screen, animate) {
      if (!screen) return;
      const src = screenSrc(screen.file);
      const label = screen.label || "";
      const caption = screen.caption || "";
      const urlLabel = screen.urlLabel || label;

      const apply = () => {
        if (img) {
          img.src = src;
          img.alt = label ? "Ansicht: " + label : img.alt;
          img.onerror = () => {
            const ph = document.createElement("div");
            ph.className = "shot-placeholder";
            ph.innerHTML =
              "<p>" +
              (label || "App") +
              '</p><span>Screenshot folgt</span>';
            img.replaceWith(ph);
          };
        }
        if (cap) cap.textContent = caption;
        if (urlEl) urlEl.textContent = urlLabel;
        if (img) img.classList.remove("is-fading");
        if (frame) frame.classList.remove("is-switching");
      };

      if (!animate || reduceMotion || !img) {
        apply();
        return;
      }
      img.classList.add("is-fading");
      if (frame) frame.classList.add("is-switching");
      window.setTimeout(apply, 180);
    }

    function renderThumbs(persona) {
      if (!thumbsHost) return;
      thumbsHost.innerHTML = "";
      (persona.screens || []).forEach((screen, i) => {
        const btn = document.createElement("button");
        btn.type = "button";
        btn.setAttribute("role", "tab");
        btn.setAttribute("aria-selected", i === screenIdx ? "true" : "false");
        if (i === screenIdx) btn.classList.add("is-active");
        btn.innerHTML = "<span>" + screen.label + "</span>";
        btn.addEventListener("click", () => {
          screenIdx = i;
          renderThumbs(persona);
          showScreen(persona, screen, true);
        });
        thumbsHost.appendChild(btn);
      });
    }

    function selectPersona(idx, animate) {
      personaIdx = idx;
      screenIdx = 0;
      const persona = catalog.personas[personaIdx];
      if (!persona) return;
      if (personaHost) {
        personaHost.querySelectorAll("button").forEach((b, i) => {
          const on = i === personaIdx;
          b.classList.toggle("is-active", on);
          b.setAttribute("aria-selected", on ? "true" : "false");
        });
      }
      if (leadEl) leadEl.textContent = persona.lead || "";
      renderThumbs(persona);
      showScreen(persona, persona.screens[0], animate);
    }

    if (personaHost) {
      catalog.personas.forEach((persona, i) => {
        const btn = document.createElement("button");
        btn.type = "button";
        btn.setAttribute("role", "tab");
        btn.setAttribute("aria-selected", i === personaIdx ? "true" : "false");
        if (i === personaIdx) btn.classList.add("is-active");
        btn.textContent = persona.label;
        btn.addEventListener("click", () => selectPersona(i, true));
        personaHost.appendChild(btn);
      });
    }

    selectPersona(personaIdx, false);
  }

  document.querySelectorAll(".btn").forEach((btn) => {
    btn.addEventListener("pointerdown", () => btn.classList.add("is-pressed"));
    ["pointerup", "pointerleave", "pointercancel"].forEach((ev) => {
      btn.addEventListener(ev, () => btn.classList.remove("is-pressed"));
    });
  });
})();
