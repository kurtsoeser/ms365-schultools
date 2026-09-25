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

  /* Soft parallax on hero atmosphere */
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

  /* Scroll reveal */
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

  /* Screenshot tabs with fade */
  const showcase = document.querySelector("[data-showcase]");
  if (showcase) {
    const frame = showcase.querySelector(".screen-main .device-frame");
    const img = showcase.querySelector("#screen-main-img");
    const cap = showcase.querySelector("#screen-main-cap");
    const urlEl = showcase.querySelector("#screen-main-url");
    const tabs = showcase.querySelectorAll(".screen-thumbs button");

    tabs.forEach((btn) => {
      btn.addEventListener("click", () => {
        tabs.forEach((b) => {
          b.classList.remove("is-active");
          b.setAttribute("aria-selected", "false");
        });
        btn.classList.add("is-active");
        btn.setAttribute("aria-selected", "true");

        const src = btn.getAttribute("data-src");
        const label = btn.getAttribute("data-label") || "";
        const caption = btn.getAttribute("data-cap") || "";
        const urlLabel = btn.getAttribute("data-url") || label;

        const apply = () => {
          if (img && src) {
            img.src = src;
            img.alt = label ? `Ansicht: ${label}` : img.alt;
            img.onerror = () => {
              const ph = document.createElement("div");
              ph.className = "shot-placeholder";
              ph.innerHTML = `<p>${label || "App"}</p><span>Screenshot folgt</span>`;
              img.replaceWith(ph);
            };
          }
          if (cap) cap.textContent = caption;
          if (urlEl) urlEl.textContent = urlLabel;
          if (img) img.classList.remove("is-fading");
          if (frame) frame.classList.remove("is-switching");
        };

        if (reduceMotion || !img) {
          apply();
          return;
        }

        img.classList.add("is-fading");
        if (frame) frame.classList.add("is-switching");
        window.setTimeout(apply, 180);
      });
    });
  }

  /* Subtle button press ripple via pointer */
  document.querySelectorAll(".btn").forEach((btn) => {
    btn.addEventListener("pointerdown", () => btn.classList.add("is-pressed"));
    ["pointerup", "pointerleave", "pointercancel"].forEach((ev) => {
      btn.addEventListener(ev, () => btn.classList.remove("is-pressed"));
    });
  });
})();
