/**
 * Product-Tour UI: Willkommens-Modal (bestehende Onboarding-Klassen) + Spotlight-Coachmarks.
 */
import { clampTourStep, isLastTourStep } from './product-tour-logic.js';

function escapeHtml(s) {
    return String(s ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function preferReducedMotion() {
    try {
        return !!(window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches);
    } catch {
        return false;
    }
}

/**
 * @typedef {{
 *   id?: string,
 *   title: string,
 *   bodyHtml: string,
 *   primaryLabel?: string
 * }} WelcomePanel
 */

/**
 * @typedef {{
 *   id?: string,
 *   selector: string,
 *   title: string,
 *   bodyHtml: string,
 *   placement?: 'auto'|'top'|'bottom'|'left'|'right',
 *   padding?: number,
 *   primaryLabel?: string,
 *   before?: () => void | Promise<void>,
 *   onPrimary?: () => void | Promise<void> | 'next' | 'finish'
 * }} SpotlightStep
 */

/**
 * @param {WelcomePanel[]} panels
 * @param {{ onComplete?: (reason: 'done'|'skipped') => void, title?: string, lead?: string }} [opts]
 */
export function openWelcomeTour(panels, opts) {
    const list = Array.isArray(panels) ? panels.filter(Boolean) : [];
    if (!list.length) {
        if (opts && typeof opts.onComplete === 'function') opts.onComplete('done');
        return { close() {} };
    }

    const existing = document.getElementById('ms365OnboardingOverlay');
    if (existing) existing.remove();

    let index = 0;
    const overlay = document.createElement('div');
    overlay.id = 'ms365OnboardingOverlay';
    overlay.className = 'ms365-onboarding-overlay is-open';
    overlay.setAttribute('role', 'dialog');
    overlay.setAttribute('aria-modal', 'true');
    overlay.setAttribute('aria-labelledby', 'ms365OnboardingTitle');

    const title = String((opts && opts.title) || 'Willkommen bei MS365-Schul-Tools');
    const lead = String(
        (opts && opts.lead) || 'Kurzer Einstieg für Schul-IT – danach starten Sie mit den Stammdaten.'
    );

    overlay.innerHTML =
        '<div class="ms365-onboarding-card">' +
        '<div class="ms365-onboarding-header">' +
        '<h2 id="ms365OnboardingTitle">' +
        escapeHtml(title) +
        '</h2>' +
        '<p>' +
        escapeHtml(lead) +
        '</p>' +
        '</div>' +
        '<div class="ms365-onboarding-steps" id="ms365OnboardingDots" aria-hidden="true"></div>' +
        '<div class="ms365-onboarding-body" id="ms365OnboardingBody"></div>' +
        '<div class="ms365-onboarding-footer">' +
        '<button type="button" class="ms365-onboarding-skip" id="ms365OnboardingSkip">Später</button>' +
        '<div class="ms365-onboarding-actions">' +
        '<button type="button" class="btn alt" id="ms365OnboardingBack" hidden>Zurück</button>' +
        '<button type="button" class="btn" id="ms365OnboardingNext">Weiter</button>' +
        '</div></div></div>';

    document.body.appendChild(overlay);

    const dots = overlay.querySelector('#ms365OnboardingDots');
    const body = overlay.querySelector('#ms365OnboardingBody');
    const btnBack = /** @type {HTMLButtonElement|null} */ (overlay.querySelector('#ms365OnboardingBack'));
    const btnNext = /** @type {HTMLButtonElement|null} */ (overlay.querySelector('#ms365OnboardingNext'));
    const btnSkip = overlay.querySelector('#ms365OnboardingSkip');

    function finish(reason) {
        overlay.classList.remove('is-open');
        overlay.remove();
        document.removeEventListener('keydown', onKey);
        if (opts && typeof opts.onComplete === 'function') opts.onComplete(reason);
    }

    function render() {
        index = clampTourStep(index, list.length);
        if (dots) {
            dots.innerHTML = list
                .map((_, i) => {
                    const cls =
                        'ms365-onboarding-step-dot' +
                        (i === index ? ' is-active' : '') +
                        (i < index ? ' is-done' : '');
                    return '<span class="' + cls + '"></span>';
                })
                .join('');
        }
        const panel = list[index];
        if (body && panel) {
            body.innerHTML =
                '<div class="ms365-onboarding-panel">' +
                '<h3>' +
                escapeHtml(panel.title) +
                '</h3>' +
                panel.bodyHtml +
                '</div>';
        }
        if (btnBack) btnBack.hidden = index <= 0;
        if (btnNext) {
            const last = isLastTourStep(index, list.length);
            btnNext.textContent = last
                ? String(panel.primaryLabel || 'Tour starten')
                : String(panel.primaryLabel || 'Weiter');
        }
        try {
            (btnNext || btnSkip)?.focus();
        } catch {
            /* ignore */
        }
    }

    function onKey(ev) {
        if (ev.key === 'Escape') {
            ev.preventDefault();
            finish('skipped');
        } else if (ev.key === 'ArrowRight') {
            ev.preventDefault();
            if (isLastTourStep(index, list.length)) finish('done');
            else {
                index += 1;
                render();
            }
        } else if (ev.key === 'ArrowLeft' && index > 0) {
            ev.preventDefault();
            index -= 1;
            render();
        }
    }

    btnSkip?.addEventListener('click', () => finish('skipped'));
    btnBack?.addEventListener('click', () => {
        index = Math.max(0, index - 1);
        render();
    });
    btnNext?.addEventListener('click', () => {
        if (isLastTourStep(index, list.length)) finish('done');
        else {
            index += 1;
            render();
        }
    });
    document.addEventListener('keydown', onKey);
    render();

    return { close: () => finish('skipped') };
}

/**
 * @param {HTMLElement} el
 * @param {number} padding
 */
function rectWithPadding(el, padding) {
    const r = el.getBoundingClientRect();
    const pad = Math.max(0, Number(padding) || 0);
    return {
        top: Math.max(8, r.top - pad),
        left: Math.max(8, r.left - pad),
        width: Math.min(window.innerWidth - 16, r.width + pad * 2),
        height: Math.min(window.innerHeight - 16, r.height + pad * 2)
    };
}

/**
 * @param {{ top: number, left: number, width: number, height: number }} hole
 * @param {DOMRect} tipRect
 * @param {'auto'|'top'|'bottom'|'left'|'right'} placement
 */
function placeTooltip(hole, tipRect, placement) {
    const gap = 14;
    const vw = window.innerWidth;
    const vh = window.innerHeight;
    let top = hole.top + hole.height + gap;
    let left = hole.left;

    const tryBottom = () => {
        top = hole.top + hole.height + gap;
        left = hole.left + hole.width / 2 - tipRect.width / 2;
    };
    const tryTop = () => {
        top = hole.top - tipRect.height - gap;
        left = hole.left + hole.width / 2 - tipRect.width / 2;
    };
    const tryRight = () => {
        top = hole.top + hole.height / 2 - tipRect.height / 2;
        left = hole.left + hole.width + gap;
    };
    const tryLeft = () => {
        top = hole.top + hole.height / 2 - tipRect.height / 2;
        left = hole.left - tipRect.width - gap;
    };

    const pref = placement || 'auto';
    if (pref === 'top') tryTop();
    else if (pref === 'left') tryLeft();
    else if (pref === 'right') tryRight();
    else if (pref === 'bottom') tryBottom();
    else {
        const spaceBelow = vh - (hole.top + hole.height);
        if (spaceBelow >= tipRect.height + gap + 20) tryBottom();
        else if (hole.top >= tipRect.height + gap + 20) tryTop();
        else tryBottom();
    }

    left = Math.min(Math.max(10, left), vw - tipRect.width - 10);
    top = Math.min(Math.max(10, top), vh - tipRect.height - 10);
    return { top, left };
}

/**
 * @param {SpotlightStep[]} steps
 * @param {{ onComplete?: (reason: 'done'|'skipped') => void, skipLabel?: string }} [opts]
 */
export function startSpotlightTour(steps, opts) {
    const list = Array.isArray(steps) ? steps.filter((s) => s && s.selector) : [];
    if (!list.length) {
        if (opts && typeof opts.onComplete === 'function') opts.onComplete('done');
        return { close() {} };
    }

    document.getElementById('ms365ProductTourRoot')?.remove();

    let index = 0;
    let closed = false;
    const root = document.createElement('div');
    root.id = 'ms365ProductTourRoot';
    root.className = 'ms365-product-tour';
    root.innerHTML =
        '<div class="ms365-product-tour__mask" aria-hidden="true">' +
        '<div class="ms365-product-tour__hole"></div>' +
        '</div>' +
        '<div class="ms365-product-tour__tip" role="dialog" aria-modal="true" aria-labelledby="ms365ProductTourTitle">' +
        '<p class="ms365-product-tour__progress" id="ms365ProductTourProgress"></p>' +
        '<h3 id="ms365ProductTourTitle"></h3>' +
        '<div class="ms365-product-tour__body" id="ms365ProductTourBody"></div>' +
        '<div class="ms365-product-tour__actions">' +
        '<button type="button" class="ms365-product-tour__skip" id="ms365ProductTourSkip">' +
        escapeHtml((opts && opts.skipLabel) || 'Tour beenden') +
        '</button>' +
        '<div class="ms365-product-tour__nav">' +
        '<button type="button" class="btn alt" id="ms365ProductTourBack" hidden>Zurück</button>' +
        '<button type="button" class="btn" id="ms365ProductTourNext">Weiter</button>' +
        '</div></div></div>';

    document.body.appendChild(root);
    document.documentElement.classList.add('ms365-product-tour-active');

    const holeEl = /** @type {HTMLElement} */ (root.querySelector('.ms365-product-tour__hole'));
    const tipEl = /** @type {HTMLElement} */ (root.querySelector('.ms365-product-tour__tip'));
    const titleEl = root.querySelector('#ms365ProductTourTitle');
    const bodyEl = root.querySelector('#ms365ProductTourBody');
    const progEl = root.querySelector('#ms365ProductTourProgress');
    const btnBack = /** @type {HTMLButtonElement} */ (root.querySelector('#ms365ProductTourBack'));
    const btnNext = /** @type {HTMLButtonElement} */ (root.querySelector('#ms365ProductTourNext'));
    const btnSkip = root.querySelector('#ms365ProductTourSkip');

    function finish(reason) {
        if (closed) return;
        closed = true;
        window.removeEventListener('resize', onReposition);
        window.removeEventListener('scroll', onReposition, true);
        document.removeEventListener('keydown', onKey);
        document.documentElement.classList.remove('ms365-product-tour-active');
        root.remove();
        if (opts && typeof opts.onComplete === 'function') opts.onComplete(reason);
    }

    function resolveTarget(selector) {
        try {
            return document.querySelector(selector);
        } catch {
            return null;
        }
    }

    function layout() {
        const step = list[index];
        if (!step) return;
        const target = resolveTarget(step.selector);
        if (!target) {
            tipEl.classList.add('is-orphan');
            holeEl.style.opacity = '0';
            if (titleEl) titleEl.textContent = step.title;
            if (bodyEl) bodyEl.innerHTML = step.bodyHtml;
            if (progEl) progEl.textContent = index + 1 + ' / ' + list.length;
            tipEl.style.top = '20%';
            tipEl.style.left = '50%';
            tipEl.style.transform = 'translateX(-50%)';
            return;
        }
        tipEl.classList.remove('is-orphan');
        tipEl.style.transform = '';
        try {
            target.scrollIntoView({
                block: 'center',
                inline: 'nearest',
                behavior: preferReducedMotion() ? 'auto' : 'smooth'
            });
        } catch {
            /* ignore */
        }
        const hole = rectWithPadding(/** @type {HTMLElement} */ (target), step.padding == null ? 8 : step.padding);
        holeEl.style.opacity = '1';
        holeEl.style.top = hole.top + 'px';
        holeEl.style.left = hole.left + 'px';
        holeEl.style.width = hole.width + 'px';
        holeEl.style.height = hole.height + 'px';

        if (titleEl) titleEl.textContent = step.title;
        if (bodyEl) bodyEl.innerHTML = step.bodyHtml;
        if (progEl) progEl.textContent = index + 1 + ' / ' + list.length;

        const tipRect = tipEl.getBoundingClientRect();
        const pos = placeTooltip(hole, tipRect, step.placement || 'auto');
        tipEl.style.top = pos.top + 'px';
        tipEl.style.left = pos.left + 'px';
    }

    async function show(i) {
        index = clampTourStep(i, list.length);
        const step = list[index];
        if (step && typeof step.before === 'function') {
            try {
                await step.before();
            } catch {
                /* ignore */
            }
        }
        const last = isLastTourStep(index, list.length);
        btnBack.hidden = index <= 0;
        btnNext.textContent = last
            ? String((step && step.primaryLabel) || 'Fertig')
            : String((step && step.primaryLabel) || 'Weiter');
        layout();
        // zweites Layout nach Scroll/Fonts
        window.setTimeout(layout, preferReducedMotion() ? 40 : 280);
        try {
            btnNext.focus();
        } catch {
            /* ignore */
        }
    }

    async function goNext() {
        const step = list[index];
        if (step && typeof step.onPrimary === 'function') {
            try {
                const r = await step.onPrimary();
                if (r === 'finish') {
                    finish('done');
                    return;
                }
            } catch {
                /* ignore */
            }
        }
        if (isLastTourStep(index, list.length)) finish('done');
        else await show(index + 1);
    }

    function onReposition() {
        if (!closed) layout();
    }

    function onKey(ev) {
        if (ev.key === 'Escape') {
            ev.preventDefault();
            finish('skipped');
        } else if (ev.key === 'ArrowRight') {
            ev.preventDefault();
            goNext();
        } else if (ev.key === 'ArrowLeft' && index > 0) {
            ev.preventDefault();
            show(index - 1);
        }
    }

    btnSkip?.addEventListener('click', () => finish('skipped'));
    btnBack.addEventListener('click', () => show(index - 1));
    btnNext.addEventListener('click', () => goNext());
    window.addEventListener('resize', onReposition);
    window.addEventListener('scroll', onReposition, true);
    document.addEventListener('keydown', onKey);

    show(0);

    return { close: () => finish('skipped'), next: goNext };
}

window.ms365ProductTour = {
    openWelcomeTour,
    startSpotlightTour
};
