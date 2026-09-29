/**
 * Wizard-Schritt-Navigation (Analyse 02 Phase C).
 * showStep-Logik bleibt verhaltensgleich; Host liefert Closures aus setup-wizard.js.
 *
 * @param {object} host
 * @param {(n: number) => void} host.onBeforeLeaveStep
 * @param {(step: number) => void} host.onEnterStep
 * @param {() => number} host.getPrevStep
 * @param {(n: number) => void} host.setPrevStep
 */
export function createShowStep(host) {
    return function showStep(n) {
        const step = Math.max(1, Math.min(11, parseInt(n, 10) || 1));
        const prev = typeof host.getPrevStep === 'function' ? host.getPrevStep() : 0;
        if (typeof host.onBeforeLeaveStep === 'function') {
            host.onBeforeLeaveStep(prev, step);
        }
        if (typeof host.setPrevStep === 'function') host.setPrevStep(step);

        for (let i = 1; i <= 11; i++) {
            const panel = document.getElementById('swStep' + i);
            if (panel) panel.style.display = i === step ? 'block' : 'none';
        }
        document.querySelectorAll('[data-sw-step]').forEach(function (btn) {
            const sn = parseInt(btn.getAttribute('data-sw-step'), 10);
            const on = sn === step;
            btn.classList.toggle('active', on);
            btn.setAttribute('aria-selected', on ? 'true' : 'false');
        });
        try {
            if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.touchWizardVisit === 'function') {
                window.ms365AppDataV2.touchWizardVisit(step);
            }
        } catch {
            // ignore
        }
        if (typeof host.onEnterStep === 'function') host.onEnterStep(step);
    };
}
