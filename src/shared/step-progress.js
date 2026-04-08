(function () {
    'use strict';

    /**
     * Setzt die CSS-Variable --step-progress-pct auf dem .steps-Container (Fortschrittsbalken).
     * @param {HTMLElement|null} stepsEl
     * @param {number|string} step
     * @param {Array<number|string>} stepOrder
     */
    window.ms365ApplyStepProgress = function ms365ApplyStepProgress(stepsEl, step, stepOrder) {
        if (!stepsEl || !Array.isArray(stepOrder) || !stepOrder.length) return;
        const idx = stepOrder.indexOf(step);
        if (idx < 0) return;
        const pct = Math.min(100, Math.max(6, Math.round(((idx + 1) / stepOrder.length) * 100)));
        stepsEl.style.setProperty('--step-progress-pct', pct + '%');
    };
})();
