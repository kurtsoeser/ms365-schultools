import {
    grantAccess,
    isAccessGranted,
    isPinGateEnabled,
    isValidPin,
    normalizePin,
    safeReturnPath
} from './pin-gate-core.js';

(function () {
    'use strict';

    const config = typeof window !== 'undefined' ? window.MS365_ACCESS_CONFIG : null;
    const params = new URLSearchParams(location.search);
    const returnTarget = safeReturnPath(params.get('return'));

    if (isAccessGranted()) {
        location.replace(returnTarget);
        return;
    }
    const form = document.getElementById('welcome-pin-form');
    const input = document.getElementById('welcome-pin-input');
    const errorEl = document.getElementById('welcome-pin-error');
    const submitBtn = document.getElementById('welcome-pin-submit');

    if (!isPinGateEnabled(config)) {
        const target = safeReturnPath(new URLSearchParams(location.search).get('return'));
        location.replace(target);
        return;
    }

    if (!form || !input) return;

    function showError(message) {
        if (!errorEl) return;
        errorEl.textContent = message;
        errorEl.hidden = !message;
    }

    form.addEventListener('submit', function (event) {
        event.preventDefault();
        showError('');

        const pin = normalizePin(input.value);
        if (!pin) {
            showError('Bitte geben Sie eine PIN ein.');
            input.focus();
            return;
        }

        if (!isValidPin(pin, config.pins)) {
            showError('Die PIN ist nicht gültig. Bitte erneut versuchen.');
            input.value = '';
            input.focus();
            return;
        }

        if (submitBtn) submitBtn.disabled = true;
        grantAccess();
        location.replace(returnTarget);
    });

    input.addEventListener('input', function () {
        showError('');
    });
})();
