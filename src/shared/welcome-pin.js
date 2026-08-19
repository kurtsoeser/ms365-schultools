import {
    grantAccess,
    isAccessGranted,
    grantAdminAccess,
    isAdminAccessGranted,
    isPinGateEnabled,
    isValidPin,
    normalizePin,
    resolveReturnUrl
} from './pin-gate-core.js';
import { getActiveUserAccessConfig } from './access-override-store.js';
import { loadReleaseNotes, getNewReleaseNotes, getLastSeenAt, setLastSeenAt } from './release-notes-store.js';

(function () {
    'use strict';

    const config = typeof window !== 'undefined' ? window.MS365_ACCESS_CONFIG : null;
    const params = new URLSearchParams(location.search);
    const returnTarget = resolveReturnUrl(params.get('return'), location.href);

    const requestedMode = String(params.get('mode') || '').toLowerCase() === 'admin' ? 'admin' : 'user';
    const isAdminTarget = /\/admin\.html(?:\?|#|$)/i.test(returnTarget || '');

    function selectedMode() {
        const radios = document.querySelectorAll('input[name="welcome-pin-mode"]');
        for (let i = 0; i < radios.length; i++) {
            if (radios[i].checked) return String(radios[i].value || 'user');
        }
        return requestedMode;
    }

    // Query-Param steuert die UI-Default-Auswahl.
    if (requestedMode === 'admin') {
        const adminRadio = document.querySelector('input[name="welcome-pin-mode"][value="admin"]');
        if (adminRadio) adminRadio.checked = true;
    }

    if (requestedMode === 'admin') {
        if (isAdminAccessGranted()) {
            location.replace('admin.html');
            return;
        }
    } else if (isAccessGranted()) {
        location.replace(returnTarget);
        return;
    }

    const form = document.getElementById('welcome-pin-form');
    const input = document.getElementById('welcome-pin-input');
    const errorEl = document.getElementById('welcome-pin-error');
    const submitBtn = document.getElementById('welcome-pin-submit');

    const userAccessConfig = getActiveUserAccessConfig(config);
    const adminPins = (() => {
        if (!config) return [];
        if (Array.isArray(config.adminPins) && config.adminPins.length) return config.adminPins;
        if (typeof config.adminPin === 'string' && config.adminPin) return [config.adminPin];
        return [];
    })();

    // Wenn weder Schul- noch Admin-Gate konfiguriert ist, einfach durchlassen.
    if (selectedMode() === 'user' && !isPinGateEnabled(userAccessConfig)) {
        location.replace(returnTarget);
        return;
    }

    if (!form || !input) return;

    function showError(message) {
        if (!errorEl) return;
        errorEl.textContent = message;
        errorEl.hidden = !message;
    }

    function hideReleaseNotes() {
        const wrap = document.getElementById('welcome-release-notes-wrap');
        if (wrap) wrap.style.display = 'none';
    }

    form.addEventListener('submit', function (event) {
        event.preventDefault();
        hideReleaseNotes();
        showError('');

        const pin = normalizePin(input.value);
        if (!pin) {
            showError('Bitte geben Sie eine PIN ein.');
            input.focus();
            return;
        }

        const mode = selectedMode();
        if (mode === 'admin') {
            if (!adminPins.length || !(config && config.enabled !== false)) {
                showError('Admin-Zugang ist nicht konfiguriert.');
                return;
            }
            if (!isValidPin(pin, adminPins)) {
                showError('Der Master-PIN ist nicht gültig.');
                input.value = '';
                input.focus();
                return;
            }
            if (submitBtn) submitBtn.disabled = true;
            grantAdminAccess();
            location.replace('admin.html');
            return;
        }

        // User-Modus, aber Ziel ist die Admin-Seite:
        if (isAdminTarget) {
            showError('Für die Admin-Seite ist der Admin-Master-PIN erforderlich.');
            input.focus();
            return;
        }

        if (!isPinGateEnabled(userAccessConfig) || !isValidPin(pin, userAccessConfig.pins)) {
            showError('Die PIN ist nicht gültig. Bitte erneut versuchen.');
            input.value = '';
            input.focus();
            return;
        }

        if (submitBtn) submitBtn.disabled = true;
        grantAccess();

        const notes = loadReleaseNotes(localStorage);
        const lastSeenAt = getLastSeenAt(localStorage);
        const newNotes = getNewReleaseNotes({ notes: notes, lastSeenAtIso: lastSeenAt });

        const wrap = document.getElementById('welcome-release-notes-wrap');
        const list = document.getElementById('welcome-release-notes-list');
        const contBtn = document.getElementById('welcome-continue-btn');

        if (!newNotes.length || !wrap || !list || !contBtn) {
            location.replace(returnTarget);
            return;
        }

        const fmt = new Intl.DateTimeFormat('de-AT', {
            day: '2-digit',
            month: '2-digit',
            year: 'numeric',
            hour: '2-digit',
            minute: '2-digit'
        });

        list.replaceChildren();
        newNotes.forEach((n) => {
            const note = document.createElement('div');
            note.className = 'note';

            const h = document.createElement('h4');
            h.textContent = n.title || '(ohne Titel)';

            const meta = document.createElement('div');
            meta.className = 'meta';
            meta.textContent = n.at && !Number.isNaN(new Date(n.at).getTime()) ? `Stand: ${fmt.format(new Date(n.at))}` : '';

            const pre = document.createElement('pre');
            pre.textContent = n.body || '';

            note.appendChild(h);
            note.appendChild(meta);
            note.appendChild(pre);
            list.appendChild(note);
        });

        wrap.style.display = 'block';
        contBtn.textContent = 'Weiter';

        contBtn.onclick = function () {
            const latestAt = newNotes.reduce(function (acc, x) {
                const t = new Date(x.at).getTime();
                const accT = acc ? new Date(acc).getTime() : -1;
                return t > accT ? x.at : acc;
            }, '');
            setLastSeenAt(latestAt, localStorage);
            location.replace(returnTarget);
        };
    });

    input.addEventListener('input', function () {
        showError('');
    });
})();
