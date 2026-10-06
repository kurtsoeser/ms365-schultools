/**
 * Speichern in der Dashboard-Header-Leiste (Stammdaten aus dem lokalen Speicher erneut schreiben).
 */
function notifySaved(msg, opts) {
    const text = String(msg || 'Stammdaten in diesem Browser gespeichert.').trim();
    const options = opts && typeof opts === 'object' ? opts : {};
    if (typeof window.ms365ShowToast === 'function') {
        window.ms365ShowToast(text, {
            kind: 'success',
            durationMs: options.quiet ? 2200 : undefined
        });
        return;
    }
    if (typeof window.ms365ToastOrAlert === 'function') {
        window.ms365ToastOrAlert(text, { kind: 'success' });
    }
}

function dashSaveFromToolbar() {
    if (typeof window.ms365TenantSettingsLoad !== 'function' || typeof window.ms365TenantSettingsSave !== 'function') {
        window.location.href = 'tenant.html';
        return;
    }
    const s = window.ms365TenantSettingsLoad();
    window.ms365TenantSettingsSave(s);
    try {
        window.dispatchEvent(
            new CustomEvent('ms365-local-data-changed', { detail: { source: 'tenant-settings' } })
        );
    } catch {
        /* ignore */
    }
    try {
        const spo = window.ms365StammdatenSpoAutoSync;
        if (spo && typeof spo.schedulePush === 'function') {
            spo.schedulePush(0);
        }
    } catch {
        /* ignore */
    }
    notifySaved();
}

function mountDashboardAutoSave() {
    if (!document.body.classList.contains('page-dashboard')) return;
    if (document.documentElement.dataset.dashAutoSaveBound === '1') return;
    document.documentElement.dataset.dashAutoSaveBound = '1';

    let timer = 0;
    let lastToastAt = 0;

    function flush() {
        if (typeof window.ms365TenantSettingsLoad !== 'function' || typeof window.ms365TenantSettingsSave !== 'function') {
            return;
        }
        const s = window.ms365TenantSettingsLoad();
        window.ms365TenantSettingsSave(s);
        try {
            const spo = window.ms365StammdatenSpoAutoSync;
            if (spo && typeof spo.schedulePush === 'function') spo.schedulePush(0);
        } catch {
            /* ignore */
        }
        const now = Date.now();
        if (now - lastToastAt > 2800) {
            notifySaved('Gespeichert', { quiet: true });
            lastToastAt = now;
        }
    }

    function schedule(ev) {
        const src = ev && ev.detail && ev.detail.source ? String(ev.detail.source) : '';
        if (src === 'dashboard-auto-save' || src === 'tenant-settings') return;
        clearTimeout(timer);
        timer = window.setTimeout(flush, 700);
    }

    window.addEventListener('ms365-local-data-changed', schedule);
    window.addEventListener('ms365-tenant-settings-changed', schedule);
}

export function mountDashboardRegisterToolbar() {
    const root = document.getElementById('dashRegisterToolbar');
    if (root && root.dataset.bound !== '1') {
        root.dataset.bound = '1';
        root.querySelectorAll('[data-dash-tenant-action="save"]').forEach(function (btn) {
            btn.addEventListener('click', dashSaveFromToolbar);
        });
    }
    mountDashboardAutoSave();
}

if (typeof document !== 'undefined') {
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', mountDashboardRegisterToolbar);
    } else {
        mountDashboardRegisterToolbar();
    }
}
