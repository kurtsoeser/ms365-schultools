(function () {
    'use strict';

    /** Ein gemeinsamer Wert für alle Module (Kursteams, Jahrgang, ARGE). */
    const STORAGE_KEY = 'ms365-school-email-domain-v1';
    const TENANT_KEY = 'ms365-tenant-settings-v1';
    const DEFAULT_DOMAIN_NO_AT = 'ms365.schule';

    const KURSTEAMS_KEY = 'webuntis-teams-creator-state-v1';
    const JAHRGANG_KEY = 'ms365-jahrgang-state-v1';
    const ARGE_KEY = 'ms365-arge-state-v1';

    function normalizeNoAt(v) {
        return String(v ?? '')
            .trim()
            .replace(/^@+/, '');
    }

    function safeJsonParse(s) {
        try {
            return JSON.parse(String(s));
        } catch {
            return null;
        }
    }

    function getInputEl() {
        return document.getElementById('schoolEmailDomain');
    }

    /** Domain aus Tenant-JSON (kanonisch), sonst null wenn nicht gesetzt. */
    function getDomainFromTenantJson() {
        try {
            const raw = localStorage.getItem(TENANT_KEY);
            if (!raw) return null;
            const o = safeJsonParse(raw);
            if (!o || typeof o !== 'object') return null;
            const d = normalizeNoAt(o.domain);
            return d || null;
        } catch {
            return null;
        }
    }

    function migrateLegacyOnce() {
        try {
            const existing = localStorage.getItem(STORAGE_KEY);
            if (existing !== null && existing !== '') {
                return normalizeNoAt(existing) || DEFAULT_DOMAIN_NO_AT;
            }
            const kt = localStorage.getItem(KURSTEAMS_KEY);
            if (kt) {
                const o = JSON.parse(kt);
                if (o.emailDomain) {
                    const n = normalizeNoAt(o.emailDomain);
                    if (n) return n;
                }
            }
        } catch {
            // ignore
        }
        try {
            const jg = localStorage.getItem(JAHRGANG_KEY);
            if (jg) {
                const o = JSON.parse(jg);
                if (o.jgDomain) {
                    const n = normalizeNoAt(o.jgDomain);
                    if (n) return n;
                }
            }
        } catch {
            // ignore
        }
        try {
            const ar = localStorage.getItem(ARGE_KEY);
            if (ar) {
                const o = JSON.parse(ar);
                if (o.argeDomain) {
                    const n = normalizeNoAt(o.argeDomain);
                    if (n) return n;
                }
            }
        } catch {
            // ignore
        }
        return DEFAULT_DOMAIN_NO_AT;
    }

    function resolveInitialDomain() {
        const fromTenant = getDomainFromTenantJson();
        if (fromTenant) return fromTenant;
        return migrateLegacyOnce();
    }

    function persistToStorage(value) {
        const n = normalizeNoAt(value);
        try {
            localStorage.setItem(STORAGE_KEY, n || DEFAULT_DOMAIN_NO_AT);
        } catch {
            // ignore
        }
    }

    function persistFromInput() {
        const el = getInputEl();
        if (!el) return;
        persistToStorage(el.value);
    }

    function initDomValue() {
        const v = resolveInitialDomain();
        persistToStorage(v);
        const el = getInputEl();
        if (el) {
            el.value = v;
        }
    }

    /**
     * Domain ohne führendes @ (für Vorschauen, Exchange, Graph).
     */
    function getSchoolDomainNoAt() {
        const el = getInputEl();
        if (el) return normalizeNoAt(el.value);
        try {
            const v = localStorage.getItem(STORAGE_KEY);
            if (v !== null && v !== '') return normalizeNoAt(v);
        } catch {
            // ignore
        }
        return normalizeNoAt(resolveInitialDomain());
    }

    /**
     * Suffix für Kursteams-Lehrer: immer mit einem @, z. B. "@ms365.schule".
     */
    function getTeacherEmailDomainSuffix() {
        const d = getSchoolDomainNoAt();
        return d ? '@' + d : '@';
    }

    function setSchoolDomainNoAt(value) {
        const n = normalizeNoAt(value);
        const final = n || DEFAULT_DOMAIN_NO_AT;
        persistToStorage(final);
        const el = getInputEl();
        if (el) el.value = final;
    }

    function isTenantSchoolDomainConfigured() {
        return !!getDomainFromTenantJson();
    }

    function tenantSettingsPageHref() {
        try {
            const path = window.location.pathname || '';
            if (path.includes('/tools/')) return '../tenant.html';
            return 'tenant.html';
        } catch {
            return 'tenant.html';
        }
    }

    function closeTenantDomainModal() {
        const wrap = document.getElementById('ms365TenantDomainModal');
        if (wrap) wrap.classList.remove('open');
    }

    function showTenantDomainRequiredModal() {
        let wrap = document.getElementById('ms365TenantDomainModal');
        if (!wrap) {
            wrap = document.createElement('div');
            wrap.id = 'ms365TenantDomainModal';
            wrap.className = 'modal-overlay';
            wrap.setAttribute('role', 'dialog');
            wrap.setAttribute('aria-modal', 'true');
            wrap.innerHTML =
                '<div class="modal-box">' +
                '<h3>Tenant-Einstellungen</h3>' +
                '<p style="margin:0 0 12px;line-height:1.5;color:#495057;">Für dieses Werkzeug muss die <strong>E-Mail-Domain der Schule</strong> (ohne @) in den Tenant-Einstellungen gespeichert sein.</p>' +
                '<p style="margin:0;line-height:1.45;color:#6c757d;font-size:0.95em;">Dort legen Sie die Domain einmal zentral fest; sie gilt für Kursteams, Jahrgangsgruppen und ARGEs.</p>' +
                '<div class="modal-actions">' +
                '<button type="button" class="btn" id="ms365TenantDomainModalLater">Schließen</button>' +
                '<button type="button" class="btn btn-success" id="ms365TenantDomainModalGo">Zu den Tenant-Einstellungen</button>' +
                '</div></div>';
            document.body.appendChild(wrap);
            wrap.querySelector('#ms365TenantDomainModalLater').addEventListener('click', closeTenantDomainModal);
            wrap.querySelector('#ms365TenantDomainModalGo').addEventListener('click', () => {
                window.location.href = tenantSettingsPageHref();
            });
            wrap.addEventListener('click', (e) => {
                if (e.target === wrap) closeTenantDomainModal();
            });
        }
        wrap.classList.add('open');
    }

    function shouldSkipTenantDomainPrompt() {
        try {
            if (document.body && document.body.getAttribute('data-ms365-skip-tenant-domain-prompt') === 'true') {
                return true;
            }
            const p = window.location.pathname || '';
            if (/tenant\.html/i.test(p)) return true;
            return false;
        } catch {
            return false;
        }
    }

    function shouldAutoPromptTenantDomain() {
        if (shouldSkipTenantDomainPrompt()) return false;
        try {
            const p = window.location.pathname || '';
            if (/\/tools\/(kursteams|arge|jahrgang)\.html/i.test(p)) return true;
            return false;
        } catch {
            return false;
        }
    }

    function scheduleTenantDomainPrompt() {
        if (!shouldAutoPromptTenantDomain()) return;
        if (isTenantSchoolDomainConfigured()) return;
        showTenantDomainRequiredModal();
    }

    window.ms365GetSchoolDomainNoAt = getSchoolDomainNoAt;
    window.ms365GetTeacherEmailDomainSuffix = getTeacherEmailDomainSuffix;
    window.ms365SetSchoolDomainNoAt = setSchoolDomainNoAt;
    window.ms365NormalizeSchoolDomainNoAt = normalizeNoAt;
    window.ms365DefaultSchoolDomainNoAt = function () {
        return DEFAULT_DOMAIN_NO_AT;
    };
    window.ms365IsTenantSchoolDomainConfigured = isTenantSchoolDomainConfigured;
    window.ms365ShowTenantDomainRequiredModal = showTenantDomainRequiredModal;

    function bindInput() {
        const el = getInputEl();
        if (!el) return;
        el.addEventListener('input', persistFromInput);
        el.addEventListener('change', persistFromInput);
    }

    initDomValue();
    bindInput();

    document.addEventListener('DOMContentLoaded', () => {
        setTimeout(scheduleTenantDomainPrompt, 0);
    });
})();
