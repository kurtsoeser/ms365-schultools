(function () {
    'use strict';

    /** Ein gemeinsamer Wert für alle Module (Kursteams, Jahrgang, ARGE). */
    const STORAGE_KEY = 'ms365-school-email-domain-v1';
    const DEFAULT_DOMAIN_NO_AT = 'ms365.schule';

    const KURSTEAMS_KEY = 'webuntis-teams-creator-state-v1';
    const JAHRGANG_KEY = 'ms365-jahrgang-state-v1';
    const ARGE_KEY = 'ms365-arge-state-v1';

    function normalizeNoAt(v) {
        return String(v ?? '')
            .trim()
            .replace(/^@+/, '');
    }

    function getInputEl() {
        return document.getElementById('schoolEmailDomain');
    }

    function persistFromInput() {
        const el = getInputEl();
        if (!el) return;
        try {
            localStorage.setItem(STORAGE_KEY, normalizeNoAt(el.value));
        } catch {
            // ignore
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

    function initDomValue() {
        const el = getInputEl();
        if (!el) return;
        const v = migrateLegacyOnce();
        el.value = v;
        try {
            localStorage.setItem(STORAGE_KEY, normalizeNoAt(v));
        } catch {
            // ignore
        }
    }

    /**
     * Domain ohne führendes @ (für Vorschauen, Exchange, Graph).
     */
    function getSchoolDomainNoAt() {
        const el = getInputEl();
        return el ? normalizeNoAt(el.value) : '';
    }

    /**
     * Suffix für Kursteams-Lehrer: immer mit einem @, z. B. "@ms365.schule".
     */
    function getTeacherEmailDomainSuffix() {
        const d = getSchoolDomainNoAt();
        return d ? '@' + d : '@';
    }

    function setSchoolDomainNoAt(value) {
        const el = getInputEl();
        if (!el) return;
        const n = normalizeNoAt(value);
        el.value = n || DEFAULT_DOMAIN_NO_AT;
        persistFromInput();
    }

    window.ms365GetSchoolDomainNoAt = getSchoolDomainNoAt;
    window.ms365GetTeacherEmailDomainSuffix = getTeacherEmailDomainSuffix;
    window.ms365SetSchoolDomainNoAt = setSchoolDomainNoAt;
    window.ms365NormalizeSchoolDomainNoAt = normalizeNoAt;
    window.ms365DefaultSchoolDomainNoAt = function () {
        return DEFAULT_DOMAIN_NO_AT;
    };

    function bindInput() {
        const el = getInputEl();
        if (!el) return;
        el.addEventListener('input', persistFromInput);
        el.addEventListener('change', persistFromInput);
    }

    // Defer-Skripte laufen nach vollständigem HTML-Parse; Eingabefeld ist schon im DOM.
    initDomValue();
    bindInput();
})();

