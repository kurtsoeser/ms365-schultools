/**
 * School e-mail pattern helpers for tenant-settings-ui.
 */
import { normStr } from './tenant-settings-ui-utils-dialog.js';

function schoolDomainForEmail() {
    if (typeof window.ms365GetSchoolDomainNoAt === 'function') {
        return normStr(window.ms365GetSchoolDomainNoAt()).replace(/^@+/, '');
    }
    return '';
}

function emailPatternFromSelect(id, fallback) {
    const el = document.getElementById(id);
    const v = el ? normStr(el.value) : '';
    return v || fallback || 'vorname.nachname';
}

/** @returns {'first'|'all'} */
function emailFirstNameModeFromSelect(id, fallback) {
    const el = document.getElementById(id);
    const raw = el ? normStr(el.value).toLowerCase() : '';
    const api = window.ms365PersonEmailFromName;
    if (api && typeof api.resolveFirstNameMode === 'function') {
        return api.resolveFirstNameMode(raw || fallback || 'first');
    }
    if (raw === 'all' || raw === 'alle') return 'all';
    if (raw === 'first' || raw === 'erster') return 'first';
    const fb = String(fallback || 'first').toLowerCase();
    return fb === 'all' ? 'all' : 'first';
}

function emailPatternFromLocalStorageFallback() {
    try {
        const raw =
            localStorage.getItem('ms365-tenant-school-email-pattern-v1') ||
            localStorage.getItem('ms365-tenant-teachers-email-pattern-v1') ||
            '';
        if (!raw) return '';
        try {
            const obj = JSON.parse(raw);
            if (obj && typeof obj === 'object' && !Array.isArray(obj)) {
                return String(obj.patternId || obj.pattern || '').trim();
            }
        } catch {
            /* Legacy-String */
        }
        return String(raw).trim();
    } catch {
        return '';
    }
}

function emailFirstNameModeFromLocalStorageFallback() {
    try {
        const raw = localStorage.getItem('ms365-tenant-school-email-pattern-v1') || '';
        if (!raw) return '';
        const obj = JSON.parse(raw);
        if (obj && typeof obj === 'object' && !Array.isArray(obj) && obj.firstNameMode != null) {
            return emailFirstNameModeFromSelect('', String(obj.firstNameMode));
        }
    } catch {
        /* ignore */
    }
    return '';
}

/** Schulweites E-Mail-Muster (Stammdaten-Tab), mit Fallback für ältere UI-IDs. */
function tenantSchoolEmailPatternId() {
    return emailPatternFromSelect(
        'tenantSchoolEmailPattern',
        emailPatternFromSelect(
            'tenantTeachersEmailPattern',
            emailPatternFromSelect(
                'tenantStudentsEmailPattern',
                emailPatternFromSelect(
                    'swTeachersEmailPattern',
                    emailPatternFromSelect(
                        'swStudentsEmailPattern',
                        emailPatternFromSelect('evImportEmailPattern', emailPatternFromLocalStorageFallback() || 'vorname.nachname')
                    )
                )
            )
        )
    );
}

function tenantSchoolEmailFirstNameMode() {
    return emailFirstNameModeFromSelect(
        'tenantSchoolEmailGivenNames',
        emailFirstNameModeFromSelect(
            'tenantTeachersEmailGivenNames',
            emailFirstNameModeFromSelect(
                'tenantStudentsEmailGivenNames',
                emailFirstNameModeFromSelect(
                    'swTeachersEmailGivenNames',
                    emailFirstNameModeFromSelect(
                        'swStudentsEmailGivenNames',
                        emailFirstNameModeFromLocalStorageFallback() || 'first'
                    )
                )
            )
        )
    );
}

export {
    schoolDomainForEmail,
    emailPatternFromSelect,
    emailFirstNameModeFromSelect,
    emailPatternFromLocalStorageFallback,
    emailFirstNameModeFromLocalStorageFallback,
    tenantSchoolEmailPatternId,
    tenantSchoolEmailFirstNameMode
};
