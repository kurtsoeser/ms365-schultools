/**
 * index.html: ruhiger Hinweis nach Redirect von school-it-access-gate
 * (?access=denied&reason=school-it).
 */

const BANNER_ID = 'dashAccessDeniedBanner';

/**
 * @param {string} [search]
 * @returns {{ access: string, reason: string }|null}
 */
export function readSchoolItAccessDeniedParams(search) {
    try {
        const q = new URLSearchParams(
            typeof search === 'string' ? search : String(window.location.search || '')
        );
        const access = String(q.get('access') || '')
            .trim()
            .toLowerCase();
        if (access !== 'denied') return null;
        return {
            access: 'denied',
            reason: String(q.get('reason') || '')
                .trim()
                .toLowerCase()
        };
    } catch {
        return null;
    }
}

function messageForReason(reason) {
    if (reason === 'school-it') {
        return {
            title: 'Nur für die Schul-IT',
            text: 'Dieses Werkzeug richtet die Schul-IT ein. Sie bleiben auf der Startseite und können die Alltagswerkzeuge weiter nutzen.'
        };
    }
    return {
        title: 'Kein Zugriff',
        text: 'Für diesen Bereich fehlt die erforderliche Rolle. Bei Bedarf die Schul-IT ansprechen.'
    };
}

function ensureBanner() {
    let el = document.getElementById(BANNER_ID);
    if (el) return el;
    el = document.createElement('div');
    el.id = BANNER_ID;
    el.className = 'dash-access-denied';
    el.setAttribute('role', 'alert');
    el.hidden = true;

    const host =
        document.getElementById('dashGreetingBand') ||
        document.querySelector('.content') ||
        document.body;
    if (host && host.id === 'dashGreetingBand' && host.parentElement) {
        host.parentElement.insertBefore(el, host.nextSibling);
    } else if (host && host.classList && host.classList.contains('content')) {
        host.insertBefore(el, host.firstChild);
    } else {
        document.body.insertBefore(el, document.body.firstChild);
    }
    return el;
}

function clearQueryParams() {
    try {
        const url = new URL(window.location.href);
        if (!url.searchParams.has('access') && !url.searchParams.has('reason')) return;
        url.searchParams.delete('access');
        url.searchParams.delete('reason');
        const next = url.pathname + (url.search ? url.search : '') + url.hash;
        window.history.replaceState({}, '', next);
    } catch {
        /* ignore */
    }
}

/**
 * @param {{ search?: string, persistQuery?: boolean }} [opts]
 * @returns {boolean} true wenn Banner gezeigt wurde
 */
export function mountSchoolItAccessDeniedBanner(opts) {
    const params = readSchoolItAccessDeniedParams(opts && opts.search);
    if (!params) return false;

    const copy = messageForReason(params.reason);
    const el = ensureBanner();
    el.hidden = false;
    el.innerHTML =
        '<div class="dash-access-denied__icon" aria-hidden="true"><i class="bi bi-shield-lock"></i></div>' +
        '<div class="dash-access-denied__body">' +
        '<p class="dash-access-denied__title">' +
        copy.title +
        '</p>' +
        '<p class="dash-access-denied__text">' +
        copy.text +
        '</p>' +
        '</div>' +
        '<button type="button" class="dash-access-denied__dismiss" aria-label="Hinweis schließen">' +
        '<i class="bi bi-x-lg" aria-hidden="true"></i>' +
        '</button>';

    const dismiss = el.querySelector('.dash-access-denied__dismiss');
    if (dismiss) {
        dismiss.addEventListener('click', function () {
            el.hidden = true;
            clearQueryParams();
        });
    }

    if (!(opts && opts.persistQuery)) {
        clearQueryParams();
    }
    return true;
}

if (typeof window !== 'undefined') {
    window.ms365SchoolItAccessDeniedUi = {
        read: readSchoolItAccessDeniedParams,
        mount: mountSchoolItAccessDeniedBanner
    };
}
