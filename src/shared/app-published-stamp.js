import {
    STAMP_ELEMENT_ID,
    createPublishedStampElement,
    formatPublishedStamp,
    parsePublishedAt
} from './app-published-stamp-logic.js';

function buildInfoUrl() {
    return new URL('../../app-build.json', import.meta.url).href;
}

function mountStamp(label) {
    if (typeof document === 'undefined' || !document.body) return;
    const existing = document.getElementById(STAMP_ELEMENT_ID);
    if (existing) {
        existing.textContent = label;
        return;
    }
    const el = createPublishedStampElement(label, document);
    if (!el) return;
    const target = document.getElementById('ms365FixedFooterRight') || document.body;
    target.appendChild(el);
}

async function loadAndMount() {
    try {
        const res = await fetch(buildInfoUrl(), { cache: 'no-store' });
        if (!res.ok) return;
        const iso = parsePublishedAt(await res.json());
        const label = formatPublishedStamp(iso);
        if (!label) return;
        mountStamp(label);
    } catch {
        /* lokal ohne Build-Info: nichts anzeigen */
    }
}

if (typeof document !== 'undefined') {
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', loadAndMount);
    } else {
        loadAndMount();
    }
}
