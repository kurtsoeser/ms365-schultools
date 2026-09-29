/**
 * Truncation-Banner: sperrt Apply wenn Mitgliedsliste unvollständig (Analyse 01/03).
 */
export function showTruncationBanner(hostOrId, opts) {
    const host = typeof hostOrId === 'string' ? document.getElementById(hostOrId) : hostOrId;
    if (!host) return null;
    const o = opts || {};
    let el = host.querySelector('[data-ms365-truncation-banner]');
    if (!el) {
        el = document.createElement('div');
        el.setAttribute('data-ms365-truncation-banner', '1');
        el.className = 'ms365-truncation-banner';
        el.setAttribute('role', 'alert');
        host.insertBefore(el, host.firstChild);
    }
    el.hidden = false;
    el.innerHTML =
        '<strong>Liste möglicherweise unvollständig.</strong> ' +
        (o.message ||
            'Die Mitgliedschaft konnte nicht vollständig geladen werden. Übernehmen (Apply) ist gesperrt – bitte erneut laden oder Limits prüfen.') +
        (o.helpHref
            ? ' <a href="' + o.helpHref + '">Hilfe</a>'
            : ' <a href="../hilfe.html#faq-truncation">Hilfe</a>');
    return el;
}

export function hideTruncationBanner(hostOrId) {
    const host = typeof hostOrId === 'string' ? document.getElementById(hostOrId) : hostOrId;
    if (!host) return;
    const el = host.querySelector('[data-ms365-truncation-banner]');
    if (el) el.hidden = true;
}

export function guardApplyAgainstTruncation(truncated, applyBtn) {
    if (!applyBtn) return !truncated;
    applyBtn.disabled = !!truncated;
    if (truncated) {
        applyBtn.title = 'Gesperrt: Mitgliedsliste unvollständig (Truncation).';
    } else {
        applyBtn.removeAttribute('title');
    }
    return !truncated;
}

if (typeof window !== 'undefined') {
    window.ms365TruncationUi = {
        showTruncationBanner: showTruncationBanner,
        hideTruncationBanner: hideTruncationBanner,
        guardApplyAgainstTruncation: guardApplyAgainstTruncation
    };
}
