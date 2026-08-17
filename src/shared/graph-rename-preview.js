(function () {
    'use strict';

    function escapeRe(s) {
        return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    }

    /**
     * Erwartetes Muster: "<Präfix> <Stufe><Kürzel>" z. B. "Klasse 1A", "Klasse 10HAK"
     * @param {string} displayName
     * @param {string} prefix z. B. "Klasse"
     * @returns {string|null}
     */
    function computeNewDisplayNamePlusOne(displayName, prefix) {
        const p = String(prefix || '').trim();
        if (!p) return null;
        const re = new RegExp('^' + escapeRe(p) + '\\s+(\\d+)([A-Za-z0-9\\-]*)$', 'i');
        const m = String(displayName || '').trim().match(re);
        if (!m) return null;
        const current = parseInt(m[1], 10);
        if (!isFinite(current)) return null;
        const next = current + 1;
        return p + ' ' + String(next) + (m[2] || '');
    }

    /**
     * Namen ohne Präfix: "1HMA" → "2HMA", "10HAK" → "11HAK".
     * @param {string} displayName
     * @returns {string|null}
     */
    function incrementLeadingGrade(displayName) {
        const s = String(displayName || '').trim();
        const m = s.match(/^(\d{1,2})([A-Za-z][A-Za-z0-9\-]*)$/);
        if (!m) return null;
        const current = parseInt(m[1], 10);
        if (!isFinite(current)) return null;
        return String(current + 1) + m[2];
    }

    /**
     * Vorschlag für den Schuljahreswechsel: zuerst Präfix-Muster, sonst führende Stufe.
     * @param {string} displayName
     * @param {string} [prefix]
     * @returns {string|null}
     */
    function suggestDisplayNamePlusOne(displayName, prefix) {
        const p = String(prefix || '').trim() || 'Klasse';
        return (
            computeNewDisplayNamePlusOne(displayName, p) ||
            computeNewDisplayNamePlusOne(displayName, 'Klasse') ||
            incrementLeadingGrade(displayName)
        );
    }

    window.ms365GraphRenamePreview = {
        computeNewDisplayNamePlusOne,
        incrementLeadingGrade,
        suggestDisplayNamePlusOne,
        escapeRe
    };
})();
