(function () {
    'use strict';

    /**
     * Einheitliche Prüfung, ob abhängige Skripte geladen sind (Ladereihenfolge in *.html).
     * @param {Record<string, unknown>} spec Name → Referenz (truthy = ok)
     * @param {string} [label] z. B. Dateiname des Aufrufers
     */
    function assertModules(spec, label) {
        const missing = [];
        Object.keys(spec).forEach(function (name) {
            if (!spec[name]) missing.push(name);
        });
        if (missing.length) {
            const hint = label ? ' [' + label + ']' : '';
            throw new Error(
                'Fehlende Module' +
                    hint +
                    ': ' +
                    missing.join(', ') +
                    '. Reihenfolge der <script>-Tags prüfen (siehe Kommentar in tools/kursteams.html).'
            );
        }
    }

    window.ms365AssertModules = assertModules;
})();
