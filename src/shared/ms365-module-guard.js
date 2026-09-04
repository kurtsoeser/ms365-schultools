/**
 * Einheitliche Prüfung, ob abhängige Skripte geladen sind.
 * Als Side-Effect-Import in kursteam-teams.js laden (erster Import), damit Vite-Bundles
 * den Guard enthalten – klassische <script defer>-Reihenfolge reicht nach dem Build nicht.
 * Kein `export`: Tests laden die Datei per vm als klassisches Script.
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
                '. Reihenfolge der <script>-Tags bzw. Imports prüfen (siehe Kommentar in tools/kursteams.html).'
        );
    }
}

window.ms365AssertModules = assertModules;
