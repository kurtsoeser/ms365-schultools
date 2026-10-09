/**
 * Graph-Zugriff für Entra-Suchdialoge: spo-graph-shared oder graph-unified-groups (z. B. Verwaltung).
 */

/**
 * @returns {{ getGraphToken: function, graphJson: function }}
 */
export function getGraphPickerApi() {
    const spo = typeof window !== 'undefined' ? window.ms365SpoGraph : null;
    if (spo && typeof spo.getGraphToken === 'function' && typeof spo.graphJson === 'function') {
        return spo;
    }
    const ug = typeof window !== 'undefined' ? window.ms365GraphUnifiedGroups : null;
    if (ug && typeof ug.getGraphToken === 'function' && typeof ug.graphJson === 'function') {
        return {
            getGraphToken(scopes) {
                return ug.getGraphToken(scopes);
            },
            graphJson(method, pathOrUrl, token, body, versionOrHeaders) {
                if (versionOrHeaders === 'v1.0' || versionOrHeaders === 'beta') {
                    return ug.graphJson(method, pathOrUrl, token, body, undefined);
                }
                return ug.graphJson(method, pathOrUrl, token, body, versionOrHeaders);
            }
        };
    }
    throw new Error(
        'Microsoft Graph ist hier nicht bereit. Seite neu laden und ggf. im Tool oben bei Microsoft anmelden.'
    );
}
