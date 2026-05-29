/**
 * PIN-Zugang zur App (Session im Browser, Tab-Sitzung).
 * Kopie als access-config.js anlegen und PINs anpassen.
 *
 * Hinweis: Bei statischem Hosting ist das nur ein Zugangshindernis im Browser
 * (kein Server-Schutz). Für echte Absicherung Hosting mit Server-Auth nutzen.
 */
window.MS365_ACCESS_CONFIG = {
    /** false = Schutz aus (z. B. lokale Entwicklung) */
    enabled: true,
    /**
     * Gültige PINs (mehrere möglich). Vergleich ohne Berücksichtigung von
     * Groß-/Kleinschreibung; Leerzeichen am Anfang/Ende werden ignoriert.
     */
    pins: ['MS365-Schule', 'IT-Team']
};
