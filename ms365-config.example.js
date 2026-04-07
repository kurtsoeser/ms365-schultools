/**
 * Diese Datei ist die Vorlage; tragen Sie die Client-ID in ms365-config.js ein (siehe gleicher Inhalt).
 *
 * === Entra ID: App-Registrierung (einmalig in IHREM Mandanten) ===
 *
 * 1) https://entra.microsoft.com → Identität → Anwendungen → App-Registrierungen → „Neue Registrierung“
 * 2) Name: z. B. „MS365 Schul-Tool ARGE“
 *    „Unterstützte Kontotypen“: „Konten in einem beliebigen Organisationsverzeichnis (beliebiger Microsoft Entra ID-Mandant – Multimandanten)“
 * 3) „Weiterleitungs-URI“: Plattform „Einzelseitenanwendung (SPA)“
 *    URI exakt so eintragen wie Ihre Seite im Browser, z. B.:
 *    https://IHRE-DOMAIN.tld/ms365-schooltool.html
 *    (lokal testen: http://localhost:PORT/ms365-schooltool.html – dieselbe URI auch in Entra eintragen)
 * 4) Registrieren → auf der Übersichtsseite „Anwendungs-ID (Client)“ kopieren → unten bei clientId einfügen
 * 5) API-Berechtigungen → Berechtigung hinzufügen → Microsoft Graph → Delegierte Berechtigungen:
 *    - Group.ReadWrite.All
 *    - User.Read.All
 *    → „Administratorzustimmung für [Organisation] erteilen“ (Global Admin o. ä.)
 * 6) Unter „Authentifizierung“ prüfen: implizite Genehmigung ist NICHT nötig; SPA + Redirect-URI reicht.
 *
 * Schul-Admins legen KEINE eigene App an – sie öffnen nur Ihre URL und melden sich an (Zustimmung ggf. einmal pro Mandant).
 */
window.MS365_MSAL_CONFIG = {
    /** Anwendungs-ID (Client) aus der App-Registrierung */
    clientId: '',
    /** Multimandanten: alle Organisationskonten */
    authority: 'https://login.microsoftonline.com/organizations',
    /**
     * Muss EXAKT einer „Weiterleitungs-URI“ in der App-Registrierung entsprechen.
     * Standard: aktuelle Seiten-URL ohne Hash (Anker).
     */
    redirectUri: typeof window !== 'undefined' ? window.location.href.split('#')[0] : ''
};
