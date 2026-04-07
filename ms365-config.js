/**
 * Tragen Sie unten Ihre Anwendungs-ID (Client) aus der Entra-App-Registrierung ein.
 * Ausführliche Schritte: siehe ms365-config.example.js (Kommentarblock oben).
 */
window.MS365_MSAL_CONFIG = {
    clientId: '',
    authority: 'https://login.microsoftonline.com/organizations',
    redirectUri: typeof window !== 'undefined' ? window.location.href.split('#')[0] : ''
};
