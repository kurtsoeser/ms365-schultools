/**
 * Tragen Sie unten Ihre Anwendungs-ID (Client) aus der Entra-App-Registrierung ein.
 * Ausführliche Schritte: siehe ms365-config.example.js (Kommentarblock oben).
 */
window.MS365_MSAL_CONFIG = {
    clientId: 'e1d877c3-004c-4040-8c3b-81a59e0c7050',
    authority: 'https://login.microsoftonline.com/organizations',
    redirectUri: typeof window !== 'undefined' ? window.location.href.split('#')[0] : ''
};
