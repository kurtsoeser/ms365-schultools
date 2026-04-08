/**
 * Tragen Sie unten Ihre Anwendungs-ID (Client) aus der Entra-App-Registrierung ein.
 * Ausführliche Schritte: siehe ms365-config.example.js (Kommentarblock oben).
 */
window.MS365_MSAL_CONFIG = {
    clientId: 'e1d877c3-004c-4040-8c3b-81a59e0c7050',
    authority: 'https://login.microsoftonline.com/organizations',
    redirectUri: typeof window !== 'undefined' ? window.location.href.split('#')[0] : ''
};

(function () {
    if (typeof document === 'undefined') return;
    function injectSiteCredit() {
        if (document.getElementById('siteCreditKurtrocks')) return;
        const p = document.createElement('p');
        p.id = 'siteCreditKurtrocks';
        p.className = 'site-credit-row';
        const a = document.createElement('a');
        a.className = 'site-credit-link';
        a.href = 'https://www.kurtrocks.com/';
        a.target = '_blank';
        a.rel = 'noopener noreferrer';
        const icon = document.createElement('i');
        icon.className = 'bi bi-info-circle';
        icon.setAttribute('aria-hidden', 'true');
        a.appendChild(icon);
        a.appendChild(document.createTextNode('kurtrocks.com'));
        p.appendChild(a);
        document.body.appendChild(p);
    }
    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', injectSiteCredit);
    else injectSiteCredit();
})();
