/*
 * Seiten-Skript für tools/pa-erst-setup.html (klassisches Skript, kein Modul).
 * Stand früher inline in der Seite; ausgelagert, damit die
 * Content-Security-Policy ohne 'unsafe-inline' auskommt.
 */
document.addEventListener('DOMContentLoaded', function () {
    if (window.ms365PaOnboarding) window.ms365PaOnboarding.mount('#paOnboardingHost');
});
