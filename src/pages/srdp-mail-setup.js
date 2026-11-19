/*
 * Seiten-Skript für tools/srdp-mail-setup.html (klassisches Skript, kein Modul).
 * Stand früher inline in der Seite; ausgelagert, damit die
 * Content-Security-Policy ohne 'unsafe-inline' auskommt.
 */
document.addEventListener('DOMContentLoaded', function () {
    if (window.ms365PaSchoolOnboarding && typeof window.ms365PaSchoolOnboarding.mount === 'function') {
        window.ms365PaSchoolOnboarding.mount(document.getElementById('paOnboardingHost'));
    }
});
