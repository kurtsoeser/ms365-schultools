/*
 * Seiten-Skript für tools/schulaktivitaeten-setup.html (klassisches Skript, kein Modul).
 * Stand früher inline in der Seite; ausgelagert, damit die
 * Content-Security-Policy ohne 'unsafe-inline' auskommt.
 */
document.addEventListener('DOMContentLoaded', function () {
    if (!window.ms365PaOnboarding) return;
    function onOnboardingChange() {
        if (window.ms365AktSetup && typeof window.ms365AktSetup.refreshGlance === 'function') {
            window.ms365AktSetup.refreshGlance();
        }
        var step = window.ms365AktSetup && window.ms365AktSetup.loadSetupStep
            ? window.ms365AktSetup.loadSetupStep() : 1;
        if (window.ms365AktSetup && window.ms365AktSetup.showSetupStep) {
            window.ms365AktSetup.showSetupStep(step);
        }
    }
    window.ms365PaOnboarding.mount('#paOnboardingHost', { compact: true, onChange: onOnboardingChange });
    onOnboardingChange();
});
