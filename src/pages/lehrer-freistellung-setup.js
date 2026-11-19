/*
 * Seiten-Skript für tools/lehrer-freistellung-setup.html (klassisches Skript, kein Modul).
 * Stand früher inline in der Seite; ausgelagert, damit die
 * Content-Security-Policy ohne 'unsafe-inline' auskommt.
 */
document.addEventListener('DOMContentLoaded', function () {
    if (!window.ms365PaOnboarding) return;
    function onOnboardingChange() {
        if (window.ms365LfrSetup && typeof window.ms365LfrSetup.refreshGlance === 'function') {
            window.ms365LfrSetup.refreshGlance();
        }
        var step = window.ms365LfrSetup && window.ms365LfrSetup.loadSetupStep
            ? window.ms365LfrSetup.loadSetupStep() : 1;
        if (window.ms365LfrSetup && window.ms365LfrSetup.showSetupStep) {
            window.ms365LfrSetup.showSetupStep(step);
        }
    }
    window.ms365PaOnboarding.mount('#paOnboardingHost', { compact: true, onChange: onOnboardingChange });
    onOnboardingChange();
});
