/*
 * Seiten-Skript für tools/freistellung-setup.html (klassisches Skript, kein Modul).
 * Stand früher inline in der Seite; ausgelagert, damit die
 * Content-Security-Policy ohne 'unsafe-inline' auskommt.
 */
document.addEventListener('DOMContentLoaded', function () {
    if (!window.ms365PaOnboarding) return;
    function onOnboardingChange(p) {
        if (window.ms365FreistellungSetup && typeof window.ms365FreistellungSetup.refreshGlance === 'function') {
            window.ms365FreistellungSetup.refreshGlance();
        }
        var step = 1;
        if (window.ms365FreistellungSetup && typeof window.ms365FreistellungSetup.loadSetupStep === 'function') {
            step = window.ms365FreistellungSetup.loadSetupStep();
        }
        if (window.ms365FreistellungSetup && typeof window.ms365FreistellungSetup.showSetupStep === 'function') {
            window.ms365FreistellungSetup.showSetupStep(step);
        } else {
            var el = document.getElementById('frPhaseHint');
            if (el && p && p.total && p.done >= p.total) {
                el.textContent = 'Vorbereitung fertig – weiter mit Konten (Schritt 2).';
            }
        }
    }
    window.ms365PaOnboarding.mount('#paOnboardingHost', { compact: true, onChange: onOnboardingChange });
    onOnboardingChange(window.ms365PaOnboarding.progress());
});
