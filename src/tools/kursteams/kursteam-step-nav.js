/**
 * Schritt-Navigation: Weiter-Buttons, Checklisten, Fokus, Scroll.
 */
const ns = (window.ms365Kursteam = window.ms365Kursteam || {});

const CONTINUE_DEFAULTS = {
    continueBtn1: 'Bitte zuerst Daten importieren (Paste oder Datei).',
    continueBtn2: 'Bitte zuerst „Filter anwenden“.',
    continueBtn2_5: 'Bitte mindestens eine Unterrichtszeile erfassen.',
    continueBtn3: 'Weiter, wenn alle Lehrer-E-Mails zugeordnet sind.',
    continueBtn4: 'Bitte zuerst „Team-Namen generieren“.'
};

ns.setContinueButton = function setContinueButton(btnId, enabled, hintText) {
    const btn = document.getElementById(btnId);
    const hint = document.getElementById(btnId + 'Hint');
    if (btn) {
        btn.disabled = !enabled;
        btn.setAttribute('aria-disabled', enabled ? 'false' : 'true');
    }
    if (hint) {
        const text = hintText !== undefined && hintText !== null
            ? hintText
            : enabled
              ? ''
              : CONTINUE_DEFAULTS[btnId] || '';
        hint.textContent = text;
        hint.hidden = !text;
    }
};

ns.scrollToContinue = function scrollToContinue(btnId) {
    const btn = document.getElementById(btnId);
    if (!btn) return;
    const anchor = btn.closest('.kt-step-footer') || btn.closest('.kt-continue-wrap') || btn;
    try {
        anchor.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    } catch {
        anchor.scrollIntoView();
    }
    if (!btn.disabled) {
        setTimeout(() => {
            try {
                btn.focus({ preventScroll: true });
            } catch {
                /* ignore */
            }
        }, 450);
    }
};

ns.focusStepHeading = function focusStepHeading(step) {
    const content = document.querySelector('#panelWebuntis .content > .step-content[data-step="' + step + '"]');
    if (!content) return;
    const h2 = content.querySelector('.kt-step-header h2');
    if (!h2) return;
    if (!h2.hasAttribute('tabindex')) h2.setAttribute('tabindex', '-1');
    try {
        h2.focus({ preventScroll: false });
    } catch {
        /* ignore */
    }
};

function setMiniCheckItem(listId, checkKey, done) {
    const list = document.getElementById(listId);
    if (!list) return;
    const item = list.querySelector('[data-kt-check="' + checkKey + '"]');
    if (!item) return;
    item.classList.toggle('is-done', !!done);
    item.classList.toggle('is-pending', !done);
    item.setAttribute('aria-checked', done ? 'true' : 'false');
}

ns.updateStep4Checklist = function updateStep4Checklist() {
    const needed = parseInt(document.getElementById('uniqueTeachersNeeded')?.textContent || '0', 10);
    const unmapped = parseInt(document.getElementById('unmappedTeachers')?.textContent || '0', 10);
    setMiniCheckItem('step4Checklist', 'teachers-loaded', needed > 0);
    setMiniCheckItem('step4Checklist', 'all-mapped', needed > 0 && unmapped === 0);
    ns.setContinueButton(
        'continueBtn3',
        needed === 0 || unmapped === 0,
        needed > 0 && unmapped > 0
            ? unmapped + ' Lehrer ohne E-Mail – bitte zuordnen oder importieren.'
            : undefined
    );
};

ns.updateStep5Checklist = function updateStep5Checklist() {
    const validTeams = (ns.teamsData || []).filter(t => t.isValid);
    const generated = !!ns.teamsGenerated && (ns.teamsData || []).length > 0;
    setMiniCheckItem('step5Checklist', 'pattern', true);
    setMiniCheckItem('step5Checklist', 'generated', generated);
    setMiniCheckItem('step5Checklist', 'valid', validTeams.length > 0);
    ns.setContinueButton(
        'continueBtn4',
        validTeams.length > 0,
        validTeams.length > 0 ? '' : undefined
    );
};

ns.initKursteamStepNav = function initKursteamStepNav() {
    ns.setContinueButton('continueBtn1', false);
    ns.setContinueButton('continueBtn2', false);
    ns.setContinueButton('continueBtn4', false);
    if (typeof ns.updateStep4Checklist === 'function') ns.updateStep4Checklist();
    if (typeof ns.updateStep5Checklist === 'function') ns.updateStep5Checklist();
};

document.addEventListener('DOMContentLoaded', () => {
    if (typeof ns.initKursteamStepNav === 'function') ns.initKursteamStepNav();
});
