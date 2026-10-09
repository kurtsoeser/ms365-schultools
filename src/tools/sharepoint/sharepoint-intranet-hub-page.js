/**
 * Wizard-Navigation für sharepoint-intranet-hub.html
 */
import {
    IH_WIZARD_STEP_COUNT,
    clampWizardStep,
    loadWizardStep,
    saveWizardStep,
    validateWizardStep,
    wizardPhaseHint
} from './sharepoint-intranet-hub-wizard.js';

function $(id) {
    return document.getElementById(id);
}

function toast(msg) {
    if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(msg);
    else window.alert(msg);
}

function wizardContextFromDom() {
    const last = $('fLastSiteUrl') ? String($('fLastSiteUrl').value || '').trim() : '';
    const manual = $('fManualUrl') ? String($('fManualUrl').value || '').trim() : '';
    return {
        siteUrl: last || manual,
        host: $('fHost') ? String($('fHost').value || '').trim() : '',
        slug: $('fSlug') ? String($('fSlug').value || '').trim() : ''
    };
}

export function showWizardStep(n) {
    const step = saveWizardStep(clampWizardStep(n));
    for (let i = 1; i <= IH_WIZARD_STEP_COUNT; i++) {
        const panel = $('ihWizardStep' + i);
        if (!panel) continue;
        const on = i === step;
        panel.hidden = !on;
        panel.setAttribute('aria-hidden', on ? 'false' : 'true');
    }
    document.querySelectorAll('#ihWizardGlance [data-ih-wizard-step]').forEach(function (btn) {
        const sn = parseInt(btn.getAttribute('data-ih-wizard-step'), 10);
        const on = sn === step;
        btn.classList.toggle('is-active', on);
        btn.setAttribute('aria-selected', on ? 'true' : 'false');
        btn.setAttribute('tabindex', on ? '0' : '-1');
    });
    const hint = $('ihPhaseHint');
    if (hint) hint.textContent = wizardPhaseHint(step);
    const back = $('ihWizardBack');
    const next = $('ihWizardNext');
    if (back) back.disabled = step <= 1;
    if (next) {
        const isLast = step >= IH_WIZARD_STEP_COUNT;
        next.innerHTML = isLast
            ? '<i class="bi bi-check2-circle"></i>Fertig'
            : 'Weiter<i class="bi bi-arrow-right" style="margin-left:6px;"></i>';
        next.classList.toggle('btn-success', isLast);
    }
    try {
        const u = new URL(window.location.href);
        u.searchParams.set('step', String(step));
        window.history.replaceState({}, '', u);
    } catch {
        /* ignore */
    }
}

function tryAdvance(fromStep) {
    const err = validateWizardStep(fromStep, wizardContextFromDom());
    if (err) {
        toast(err);
        return false;
    }
    return true;
}

function wireWizard() {
    document.querySelectorAll('[data-ih-wizard-step]').forEach(function (btn) {
        btn.addEventListener('click', function () {
            const target = parseInt(btn.getAttribute('data-ih-wizard-step'), 10);
            const cur = loadWizardStep();
            if (target > cur) {
                for (let s = cur; s < target; s++) {
                    if (!tryAdvance(s)) {
                        showWizardStep(s);
                        return;
                    }
                }
            }
            showWizardStep(target);
        });
    });
    const back = $('ihWizardBack');
    if (back) {
        back.addEventListener('click', function () {
            const cur = loadWizardStep();
            if (cur > 1) showWizardStep(cur - 1);
        });
    }
    const next = $('ihWizardNext');
    if (next) {
        next.addEventListener('click', function () {
            const cur = loadWizardStep();
            if (cur >= IH_WIZARD_STEP_COUNT) {
                showWizardStep(cur);
                return;
            }
            if (!tryAdvance(cur)) return;
            showWizardStep(cur + 1);
        });
    }
    window.addEventListener('ms365-ih-site-created', function () {
        showWizardStep(2);
    });
    window.addEventListener('ms365-ih-hub-registered', function () {
        const cur = loadWizardStep();
        if (cur < 3) showWizardStep(3);
    });
}

document.addEventListener('DOMContentLoaded', function () {
    wireWizard();
    showWizardStep(loadWizardStep());
});
