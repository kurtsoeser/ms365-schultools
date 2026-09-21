(function (global) {
    'use strict';

    var STORAGE_KEY = 'ms365-pa-onboarding-v1';

    var STEPS = [
        {
            id: 'person',
            title: 'Eine IT-Person festlegen',
            why: 'Nur jemand mit Admin- oder Maker-Rechten kann Environments prüfen und Flows importieren.',
            how: 'Am besten die schulische IT / der Tenant-Admin. Lehrkräfte brauchen das später nicht.',
            links: []
        },
        {
            id: 'license',
            title: 'Lizenz kurz checken',
            why: 'Mit Microsoft 365 A3 (oder O365 A3) reicht Power Automate für Standard-Flows (SharePoint, Approvals, Outlook).',
            how: 'Wenn die Schule schon Teams und SharePoint nutzt, ist das meist erledigt. Extra „Power Automate Premium“ braucht ihr für Freistellungen typischerweise nicht.',
            links: [
                {
                    href: 'https://admin.microsoft.com/#/licenses',
                    label: 'Lizenzen im Admin Center'
                }
            ]
        },
        {
            id: 'dataverse',
            title: 'Default-Environment / Dataverse prüfen',
            why: 'Solutions und saubere Flow-Verwaltung brauchen eine Power-Platform-Umgebung mit Dataverse. Die Default-Umgebung hat das meist schon.',
            how:
                '1) Link öffnen → mit Schul-Admin anmelden.\n' +
                '2) Links „Environments“ / Umgebungen.\n' +
                '3) „Default“ oder „(default)“ anklicken.\n' +
                '4) Steht dort etwas zu Dataverse / Datenbank? → gut, Haken setzen.\n' +
                '5) Wenn unsicher: Screenshot an die begleitende IT – nicht selbst rumexperimentieren.',
            links: [
                {
                    href: 'https://admin.powerplatform.microsoft.com/environments',
                    label: 'Power Platform – Environments'
                }
            ]
        },
        {
            id: 'maker',
            title: 'Maker-Rechte für die IT-Person',
            why: 'Ohne Rolle „Environment Maker“ (oder Admin) kann man keine Solutions/Flows in der Umgebung anlegen oder importieren.',
            how:
                'Im gleichen Environment → „Settings“ / Sicherheit → Benutzer → die IT-Person suchen → Rolle „Environment Maker“ zuweisen (oder System Administrator, wenn ohnehin Admin).',
            links: [
                {
                    href: 'https://admin.powerplatform.microsoft.com/environments',
                    label: 'Environments öffnen'
                }
            ]
        },
        {
            id: 'mailbox',
            title: 'Postfächer: pro Workflow, nicht „eins für alles“',
            why: 'Jeder Flow kann ein anderes freigegebenes Postfach brauchen – Freistellungen z. B. automate@…, Kalender z. B. kalender@…, Anträge etwas anderes.',
            how:
                'Hier noch kein konkretes Postfach anlegen. Merken: Absender und Kalender-Mailbox stellt ihr erst im jeweiligen Automations-Tool ein (Freistellungen, Termine→Kalender, …). Die Flow-Besitzerin braucht dort „Senden als“ / Zugriff auf genau dieses Postfach. Übersicht und Anlage: Postfächer-Tool oder Admin Center.',
            links: [
                {
                    href: 'postfaecher.html',
                    label: 'Postfächer-Tool',
                    internal: true
                },
                {
                    href: 'https://admin.microsoft.com/#/mailboxes',
                    label: 'Postfächer (Admin Center)'
                }
            ]
        },
        {
            id: 'ready',
            title: 'Bereit für die Automations-Tools',
            why: 'Environment und Rechte sitzen – die fachlichen Einstellungen (Liste, Postfach, Genehmiger) macht jedes Tool selbst.',
            how: 'Weiter mit Freistellungen, Termine→Kalender oder der Automationen-Übersicht. Pro Workflow die dortigen Felder ausfüllen; beim Import Connections der Schule zuweisen.',
            links: [
                {
                    href: 'power-automate-rezepte.html',
                    label: 'Automationen-Übersicht',
                    internal: true
                },
                {
                    href: 'freistellung-setup.html',
                    label: 'Freistellungen Setup',
                    internal: true
                },
                {
                    href: 'https://make.powerautomate.com/',
                    label: 'Power Automate öffnen'
                }
            ]
        }
    ];

    function loadState() {
        try {
            return JSON.parse(localStorage.getItem(STORAGE_KEY) || '{}') || {};
        } catch (e) {
            return {};
        }
    }

    function saveState(state) {
        try {
            localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
        } catch (e) {}
    }

    function escapeHtml(s) {
        return String(s || '')
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;')
            .replace(/\n/g, '<br>');
    }

    function progress(state) {
        var n = 0;
        STEPS.forEach(function (s) {
            if (state[s.id]) n++;
        });
        return { done: n, total: STEPS.length };
    }

    /**
     * @param {HTMLElement} host
     * @param {{ compact?: boolean, onChange?: function }} opts
     */
    function render(host, opts) {
        if (!host) return;
        opts = opts || {};
        var state = loadState();
        var p = progress(state);

        var html = '';
        html +=
            '<div class="pa-onb" data-pa-onboarding>' +
            '<div class="pa-onb__progress" role="status">' +
            '<strong>' +
            p.done +
            ' von ' +
            p.total +
            '</strong> Einrichtungsschritte erledigt' +
            '<div class="pa-onb__bar" aria-hidden="true"><span style="width:' +
            Math.round((100 * p.done) / p.total) +
            '%"></span></div>' +
            '</div>';

        STEPS.forEach(function (step, idx) {
            var on = !!state[step.id];
            html +=
                '<details class="pa-onb__step' +
                (on ? ' is-done' : '') +
                '"' +
                (opts.compact && on ? '' : idx === firstOpenIndex(state) ? ' open' : '') +
                '>' +
                '<summary>' +
                '<span class="pa-onb__num">' +
                (idx + 1) +
                '</span>' +
                '<span class="pa-onb__sum-title">' +
                escapeHtml(step.title) +
                '</span>' +
                (on ? '<span class="pa-onb__badge">erledigt</span>' : '') +
                '</summary>' +
                '<div class="pa-onb__body">' +
                '<p class="pa-onb__why"><strong>Warum:</strong> ' +
                escapeHtml(step.why) +
                '</p>' +
                '<p class="pa-onb__how"><strong>So geht’s:</strong><br>' +
                escapeHtml(step.how) +
                '</p>';

            if (step.links && step.links.length) {
                html += '<div class="pa-onb__links">';
                step.links.forEach(function (lnk) {
                    var ext = lnk.internal ? '' : ' target="_blank" rel="noopener"';
                    html +=
                        '<a class="btn btn-sm" href="' +
                        escapeHtml(lnk.href) +
                        '"' +
                        ext +
                        '><i class="bi bi-box-arrow-up-right"></i>' +
                        escapeHtml(lnk.label) +
                        '</a>';
                });
                html += '</div>';
            }

            html +=
                '<label class="checkbox-label pa-onb__check">' +
                '<input type="checkbox" data-pa-onb="' +
                escapeHtml(step.id) +
                '"' +
                (on ? ' checked' : '') +
                '> Diesen Schritt erledigt</label>' +
                '</div></details>';
        });

        html +=
            '<p class="muted pa-onb__foot">Fortschritt wird nur in diesem Browser gespeichert. ' +
            '<button type="button" class="btn btn-sm" data-pa-onb-reset>Zurücksetzen</button></p>' +
            '</div>';

        host.innerHTML = html;

        host.querySelectorAll('[data-pa-onb]').forEach(function (el) {
            el.addEventListener('change', function () {
                var id = el.getAttribute('data-pa-onb');
                var st = loadState();
                st[id] = !!el.checked;
                saveState(st);
                render(host, opts);
                if (typeof opts.onChange === 'function') opts.onChange(progress(st));
            });
        });

        var reset = host.querySelector('[data-pa-onb-reset]');
        if (reset) {
            reset.addEventListener('click', function () {
                if (!confirm('Einrichtungs-Fortschritt zurücksetzen?')) return;
                saveState({});
                render(host, opts);
                if (typeof opts.onChange === 'function') opts.onChange(progress({}));
            });
        }
    }

    function firstOpenIndex(state) {
        for (var i = 0; i < STEPS.length; i++) {
            if (!state[STEPS[i].id]) return i;
        }
        return STEPS.length - 1;
    }

    function injectStylesOnce() {
        if (document.getElementById('pa-onb-styles')) return;
        var css =
            '.pa-onb{display:grid;gap:10px;}' +
            '.pa-onb__progress{font-size:.95em;margin-bottom:4px;}' +
            '.pa-onb__bar{height:8px;background:var(--surface-2,#e8ecf0);border-radius:99px;margin-top:8px;overflow:hidden;}' +
            '.pa-onb__bar>span{display:block;height:100%;background:var(--accent,#0b6a9f);border-radius:99px;transition:width .2s;}' +
            '.pa-onb__step{border:1px solid var(--border);border-radius:12px;background:var(--card);padding:0;}' +
            '.pa-onb__step.is-done{border-color:color-mix(in srgb, var(--success,#0a7a3e) 45%, var(--border));}' +
            '.pa-onb__step>summary{cursor:pointer;list-style:none;display:flex;align-items:center;gap:10px;padding:12px 14px;font-weight:600;}' +
            '.pa-onb__step>summary::-webkit-details-marker{display:none;}' +
            '.pa-onb__num{flex:0 0 auto;width:1.7em;height:1.7em;border-radius:50%;display:inline-grid;place-items:center;background:var(--surface-2,#e8ecf0);font-size:.85em;}' +
            '.pa-onb__step.is-done .pa-onb__num{background:color-mix(in srgb, var(--success,#0a7a3e) 25%, transparent);}' +
            '.pa-onb__sum-title{flex:1;}' +
            '.pa-onb__badge{font-size:.75em;font-weight:600;color:var(--success,#0a7a3e);}' +
            '.pa-onb__body{padding:0 14px 14px 14px;font-size:.94em;line-height:1.5;}' +
            '.pa-onb__why,.pa-onb__how{margin:0 0 10px;}' +
            '.pa-onb__links{display:flex;flex-wrap:wrap;gap:8px;margin:0 0 12px;}' +
            '.pa-onb__check{display:block;margin:0;}' +
            '.pa-onb__foot{margin:8px 0 0;font-size:.88em;}';
        var style = document.createElement('style');
        style.id = 'pa-onb-styles';
        style.textContent = css;
        document.head.appendChild(style);
    }

    /** Mount into #paOnboardingHost or [data-pa-onboarding-host] */
    function mount(selector, opts) {
        injectStylesOnce();
        var host =
            typeof selector === 'string'
                ? document.querySelector(selector)
                : selector;
        if (!host) return null;
        render(host, opts || {});
        return host;
    }

    function isComplete() {
        var st = loadState();
        return progress(st).done >= STEPS.length;
    }

    global.ms365PaOnboarding = {
        steps: STEPS,
        mount: mount,
        loadState: loadState,
        progress: function () {
            return progress(loadState());
        },
        isComplete: isComplete,
        storageKey: STORAGE_KEY
    };
})(window);
