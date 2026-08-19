(function () {
    'use strict';

    var STORAGE_KEY = 'ms365-onboarding-welcome-v1';

    var STEPS = [
        {
            title: 'Daten bleiben in diesem Browser',
            body: 'Die App speichert Schul-Stammdaten lokal in Ihrem Browser – es gibt keinen Schul-Server dahinter. Nutzen Sie möglichst immer denselben Browser (nicht privat/inkognito).',
            list: ['Nach größeren Änderungen: Browser-Backup exportieren (unten auf dem Dashboard).']
        },
        {
            title: 'Microsoft-Anmeldung',
            body: 'Oben rechts bei Microsoft 365 anmelden – mit einem Konto, das Gruppen und Benutzer verwalten darf. Ohne Anmeldung können Sie Stammdaten pflegen, aber keine Gruppen in Microsoft 365 anlegen.',
            list: null
        },
        {
            title: 'Erster Schritt',
            body: 'Starten Sie mit der geführten Einrichtung – oder probieren Sie die Demo mit Beispieldaten, ohne etwas anzulegen.',
            list: ['Einrichtung: Domain, Listen, Übersicht', 'Demo: MS365 Musterschule zum Ausprobieren']
        }
    ];

    function hasSeen() {
        try {
            return localStorage.getItem(STORAGE_KEY) === 'seen';
        } catch {
            return true;
        }
    }

    function markSeen() {
        try {
            localStorage.setItem(STORAGE_KEY, 'seen');
        } catch {
            /* ignore */
        }
    }

    function shouldOffer() {
        if (hasSeen()) return false;
        return !!document.getElementById('dashboard-tasks');
    }

    function buildOverlay() {
        var overlay = document.createElement('div');
        overlay.className = 'ms365-onboarding-overlay';
        overlay.id = 'ms365OnboardingOverlay';
        overlay.setAttribute('role', 'dialog');
        overlay.setAttribute('aria-modal', 'true');
        overlay.setAttribute('aria-labelledby', 'ms365OnboardingTitle');

        var card = document.createElement('div');
        card.className = 'ms365-onboarding-card';

        var header = document.createElement('header');
        header.className = 'ms365-onboarding-header';
        header.innerHTML =
            '<h2 id="ms365OnboardingTitle">Willkommen in der MS365-Schulverwaltung</h2>' +
            '<p>In drei kurzen Schritten – danach wissen Sie, wo Sie starten.</p>';

        var dots = document.createElement('div');
        dots.className = 'ms365-onboarding-steps';
        dots.setAttribute('aria-hidden', 'true');
        STEPS.forEach(function (_, i) {
            var dot = document.createElement('span');
            dot.className = 'ms365-onboarding-step-dot' + (i === 0 ? ' is-active' : '');
            dot.setAttribute('data-step-dot', String(i));
            dots.appendChild(dot);
        });

        var body = document.createElement('div');
        body.className = 'ms365-onboarding-body';
        STEPS.forEach(function (step, i) {
            var panel = document.createElement('div');
            panel.className = 'ms365-onboarding-panel';
            panel.setAttribute('data-step-panel', String(i));
            if (i !== 0) panel.hidden = true;

            var h3 = document.createElement('h3');
            h3.textContent = step.title;
            panel.appendChild(h3);

            var p = document.createElement('p');
            p.textContent = step.body;
            panel.appendChild(p);

            if (step.list && step.list.length) {
                var ul = document.createElement('ul');
                step.list.forEach(function (item) {
                    var li = document.createElement('li');
                    li.textContent = item;
                    ul.appendChild(li);
                });
                panel.appendChild(ul);
            }
            body.appendChild(panel);
        });

        var footer = document.createElement('footer');
        footer.className = 'ms365-onboarding-footer';

        var skip = document.createElement('button');
        skip.type = 'button';
        skip.className = 'ms365-onboarding-skip';
        skip.textContent = 'Nicht mehr anzeigen';

        var actions = document.createElement('div');
        actions.className = 'ms365-onboarding-actions';

        var backBtn = document.createElement('button');
        backBtn.type = 'button';
        backBtn.className = 'btn btn-ghost';
        backBtn.id = 'ms365OnboardingBack';
        backBtn.textContent = 'Zurück';
        backBtn.hidden = true;

        var nextBtn = document.createElement('button');
        nextBtn.type = 'button';
        nextBtn.className = 'btn btn-success';
        nextBtn.id = 'ms365OnboardingNext';
        nextBtn.textContent = 'Weiter';

        actions.appendChild(backBtn);
        actions.appendChild(nextBtn);
        footer.appendChild(skip);
        footer.appendChild(actions);

        card.appendChild(header);
        card.appendChild(dots);
        card.appendChild(body);
        card.appendChild(footer);
        overlay.appendChild(card);

        return {
            overlay: overlay,
            backBtn: backBtn,
            nextBtn: nextBtn,
            skipBtn: skip,
            panels: body.querySelectorAll('[data-step-panel]'),
            dots: dots.querySelectorAll('[data-step-dot]')
        };
    }

    function close(ui) {
        ui.overlay.classList.remove('is-open');
        document.body.style.overflow = '';
    }

    function showStep(ui, index) {
        ui.panels.forEach(function (panel, i) {
            panel.hidden = i !== index;
        });
        ui.dots.forEach(function (dot, i) {
            dot.classList.toggle('is-active', i === index);
            dot.classList.toggle('is-done', i < index);
        });
        ui.backBtn.hidden = index <= 0;
        ui.nextBtn.textContent = index >= STEPS.length - 1 ? 'Los geht\'s' : 'Weiter';
        ui._step = index;
    }

    function openOnboarding() {
        if (!shouldOffer()) return false;

        var ui = buildOverlay();
        ui._step = 0;
        document.body.appendChild(ui.overlay);

        function finish() {
            markSeen();
            close(ui);
        }

        ui.skipBtn.addEventListener('click', finish);
        ui.backBtn.addEventListener('click', function () {
            if (ui._step > 0) showStep(ui, ui._step - 1);
        });
        ui.nextBtn.addEventListener('click', function () {
            if (ui._step >= STEPS.length - 1) finish();
            else showStep(ui, ui._step + 1);
        });
        ui.overlay.addEventListener('click', function (e) {
            if (e.target === ui.overlay) finish();
        });

        document.body.style.overflow = 'hidden';
        ui.overlay.classList.add('is-open');
        ui.nextBtn.focus();
        return true;
    }

    window.ms365OnboardingWelcome = {
        STORAGE_KEY: STORAGE_KEY,
        hasSeen: hasSeen,
        markSeen: markSeen,
        open: openOnboarding
    };

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', function () {
            setTimeout(openOnboarding, 400);
        });
    } else {
        setTimeout(openOnboarding, 400);
    }
})();
