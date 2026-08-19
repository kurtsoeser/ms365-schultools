(function () {
    'use strict';

    /** @type {Record<string, string>} */
    var TIPS = {
        'task-setup':
            'Einmaliger Assistent: Domain, Listen und Übersicht anlegen. Danach nutzen alle Werkzeuge dieselben Stammdaten in diesem Browser.',
        'task-unterricht':
            'Wann? Zu Schuljahresbeginn oder bei neuen Klassen/Kursen. Klassengruppen = eine Microsoft-365-Gruppe pro Klasse. Unterrichtsteams = Teams pro Fach/Kurs aus dem Stundenplan.',
        'task-gruppen':
            'Wann? Wenn alle Schülerinnen, Lehrkräfte oder Verwaltung in großen Gruppen sein sollen (E-Mail, Berechtigungen). Das sind Sammelgruppen in Microsoft 365.',
        'task-personen':
            'Wann? Einzelne Konten suchen, ein neues Konto anlegen oder externe Personen (Gäste) einladen – z. B. Eltern oder Projektpartner.',
        'task-schuljahr':
            'Wann? Einmal pro Jahr beim Wechsel (z. B. im Sommer). Die Checkliste führt durch Schuljahr, Anzeigenamen, Abschlussjahrgang und Teams.',
        'task-ordnung':
            'Wann? Regelmäßig aufräumen: leere Gruppen, fehlende Besitzer und einen Überblick über alle Gruppen der Schule.',

        'prog-unterricht':
            '„Verknüpft“ heißt: Die App kennt die passende Microsoft-365-Gruppe für diese Klasse. Noch nicht verknüpft? Im Werkzeug Klassengruppen die bestehende Gruppe zuordnen oder neu anlegen.',
        'prog-gruppen':
            'Zeigt, wie viele der großen Sammelgruppen (Schüler, Lehrkräfte, Verwaltung) und Fach-/ARGE-Gruppen mit Microsoft 365 verbunden sind.',
        'prog-schuljahr':
            'Fortschritt der Checkliste „Schuljahr wechseln“. Erscheint erst nach abgeschlossener Ersteinrichtung.',

        stammdaten:
            'Listen, die alle Werkzeuge nutzen: Domain, Fächer, Lehrkräfte, Klassen und Schülerinnen. Die Daten bleiben in diesem Browser – nichts wird auf einen Schul-Server geschickt.',
        datenhygiene:
            'Abgleich mit Microsoft 365: Vergleicht Ihre lokalen Stammlisten mit den Gruppen in Microsoft 365. So erkennen Sie veraltete oder falsche Mitgliedschaften.',

        jahrgang:
            'Microsoft-365-Gruppe pro Klasse anlegen oder eine bestehende zuordnen. Brauchen Sie das? Fast immer – Basis für Klassen-E-Mail und Teams.',
        kursteams:
            'Teams für Unterrichtsfächer aus dem Stundenplan (z. B. WebUntis-Export). Sinnvoll, wenn Lehrkräfte pro Fach/Kurs ein eigenes Team nutzen.',
        'arge-fachgruppen':
            'Eine Gruppe pro Fach oder Arbeitsgemeinschaft (ARGE). Voraussetzung: Fächer und ARGEs in den Stammdaten gepflegt.',
        'personen-verwaltung':
            'Personen der Schule im Microsoft-Tenant suchen, Lizenzen prüfen und sehen, in welchen Gruppen sie sind.',
        'gaeste-verwalten':
            'Externe Personen einladen (Gast-Konten) und festlegen, wer einladen darf – z. B. für Eltern oder Kooperationspartner.',
        slg: 'Zwei zentrale Sammelgruppen: alle Schülerinnen und alle Lehrkräfte – für E-Mail-Verteiler und Berechtigungen.',
        verwaltung: 'Gruppe für Sekretariat, Direktion und weitere Verwaltungsrollen – oft als Besitzerin anderer Gruppen.',
        'organisations-assistent':
            'Geführte Checkliste beim Schuljahreswechsel: Schuljahr, Namen, Abschlussjahrgang, Schülerliste, Unterrichtsteams.',
        postfaecher:
            'Gemeinsame Postfächer in Exchange – z. B. sekretariat@… – die mehrere Personen nutzen.',
        verteilerlisten: 'Klassische E-Mail-Verteilerlisten anzeigen und per Skript anlegen oder ändern.',
        'eltern-verteiler':
            'Erziehungsberechtigte den Schülerinnen zuordnen und Klassen-/Jahrgangs-Verteiler erzeugen.',
        'sharepoint-intranet-hub':
            'Schulwebsite/Intranet auf SharePoint: Kommunikationssite, Hub und Startlisten anlegen.',
        'sharepoint-liste-lehrer': 'Lehrkräfte aus den Stammdaten als Liste auf der Schul-Website veröffentlichen.',
        'sharepoint-liste-schultermine':
            'Leere Termineliste auf der Website anlegen – Befüllung später möglich.',
        gruppenerstellung:
            'Festlegen, wer an der Schule neue Teams in Microsoft 365 anlegen darf – verhindert Gruppen-Chaos.',
        'sharepoint-mandant-website': 'Regeln, ob und wie neue SharePoint-Websites an der Schule entstehen dürfen.',
        'sharepoint-mandant-teilen':
            'Freigabe-Regeln: wie stark Dateien innerhalb und außerhalb der Schule geteilt werden dürfen.',
        'schulstruktur-sync':
            'Alle Gruppen und Teams der Schule im Überblick – pflegen, archivieren, Zusatz-Teams anlegen.',
        'leere-gruppen-report':
            'Gruppen ohne Besitzer oder ohne Mitglieder finden – zum Aufräumen und für die IT-Hygiene.'
    };

    var mounted = false;

    function escapeAttr(s) {
        return String(s || '')
            .replace(/&/g, '&amp;')
            .replace(/"/g, '&quot;')
            .replace(/</g, '&lt;');
    }

    function createTipElement(tipId, label, alignRight) {
        var text = TIPS[tipId];
        if (!text) return null;

        var wrap = document.createElement('span');
        wrap.className = 'context-tip' + (alignRight ? ' context-tip--align-right' : '');
        wrap.setAttribute('data-context-tip-mounted', tipId);

        var bubbleId = 'context-tip-bubble-' + tipId.replace(/[^a-z0-9-]/gi, '-');

        var btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'context-tip__btn';
        btn.setAttribute('aria-label', label || 'Kurzer Hinweis');
        btn.setAttribute('aria-describedby', bubbleId);
        btn.innerHTML = '<i class="bi bi-info-circle" aria-hidden="true"></i>';

        var bubble = document.createElement('span');
        bubble.id = bubbleId;
        bubble.className = 'context-tip__bubble';
        bubble.setAttribute('role', 'tooltip');
        bubble.textContent = text;

        wrap.appendChild(btn);
        wrap.appendChild(bubble);
        return wrap;
    }

    function appendTipToHeading(container, tipId, headingSelector, label) {
        if (!container || container.querySelector('[data-context-tip-mounted="' + tipId + '"]')) return;
        var heading = container.querySelector(headingSelector);
        if (!heading) return;
        var tip = createTipElement(tipId, label);
        if (tip) heading.appendChild(tip);
    }

    function mountTaskTips(root) {
        var tasks = (root || document).querySelectorAll('.dash-task[data-context-tip]');
        tasks.forEach(function (section) {
            var tipId = section.getAttribute('data-context-tip');
            if (!tipId) return;
            var title = section.querySelector('h3');
            var label = title ? 'Hinweis: ' + title.textContent.trim() : 'Kurzer Hinweis';
            appendTipToHeading(section, tipId, 'h3', label);
        });
    }

    function mountToolTips(root) {
        var cards = (root || document).querySelectorAll('#dashboard-tools [data-tool-id]');
        cards.forEach(function (card) {
            var tipId = card.getAttribute('data-tool-id');
            if (!tipId || !TIPS[tipId]) return;
            var titleEl = card.querySelector('h2');
            var label = titleEl ? 'Hinweis: ' + titleEl.textContent.replace(/\s+/g, ' ').trim() : 'Kurzer Hinweis';
            appendTipToHeading(card, tipId, 'h2', label);
        });
    }

    function mountStandaloneTips(root) {
        var nodes = (root || document).querySelectorAll('[data-context-tip]:not(.dash-task):not([data-tool-id])');
        nodes.forEach(function (node) {
            var tipId = node.getAttribute('data-context-tip');
            if (!tipId || node.querySelector('[data-context-tip-mounted="' + tipId + '"]')) return;
            if (node.getAttribute('data-context-tip-target') === 'heading') {
                appendTipToHeading(node.closest('section') || node.parentElement, tipId, 'h2, h3', 'Kurzer Hinweis');
                return;
            }
            var tip = createTipElement(tipId, 'Kurzer Hinweis', node.classList.contains('context-tip--align-right'));
            if (tip) node.appendChild(tip);
        });
    }

    function mountProgressTips(root) {
        var rows = (root || document).querySelectorAll('.dash-task-progress-row[data-context-tip]');
        rows.forEach(function (row) {
            var tipId = row.getAttribute('data-context-tip');
            if (!tipId || row.querySelector('[data-context-tip-mounted="' + tipId + '"]')) return;
            var tip = createTipElement(tipId, 'Was bedeutet dieser Fortschritt?', true);
            if (tip) {
                tip.classList.add('context-tip--inline');
                row.appendChild(tip);
            }
        });
    }

    function mountStandTips(root) {
        var stand = (root || document).querySelector('.dash-stand');
        if (stand) {
            appendTipToHeading(stand, 'stammdaten', '.dash-stand__title', 'Hinweis zu Ihrem Stand');
        }
        var hygieneHead = (root || document).querySelector('.dash-stand__hygiene-head h3');
        if (hygieneHead && !hygieneHead.querySelector('[data-context-tip-mounted="datenhygiene"]')) {
            var tip = createTipElement('datenhygiene', 'Hinweis zum Gruppenabgleich');
            if (tip) hygieneHead.appendChild(tip);
        }
    }

    function mountContextTips(root) {
        mountTaskTips(root);
        mountToolTips(root);
        mountStandTips(root);
        mountProgressTips(root);
        mountStandaloneTips(root);
        mounted = true;
    }

    window.ms365ContextTips = {
        TIPS: TIPS,
        mount: mountContextTips,
        createTipElement: createTipElement
    };

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', function () {
            mountContextTips(document);
        });
    } else {
        mountContextTips(document);
    }
})();
