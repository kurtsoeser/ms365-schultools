/**
 * Wave 3: schreibt ausgearbeitete Playbook-HTML-Seiten (Apps + Härtung).
 * Einmalig ausführen: node scripts/_gen-playbook-wave3.mjs
 */
import fs from 'node:fs';
import path from 'node:path';

function esc(s) {
    return String(s || '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;');
}

function stepHtml(step, index) {
    const n = index + 1;
    const optional = step.optional ? ' pb-step--optional cp-step--optional' : '';
    const pill = step.pill
        ? ` <span class="pb-pill pb-pill--${step.pill.tone || 'must'}">${esc(step.pill.label)}</span>`
        : '';
    const actions = Array.isArray(step.actions) && step.actions.length
        ? `<div class="pb-step__actions">${step.actions
              .map((a, i) => {
                  const primary = i === 0 ? ' btn-primary' : '';
                  return `<a class="btn${primary}" href="${a.href}"><i class="bi ${a.icon || 'bi-box-arrow-up-right'}"></i>${esc(a.label)}</a>`;
              })
              .join('')}</div>`
        : step.href
          ? `<a class="btn" href="${step.href}"><i class="bi ${step.icon || 'bi-box-arrow-up-right'}"></i>${esc(step.cta || 'Öffnen')}</a>`
          : '';
    return `<article class="pb-step cp-step${optional}" data-cp-step="${step.id}">
                        <div class="pb-step__n">${n}</div>
                        <div>
                            <h3>${esc(step.title)}${pill}</h3>
                            <p>${step.body}</p>
                            <p class="pb-step__ready" data-cp-ready-hint hidden></p>
                            <label class="pb-check"><input type="checkbox" data-cp-check="${step.id}"> Erledigt</label>
                        </div>
                        ${actions}
                    </article>`;
}

function pageHtml(def) {
    const steps = def.steps || [];
    const stepIds = steps.map((s) => s.id);
    const toolbar = (def.toolbar || [])
        .map((t) => `<a class="btn" href="${t.href}"><i class="bi ${t.icon || 'bi-link-45deg'}"></i>${esc(t.label)}</a>`)
        .join('\n                ');
    const lead = def.lead ? `<p class="pb-lead">${def.lead}</p>` : '';
    const readyImport = def.readyModule
        ? `import { ${def.readyFn} } from '../src/shared/${def.readyModule}';`
        : '';
    const readyWire = def.readyFn
        ? `,
            evaluateReadyChecks: function () {
                return ${def.readyFn}(loadTenantSettings());
            }`
        : '';
    const loadSettings = def.readyFn
        ? `
        function loadTenantSettings() {
            try {
                if (typeof window.ms365TenantSettingsLoad === 'function') {
                    return window.ms365TenantSettingsLoad();
                }
            } catch {
                /* ignore */
            }
            return null;
        }
`
        : '';

    return `<!DOCTYPE html>
<html lang="de">
<head>
    <script src="../src/shared/theme-boot.js?v=4"></script>
    <meta charset="UTF-8">
    <script src="../access-config.js"></script>
    <script src="../src/shared/pin-gate.js"></script>
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>MS365-Schul-Tools – ${esc(def.title)}</title>
    <link rel="stylesheet" href="../app.css?v=14">
    <link rel="stylesheet" href="../src/shared/design-tokens.css?v=2">
    <link rel="stylesheet" href="../src/shared/shell-fluent.css?v=10">
    <link rel="stylesheet" href="../src/shared/fluent-blocks.css?v=8">
    <link rel="stylesheet" href="../src/shared/tool-modern.css">
    <link rel="stylesheet" href="../src/shared/playbook-shell.css?v=1">
    <link rel="stylesheet" href="../vendor/bootstrap-icons/bootstrap-icons.min.css" vite-ignore>
</head>
<body class="app-shell-chrome">
    <div class="container page-card tm-page pb-page">
        <div class="header">
            <h1>MS365-Schul-Tools</h1>
            <div class="header-tool-indicator" aria-label="Aktuelles Werkzeug">
                <span class="header-tool-indicator__label">Werkzeug:</span>
                <span class="header-tool-indicator__name"><i class="bi ${def.icon}"></i>${esc(def.toolName)}</span>
            </div>
            <p>${def.intro}</p>
            ${
                def.help
                    ? `<p class="header-help-row"><a href="../hilfe.html#${def.help}" class="header-help-link"><i class="bi bi-question-circle"></i>Hilfe</a></p>`
                    : ''
            }
            <div class="toolbar">
                <a class="btn" href="../index.html"><i class="bi bi-arrow-left"></i>Dashboard</a>
                ${toolbar}
            </div>
        </div>
        <div class="content">
            <section class="tm-panel">
                <div class="tm-panel__head">
                    <i class="bi bi-signpost-2" aria-hidden="true"></i>
                    <div><h3>${esc(def.panelTitle || 'Schritte')}</h3><p id="cpProgress" class="pb-progress">0 / ${steps.length} erledigt</p></div>
                </div>
                ${lead}
                <div class="pb-steps">
                    ${steps.map(stepHtml).join('\n                    ')}
                </div>
                <div class="pb-actions tm-actions">
                    <button type="button" class="fb-btn fb-btn--ghost" id="cpReset"><i class="bi bi-arrow-counterclockwise"></i>Haken zurücksetzen</button>
                </div>
            </section>
        </div>
    </div>
    <script src="../src/shared/theme-toggle.js?v=4" defer></script>
    <script src="../src/shared/app-dialog.js" defer></script>
    <script type="module">
        import { wirePlaybookPage } from '../src/shared/playbook-store.js';
        ${readyImport}
${loadSettings}
        wirePlaybookPage({
            storageKey: '${def.storageKey}',
            stepIds: ${JSON.stringify(stepIds)}${readyWire}
        });
    </script>
</body>
</html>
`;
}

/** @type {Record<string, any>} */
const pages = {
    'playbook-schulaktivitaeten.html': {
        title: 'Playbook Schulaktivitäten',
        toolName: 'Playbook Schulaktivitäten',
        icon: 'bi-calendar2-check',
        intro:
            'Nach der Blaupause: Primär-Site, AKT-Listen, Rechte und Testantrag – bis Lehrkräfte den Schulaktivitäten-Planer nutzen.',
        help: 'tool-playbook-schulaktivitaeten',
        storageKey: 'ms365-playbook-schulaktivitaeten-v1',
        readyModule: 'playbook-tool-ready.js',
        readyFn: 'evaluateAppPlaybookReadyChecks',
        lead: 'Kern: Primär-Site → Listen → Setup/Rechte → Test → Einweisung.',
        toolbar: [
            { href: 'schulaktivitaeten-planer.html', label: 'Planer', icon: 'bi-calendar2-check' },
            { href: 'schulaktivitaeten-setup.html', label: 'Setup', icon: 'bi-gear' }
        ],
        steps: [
            {
                id: 'site',
                title: 'Primär-Site prüfen',
                pill: { label: 'Pflicht', tone: 'must' },
                body: 'Tool-SharePoint aus der Blaupause – Listen liegen auf der Primär-Site (Planer &amp; Apps).',
                href: '../tenant.html#spo-planer',
                icon: 'bi-grid-1x2',
                cta: 'Planer & Apps'
            },
            {
                id: 'listen',
                title: 'SharePoint-Listen anlegen',
                pill: { label: 'Pflicht', tone: 'must' },
                body: 'Schulaktivitäten-Listen auf der Primär-Site anlegen und mit dem Planer verbinden.',
                href: 'sharepoint-liste-schulaktivitaeten.html',
                icon: 'bi-list-ul',
                cta: 'AKT-Listen'
            },
            {
                id: 'rechte',
                title: 'Rechte &amp; Rollen',
                pill: { label: 'Pflicht', tone: 'must' },
                body: 'Wer darf beantragen, genehmigen, administrieren – im Setup und Planer festlegen.',
                href: 'schulaktivitaeten-setup.html',
                icon: 'bi-shield-check',
                cta: 'Setup'
            },
            {
                id: 'test',
                title: 'Testantrag durchspielen',
                pill: { label: 'Pflicht', tone: 'must' },
                body: 'Einen Probe-Antrag anlegen und den Genehmigungsweg prüfen.',
                href: 'schulaktivitaeten-planer.html',
                icon: 'bi-journal-check',
                cta: 'Planer'
            },
            {
                id: 'schulung',
                title: 'Lehrkräfte einweisen',
                pill: { label: 'Optional', tone: 'later' },
                optional: true,
                body: 'Kurzhinweis im Intranet oder Teams – was Lehrkräfte im Planer sehen und tun.',
                href: 'sharepoint-intranet-hub.html',
                icon: 'bi-house-door',
                cta: 'Intranet'
            }
        ]
    },
    'playbook-projektwochen.html': {
        title: 'Playbook Projektwochen',
        toolName: 'Playbook Projektwochen',
        icon: 'bi-calendar2-week',
        intro:
            'Listen, Bookings/Angebote und Freigabe – bis Projektwochen im Alltag laufen. Primär-Site kommt aus der Blaupause.',
        help: 'tool-playbook-projektwochen',
        storageKey: 'ms365-playbook-projektwochen-v1',
        readyModule: 'playbook-tool-ready.js',
        readyFn: 'evaluateAppPlaybookReadyChecks',
        lead: 'Kern: Site → PW-Listen → Bookings/Angebote → Test → Einweisung.',
        toolbar: [
            { href: 'projektwochen.html', label: 'Projektwochen', icon: 'bi-calendar2-week' },
            { href: 'sharepoint-liste-projektwochen.html', label: 'PW-Listen', icon: 'bi-list-ul' }
        ],
        steps: [
            {
                id: 'site',
                title: 'Primär-Site prüfen',
                pill: { label: 'Pflicht', tone: 'must' },
                body: 'Projektwochen-Listen brauchen die gemeinsame Tool-Site (Stammdaten → Planer &amp; Apps).',
                href: '../tenant.html#spo-planer',
                icon: 'bi-grid-1x2',
                cta: 'Planer & Apps'
            },
            {
                id: 'listen',
                title: 'Projektwochen-Listen',
                pill: { label: 'Pflicht', tone: 'must' },
                body: 'Aktionen und Angebote auf der Primär-Site anlegen.',
                href: 'sharepoint-liste-projektwochen.html',
                icon: 'bi-list-ul',
                cta: 'PW-Listen'
            },
            {
                id: 'bookings',
                title: 'Bookings / Angebote',
                pill: { label: 'Pflicht', tone: 'must' },
                body: 'Buchungs- oder Angebotsfluss im Tool prüfen – je nach Schulmodell Bookings oder Listen.',
                href: 'projektwochen.html',
                icon: 'bi-calendar2-check',
                cta: 'Projektwochen'
            },
            {
                id: 'test',
                title: 'Probebuchung / Testangebot',
                pill: { label: 'Pflicht', tone: 'must' },
                body: 'Einen Durchlauf als Lehrkraft/Schüler:in testen, bevor es live geht.',
                href: 'projektwochen.html',
                icon: 'bi-journal-check',
                cta: 'Testen'
            },
            {
                id: 'schulung',
                title: 'Kollegium informieren',
                pill: { label: 'Optional', tone: 'later' },
                optional: true,
                body: 'Kurzanleitung und Link im Intranet oder als Mail.',
                href: 'sharepoint-intranet-hub.html',
                icon: 'bi-house-door',
                cta: 'Intranet'
            }
        ]
    },
    'playbook-einsammler.html': {
        title: 'Playbook Einsammler',
        toolName: 'Playbook Einsammler',
        icon: 'bi-collection',
        intro: 'SharePoint-Liste, Rechte und Alltag – bis Lehrkräfte Einsammlungen starten und Antworten sehen.',
        help: 'tool-playbook-einsammler',
        storageKey: 'ms365-playbook-einsammler-v1',
        readyModule: 'playbook-tool-ready.js',
        readyFn: 'evaluateAppPlaybookReadyChecks',
        lead: 'Kern: Listen → Rechte → Test-Einsammlung → Einweisung.',
        toolbar: [
            { href: 'einsammler.html', label: 'Einsammler', icon: 'bi-collection' },
            { href: 'sharepoint-liste-einsammler.html', label: 'Listen', icon: 'bi-list-ul' }
        ],
        steps: [
            {
                id: 'listen',
                title: 'SharePoint-Liste anlegen',
                pill: { label: 'Pflicht', tone: 'must' },
                body: 'Einsammler-Liste auf der Primär-Site anlegen und im Tool verbinden.',
                href: 'sharepoint-liste-einsammler.html',
                icon: 'bi-list-ul',
                cta: 'Listen'
            },
            {
                id: 'rechte',
                title: 'Rechte prüfen',
                pill: { label: 'Pflicht', tone: 'must' },
                body: 'Wer darf Einsammlungen starten und Antworten sehen – Gruppen/Listenrechte abstimmen.',
                href: 'einsammler.html',
                icon: 'bi-shield-check',
                cta: 'Einsammler'
            },
            {
                id: 'test',
                title: 'Test-Einsammlung',
                pill: { label: 'Pflicht', tone: 'must' },
                body: 'Probe-Einsammlung anlegen, Antwort abgeben und Rücklauf prüfen.',
                href: 'einsammler.html',
                icon: 'bi-journal-check',
                cta: 'Testen'
            },
            {
                id: 'schulung',
                title: 'Lehrkräfte einweisen',
                pill: { label: 'Optional', tone: 'later' },
                optional: true,
                body: 'Kurz: Link, Ablauf, was Schüler:innen sehen.',
                href: 'sharepoint-intranet-hub.html',
                icon: 'bi-house-door',
                cta: 'Intranet'
            }
        ]
    },
    'playbook-elternbrief.html': {
        title: 'Playbook Elternbrief',
        toolName: 'Playbook Elternbrief',
        icon: 'bi-envelope-paper',
        intro:
            'Liste, Verteiler und Versand – getrennt vom allgemeinen Eltern-Playbook und vom Elternsprechtag.',
        help: 'tool-playbook-elternbrief',
        storageKey: 'ms365-playbook-elternbrief-v1',
        readyModule: 'playbook-tool-ready.js',
        readyFn: 'evaluateAppPlaybookReadyChecks',
        lead: 'Kern: Liste → Verteiler → Probebrief → Einweisung.',
        toolbar: [
            { href: 'elternbrief.html', label: 'Elternbrief', icon: 'bi-envelope-paper' },
            { href: 'eltern-verteiler.html', label: 'Verteiler', icon: 'bi-people' }
        ],
        steps: [
            {
                id: 'listen',
                title: 'SharePoint-Liste',
                pill: { label: 'Pflicht', tone: 'must' },
                body: 'Elternbrief-Liste auf der Primär-Site anlegen.',
                href: 'sharepoint-liste-elternbrief.html',
                icon: 'bi-list-ul',
                cta: 'Listen'
            },
            {
                id: 'verteiler',
                title: 'Eltern-Verteiler prüfen',
                pill: { label: 'Pflicht', tone: 'must' },
                body: 'Klassen-/Jahrgangsverteiler aus Stammdaten – ohne Guardians bleibt der Versand leer.',
                href: 'eltern-verteiler.html',
                icon: 'bi-people',
                cta: 'Verteiler'
            },
            {
                id: 'test',
                title: 'Probebrief senden',
                pill: { label: 'Pflicht', tone: 'must' },
                body: 'An eine Testklasse oder sich selbst – Lesebestätigung und Darstellung prüfen.',
                href: 'elternbrief.html',
                icon: 'bi-send',
                cta: 'Elternbrief'
            },
            {
                id: 'schulung',
                title: 'KV / Sekretariat einweisen',
                pill: { label: 'Optional', tone: 'later' },
                optional: true,
                body: 'Wer darf Briefe starten, welche Verteiler sind Standard.',
                href: 'playbook-eltern.html',
                icon: 'bi-chat-heart',
                cta: 'Elternkanal'
            }
        ]
    },
    'playbook-krankmeldung.html': {
        title: 'Playbook Krankmeldung',
        toolName: 'Playbook Krankmeldung',
        icon: 'bi-heart-pulse',
        intro: 'Liste, Rechte und Testmeldung – bis Eltern/Schüler melden und die Schule den Eingang sieht.',
        help: 'tool-playbook-krankmeldung',
        storageKey: 'ms365-playbook-krankmeldung-v1',
        readyModule: 'playbook-tool-ready.js',
        readyFn: 'evaluateAppPlaybookReadyChecks',
        lead: 'Kern: Liste → Rechte → Testmeldung → Einweisung.',
        toolbar: [
            { href: 'krankmeldung.html', label: 'Krankmeldung', icon: 'bi-heart-pulse' },
            { href: 'sharepoint-liste-krankmeldung.html', label: 'Listen', icon: 'bi-list-ul' }
        ],
        steps: [
            {
                id: 'listen',
                title: 'SharePoint-Liste',
                pill: { label: 'Pflicht', tone: 'must' },
                body: 'Krankmeldungs-Liste auf der Primär-Site anlegen und verbinden.',
                href: 'sharepoint-liste-krankmeldung.html',
                icon: 'bi-list-ul',
                cta: 'Listen'
            },
            {
                id: 'rechte',
                title: 'Sichtbarkeit &amp; Rechte',
                pill: { label: 'Pflicht', tone: 'must' },
                body: 'Wer sieht Eingänge (KV, Sekretariat)? Formular-Link für Eltern/Schüler klären.',
                href: 'krankmeldung.html',
                icon: 'bi-shield-check',
                cta: 'Krankmeldung'
            },
            {
                id: 'test',
                title: 'Testmeldung',
                pill: { label: 'Pflicht', tone: 'must' },
                body: 'Formular ausfüllen und prüfen, dass der Eintrag in der Liste erscheint.',
                href: 'krankmeldung-formular.html',
                icon: 'bi-journal-check',
                cta: 'Formular'
            },
            {
                id: 'schulung',
                title: 'Hinweise veröffentlichen',
                pill: { label: 'Optional', tone: 'later' },
                optional: true,
                body: 'Link und Ablauf im Intranet oder Elternkanal.',
                href: 'sharepoint-intranet-hub.html',
                icon: 'bi-house-door',
                cta: 'Intranet'
            }
        ]
    },
    'playbook-srdp.html': {
        title: 'Playbook SRDP-Anmeldung',
        toolName: 'Playbook SRDP-Anmeldung',
        icon: 'bi-mortarboard',
        intro: 'Listen, Rechte, Testanmeldung und optionales Mail-Setup für die SRDP-Anmeldung.',
        help: 'tool-playbook-srdp',
        storageKey: 'ms365-playbook-srdp-v1',
        readyModule: 'playbook-tool-ready.js',
        readyFn: 'evaluateAppPlaybookReadyChecks',
        lead: 'Kern: Listen → Rechte → Test → Mail (optional) → Einweisung.',
        toolbar: [
            { href: 'srdp-anmeldung.html', label: 'SRDP', icon: 'bi-mortarboard' },
            { href: 'sharepoint-liste-srdp.html', label: 'Listen', icon: 'bi-list-ul' }
        ],
        steps: [
            {
                id: 'listen',
                title: 'SRDP-Listen anlegen',
                pill: { label: 'Pflicht', tone: 'must' },
                body: 'Anmeldelisten auf der Primär-Site anlegen.',
                href: 'sharepoint-liste-srdp.html',
                icon: 'bi-list-ul',
                cta: 'Listen'
            },
            {
                id: 'rechte',
                title: 'Rechte &amp; Fristen',
                pill: { label: 'Pflicht', tone: 'must' },
                body: 'Wer darf anmelden, wer sieht Auswertungen – Fristen im Tool setzen.',
                href: 'srdp-anmeldung.html',
                icon: 'bi-shield-check',
                cta: 'SRDP'
            },
            {
                id: 'test',
                title: 'Testanmeldung',
                pill: { label: 'Pflicht', tone: 'must' },
                body: 'Probe-Anmeldung als Schüler:in und Sicht in der Verwaltung prüfen.',
                href: 'srdp-anmeldung.html',
                icon: 'bi-journal-check',
                cta: 'Testen'
            },
            {
                id: 'mail',
                title: 'Mail-Setup',
                pill: { label: 'Optional', tone: 'later' },
                optional: true,
                body: 'Bestätigungs- oder Erinnerungsmails, falls die Schule das nutzt.',
                href: 'srdp-mail-setup.html',
                icon: 'bi-envelope',
                cta: 'Mail'
            },
            {
                id: 'schulung',
                title: 'Klassenlehrer:innen informieren',
                pill: { label: 'Optional', tone: 'later' },
                optional: true,
                body: 'Link, Fristen und Support-Weg kurz kommunizieren.',
                href: 'sharepoint-intranet-hub.html',
                icon: 'bi-house-door',
                cta: 'Intranet'
            }
        ]
    },
    'playbook-wahlen.html': {
        title: 'Playbook Wahlen',
        toolName: 'Playbook Wahlen (Schulsprecher &amp; Wahlfach)',
        icon: 'bi-ui-checks-grid',
        intro:
            'Zwei verwandte Wahl-Tools in einem Playbook: Schulsprecherwahl und Wahlfach-Zuteilung – Konzept, Listen, Test.',
        help: 'tool-playbook-wahlen',
        storageKey: 'ms365-playbook-wahlen-v1',
        readyModule: 'playbook-tool-ready.js',
        readyFn: 'evaluateAppPlaybookReadyChecks',
        lead: 'Kern: Konzept → Listen → Rechte → Test → Einweisung. Schulsprecher und Wahlfach getrennt öffnen.',
        toolbar: [
            { href: 'schulsprecherwahl.html', label: 'Schulsprecher', icon: 'bi-people' },
            { href: 'wahlfach-zuteilung.html', label: 'Wahlfach', icon: 'bi-ui-checks' }
        ],
        steps: [
            {
                id: 'konzept',
                title: 'Wahlmodell klären',
                pill: { label: 'Pflicht', tone: 'must' },
                body: 'Welche Wahl läuft wann? Schulsprecher und/oder Wahlfach – Fristen und Zielgruppen festlegen.',
                href: 'schulsprecherwahl.html',
                icon: 'bi-diagram-3',
                cta: 'Schulsprecher'
            },
            {
                id: 'listen',
                title: 'Listen anlegen',
                pill: { label: 'Pflicht', tone: 'must' },
                body: 'Wahlfach-Listen (und ggf. Schulsprecher-Setup) auf der Primär-Site.',
                actions: [
                    {
                        href: 'sharepoint-liste-wahlfach.html',
                        label: 'Wahlfach-Listen',
                        icon: 'bi-list-ul'
                    },
                    { href: 'schulsprecherwahl.html', label: 'Schulsprecher', icon: 'bi-people' }
                ]
            },
            {
                id: 'rechte',
                title: 'Rechte &amp; Fenster',
                pill: { label: 'Pflicht', tone: 'must' },
                body: 'Wer sieht Ergebnisse, wann ist die Stimmabgabe offen.',
                actions: [
                    {
                        href: 'wahlfach-zuteilung.html',
                        label: 'Wahlfach',
                        icon: 'bi-ui-checks'
                    },
                    {
                        href: 'schulsprecherwahl.html',
                        label: 'Schulsprecher',
                        icon: 'bi-people'
                    }
                ]
            },
            {
                id: 'test',
                title: 'Probeabstimmung',
                pill: { label: 'Pflicht', tone: 'must' },
                body: 'Eine Stimme bzw. Wahlfach-Anmeldung testen und Auszählung prüfen.',
                href: 'schulsprecherwahl-stimme.html',
                icon: 'bi-journal-check',
                cta: 'Stimme testen'
            },
            {
                id: 'schulung',
                title: 'Klassen informieren',
                pill: { label: 'Optional', tone: 'later' },
                optional: true,
                body: 'Link und Zeitraum an KV/Schüler:innen – Intranet oder Teams.',
                href: 'sharepoint-intranet-hub.html',
                icon: 'bi-house-door',
                cta: 'Intranet'
            }
        ]
    },
    'playbook-intranet.html': {
        title: 'Playbook Intranet',
        toolName: 'Playbook Intranet',
        icon: 'bi-house-door',
        intro:
            'Hub und Listen einmal aufsetzen – danach Planer nutzen. Primär-Site und Blaupause sollten stehen; SA/PW haben eigene Playbooks für den Feinschliff.',
        help: 'tool-playbook-intranet',
        storageKey: 'ms365-playbook-intranet-v1',
        readyModule: 'playbook-tool-ready.js',
        readyFn: 'evaluateAppPlaybookReadyChecks',
        lead: 'Empfohlen nach Blaupause: Hub → Basislisten → optional SA/PW-Listen → Planer öffnen.',
        toolbar: [
            { href: 'sharepoint-intranet-hub.html', label: 'Intranet-Hub', icon: 'bi-house-door' },
            { href: '../tenant.html#spo-hub', label: 'Hub-URL', icon: 'bi-link-45deg' }
        ],
        steps: [
            {
                id: 'hub',
                title: 'Schul-Intranet / Hub',
                pill: { label: 'Pflicht', tone: 'must' },
                body: 'Kommunikationssite und Hub anlegen oder verbinden – Einstieg für News und Links.',
                href: 'sharepoint-intranet-hub.html',
                icon: 'bi-house-door',
                cta: 'Hub'
            },
            {
                id: 'lehrer',
                title: 'Lehrerliste',
                pill: { label: 'Pflicht', tone: 'must' },
                body: 'Lehrkräfte aus Stammdaten als SharePoint-Liste für Intranet und Planer.',
                href: 'sharepoint-liste-lehrer.html',
                icon: 'bi-table',
                cta: 'Lehrerliste'
            },
            {
                id: 'termine',
                title: 'Schultermine',
                pill: { label: 'Pflicht', tone: 'must' },
                body: 'Terminliste inkl. Sync-Hinweise – Grundlage für Kalender und Sprechtag.',
                href: 'sharepoint-liste-schultermine.html',
                icon: 'bi-calendar3',
                cta: 'Termine'
            },
            {
                id: 'sa',
                title: 'Schularbeiten-Listen',
                pill: { label: 'Optional', tone: 'later' },
                optional: true,
                body: 'Paket anlegen – Detail-Setup im Schularbeiten-Playbook.',
                actions: [
                    {
                        href: 'sharepoint-liste-schularbeiten.html',
                        label: 'SA-Listen',
                        icon: 'bi-list-ul'
                    },
                    {
                        href: 'playbook-schularbeiten-planer.html',
                        label: 'SA-Playbook',
                        icon: 'bi-journal-check'
                    }
                ]
            },
            {
                id: 'pw',
                title: 'Projektwochen-Listen',
                pill: { label: 'Optional', tone: 'later' },
                optional: true,
                body: 'Aktionen/Angebote – Feinschliff im Projektwochen-Playbook.',
                actions: [
                    {
                        href: 'sharepoint-liste-projektwochen.html',
                        label: 'PW-Listen',
                        icon: 'bi-calendar2-week'
                    },
                    {
                        href: 'playbook-projektwochen.html',
                        label: 'PW-Playbook',
                        icon: 'bi-journal-check'
                    }
                ]
            },
            {
                id: 'vertretung',
                title: 'Vertretungsplan-Liste',
                pill: { label: 'Optional', tone: 'later' },
                optional: true,
                body: 'Tägliche Kurzinfo für Lehrkräfte, falls die Schule das nutzt.',
                href: 'sharepoint-liste-vertretung.html',
                icon: 'bi-arrow-repeat',
                cta: 'Vertretung'
            },
            {
                id: 'nutzen',
                title: 'Planer öffnen',
                pill: { label: 'Pflicht', tone: 'must' },
                body: 'Nach den Listen: Schularbeiten- oder Projektwochen-Planer im Alltag nutzen.',
                actions: [
                    {
                        href: 'schularbeiten-planer.html',
                        label: 'SA-Planer',
                        icon: 'bi-journal-check'
                    },
                    {
                        href: 'projektwochen.html',
                        label: 'Projektwochen',
                        icon: 'bi-calendar2-week'
                    }
                ]
            }
        ]
    },
    'playbook-kursteams.html': {
        title: 'Playbook Unterrichtsteams',
        toolName: 'Playbook Unterrichtsteams',
        icon: 'bi-mortarboard',
        intro:
            'Nach Import und Klassengruppen: Kursteams anlegen, Vorlagen, Monitor und Katalog. Am Schuljahresanfang oft parallel zum Schuljahresstart-Playbook.',
        help: 'tool-playbook-kursteams',
        storageKey: 'ms365-playbook-kursteams-v1',
        readyModule: 'playbook-tool-ready.js',
        readyFn: 'evaluateKursteamsPlaybookReadyChecks',
        lead: 'Kern: Stammdaten/Import → SLG/Klassen → Kursteams → Vorlagen → Monitor → Katalog.',
        toolbar: [
            { href: 'kursteams.html', label: 'Kursteams', icon: 'bi-rocket-takeoff' },
            {
                href: 'playbook-schuljahresstart.html',
                label: 'Schuljahresstart',
                icon: 'bi-rocket-takeoff'
            }
        ],
        steps: [
            {
                id: 'import',
                title: 'Stammdaten &amp; Belegung',
                pill: { label: 'Pflicht', tone: 'must' },
                body: 'Lehrkräfte, Fächer und Klassen importieren oder pflegen – Grundlage für Teams.',
                href: 'webuntis-stammdaten-import.html',
                icon: 'bi-box-arrow-in-down',
                cta: 'Import'
            },
            {
                id: 'klassen',
                title: 'Klassengruppen &amp; SLG',
                pill: { label: 'Pflicht', tone: 'must' },
                body: 'Schüler in Klassen-Teams – Basis für Kursteams-Mitgliedschaft.',
                href: 'schueler-lehrer-gruppen.html',
                icon: 'bi-people',
                cta: 'SLG'
            },
            {
                id: 'kursteams',
                title: 'Kursteams anlegen',
                pill: { label: 'Pflicht', tone: 'must' },
                body: 'Massenanlage aus Stundenplan oder einzelne Teams nachziehen.',
                href: 'kursteams.html',
                icon: 'bi-rocket-takeoff',
                cta: 'Anlegen'
            },
            {
                id: 'templates',
                title: 'Vorlagen anwenden',
                pill: { label: 'Optional', tone: 'later' },
                optional: true,
                body: 'Kanal-Struktur je Fach/Schulstufe auf Teams übertragen.',
                href: 'kursteam-templates.html',
                icon: 'bi-layout-text-sidebar-reverse',
                cta: 'Vorlagen'
            },
            {
                id: 'monitor',
                title: 'Sync-Monitor &amp; Lücken',
                pill: { label: 'Pflicht', tone: 'must' },
                body: 'Fehlende Teams, Besitzer und Lehrer-Matches prüfen.',
                href: 'webuntis-sync-monitor.html',
                icon: 'bi-activity',
                cta: 'Monitor'
            },
            {
                id: 'katalog',
                title: 'Katalog pflegen',
                pill: { label: 'Optional', tone: 'later' },
                optional: true,
                body: 'Namen, Filter und optional Anzeigenamen in Entra aktualisieren.',
                href: 'unterrichtsteams-katalog.html',
                icon: 'bi-table',
                cta: 'Katalog'
            }
        ]
    },
    'playbook-eltern.html': {
        title: 'Playbook Elternkommunikation',
        toolName: 'Playbook Elternkommunikation',
        icon: 'bi-chat-heart',
        intro:
            'Kanal-Überblick: Guardians, Verteiler, optional Sprechtag und Intranet. Der Elternsprechtag hat ein eigenes Playbook; Elternbrief ist eine eigene App-Checkliste.',
        help: 'tool-playbook-eltern',
        storageKey: 'ms365-playbook-eltern-v1',
        readyModule: 'playbook-tool-ready.js',
        readyFn: 'evaluateElternPlaybookReadyChecks',
        lead: 'Kern: Guardians → Verteiler. Sprechtag, Intranet und Automationen optional.',
        toolbar: [
            { href: 'eltern-verteiler.html', label: 'Eltern-Verteiler', icon: 'bi-people' },
            {
                href: 'playbook-elternsprechtag.html',
                label: 'Sprechtag-Playbook',
                icon: 'bi-calendar2-heart'
            },
            { href: 'playbook-elternbrief.html', label: 'Elternbrief', icon: 'bi-envelope-paper' }
        ],
        steps: [
            {
                id: 'guardians',
                title: 'Erziehungsberechtigte zuordnen',
                pill: { label: 'Pflicht', tone: 'must' },
                body: 'Guardians in Stammdaten/Import prüfen – Basis für Verteiler und Briefe.',
                href: '../tenant.html#schueler',
                icon: 'bi-person-lines-fill',
                cta: 'Stammdaten'
            },
            {
                id: 'verteiler',
                title: 'Eltern-Verteiler',
                pill: { label: 'Pflicht', tone: 'must' },
                body: 'Klassen- und Jahrgangsverteiler erzeugen.',
                href: 'eltern-verteiler.html',
                icon: 'bi-people',
                cta: 'Verteiler'
            },
            {
                id: 'bookings',
                title: 'Elternsprechtag (Bookings)',
                pill: { label: 'Optional', tone: 'later' },
                optional: true,
                body: 'Detailablauf im eigenen Sprechtag-Playbook – hier nur Einstieg.',
                actions: [
                    {
                        href: 'playbook-elternsprechtag.html',
                        label: 'Sprechtag-Playbook',
                        icon: 'bi-calendar2-heart'
                    },
                    {
                        href: 'elternsprechtag-bookings.html',
                        label: 'Bookings',
                        icon: 'bi-calendar2-check'
                    }
                ]
            },
            {
                id: 'intranet',
                title: 'Intranet-News / Seiten',
                pill: { label: 'Optional', tone: 'later' },
                optional: true,
                body: 'Elterninfos auf der Schulwebsite pflegen.',
                href: 'sharepoint-intranet-hub.html',
                icon: 'bi-house-door',
                cta: 'Intranet'
            },
            {
                id: 'pa',
                title: 'Erinnerungs-Automation',
                pill: { label: 'Optional', tone: 'later' },
                optional: true,
                body: 'Power-Automate-Rezepte für Termine/Mails, falls gewünscht.',
                href: 'power-automate-rezepte.html',
                icon: 'bi-lightning-charge',
                cta: 'Automationen'
            }
        ]
    },
    'playbook-elternsprechtag.html': {
        title: 'Playbook Elternsprechtag',
        toolName: 'Playbook Elternsprechtag',
        icon: 'bi-calendar2-heart',
        intro:
            'Fokussiert auf den Sprechtag: Termin, Guardians, Verteiler, Bookings, Info und Nachbereitung – getrennt vom allgemeinen Elternkanal.',
        help: 'tool-playbook-elternsprechtag',
        storageKey: 'ms365-playbook-elternsprechtag-v1',
        readyModule: 'playbook-tool-ready.js',
        readyFn: 'evaluateElternPlaybookReadyChecks',
        lead: 'Kern: Termin → Guardians → Verteiler → Bookings → Info. Nachbereitung optional.',
        toolbar: [
            {
                href: 'elternsprechtag-bookings.html',
                label: 'Bookings',
                icon: 'bi-calendar2-heart'
            },
            { href: 'playbook-eltern.html', label: 'Elternkommunikation', icon: 'bi-chat-heart' }
        ],
        steps: [
            {
                id: 'termin',
                title: 'Termin &amp; Ablauf planen',
                pill: { label: 'Pflicht', tone: 'must' },
                body: 'Datum, Zeitfenster, Raumlogik und Zuständigkeiten festhalten.',
                href: 'sharepoint-liste-schultermine.html',
                icon: 'bi-calendar3',
                cta: 'Schultermine'
            },
            {
                id: 'guardians',
                title: 'Erziehungsberechtigte prüfen',
                pill: { label: 'Pflicht', tone: 'must' },
                body: 'Guardians in Stammdaten – Basis für Verteiler und Ansprache.',
                href: '../tenant.html#schueler',
                icon: 'bi-person-lines-fill',
                cta: 'Stammdaten'
            },
            {
                id: 'verteiler',
                title: 'Eltern-Verteiler',
                pill: { label: 'Pflicht', tone: 'must' },
                body: 'Klassen- und Jahrgangsverteiler für Einladungen vorbereiten.',
                href: 'eltern-verteiler.html',
                icon: 'bi-people',
                cta: 'Verteiler'
            },
            {
                id: 'bookings',
                title: 'Bookings einrichten',
                pill: { label: 'Pflicht', tone: 'must' },
                body: 'Eine Umgebung, 10‑Min-Dienst, Lehrkräfte als Mitarbeiter.',
                href: 'elternsprechtag-bookings.html',
                icon: 'bi-calendar2-heart',
                cta: 'Bookings'
            },
            {
                id: 'lehrer',
                title: 'Lehrkräfte einbinden',
                pill: { label: 'Pflicht', tone: 'must' },
                body: 'Kurz-Anleitung: Kalender, Buchungslink, Vertretungen.',
                href: 'sharepoint-intranet-hub.html',
                icon: 'bi-house-door',
                cta: 'Intranet'
            },
            {
                id: 'kommunikation',
                title: 'Eltern informieren',
                pill: { label: 'Pflicht', tone: 'must' },
                body: 'News auf dem Intranet, optional Verteiler-Mail mit Buchungslink.',
                href: 'playbook-eltern.html',
                icon: 'bi-chat-heart',
                cta: 'Elternkanal'
            },
            {
                id: 'nachbereitung',
                title: 'Nach dem Sprechtag',
                pill: { label: 'Optional', tone: 'later' },
                optional: true,
                body: 'Auslastung prüfen, Bookings fürs nächste Mal notieren, ggf. Automation anpassen.',
                href: 'power-automate-rezepte.html',
                icon: 'bi-lightning-charge',
                cta: 'Automationen'
            }
        ]
    }
};

const toolsDir = path.resolve('tools');
for (const [file, def] of Object.entries(pages)) {
    const out = path.join(toolsDir, file);
    fs.writeFileSync(out, pageHtml(def));
    console.log('wrote', file);
}
console.log('done', Object.keys(pages).length);
