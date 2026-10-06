/**
 * Schulpraxis-Texte für „Was möchten Sie tun?“ (Tool-Karten & Szenario-Einstieg).
 */
import {
    catalogCardDescription,
    catalogToolActionLabel,
    truncateCatalogDescription
} from './dashboard-tool-card.js';
import { specToolDescription } from './dashboard-tool-spec-copy.js';

/** @type {Record<string, string>} */
const TASK_ROW_TOOL_ID_ALIASES = {
    'schueler-sammelgruppe': 'slg-schueler',
    'lehrer-sammelgruppe': 'slg-lehrer',
    'jahrgangsgruppen': 'jahrgang',
    'kursteams': 'kursteams',
    'datenhygiene': 'datenhygiene',
    'verwaltung': 'verwaltung'
};

/** @param {string} href */
export function normalizeToolHref(href) {
    return String(href || '')
        .split('?')[0]
        .split('#')[0]
        .replace(/^\.\//, '')
        .trim()
        .toLowerCase();
}

/** @type {Record<string, { title: string, desc: string }>} */
export const TOOL_PRACTICE_COPY = {
    'tools/playbook-daten-import-verknuepfen.html': {
        title: 'Playbook: Daten importieren',
        desc: 'Schritt für Schritt von Export bis M365-Verknüpfung'
    },
    'tools/stammdaten-quelle-waehlen.html': {
        title: 'Datenquelle wählen',
        desc: 'Datei, Bildungsportal oder anderes Format festlegen'
    },
    'tools/webuntis-stammdaten-import.html': {
        title: 'Stammdaten aus WebUntis',
        desc: 'Export einlesen und im Schulregister prüfen'
    },
    'tools/bildungsportal-stammdaten.html': {
        title: 'Bildungsportal',
        desc: 'Anbindung und Roadmap für Ihr Bundesland'
    },
    'tools/schueler-sammelgruppe.html': {
        title: 'Alle Schüler:innen in M365',
        desc: 'Sammelgruppe pflegen und mit Microsoft abgleichen'
    },
    'tools/lehrer-sammelgruppe.html': {
        title: 'Alle Lehrkräfte in M365',
        desc: 'Lehrkräfte-Sammelgruppe verwalten und synchronisieren'
    },
    'tools/verwaltung.html': {
        title: 'Verwaltung & Sekretariat',
        desc: 'Verwaltungsgruppe für Office und Teams'
    },
    'tools/klassenvorstaende.html': {
        title: 'Klassenvorstände',
        desc: 'Verantwortliche pro Klasse zuordnen'
    },
    'tools/klassenchats.html': {
        title: 'KlassenChats',
        desc: 'Chats für Klassen anlegen und pflegen'
    },
    'tools/arge-fachgruppen.html': {
        title: 'Fächer und Fachgruppen',
        desc: 'ARGE-Teams zu Fächern verknüpfen'
    },
    'tools/datenhygiene.html': {
        title: 'Datenhygiene',
        desc: 'Stammlisten mit Microsoft 365 abgleichen – Konsistent oder Abweichung'
    },
    'tools/weitere-teams-gruppen.html': {
        title: 'Weitere Teams & Gruppen',
        desc: 'Sonderfälle und zusätzliche Gruppen anlegen'
    },
    'tools/jahrgangsgruppen.html': {
        title: 'Klassengruppen',
        desc: 'Klassen mit M365-Teams verknüpfen'
    },
    'tools/kursteams.html': {
        title: 'Unterrichtsteams',
        desc: 'Teams aus dem Stundenplan anlegen'
    },
    'tools/unterrichtsteams-katalog.html': {
        title: 'Unterrichtsteams-Katalog',
        desc: 'Alle Teams ansehen und nachziehen'
    },
    'tools/webuntis-sync-monitor.html': {
        title: 'Stundenplan-Abgleich',
        desc: 'WebUntis-Sync prüfen und überwachen'
    },
    'tools/klassen-merge.html': {
        title: 'Klassen zusammenlegen',
        desc: 'Zwei Klassen in einer Gruppe zusammenführen'
    },
    'tools/onenote-verteilung.html': {
        title: 'OneNote-Inhalte verteilen',
        desc: 'Inhalte an Klassen-Notebooks ausrollen'
    },
    'tools/personen-verwaltung.html': {
        title: 'Personen suchen',
        desc: 'Konten in Microsoft 365 finden'
    },
    'tools/gaeste-verwalten.html': {
        title: 'Gäste verwalten',
        desc: 'Externe Zugänge prüfen und einladen'
    },
    'tools/lizenzverwaltung.html': {
        title: 'Lizenzen zuweisen',
        desc: 'Microsoft-Lizenzen für Personen verwalten'
    },
    'tools/namenskonvention-audit.html': {
        title: 'Einheitliche Benennung',
        desc: 'Anzeigenamen und Kürzel prüfen'
    },
    'tools/schueler-lifecycle.html': {
        title: 'Schüler-Lifecycle',
        desc: 'Zu- und Abgänge im Schuljahr nachverfolgen'
    },
    'tools/playbook-schuljahresstart.html': {
        title: 'Playbook Schuljahresstart',
        desc: 'Checkliste für den Wechsel ins neue Jahr'
    },
    'tools/organisations-assistent.html': {
        title: 'Schuljahr-Assistent',
        desc: 'Jahreswechsel strukturiert vorbereiten'
    },
    'tools/klassen-umbenennen.html': {
        title: 'Klassen umbenennen',
        desc: 'Anzeigenamen für das neue Schuljahr anpassen'
    },
    'tools/cleanup-playbook.html': {
        title: 'Aufräumen zum Schuljahr',
        desc: 'Alte Gruppen und Teams geordnet abschließen'
    },
    'tools/sharepoint-intranet-hub.html': {
        title: 'Schul-Intranet',
        desc: 'Hub und Startseite für die Schule'
    },
    'tools/schularbeiten-planer.html': {
        title: 'Schularbeiten-Termine',
        desc: 'Prüfungstermine planen und veröffentlichen'
    },
    'tools/freistellung-planer.html': {
        title: 'Freistellungsanträge',
        desc: 'Anträge und Genehmigungen organisieren'
    },
    'tools/playbook-intranet.html': {
        title: 'Playbook Intranet',
        desc: 'Intranet Schritt für Schritt einrichten'
    },
    'tools/schulaktivitaeten-planer.html': {
        title: 'Schulaktivitäten',
        desc: 'Veranstaltungen und Termine koordinieren'
    },
    'tools/projektwochen.html': {
        title: 'Projektwochen',
        desc: 'Projektphasen und Teams planen'
    },
    'tools/sharepoint-liste-stammdaten.html': {
        title: 'Stammdaten-Listen (IT)',
        desc: 'SharePoint-Listen mit dem Register synchronisieren'
    },
    'tools/leere-gruppen-report.html': {
        title: 'Leere Gruppen prüfen',
        desc: 'Gruppen ohne Besitzer oder Mitglieder finden'
    },
    'tools/teams-archiv.html': {
        title: 'Teams archivieren',
        desc: 'Alte Teams gesammelt archivieren'
    },
    'tools/schulstruktur-sync.html': {
        title: 'Alle Gruppen',
        desc: 'Gesamtüberblick über Teams und Gruppen'
    },
    'tools/datenlandkarte.html': {
        title: 'Datenlandkarte',
        desc: 'Wo welche Daten liegen – auf einen Blick'
    },
    'tools/stammdaten-backup-abgleich.html': {
        title: 'Backup-Abgleich',
        desc: 'Register mit IT-Sicherung vergleichen'
    },
    'tenant.html': {
        title: 'Schulregister',
        desc: 'Stammdaten, Klassen und Verknüpfungen zentral pflegen'
    }
};

/** Spezialfälle mit Query-Parametern */
const TOOL_PRACTICE_COPY_QUERY = {
    'tools/kursteams.html?mode=single': {
        title: 'Einzelne Teams nachziehen',
        desc: 'Ein Unterrichtsteam manuell anlegen oder reparieren'
    },
    'tools/personen-verwaltung.html?create=1': {
        title: 'Konto anlegen',
        desc: 'Neues Konto für Lehrkraft oder Schüler:in erstellen'
    },
    'tools/schulstruktur-sync.html?mode=struktur': {
        title: 'Strukturplan (SOLL-Baum)',
        desc: 'Zielstruktur der Gruppen modellieren'
    },
    'tools/leere-gruppen-report.html?problem=no-owners': {
        title: 'Gruppen ohne Besitzer',
        desc: 'Verwaiste Gruppen finden und zuweisen'
    }
};

/** @param {string} href */
export function lookupToolCopy(href) {
    const raw = String(href || '').trim();
    const withQuery = raw.replace(/^\.\//, '').toLowerCase();
    if (TOOL_PRACTICE_COPY_QUERY[withQuery]) return TOOL_PRACTICE_COPY_QUERY[withQuery];
    const base = normalizeToolHref(raw);
    return TOOL_PRACTICE_COPY[base] || null;
}

/**
 * @param {HTMLElement} row
 */
export function taskRowToolId(row) {
    if (!row || !row.getAttribute) return '';
    const hygieneId = String(row.getAttribute('data-dash-hygiene-id') || '').trim();
    if (hygieneId) return hygieneId;
    const href = row.getAttribute('href') || '';
    const slug = normalizeToolHref(href).replace(/^tools\//, '').replace(/\.html$/, '');
    return TASK_ROW_TOOL_ID_ALIASES[slug] || slug;
}

/**
 * @param {string} href
 * @param {string} [toolId]
 * @param {string} [titleFallback]
 */
export function taskRowDescription(href, toolId, titleFallback) {
    const id = String(toolId || '').trim();
    const fromSpec = specToolDescription(id);
    if (fromSpec) return truncateCatalogDescription(fromSpec, 100);
    const copy = lookupToolCopy(href);
    if (copy && copy.desc) return copy.desc;
    const fromTip = id ? catalogCardDescription(id, '') : '';
    if (fromTip) return truncateCatalogDescription(fromTip, 100);
    const fb = String(titleFallback || '').trim();
    return fb ? truncateCatalogDescription(fb, 100) : '';
}

/**
 * @param {HTMLElement} label
 */
function plainLabelTitle(label) {
    if (!label) return '';
    const clone = label.cloneNode(true);
    clone.querySelectorAll('.bi, .dash-task-row__open-badge, .context-tip').forEach(function (n) {
        n.remove();
    });
    return stripHygieneNoiseFromTitle(String(clone.textContent || ''));
}

/** Entfernt angehängte Hygiene-Badge-Texte (Akku bei wiederholtem Refresh). */
export function stripHygieneNoiseFromTitle(text) {
    return String(text || '')
        .replace(/(?:⚠\s*)+(?:Abweichung|Offen|Konsistent|Nicht verknüpft|Leere Liste|Abgleich offen)\s*/gi, ' ')
        .replace(/\s+/g, ' ')
        .trim();
}

/**
 * @param {HTMLElement} row
 */
export function plainTaskRowTitle(row) {
    if (!row) return '';
    const titleSpan = row.querySelector('.dash-task-row__title');
    if (titleSpan) {
        const clone = titleSpan.cloneNode(true);
        clone.querySelectorAll('.dash-task-row__open-badge, .context-tip').forEach(function (n) {
            n.remove();
        });
        return stripHygieneNoiseFromTitle(clone.textContent || '');
    }
    return plainLabelTitle(row.querySelector('.dash-task-row__label'));
}

/**
 * @param {HTMLElement} row
 */
function ensureTaskRowBody(row) {
    if (!row) return null;
    let body = row.querySelector('.dash-task-row__body');
    if (body) return body;

    const label = row.querySelector('.dash-task-row__label');
    if (!label) return null;

    body = document.createElement('div');
    body.className = 'dash-task-row__body';

    const head = document.createElement('div');
    head.className = 'dash-task-row__head';

    const icon = label.querySelector('.bi');
    if (icon) head.appendChild(icon);

    let titleSpan = label.querySelector('.dash-task-row__title');
    if (!titleSpan) {
        titleSpan = document.createElement('span');
        titleSpan.className = 'dash-task-row__title';
        titleSpan.textContent = plainLabelTitle(label);
    }
    const openBadge = label.querySelector('.dash-task-row__open-badge');
    if (openBadge) titleSpan.appendChild(openBadge);
    head.appendChild(titleSpan);
    body.appendChild(head);

    let desc = row.querySelector('.dash-task-row__desc');
    if (!desc) {
        desc = document.createElement('span');
        desc.className = 'dash-task-row__desc';
    }
    body.appendChild(desc);

    const deviation = row.querySelector('.dash-task-row__deviation-hint');
    if (deviation) body.appendChild(deviation);

    const foot = document.createElement('div');
    foot.className = 'dash-task-row__foot';

    let status = row.querySelector('.dash-task-row__status');
    if (!status) {
        status = document.createElement('span');
        status.className = 'dash-task-row__status';
    }
    foot.appendChild(status);

    let action = row.querySelector('.dash-task-row__action');
    if (!action) {
        action = document.createElement('span');
        action.className = 'dash-task-row__action';
    }
    foot.appendChild(action);
    body.appendChild(foot);

    row.querySelectorAll('.bi-chevron-right').forEach(function (c) {
        c.remove();
    });
    label.remove();
    row.insertBefore(body, row.firstChild);
    return body;
}

/**
 * @param {HTMLElement} row
 */
function syncTaskRowStatusFromAttrs(row) {
    if (!row) return;
    const hint = String(row.getAttribute('data-hygiene-hint') || '').trim();
    const tone = String(row.getAttribute('data-hygiene-tone') || '').trim();
    const statusEl = row.querySelector('.dash-task-row__status');
    if (!statusEl) return;
    if (!hint) {
        statusEl.textContent = '';
        statusEl.hidden = true;
        statusEl.removeAttribute('data-tone');
        return;
    }
    statusEl.textContent = hint;
    statusEl.hidden = false;
    if (tone) statusEl.setAttribute('data-tone', tone);
    else statusEl.removeAttribute('data-tone');
}

/** @type {Array<{ icon: string, situation: string, taskId: string, href: string, toolLabel: string }>} */
export const DASH_TASK_SCENARIOS = [
    {
        icon: 'bi-person-plus',
        situation: 'Neue Schüler:in aufnehmen',
        taskId: 'dashTaskPersonen',
        href: 'tools/personen-verwaltung.html?create=1',
        toolLabel: 'Konto anlegen'
    },
    {
        icon: 'bi-person-workspace',
        situation: 'Lehrkraft hat keinen Zugang',
        taskId: 'dashTaskPersonen',
        href: 'tools/personen-verwaltung.html?create=1',
        toolLabel: 'Konto anlegen'
    },
    {
        icon: 'bi-calendar2-range',
        situation: 'Neues Schuljahr startet',
        taskId: 'dashTaskSchuljahr',
        href: 'tools/playbook-schuljahresstart.html',
        toolLabel: 'Schuljahresstart'
    },
    {
        icon: 'bi-diagram-3',
        situation: 'Teams-Gruppe für Klasse fehlt',
        taskId: 'dashTaskUnterricht',
        href: 'tools/jahrgangsgruppen.html',
        toolLabel: 'Klassengruppen'
    },
    {
        icon: 'bi-people',
        situation: 'Sammelgruppe stimmt nicht',
        taskId: 'dashTaskGruppen',
        href: 'tools/schueler-sammelgruppe.html',
        toolLabel: 'Schüler:innen-Sammelgruppe'
    }
];

/**
 * @param {HTMLElement} row
 */
export function applyToolPracticeCopyToRow(row) {
    if (!row || !row.getAttribute) return;
    ensureTaskRowBody(row);

    const href = row.getAttribute('href') || '';
    const copy = lookupToolCopy(href);
    const titleSpan = row.querySelector('.dash-task-row__title');
    const plain = plainTaskRowTitle(row);
    const titleText = copy && copy.title ? copy.title : plain;
    if (titleSpan && titleText) {
        titleSpan.querySelectorAll('.dash-task-row__open-badge').forEach(function (n) {
            n.remove();
        });
        titleSpan.textContent = titleText;
    }

    const toolId = taskRowToolId(row);
    const descEl = row.querySelector('.dash-task-row__desc');
    const factDesc = String(row.getAttribute('data-dash-fact-desc') || '').trim();
    const descText = factDesc || taskRowDescription(href, toolId, titleText || plain);
    if (descEl) {
        if (descText) {
            descEl.textContent = descText;
            descEl.hidden = false;
            descEl.classList.toggle('dash-task-row__desc--metric', !!factDesc);
        } else {
            descEl.textContent = '';
            descEl.hidden = true;
            descEl.classList.remove('dash-task-row__desc--metric');
        }
    }
    row.classList.toggle('dash-task-row--live-metric', !!factDesc);

    const actionEl = row.querySelector('.dash-task-row__action');
    if (actionEl) {
        actionEl.textContent = catalogToolActionLabel(toolId, titleText || plain);
    }

    syncTaskRowStatusFromAttrs(row);
}

/**
 * @param {HTMLElement} row
 */
export function syncToolDeviationHint(row) {
    if (!row) return;
    const tone = row.getAttribute('data-hygiene-tone');
    const isWarn = tone === 'warn' || tone === 'mismatch';
    let hint = row.querySelector('.dash-task-row__deviation-hint');

    row.querySelectorAll('.dash-task-row__open-badge').forEach(function (n) {
        n.remove();
    });

    if (!isWarn) {
        if (hint) hint.remove();
        row.classList.remove('dash-split-tool-card--fixes-deviation');
        return;
    }

    row.classList.add('dash-split-tool-card--fixes-deviation');

    if (!hint) {
        hint = document.createElement('span');
        hint.className = 'dash-task-row__deviation-hint';
        const foot = row.querySelector('.dash-task-row__foot');
        const body = row.querySelector('.dash-task-row__body');
        if (foot && body) {
            body.insertBefore(hint, foot);
        } else {
            const desc = row.querySelector('.dash-task-row__desc');
            if (desc && desc.parentElement) {
                desc.parentElement.insertBefore(hint, desc.nextSibling);
            } else {
                row.appendChild(hint);
            }
        }
    }
    hint.textContent = '→ Behebt aktuelle Abweichung';
}

/** @param {HTMLElement} task */
export function enrichTaskToolRows(task) {
    if (!task) return;
    task.querySelectorAll('.dash-task-row').forEach(function (row) {
        applyToolPracticeCopyToRow(row);
        syncTaskRowStatusFromAttrs(row);
        syncToolDeviationHint(row);
    });
}

/**
 * @param {HTMLElement} splitRoot
 * @param {{ select: (taskId: string, opts?: object) => void }} api
 */
export function mountTaskScenarioGuide(splitRoot, api) {
    if (!splitRoot || splitRoot.querySelector('.dash-tasks-split__scenarios')) return;

    const details = document.createElement('details');
    details.className = 'dash-tasks-split__scenarios';
    details.innerHTML =
        '<summary class="dash-tasks-split__scenarios-summary">' +
        '<i class="bi bi-chat-left-text" aria-hidden="true"></i>' +
        '<span>Was ist Ihre Situation?</span>' +
        '<i class="bi bi-chevron-down dash-tasks-split__scenarios-chevron" aria-hidden="true"></i>' +
        '</summary>';

    const list = document.createElement('ul');
    list.className = 'dash-tasks-split__scenarios-list';

    DASH_TASK_SCENARIOS.forEach(function (item) {
        const li = document.createElement('li');
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'dash-tasks-split__scenario-btn';
        btn.innerHTML =
            '<span class="dash-tasks-split__scenario-situation">' +
            '<i class="bi ' +
            item.icon +
            '" aria-hidden="true"></i>' +
            item.situation +
            '</span>' +
            '<span class="dash-tasks-split__scenario-arrow" aria-hidden="true">→</span>' +
            '<span class="dash-tasks-split__scenario-tool">' +
            item.toolLabel +
            '</span>';
        btn.addEventListener('click', function () {
            if (api && typeof api.select === 'function') {
                api.select(item.taskId, { pulse: true });
            }
            window.setTimeout(function () {
                const panel = document.getElementById(item.taskId);
                if (!panel) return;
                const want = item.href.toLowerCase();
                const target = Array.from(panel.querySelectorAll('a.dash-task-row')).find(function (a) {
                    const h = String(a.getAttribute('href') || '').toLowerCase();
                    return h === want;
                });
                if (!target) return;
                try {
                    target.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
                } catch {
                    /* ignore */
                }
                target.classList.add('dash-task--pulse');
                window.setTimeout(function () {
                    target.classList.remove('dash-task--pulse');
                }, 1200);
            }, 80);
            details.open = false;
        });
        li.appendChild(btn);
        list.appendChild(li);
    });

    details.appendChild(list);
    splitRoot.insertBefore(details, splitRoot.firstChild);
}
