/**
 * Reine Playbook-Logik für „Schuljahr wechseln“.
 * Kein DOM, kein localStorage, kein Graph.
 */

export const PLAYBOOK_STEP_IDS = [
    'year',
    'names',
    'graduates',
    'students',
    'kursteams',
    'subjects',
    'expert'
];

/** Schritte, die in der Alltags-Checkliste zählen (ohne Experten-Baum). */
export const PLAYBOOK_REQUIRED_IDS = PLAYBOOK_STEP_IDS.filter((id) => id !== 'expert');

/**
 * Statische Schritt-Definitionen. `href` ist relativ zu `tools/`.
 * @returns {Array<{ id: string, title: string, blurb: string, href: string, hrefLabel: string, optional?: boolean, later?: boolean }>}
 */
export function playbookStepDefs() {
    return [
        {
            id: 'year',
            title: 'Neues Schuljahr aktivieren',
            blurb:
                'Setzt das aktuelle Schuljahr in den Stammdaten. Optional werden Schüler- und Klassenliste aus dem Vorjahr kopiert. Erste Klassen ergänzen Sie danach in den Stammdaten.',
            href: '#year',
            hrefLabel: 'Schritt aufklappen'
        },
        {
            id: 'names',
            title: 'Klassen-Anzeigenamen anpassen',
            blurb:
                'Die Klassengruppe bleibt dieselbe (Mail-Alias z. B. jg2031hma). Nur der Anzeigename wechselt, etwa 1HMA → 2HMA. Schreiben nach Microsoft 365 ändert nur den Anzeigenamen.',
            href: '#namen',
            hrefLabel: 'Zum Formular'
        },
        {
            id: 'graduates',
            title: 'Abschlussjahrgang prüfen',
            blurb:
                'Klassengruppen bleiben bestehen. Zusätzlich können Sie eine Gruppe für alle Klassen desselben Abschlussjahrs anlegen oder verknüpfen. „Abschlussjahrgang“ ist diese Sammelgruppe — nicht die einzelne Klassengruppe.',
            href: '#abschluss',
            hrefLabel: 'Zum Abschlussjahrgang'
        },
        {
            id: 'students',
            title: 'Schülerliste und Sammelgruppe',
            blurb:
                'Schülerinnen und Schüler für das neue Jahr in den Stammdaten pflegen, danach die Sammelgruppe „alle Schülerinnen“ abgleichen. Ist die Sammelgruppe verknüpft, wird dieser Schritt automatisch abgehakt.',
            href: 'schueler-lehrer-gruppen.html',
            hrefLabel: 'Schüler und Lehrkräfte'
        },
        {
            id: 'kursteams',
            title: 'Unterrichtsteams neu anlegen',
            blurb:
                'Unterrichtsteams hängen am Stundenplan und sind jahresgebunden – typisch neu importieren und per CSV/CMD anlegen. Alte Teams nicht löschen, sondern später im Team-Archiv als Sammel-Aktion archivieren (Suchbegriff, z. B. Schuljahr-Präfix). Nach der Anlage hier manuell abhaken.',
            href: 'kursteams.html',
            hrefLabel: 'Unterrichtsteams'
        },
        {
            id: 'subjects',
            title: 'Fachgruppen und ARGEs prüfen',
            blurb:
                'Fach- und ARGE-Gruppen ändern sich nicht mit der Schulstufe. Prüfen Sie, ob alle Fächer und ARGEs aus den Stammdaten eine Microsoft-365-Gruppe haben. Sind alle verknüpft, wird dieser Schritt automatisch abgehakt.',
            href: 'arge-fachgruppen.html',
            hrefLabel: 'Fächer und ARGEs'
        },
        {
            id: 'expert',
            title: 'Optional: SOLL-Struktur (Experten)',
            blurb:
                'Nur wenn Sie den Baum in der Gruppenverwaltung nutzen. Für Klassengruppen den Baum nicht duplizieren – der Mail-Alias soll bleiben. Der Schuljahreswechsel läuft über Anzeigenamen auf dieser Seite.',
            href: 'schulstruktur-sync.html?mode=struktur',
            hrefLabel: 'Gruppenverwaltung',
            optional: true
        }
    ];
}

export function parseSchoolYearStartYear(label) {
    const m = String(label || '')
        .trim()
        .match(/^(\d{4})\s*\/\s*(\d{2}|\d{4})/);
    return m ? parseInt(m[1], 10) : NaN;
}

export function currentSchoolYearLabel(now) {
    const d = now instanceof Date ? now : new Date();
    const y = d.getFullYear();
    return String(y) + '/' + String(y + 1).slice(2);
}

export function nextSchoolYearLabel(cur) {
    const y = parseSchoolYearStartYear(cur);
    if (!isFinite(y)) return currentSchoolYearLabel();
    return String(y + 1) + '/' + String(y + 2).slice(2);
}

export function isSchoolYearLabel(label) {
    return isFinite(parseSchoolYearStartYear(label));
}

function emptyDone() {
    const done = {};
    PLAYBOOK_STEP_IDS.forEach((id) => {
        done[id] = false;
    });
    return done;
}

/**
 * @param {unknown} raw
 * @param {string} [currentYear]
 */
export function normalizePlaybook(raw, currentYear) {
    const src = raw && typeof raw === 'object' ? raw : {};
    const targetYear = String(src.targetYear || currentYear || '').trim();
    const doneIn = src.done && typeof src.done === 'object' ? src.done : {};
    const done = emptyDone();
    PLAYBOOK_STEP_IDS.forEach((id) => {
        done[id] = doneIn[id] === true;
    });
    if (targetYear && currentYear && targetYear === String(currentYear).trim()) {
        done.year = true;
    }
    return { targetYear, done };
}

export function markPlaybookStep(playbook, stepId, isDone) {
    const cur = normalizePlaybook(playbook);
    if (PLAYBOOK_STEP_IDS.indexOf(stepId) < 0) return cur;
    const done = Object.assign({}, cur.done);
    done[stepId] = !!isDone;
    return { targetYear: cur.targetYear, done };
}

/**
 * Nach dem Aktivieren eines Schuljahrs: Jahr-Schritt erledigt.
 * Wechselt das Zieljahr, werden die übrigen Häkchen zurückgesetzt.
 */
export function applyYearActivated(playbook, newYear) {
    const year = String(newYear || '').trim();
    const cur = normalizePlaybook(playbook);
    const done = emptyDone();
    if (cur.targetYear && cur.targetYear === year) {
        Object.assign(done, cur.done);
    }
    done.year = true;
    return { targetYear: year, done };
}

export function playbookProgress(playbook, ids) {
    const cur = normalizePlaybook(playbook);
    const list = Array.isArray(ids) && ids.length ? ids : PLAYBOOK_REQUIRED_IDS;
    let doneN = 0;
    list.forEach((id) => {
        if (cur.done[id]) doneN += 1;
    });
    const total = list.length;
    const pct = total ? Math.round((doneN / total) * 100) : 0;
    return { done: doneN, total, pct };
}

function normCatalogCode(v) {
    return String(v || '')
        .trim()
        .toLowerCase()
        .replace(/\s+/g, '');
}

function catalogLinkMatched(links, kind, code) {
    const want = kind === 'sammelgruppe' ? String(code || '').trim() : normCatalogCode(code);
    if (!want) return false;
    return (Array.isArray(links) ? links : []).some(function (l) {
        if (!l || l.kind !== kind || !l.graphGroupId) return false;
        if (kind === 'sammelgruppe') return String(l.code || '').trim() === want;
        return normCatalogCode(l.code) === want;
    });
}

/**
 * Leitet Statushinweise und Auto-Häkchen aus Stammdaten ab (ohne Graph).
 * @param {object|null|undefined} container app-data-v2-Container
 * @returns {{ students: boolean, subjects: boolean, hints: Record<string, string> }}
 */
export function inferPlaybookFromContainer(container) {
    const hints = {
        students: 'Noch keine Schülerliste bzw. Sammelgruppe für dieses Jahr erkannt.',
        subjects: 'Noch keine Fächer/ARGEs in den Stammdaten – oder noch nicht alle verknüpft.',
        kursteams:
            'Unterrichtsteams werden per CSV/CMD angelegt. Nach erfolgreicher Anlage diesen Schritt manuell abhaken.'
    };
    let students = false;
    let subjects = false;
    if (!container || typeof container !== 'object') {
        return { students, subjects, hints };
    }

    const setup = container.setup && typeof container.setup === 'object' ? container.setup : {};
    const matched = setup.matched && typeof setup.matched === 'object' ? setup.matched : {};
    const links = Array.isArray(setup.catalogLinks) ? setup.catalogLinks : [];
    const core = container.core && typeof container.core === 'object' ? container.core : {};
    const years = container.years && typeof container.years === 'object' ? container.years : {};
    const current = String(years.current || '').trim();
    const bucket =
        current && years.byLabel && typeof years.byLabel === 'object' ? years.byLabel[current] : null;
    const studentCount = bucket && Array.isArray(bucket.students) ? bucket.students.length : 0;

    const schuelerLinked =
        !!(matched.schuelerGroupId && String(matched.schuelerGroupId).trim()) ||
        catalogLinkMatched(links, 'sammelgruppe', 'schueler');
    if (schuelerLinked) {
        students = true;
        hints.students = studentCount
            ? 'Erkannt: Sammelgruppe verknüpft, ' + studentCount + ' Schüler:innen in den Stammdaten.'
            : 'Erkannt: Sammelgruppe „alle Schülerinnen“ ist verknüpft. Schülerliste ggf. noch aktualisieren.';
    } else if (studentCount) {
        hints.students =
            studentCount +
            ' Schüler:innen in den Stammdaten – Sammelgruppe noch nicht verknüpft. Werkzeug öffnen und abgleichen.';
    }

    const subjectRows = Array.isArray(core.subjects) ? core.subjects : [];
    const argeRows = Array.isArray(core.arges) ? core.arges : [];
    const catalogItems = [];
    subjectRows.forEach(function (row) {
        const code = row && typeof row === 'object' ? row.code : row;
        if (normCatalogCode(code)) catalogItems.push({ kind: 'subject', code: code });
    });
    argeRows.forEach(function (row) {
        const code = row && typeof row === 'object' ? row.code : row;
        if (normCatalogCode(code)) catalogItems.push({ kind: 'arge', code: code });
    });
    if (catalogItems.length) {
        let linkedN = 0;
        catalogItems.forEach(function (item) {
            if (catalogLinkMatched(links, item.kind, item.code)) linkedN += 1;
        });
        if (linkedN === catalogItems.length) {
            subjects = true;
            hints.subjects =
                'Erkannt: alle ' + catalogItems.length + ' Fächer/ARGEs sind mit einer Microsoft-365-Gruppe verknüpft.';
        } else {
            hints.subjects =
                linkedN +
                ' von ' +
                catalogItems.length +
                ' Fächern/ARGEs verknüpft – Rest unter „Fächer und ARGEs“ nachziehen.';
        }
    }

    return { students, subjects, hints };
}

/**
 * Setzt Auto-Häkchen für Schüler- und Fach-Schritte, wenn Stammdaten das erlauben.
 * @param {unknown} playbook
 * @param {{ students?: boolean, subjects?: boolean }|null|undefined} inferred
 * @param {string} [currentYear]
 */
export function applyInferredPlaybookDone(playbook, inferred, currentYear) {
    const cur = normalizePlaybook(playbook, currentYear);
    const done = Object.assign({}, cur.done);
    if (inferred && inferred.students) done.students = true;
    if (inferred && inferred.subjects) done.subjects = true;
    return { targetYear: cur.targetYear, done };
}

/**
 * Offene Punkte nach dem Schuljahreswechsel (lokal, ohne Graph).
 * @param {object|null|undefined} container
 * @param {unknown} playbook
 * @returns {{ id: string, text: string, href: string }[]}
 */
export function openItemsFromContainer(container, playbook) {
    const items = [];
    const c = container && typeof container === 'object' ? container : {};
    const inferred = inferPlaybookFromContainer(c);
    const currentYear = c.years ? String(c.years.current || '').trim() : '';
    const pb = applyInferredPlaybookDone(normalizePlaybook(playbook, currentYear), inferred, currentYear);
    const defs = playbookStepDefs();
    PLAYBOOK_REQUIRED_IDS.forEach(function (id) {
        if (pb.done[id]) return;
        const def = defs.find(function (d) {
            return d.id === id;
        });
        items.push({
            id: 'step-' + id,
            text: 'Checkliste: ' + (def ? def.title : id) + ' noch offen',
            href: def && def.href ? def.href : '#year'
        });
    });

    const years = c.years && typeof c.years === 'object' ? c.years : {};
    const bucket =
        currentYear && years.byLabel && typeof years.byLabel === 'object' ? years.byLabel[currentYear] : null;
    const classes = bucket && Array.isArray(bucket.classes) ? bucket.classes : [];
    const teams = c.core && Array.isArray(c.core.classTeams) ? c.core.classTeams : [];
    function classMatched(code) {
        const want = String(code || '')
            .trim()
            .toLowerCase();
        return teams.some(function (t) {
            const cc = String((t && (t.classCode || t.code)) || '')
                .trim()
                .toLowerCase();
            return cc && cc === want && t.graphGroupId;
        });
    }
    let unmatched = 0;
    classes.forEach(function (cl) {
        if (!classMatched(cl && cl.code)) unmatched += 1;
    });
    if (unmatched) {
        items.push({
            id: 'unmatched-classes',
            text: unmatched + ' Klasse(n) ohne verknüpfte Microsoft-365-Gruppe',
            href: 'jahrgangsgruppen.html'
        });
    }

    const setup = c.setup && typeof c.setup === 'object' ? c.setup : {};
    const matched = setup.matched && typeof setup.matched === 'object' ? setup.matched : {};
    if (!matched.schuelerGroupId && !inferred.students) {
        items.push({
            id: 'sammel-schueler',
            text: 'Sammelgruppe „alle Schülerinnen“ noch nicht verknüpft',
            href: 'schueler-lehrer-gruppen.html'
        });
    }

    items.push({
        id: 'archive-kursteams',
        text: 'Alte Unterrichtsteams archivieren (nicht löschen) – auch als Sammel-Archivierung möglich',
        href: 'teams-archiv.html'
    });

    return items;
}

function incrementLeadingGrade(displayName) {
    const s = String(displayName || '').trim();
    const m = s.match(/^(\d{1,2})([A-Za-z][A-Za-z0-9\-]*)$/);
    if (!m) return '';
    const current = parseInt(m[1], 10);
    if (!isFinite(current)) return '';
    return String(current + 1) + m[2];
}

function plusOneDisplayName(displayName) {
    const s = String(displayName || '').trim();
    const prefixed = s.match(/^(Klasse)\s+(\d{1,2})([A-Za-z0-9\-]*)$/i);
    if (prefixed) {
        const n = parseInt(prefixed[2], 10);
        if (isFinite(n)) return 'Klasse ' + String(n + 1) + (prefixed[3] || '');
    }
    return incrementLeadingGrade(s);
}

/**
 * Geplante Reihenfolge des Schuljahreswechsels — nur Lesen, kein Graph.
 * @param {object|null|undefined} container
 * @param {unknown} playbook
 * @returns {{ currentYear: string, nextYear: string, actions: { id: string, title: string, kind: string, status: string, detail: string }[] }}
 */
export function buildSchoolYearRunPreview(container, playbook) {
    const c = container && typeof container === 'object' ? container : {};
    const currentYear = c.years ? String(c.years.current || '').trim() : '';
    const nextYear = currentYear ? nextSchoolYearLabel(currentYear) : currentSchoolYearLabel();
    const inferred = inferPlaybookFromContainer(c);
    const pb = applyInferredPlaybookDone(normalizePlaybook(playbook, currentYear), inferred, currentYear);
    const actions = [];

    actions.push({
        id: 'year',
        title: 'Schuljahr aktivieren',
        kind: 'local',
        status: pb.done.year ? 'done' : 'ready',
        detail: currentYear
            ? 'Aktuell ' + currentYear + '. Nächstes Jahr typisch ' + nextYear + '.'
            : 'Noch kein Schuljahr gesetzt — zuerst anlegen (z. B. ' + nextYear + ').'
    });

    const teams = c.core && Array.isArray(c.core.classTeams) ? c.core.classTeams : [];
    const nameBits = [];
    let namesReady = 0;
    let namesBlocked = 0;
    teams.forEach(function (t) {
        const from = String((t && (t.displayName || t.classCode || t.code)) || '').trim();
        if (!from) return;
        const to = plusOneDisplayName(from);
        const gid = t && t.graphGroupId ? String(t.graphGroupId).trim() : '';
        if (to && gid) namesReady += 1;
        else if (to) namesBlocked += 1;
        nameBits.push(from + (to ? ' → ' + to : '') + (gid ? '' : ' (noch nicht verknüpft)'));
    });
    actions.push({
        id: 'names',
        title: 'Klassen-Anzeigenamen +1',
        kind: 'graph',
        status: pb.done.names ? 'done' : namesReady ? 'ready' : teams.length ? 'blocked' : 'later',
        detail: nameBits.length
            ? namesReady +
              ' mit Microsoft-365-Gruppe schreibbar, ' +
              namesBlocked +
              ' nur lokal. ' +
              nameBits.slice(0, 4).join('; ') +
              (nameBits.length > 4 ? ' …' : '')
            : 'Keine Klassengruppen in den Stammdaten.'
    });

    const years = c.years && typeof c.years === 'object' ? c.years : {};
    const bucket =
        currentYear && years.byLabel && typeof years.byLabel === 'object' ? years.byLabel[currentYear] : null;
    const classes = bucket && Array.isArray(bucket.classes) ? bucket.classes : [];
    const cohortYears = {};
    classes.forEach(function (cl) {
        const y = String((cl && cl.year) || '').trim();
        if (/^\d{4}$/.test(y)) cohortYears[y] = (cohortYears[y] || 0) + 1;
    });
    const links = c.setup && Array.isArray(c.setup.catalogLinks) ? c.setup.catalogLinks : [];
    const linkedCohorts = links.filter(function (L) {
        return L && L.kind === 'cohort' && L.graphGroupId;
    }).length;
    const cohortN = Object.keys(cohortYears).length;
    actions.push({
        id: 'graduates',
        title: 'Abschlussjahrgänge prüfen',
        kind: 'graph',
        status: pb.done.graduates ? 'done' : cohortN ? 'ready' : 'later',
        detail:
            cohortN
                ? cohortN +
                  ' Abschlussjahr(e) in der Klassenliste, ' +
                  linkedCohorts +
                  ' mit Microsoft-365-Gruppe. Das ist nicht die Klassengruppe.'
                : 'Keine Abschlussjahre in der Klassenliste.'
    });

    actions.push({
        id: 'students',
        title: 'Schülerliste und Sammelgruppe',
        kind: 'manual',
        status: pb.done.students ? 'done' : 'ready',
        detail: inferred.hints && inferred.hints.students ? inferred.hints.students : ''
    });

    actions.push({
        id: 'kursteams',
        title: 'Unterrichtsteams neu anlegen',
        kind: 'manual',
        status: pb.done.kursteams ? 'done' : 'later',
        detail: 'Stundenplan importieren, per CSV/CMD anlegen. Alte Teams nicht löschen.'
    });

    actions.push({
        id: 'archive',
        title: 'Alte Unterrichtsteams archivieren',
        kind: 'manual',
        status: 'later',
        detail: 'Nach der Neuanlage: Archiv-Werkzeug, nicht löschen. Dort per Suchbegriff (z. B. Schuljahr-Präfix) alle Kursteams des abgelaufenen Schuljahres auf einmal archivieren.'
    });

    actions.push({
        id: 'subjects',
        title: 'Fachgruppen und ARGEs prüfen',
        kind: 'manual',
        status: pb.done.subjects ? 'done' : 'ready',
        detail: inferred.hints && inferred.hints.subjects ? inferred.hints.subjects : ''
    });

    return { currentYear: currentYear, nextYear: nextYear, actions: actions };
}

export function normalizeRunLog(raw) {
    const list = Array.isArray(raw) ? raw : [];
    return list
        .filter(function (e) {
            return e && typeof e === 'object';
        })
        .slice(-8)
        .map(function (e) {
            return {
                at: String(e.at || ''),
                mode: e.mode === 'preview' ? 'preview' : 'note',
                summary: String(e.summary || '').slice(0, 400)
            };
        });
}

export function appendRunLogEntry(log, entry) {
    const cur = normalizeRunLog(log);
    cur.push({
        at: entry && entry.at ? String(entry.at) : new Date().toISOString(),
        mode: entry && entry.mode === 'preview' ? 'preview' : 'note',
        summary: String((entry && entry.summary) || '').slice(0, 400)
    });
    return cur.slice(-8);
}

export function summarizeRunPreview(preview) {
    const actions = preview && Array.isArray(preview.actions) ? preview.actions : [];
    const ready = actions.filter(function (a) {
        return a.status === 'ready';
    }).length;
    const done = actions.filter(function (a) {
        return a.status === 'done';
    }).length;
    return {
        total: actions.length,
        ready: ready,
        done: done,
        nextYear: preview && preview.nextYear ? preview.nextYear : ''
    };
}

export function yearAlreadyExists(years, label) {
    const y = String(label || '').trim();
    if (!y) return false;
    return (Array.isArray(years) ? years : []).some((x) => String(x).trim() === y);
}

/** Kalenderjahr am Ende des Schuljahrs: 2025/26 → 2026. */
export function schoolYearEndYear(label) {
    const y = parseSchoolYearStartYear(label);
    return isFinite(y) ? y + 1 : NaN;
}

export function cohortDisplayName(year) {
    return 'Abschlussjahrgang ' + String(year || '').trim();
}

export function cohortMailNickname(year) {
    const y = String(year || '').trim();
    if (!/^\d{4}$/.test(y)) return '';
    return ('maturajg' + y).toLowerCase().slice(0, 60);
}

/**
 * Phase relativ zum aktuellen Schuljahr.
 * In 2026/27: Abschlussjahr 2026 = gerade abgeschlossen, 2027 = aktueller Maturajahrgang.
 * @returns {'vergangen'|'gerade-abgeschlossen'|'aktuell'|'kommend'|'offen'}
 */
export function cohortPhase(graduationYear, schoolYearLabel) {
    const start = parseSchoolYearStartYear(schoolYearLabel);
    const gy = parseInt(String(graduationYear || '').trim(), 10);
    if (!isFinite(start) || !isFinite(gy)) return 'offen';
    if (gy < start) return 'vergangen';
    if (gy === start) return 'gerade-abgeschlossen';
    if (gy === start + 1) return 'aktuell';
    return 'kommend';
}

function classLabel(cl) {
    const code = String((cl && cl.code) || '').trim();
    const name = String((cl && cl.name) || '').trim();
    if (code && name && name.toUpperCase() !== code.toUpperCase()) return name + ' (' + code + ')';
    return name || code;
}

/**
 * Eine Zeile je Abschlussjahr: aus Klassenliste, bestehenden catalogLinks und alter Merkliste.
 * Eltern-Merkungen (kind eltern) gehören nicht hierher – Erziehungsberechtigte werden nicht als Microsoft-365-Benutzer angelegt.
 * @param {Array<{ year?: string, code?: string, name?: string }>} classes
 * @param {Array<{ kind?: string, code?: string, graphGroupId?: string, mailNickname?: string, mode?: string, displayName?: string }>} catalogLinks
 * @param {string} schoolYearLabel
 * @param {Array<{ kind?: string, graduationYear?: string, mailNickname?: string, displayName?: string }>} [legacyPlans]
 */
export function buildCohortRows(classes, catalogLinks, schoolYearLabel, legacyPlans) {
    /** @type {Record<string, { year: string, classes: { code: string, name: string }[], mailNickname: string, displayName: string }>} */
    const byYear = {};

    function ensure(year) {
        const y = String(year || '').trim();
        if (!/^\d{4}$/.test(y)) return null;
        if (!byYear[y]) {
            byYear[y] = {
                year: y,
                classes: [],
                mailNickname: cohortMailNickname(y),
                displayName: cohortDisplayName(y)
            };
        }
        return byYear[y];
    }

    (Array.isArray(classes) ? classes : []).forEach(function (cl) {
        const row = ensure(cl && cl.year);
        if (!row) return;
        row.classes.push({
            code: String((cl && cl.code) || '').trim(),
            name: String((cl && cl.name) || '').trim()
        });
    });

    (Array.isArray(catalogLinks) ? catalogLinks : []).forEach(function (L) {
        if (!L || L.kind !== 'cohort') return;
        const row = ensure(L.code);
        if (!row) return;
        if (L.mailNickname) row.mailNickname = String(L.mailNickname).trim();
        if (L.displayName) row.displayName = String(L.displayName).trim();
    });

    (Array.isArray(legacyPlans) ? legacyPlans : []).forEach(function (p) {
        if (!p || String(p.kind || '') === 'eltern') return;
        const row = ensure(p.graduationYear);
        if (!row) return;
        if (p.mailNickname) row.mailNickname = String(p.mailNickname).trim();
        if (p.displayName) row.displayName = String(p.displayName).trim();
    });

    const links = Array.isArray(catalogLinks) ? catalogLinks : [];
    return Object.keys(byYear)
        .sort()
        .map(function (y) {
            const row = byYear[y];
            const link = links.find(function (L) {
                return L && L.kind === 'cohort' && String(L.code) === y;
            });
            const labels = row.classes.map(classLabel).filter(Boolean);
            return {
                year: y,
                code: y,
                kind: 'cohort',
                displayName: row.displayName,
                mailNickname: row.mailNickname || cohortMailNickname(y),
                classLabels: labels,
                classCount: row.classes.length,
                phase: cohortPhase(y, schoolYearLabel),
                graphGroupId: link && link.graphGroupId ? String(link.graphGroupId).trim() : '',
                mode: link && link.mode ? String(link.mode) : ''
            };
        });
}
