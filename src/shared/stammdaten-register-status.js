/**
 * Stammdaten: Ampel für Arbeitskopie, IT-Bibliothek, Intranet, Schuljahr.
 */
import { formatSyncStatusDe, isItLibraryConfigured } from './stammdaten-sharepoint-sync-logic.js';
import { loadItMeta, loadLocalSyncMeta, isReady } from './stammdaten-sharepoint-sync-api.js';

function escapeHtml(s) {
    return String(s ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function currentSchoolYearLabel() {
    try {
        const api = typeof window !== 'undefined' ? window.ms365AppDataV2 : null;
        if (api && typeof api.getContainer === 'function') {
            const c = api.getContainer();
            if (c && c.years && c.years.current) return String(c.years.current);
        }
    } catch {
        /* ignore */
    }
    return '';
}

function yearCounts(yearLabel) {
    const out = { classes: 0, students: 0, teachers: 0, subjects: 0 };
    try {
        const api = typeof window !== 'undefined' ? window.ms365AppDataV2 : null;
        if (!api || typeof api.getYearBucket !== 'function') return out;
        const y = yearLabel || currentSchoolYearLabel();
        if (!y) return out;
        const { bucket } = api.getYearBucket(y);
        out.classes = Array.isArray(bucket.classes) ? bucket.classes.length : 0;
        out.students = Array.isArray(bucket.students) ? bucket.students.length : 0;
    } catch {
        /* ignore */
    }
    try {
        if (typeof window.ms365TenantSettingsLoad === 'function') {
            const s = window.ms365TenantSettingsLoad();
            out.teachers = Array.isArray(s && s.teachers) ? s.teachers.length : 0;
            out.subjects = Array.isArray(s && s.subjects) ? s.subjects.length : 0;
        }
    } catch {
        /* ignore */
    }
    return out;
}

/**
 * @returns {Array<{ icon: string, label: string, kind: string, title: string, href?: string }>}
 */
export function collectRegisterLayerChips() {
    const chips = [];
    const year = currentSchoolYearLabel();
    const counts = yearCounts(year);

    if (!year) {
        chips.push({
            icon: 'bi-calendar-x',
            label: 'Kein Schuljahr gewählt',
            kind: 'warn',
            title: 'Im Konto-Menü oben rechts ein Schuljahr wählen oder anlegen.',
            href: ''
        });
    } else {
        chips.push({
            icon: 'bi-calendar3',
            label:
                'Schuljahr ' +
                year +
                ' · ' +
                counts.classes +
                ' Kl. · ' +
                counts.students +
                ' Schüler',
            kind: counts.classes > 0 || counts.students > 0 ? 'ok' : 'warn',
            title:
                'Aktives Schuljahr im Browser. Klassen: ' +
                counts.classes +
                ', Schüler:innen: ' +
                counts.students +
                ', Lehrkräfte (schulweit): ' +
                counts.teachers +
                ', Fächer: ' +
                counts.subjects +
                '.',
            href: 'tenant.html#klassen'
        });
    }

    chips.push({
        icon: 'bi-laptop',
        label: 'Arbeitskopie: dieser Browser',
        kind: 'muted',
        title:
            'Hier bearbeiten Sie die Daten. Für einen anderen PC: Browser-Backup oder IT-Bibliothek (Synchron-Modus).'
    });

    let spoReady = false;
    let spoKind = 'warn';
    let spoLabel = 'IT-Bibliothek: nicht eingerichtet';
    let spoTitle =
        'Vollbackup aller App-Daten in SharePoint – IT-Sicherungsbibliothek unter Synchron oder Stammdaten-Übergabe.';
    try {
        spoReady = isReady();
        const it = loadItMeta();
        const meta = loadLocalSyncMeta();
        const st = {
            ready: spoReady,
            phase: 'idle',
            dirty: !!(meta && meta.dirty),
            lastAt: meta && meta.at,
            lastDirection: meta && meta.direction,
            error: meta && meta.pendingError,
            libraryTitle: (it && it.listTitle) || undefined
        };
        if (typeof window !== 'undefined') {
            try {
                const auto = window.ms365StammdatenSpoAutoSync;
                if (auto && typeof auto.getStatus === 'function') {
                    const live = auto.getStatus();
                    if (live) {
                        st.phase = live.phase || st.phase;
                        st.dirty = live.dirty != null ? live.dirty : st.dirty;
                        st.lastAt = live.lastAt || st.lastAt;
                        st.lastDirection = live.lastDirection || st.lastDirection;
                        st.error = live.error || st.error;
                    }
                }
            } catch {
                /* ignore */
            }
        }
        spoTitle = formatSyncStatusDe(st);
        if (!spoReady && !isItLibraryConfigured(it)) {
            spoKind = 'warn';
            spoLabel = 'IT-Bibliothek: einrichten';
        } else if (st.error) {
            spoKind = 'error';
            spoLabel = 'IT-Bibliothek: Sync-Fehler';
        } else if (st.dirty) {
            spoKind = 'warn';
            spoLabel = 'IT-Bibliothek: Änderungen offen';
        } else if (st.lastAt) {
            spoKind = 'ok';
            spoLabel = 'IT-Bibliothek: gesichert';
        } else {
            spoKind = 'muted';
            spoLabel = 'IT-Bibliothek: bereit';
        }
    } catch {
        /* defaults */
    }
    chips.push({
        icon: 'bi-shield-lock',
        label: spoLabel,
        kind: spoKind,
        title: spoTitle,
        href: 'tenant.html#stammdaten'
    });

    let intranetUrl = '';
    try {
        const api = typeof window !== 'undefined' ? window.ms365AppDataV2 : null;
        const setup = api && typeof api.getSetup === 'function' ? api.getSetup() : {};
        intranetUrl = setup && setup.intranetSiteUrl ? String(setup.intranetSiteUrl).trim() : '';
    } catch {
        /* ignore */
    }
    if (!intranetUrl) {
        chips.push({
            icon: 'bi-globe2',
            label: 'Intranet: Site-URL fehlt',
            kind: 'warn',
            title:
                'Öffentliche SharePoint-Listen brauchen die Intranet-Site-URL (Stammdaten oder Synchron-Modus).',
            href: 'tenant.html#stammdaten'
        });
    } else {
        chips.push({
            icon: 'bi-globe2',
            label: 'Intranet: Listen aus Stammdaten',
            kind: 'ok',
            title:
                'Nach dem Pflegen: unter Synchron die Stammdaten-Listen auf der Schulwebsite aktualisieren. Site: ' +
                intranetUrl,
            href: 'tenant.html#stammdaten'
        });
    }

    chips.push({
        icon: 'bi-intersect',
        label: 'Entra-Abgleich',
        kind: 'muted',
        title: 'Mitgliedschaften und Klassenfelder prüfen – Werkzeug Datenhygiene / MS-365-Gruppenverwaltung.',
        href: 'tools/datenhygiene.html'
    });

    return chips;
}

/**
 * @param {HTMLElement|null} grid
 * @param {{ compact?: boolean }} [opts]
 */
export function renderRegisterLayerStatus(grid, opts) {
    if (!grid) return;
    const compact = !!(opts && opts.compact);
    const chips = collectRegisterLayerChips();
    grid.replaceChildren();
    chips.forEach(function (c) {
        let el;
        if (compact && c.href) {
            el = document.createElement('a');
            el.href = c.href;
            el.classList.add('ts-status-chip--link');
        } else {
            el = document.createElement('span');
            if (c.href) {
                el.style.cursor = 'pointer';
                el.addEventListener('click', function () {
                    window.location.href = c.href;
                });
            }
        }
        el.className = 'ts-status-chip ts-status-chip--' + c.kind;
        if (c.title) {
            el.title = c.title;
            el.setAttribute('aria-label', c.label + ': ' + c.title);
        }
        el.innerHTML =
            '<i class="bi ' + escapeHtml(c.icon) + '" aria-hidden="true"></i> ' + escapeHtml(c.label);
        grid.appendChild(el);
    });
}
