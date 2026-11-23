/**
 * Stammdaten-Workbench: Statusbegriffe und Zeilenmenü.
 * Keine Datenlogik – nur Darstellung und Aktionshierarchie.
 */

/**
 * Sichtbarer Text ohne null/undefined. Leere Werte werden zu einem leeren String.
 * Das Wort „undefined“ wird entfernt; „Null“ als Name bleibt stehen.
 * @param {unknown} value
 * @returns {string}
 */
export function presentValue(value) {
    let s = String(value ?? '').trim();
    if (!s || /^undefined$/i.test(s) || /^null$/i.test(s)) return '';
    s = s.replace(/\bundefined\b/gi, ' ').replace(/\s+/g, ' ').trim();
    return s;
}

/**
 * @param {unknown} value
 * @param {string} [empty]
 * @returns {string}
 */
export function presentText(value, empty) {
    const v = presentValue(value);
    return v || (empty == null ? '–' : empty);
}

/**
 * Leeren Besetzungsplatz als klickbare Zeile kennzeichnen.
 * @param {HTMLElement} preview
 */
export function markEmptyOccupancy(preview) {
    if (!preview) return;
    preview.textContent = '';
    const label = document.createElement('span');
    label.textContent = 'Nicht besetzt';
    const act = document.createElement('span');
    act.className = 'ts-slot-action';
    act.textContent = '+ Besetzen';
    preview.append(label, act);
}

/**
 * @param {string} email
 * @param {{ graphUserId?: string, displayName?: string, userPrincipalName?: string, notFound?: boolean, error?: string, diff?: boolean, syncStatus?: string }|null|undefined} match
 */
export function directoryStatusView(email, match) {
    const em = String(email || '').trim().toLowerCase();
    if (!em || em.indexOf('@') === -1) {
        return {
            kind: 'unlinked',
            label: 'Nicht verknüpft',
            title: 'E-Mail nötig für den Abgleich mit Microsoft Entra'
        };
    }
    if (match && (match.diff || match.syncStatus === 'mismatch')) {
        return {
            kind: 'mismatch',
            label: 'Abweichung',
            title: 'Microsoft 365 weicht von den Stammdaten ab'
        };
    }
    if (match && match.error) {
        return {
            kind: 'error',
            label: 'Fehler',
            title: 'Fehler: ' + String(match.error)
        };
    }
    if (match && match.graphUserId) {
        const gid = String(match.graphUserId);
        const lines = ['Status: Verknüpft', 'Object ID: ' + gid];
        if (match.displayName) lines.push(String(match.displayName));
        if (match.userPrincipalName) lines.push(String(match.userPrincipalName));
        return { kind: 'linked', label: 'Verknüpft', title: lines.join('\n') };
    }
    if (match && match.notFound) {
        return {
            kind: 'unlinked',
            label: 'Nicht verknüpft',
            title: 'Kein Benutzer mit mail oder UPN gleich dieser E-Mail'
        };
    }
    return {
        kind: 'pending',
        label: 'Abgleich offen',
        title: 'Noch nicht mit Microsoft Entra abgeglichen'
    };
}

/**
 * @param {{ graphGroupId?: string, groupId?: string, displayName?: string, mailNickname?: string, mail?: string, notFound?: boolean, syncStatus?: string, diff?: boolean }|null|undefined} link
 * @param {string} [missingTitle]
 */
export function groupLinkStatusView(link, missingTitle) {
    if (link && (link.diff || link.syncStatus === 'mismatch')) {
        return {
            kind: 'mismatch',
            label: 'Abweichung',
            title: 'Die verknüpfte Microsoft-365-Gruppe weicht ab'
        };
    }
    if (link && (link.notFound || link.syncStatus === 'notFound')) {
        return {
            kind: 'unlinked',
            label: 'Nicht verknüpft',
            title: missingTitle || 'Keine passende Microsoft-365-Gruppe gefunden'
        };
    }
    const id = String((link && (link.graphGroupId || link.groupId)) || '').trim();
    if (id) {
        const lines = ['Status: Verknüpft', 'Object ID: ' + id];
        if (link.displayName) lines.push(String(link.displayName));
        if (link.mailNickname) lines.push('Alias: ' + String(link.mailNickname));
        if (link.mail) lines.push('Mail: ' + String(link.mail));
        return { kind: 'linked', label: 'Verknüpft', title: lines.join('\n') };
    }
    return {
        kind: 'pending',
        label: 'Abgleich offen',
        title: missingTitle || 'Noch nicht mit Microsoft 365 abgeglichen'
    };
}

const STATUS_MARK = {
    linked: '✓',
    unlinked: '–',
    mismatch: '!',
    pending: '○',
    error: '!'
};

/**
 * @param {HTMLElement} td
 * @param {{ kind: string, label: string, title: string }} view
 * @param {HTMLElement} [tr]
 */
export function paintStatusCell(td, view, tr) {
    if (!td || !view) return;
    if (tr) tr.style.background = '';
    const kind = view.kind || 'pending';
    const mark = STATUS_MARK[kind] || '○';
    td.className = 'ts-status-cell';
    td.style.fontSize = '';
    td.style.lineHeight = '';
    td.style.color = '';
    td.title = view.title || view.label;
    td.innerHTML =
        '<span class="ts-status ts-status--' +
        kind +
        '"><span class="ts-status__mark" aria-hidden="true">' +
        mark +
        '</span> ' +
        view.label +
        '</span>';
}

function menuLabel(btn) {
    const title = (btn.getAttribute('title') || '').trim();
    const text = (btn.textContent || '').replace(/\s+/g, ' ').trim();
    if (text === '✕' || text === '×' || /löschen|entfernen/i.test(title + ' ' + text)) return 'Löschen';
    if (title) return title;
    return text || 'Aktion';
}

function isDangerAction(btn, label) {
    return /löschen|entfernen|trennen/i.test(label) || (btn.textContent || '').trim() === '✕';
}

/**
 * Sichtbare Farbbuttons einer Aktionszelle in ein Overflow-Menü legen.
 * Vorhandene Click-Handler bleiben an den Buttons.
 * @param {HTMLElement|null|undefined} td
 * @param {{ onEdit?: () => void }} [opts]
 */
export function collapseActionCell(td, opts) {
    if (!td || td.dataset.menuReady === '1') return;
    const buttons = Array.from(td.querySelectorAll('button'));
    td.classList.add('ts-row-menu-cell');
    td.style.whiteSpace = '';
    td.style.display = '';
    td.style.gap = '';
    td.style.alignItems = '';

    const details = document.createElement('details');
    details.className = 'ts-row-menu';
    const summary = document.createElement('summary');
    summary.className = 'ts-row-menu__trigger';
    summary.setAttribute('aria-label', 'Weitere Aktionen');
    summary.innerHTML = '<span aria-hidden="true">···</span>';
    const panel = document.createElement('div');
    panel.className = 'ts-row-menu__panel';
    panel.setAttribute('role', 'menu');

    if (opts && typeof opts.onEdit === 'function') {
        const editItem = document.createElement('button');
        editItem.type = 'button';
        editItem.className = 'ts-row-menu__item';
        editItem.textContent = 'Bearbeiten';
        editItem.setAttribute('role', 'menuitem');
        editItem.addEventListener('click', function () {
            details.open = false;
            opts.onEdit();
        });
        panel.appendChild(editItem);
    }

    buttons.forEach(function (btn) {
        const label = menuLabel(btn);
        btn.classList.remove('mini-btn');
        btn.classList.add('ts-row-menu__item');
        btn.setAttribute('role', 'menuitem');
        btn.style.background = '';
        btn.style.opacity = '';
        btn.style.color = '';
        if (isDangerAction(btn, label)) btn.classList.add('is-danger');
        const icon = btn.querySelector('i');
        btn.textContent = '';
        if (icon) btn.appendChild(icon);
        const span = document.createElement('span');
        span.textContent = label;
        btn.appendChild(span);
        btn.addEventListener('click', function () {
            details.open = false;
        });
        panel.appendChild(btn);
    });

    details.append(summary, panel);
    td.appendChild(details);
    td.dataset.menuReady = '1';

    details.addEventListener('toggle', function () {
        if (!details.open) return;
        document.querySelectorAll('details.ts-row-menu[open]').forEach(function (el) {
            if (el !== details) el.open = false;
        });
    });
    if (typeof document !== 'undefined' && !document.documentElement.dataset.rowMenuBound) {
        document.documentElement.dataset.rowMenuBound = '1';
        document.addEventListener('click', function (ev) {
            document.querySelectorAll('details.ts-row-menu[open]').forEach(function (el) {
                if (!el.contains(ev.target)) el.open = false;
            });
        });
    }
}

/**
 * Besetzung zuerst als Text. Eingabe erst nach Klick oder Tastatur.
 * @param {HTMLElement} wrap
 * @param {HTMLElement} preview
 * @param {HTMLElement|null|undefined} focusEl
 */
export function bindOccupancyEdit(wrap, preview, focusEl) {
    if (!wrap || !preview) return;
    preview.tabIndex = 0;
    preview.setAttribute('role', 'button');
    const label = (preview.textContent || 'Platz').replace(/\s+/g, ' ').trim();
    preview.setAttribute('aria-label', label + ' bearbeiten');
    function setFieldsTab(on) {
        wrap.querySelectorAll('input, select').forEach(function (el) {
            el.tabIndex = on ? 0 : -1;
        });
    }
    setFieldsTab(false);
    function open() {
        wrap.classList.add('is-editing');
        setFieldsTab(true);
        if (focusEl && typeof focusEl.focus === 'function') focusEl.focus();
    }
    preview.addEventListener('click', open);
    preview.addEventListener('keydown', function (e) {
        if (e.key === 'Enter' || e.key === ' ') {
            e.preventDefault();
            open();
        }
    });
    wrap.addEventListener('focusout', function (e) {
        if (!wrap.contains(e.relatedTarget)) {
            wrap.classList.remove('is-editing');
            setFieldsTab(false);
        }
    });
}

/**
 * Erste bearbeitbare Zelle der Zeile öffnen.
 * @param {HTMLElement|null|undefined} tr
 */
export function editTableRow(tr) {
    if (!tr) return;
    const cells = Array.from(tr.children).filter(function (td) {
        if (!td || td.tagName !== 'TD') return false;
        if (td.classList.contains('action-cell') || td.classList.contains('ts-row-menu-cell')) return false;
        const title = td.getAttribute('title') || '';
        return (
            td.classList.contains('is-editable-cell') ||
            title.indexOf('Doppelklick') === 0 ||
            title.indexOf('Klicken') === 0
        );
    });
    const named = cells.find(function (td) {
        return td.dataset.classField === 'name' || td.dataset.argeField === 'name';
    });
    const target = named || cells[1] || cells[0];
    if (target && typeof target.click === 'function') target.click();
}
