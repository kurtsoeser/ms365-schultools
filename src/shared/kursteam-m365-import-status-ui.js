/**
 * Gemeinsame Status-Darstellung nach „Bestehende Teams importieren“
 * (Katalog ↔ / Plan ↔ Microsoft 365).
 */

function escapeHtml(text) {
    return String(text == null ? '' : text)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

/**
 * @param {object|null|undefined} report buildKursteamAbgleichReport-Ergebnis
 * @param {{
 *   planLabel?: string,
 *   emptyHint?: string,
 *   noPlanHint?: string,
 *   openDetailsUnder?: number
 * }} [options]
 * @returns {string} HTML
 */
export function renderKursteamM365ImportStatusHtml(report, options) {
    const opts = options && typeof options === 'object' ? options : {};
    if (!report || !report.counts) {
        return (
            '<p class="muted kt-m365-status-empty" style="margin:0;">' +
            escapeHtml(
                opts.emptyHint ||
                    'Zuerst „Teams aus M365 laden“ – danach sehen Sie hier passend / fehlen / nur in M365.'
            ) +
            '</p>'
        );
    }

    const c = report.counts;
    const planHint = opts.planLabel || 'Katalog';
    const openUnder = Number.isFinite(opts.openDetailsUnder) ? opts.openDetailsUnder : 8;
    const noPlan = c.planned === 0;

    function tableSection(title, rows, cols) {
        if (!rows || !rows.length) return '';
        const head = cols.map((col) => '<th>' + col.label + '</th>').join('');
        const body = rows
            .map((row) => {
                const tds = cols
                    .map((col) => {
                        if (col.render) return '<td>' + col.render(row) + '</td>';
                        return '<td>' + escapeHtml(String(row[col.key] || '—')) + '</td>';
                    })
                    .join('');
                return '<tr>' + tds + '</tr>';
            })
            .join('');
        return (
            '<details class="kt-m365-status-details"' +
            (rows.length <= openUnder ? ' open' : '') +
            '>' +
            '<summary>' +
            escapeHtml(title) +
            ' (' +
            rows.length +
            ')</summary>' +
            '<div class="kt-m365-status-table-wrap">' +
            '<table class="data-table kt-m365-status-table"><thead><tr>' +
            head +
            '</tr></thead><tbody>' +
            body +
            '</tbody></table></div></details>'
        );
    }

    const code = (v) =>
        v ? '<code class="kt-m365-status-code">' + escapeHtml(v) + '</code>' : '—';

    return (
        '<div class="kt-m365-status">' +
        '<p class="kt-m365-status__title">' +
        escapeHtml(planHint) +
        ' ↔ Microsoft 365</p>' +
        '<p class="kt-m365-status__counts">' +
        '<span class="kt-m365-status__ok">' +
        c.matched +
        ' passend</span> · ' +
        '<span class="kt-m365-status__miss">' +
        c.missingInM365 +
        ' fehlen in M365</span> · ' +
        '<span class="kt-m365-status__extra">' +
        c.onlyInM365 +
        ' nur in M365</span>' +
        (c.mailConflict ? ' · ' + c.mailConflict + ' Mail-Abweichung' : '') +
        '</p>' +
        (noPlan
            ? '<p class="muted kt-m365-status__hint">' +
              escapeHtml(
                  opts.noPlanHint ||
                      'Kein Katalog-/Plan-Stand – zuerst Teams in der App oder Stundenplan-Daten laden.'
              ) +
              '</p>'
            : '') +
        tableSection('Fehlt in M365', report.missingInM365 || [], [
            { key: 'klasse', label: 'Klasse' },
            { key: 'fach', label: 'Fach' },
            { key: 'lehrerCode', label: 'LK' },
            {
                key: 'gruppenmail',
                label: 'Geplante Mail',
                render: (r) => code(r.gruppenmail)
            }
        ]) +
        tableSection('Nur in M365 (nicht im Katalog/Plan)', report.onlyInM365 || [], [
            { key: 'klasse', label: 'Klasse' },
            { key: 'fach', label: 'Fach' },
            { key: 'lehrerCode', label: 'LK' },
            {
                key: 'gruppenmail',
                label: 'Mail-Nickname',
                render: (r) => code(r.gruppenmail)
            }
        ]) +
        tableSection('Passend verknüpft', report.matched || [], [
            { key: 'klasse', label: 'Klasse' },
            { key: 'fach', label: 'Fach' },
            { key: 'lehrerCode', label: 'LK' },
            {
                key: 'matchBy',
                label: 'Match',
                render: (r) => (r.matchBy === 'gruppenmail' ? 'Mail' : 'Unterricht')
            }
        ]) +
        tableSection('Gleicher Unterricht, andere Mail', report.mailConflict || [], [
            { key: 'klasse', label: 'Klasse' },
            { key: 'fach', label: 'Fach' },
            {
                key: 'planMail',
                label: 'Plan',
                render: (r) => code(r.planMail)
            },
            {
                key: 'm365Mail',
                label: 'M365',
                render: (r) => code(r.m365Mail)
            }
        ]) +
        '</div>'
    );
}

/**
 * Kurzer Einzeiler für Button-Badges o. Ä.
 * @param {object|null|undefined} report
 * @returns {{ text: string, tone: 'ok'|'warn'|'none' }}
 */
export function summarizeKursteamM365ImportHint(report) {
    if (!report || !report.counts) return { text: '', tone: 'none' };
    const missing = Number(report.counts.missingInM365 || 0);
    const only = Number(report.counts.onlyInM365 || 0);
    const matched = Number(report.counts.matched || 0);
    if (missing > 0) return { text: missing + ' fehlen', tone: 'warn' };
    if (only > 0) return { text: only + ' neu', tone: 'warn' };
    if (matched > 0) return { text: matched + ' ok', tone: 'ok' };
    return { text: 'geladen', tone: 'ok' };
}
