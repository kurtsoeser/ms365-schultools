/**
 * Wiederverwendbarer Activity Detail-Drawer (Übersicht / Liste / Kalender).
 */
import { klasseCodesOf, schoolYearLabel, schoolYearStartFromDate } from './schulaktivitaeten-planer-logic.js';
import { formatUeZeitraum } from './schulaktivitaeten-planer-zeitraster.js';
import {
    labelMaps,
    formatDeDate,
    statusLabel,
    typLabel,
    teilnahmeKurz,
    teilnahmeModusLabel,
    canEditItem,
    canDecide,
    canKvStellungnahme,
    canMarkUntis,
    studentCountForKlasse,
    matchTeacherByEmail,
    resolveUserRelations,
    formatRelationSummary,
    hasAktKvCapability,
    kvKlasseCodesForEmail,
    itemTouchesKvKlassen,
    scheduleStatusOf,
    scheduleStatusLabel,
    effectiveAccountEmail,
    effectiveTeacherCode
} from './schulaktivitaeten-planer-state.js';
import {
    getAktPersonalCalendarEventId,
    previewAktInviteTargets
} from './schulaktivitaeten-planer-calendar.js';
import {
    parseEuroFromText,
    classifyKostenAmpel,
    kostenAmpelLabel,
    kostenLimitForTyp
} from './schulaktivitaeten-planer-logic.js';

function esc(s) {
    return String(s ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

/**
 * Redaktionelles Datum: „Do, 26. November 2026“ bzw. „26.–28. November 2026“.
 * @param {string} startIso
 * @param {string} [endIso]
 */
export function formatReadableDateRange(startIso, endIso) {
    const start = String(startIso || '').trim();
    const end = String(endIso || '').trim() || start;
    if (!start) return '–';
    const parse = (iso) => {
        const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(iso);
        if (!m) return null;
        return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    };
    const a = parse(start);
    const b = parse(end);
    if (!a) return formatDeDate(start);
    const months = [
        'Januar',
        'Februar',
        'März',
        'April',
        'Mai',
        'Juni',
        'Juli',
        'August',
        'September',
        'Oktober',
        'November',
        'Dezember'
    ];
    const weekdays = ['So', 'Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa'];
    if (!b || start === end) {
        return `${weekdays[a.getDay()]}, ${a.getDate()}. ${months[a.getMonth()]} ${a.getFullYear()}`;
    }
    if (a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth()) {
        return `${a.getDate()}.–${b.getDate()}. ${months[a.getMonth()]} ${a.getFullYear()}`;
    }
    if (a.getFullYear() === b.getFullYear()) {
        return `${a.getDate()}. ${months[a.getMonth()]} – ${b.getDate()}. ${months[b.getMonth()]} ${a.getFullYear()}`;
    }
    return `${a.getDate()}. ${months[a.getMonth()]} ${a.getFullYear()} – ${b.getDate()}. ${months[b.getMonth()]} ${b.getFullYear()}`;
}

function statusBadge(status) {
    const s = String(status || 'beantragt').toLowerCase();
    let mod = 'warn';
    if (s === 'genehmigt') mod = 'ok';
    else if (s === 'abgelehnt' || s === 'abgesagt') mod = 'bad';
    else if (s === 'nachforderung') mod = 'warn';
    else if (s === 'kv_stellungnahme' || s === 'beantragt') mod = 'info';
    else if (s === 'entwurf') mod = 'muted';
    return `<span class="akt-status akt-status--${mod}" data-status="${esc(s)}"><span class="akt-status__dot" aria-hidden="true"></span>${esc(
        statusLabel(s)
    )}</span>`;
}

function scheduleBadge(item) {
    const code = scheduleStatusOf(item);
    const label = scheduleStatusLabel(code);
    if (!label) return '';
    const mod = code === 'pending' ? 'warn' : 'ok';
    return `<span class="akt-status akt-status--${mod} akt-status--schedule"><span class="akt-status__dot" aria-hidden="true"></span>${esc(
        label
    )}</span>`;
}

/**
 * Fachlichen Hinweis bereinigen: Demo-Seed-Tags und System-IDs entfernen.
 * @param {string} raw
 */
export function sanitizeFachNotiz(raw) {
    let s = String(raw || '');
    s = s.replace(/\s*·?\s*akt-demo-[\w.-]+/gi, '');
    s = s.replace(/\s*·?\s*SJ\s*\d{4}\/\d{2}/gi, '');
    s = s.replace(/\s{2,}/g, ' ').replace(/^[·\s]+|[·\s]+$/g, '');
    return s.trim();
}

/**
 * Name primär, E-Mail optional sekundär – Rollen nicht zusammenlegen.
 * @param {object} state
 * @param {{ code?: string, email?: string, name?: string }} person
 */
function resolvePerson(state, person) {
    const teachers = (state.stammdaten && state.stammdaten.teachers) || [];
    const code = String((person && person.code) || '')
        .trim()
        .toUpperCase();
    const email = String((person && person.email) || '')
        .trim()
        .toLowerCase();
    const givenName = String((person && person.name) || '').trim();
    const byCode = code ? teachers.find((t) => String(t.code || '').toUpperCase() === code) : null;
    const byEmail =
        email &&
        teachers.find(
            (t) =>
                String(t.email || '')
                    .trim()
                    .toLowerCase() === email
        );
    const t = byCode || byEmail || null;
    const name =
        givenName && !givenName.includes('@')
            ? givenName
            : (t && t.name) || (code && !code.includes('@') ? code : '') || '';
    const mail = email || (t && t.email) || (givenName.includes('@') ? givenName : '');
    return { name: name || '', email: mail || '' };
}

/**
 * Kompakte User-Kachel: Name sichtbar, E-Mail per Hover/title.
 * @param {object} state
 * @param {{ code?: string, email?: string, name?: string }} person
 * @param {string} missingLabel
 */
function renderPersonTile(state, person, missingLabel) {
    const p = resolvePerson(state, person);
    if (!p.name && !p.email) {
        return `<span class="akt-missing">${esc(missingLabel || 'Nicht erfasst')}</span>`;
    }
    const label = p.name || p.email;
    const mail =
        p.email && p.email.toLowerCase() !== String(label).toLowerCase() ? p.email : '';
    const initial = String(label)
        .trim()
        .charAt(0)
        .toUpperCase();
    return `<span class="akt-user-tile"${mail ? ` title="${esc(mail)}"` : ''}>
      <span class="akt-user-tile__av" aria-hidden="true">${esc(initial || '?')}</span>
      <span class="akt-user-tile__meta">
        <span class="akt-user-tile__name">${esc(label)}</span>
        ${mail ? `<span class="akt-user-tile__email">${esc(mail)}</span>` : ''}
      </span>
    </span>`;
}

function renderPersonLine(state, person, missingLabel) {
    return renderPersonTile(state, person, missingLabel);
}

function formatVerlaufWhen(raw) {
    const s = String(raw || '').trim();
    if (!s) return '';
    const iso = /^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2}))?/.exec(s);
    if (iso) {
        const day = `${iso[3]}.${iso[2]}.${iso[1]}`;
        return iso[4] != null ? `${day}, ${iso[4]}:${iso[5]}` : day;
    }
    return s;
}

/**
 * Verlaufseinträge – Zeilen oder Power-Automate-Blöcke (--- / Datum:/Kommentar:).
 * @param {string} aenderungsLog
 * @returns {{ when: string, text: string }[]}
 */
export function parseVerlauf(aenderungsLog) {
    const raw = String(aenderungsLog || '').trim();
    if (!raw) return [];
    const looksBlocked =
        /(^|\n)\s*---\s*(\n|$)/.test(raw) ||
        /Datum:\s*\d{4}-\d{2}-\d{2}/i.test(raw) ||
        /Kommentar:\s*/i.test(raw);
    if (looksBlocked) {
        const blocks = raw
            .split(/\n\s*---\s*\n?/)
            .map((b) => b.replace(/^---\s*/, '').trim())
            .filter(Boolean);
        return blocks
            .map((block) => {
                const dateM = /Datum:\s*([^\n]+)/i.exec(block);
                const when = formatVerlaufWhen(dateM ? dateM[1] : '');
                const text = block
                    .replace(/Datum:\s*[^\n]+\n?/i, '')
                    .replace(/^\s*\.\s*$/gm, '')
                    .trim();
                return { when, text: text || block };
            })
            .filter((e) => e.text)
            .slice(-8)
            .reverse();
    }
    return raw
        .split(/\n+/)
        .map((line) => line.trim())
        .filter(Boolean)
        .map((line) => {
            const m = /^(\d{4}-\d{2}-\d{2}|\d{1,2}\.\d{1,2}\.\d{2,4})?\s*[·.]?\s*(.*)$/.exec(line);
            if (m && m[1]) {
                return { when: formatVerlaufWhen(m[1]), text: m[2] || line };
            }
            return { when: '', text: line };
        })
        .slice(-8)
        .reverse();
}

function renderInviteTargetList(targets) {
    if (!targets.length) {
        return `<p class="muted">Keine Klassen oder Begleitung hinterlegt – Einladung nur manuell möglich.</p>`;
    }
    const items = targets
        .map(
            (t) =>
                `<li class="${t.ok ? 'is-ok' : 'is-warn'}"><strong>${esc(t.label)}</strong> · ${esc(
                    t.detail
                )}</li>`
        )
        .join('');
    const list = `<ul class="akt-drawer-cal__targets">${items}</ul>`;
    const warnN = targets.filter((t) => !t.ok).length;
    const summaryBits = [`${targets.length} Empfänger`];
    if (warnN) summaryBits.push(`${warnN} ohne Adresse`);
    /* Bei vielen Zielen einklappbar, damit die Aktionsbuttons sichtbar bleiben. */
    if (targets.length > 4) {
        return `<details class="akt-drawer-cal__targets-wrap">
      <summary>${esc(summaryBits.join(' · '))}</summary>
      ${list}
    </details>`;
    }
    return list;
}

function renderCalendarBlock(it, state) {
    const st = String(it.status || '').toLowerCase();
    if (st !== 'genehmigt') return '';
    const id = it.aktivitaetId || it.itemId;
    const email = String(effectiveAccountEmail(state) || state.accountEmail || '')
        .trim()
        .toLowerCase();
    const linked = !!(email && getAktPersonalCalendarEventId(id, email));
    const targets = previewAktInviteTargets(it, state.stammdaten);
    const okN = targets.filter((t) => t.ok).length;
    return `<section class="akt-drawer__section akt-drawer-cal">
      <details class="akt-drawer-cal__details">
        <summary>
          <span class="akt-drawer-cal__summary-title">Kalender &amp; Einladung</span>
          <span class="akt-drawer-cal__summary-meta muted">${
              okN ? `${okN} Empfänger` : 'Outlook / ICS'
          }</span>
        </summary>
        <p class="akt-drawer__lead">Termin in Outlook übernehmen oder Klasse / Begleitung per Microsoft&nbsp;365 einladen.</p>
        ${
            linked
                ? '<p class="akt-drawer-cal__linked"><i class="bi bi-check2-circle" aria-hidden="true"></i> Bereits mit deinem Outlook-Kalender verknüpft.</p>'
                : ''
        }
        <div class="akt-drawer-cal__actions">
          <button type="button" class="btn akt-btn-primary akt-btn-compact" data-akt-cal-mine="${esc(
              id
          )}"><i class="bi bi-calendar-plus" aria-hidden="true"></i>In meinen Kalender</button>
          <button type="button" class="btn akt-btn-compact" data-akt-cal-invite="${esc(
              id
          )}" ${okN ? '' : 'disabled title="Keine einladbaren Adressen gefunden"'}><i class="bi bi-envelope-plus" aria-hidden="true"></i>Einladen${
              okN ? ` (${okN})` : ''
          }</button>
          <button type="button" class="btn akt-btn-ghost akt-btn-compact" data-akt-cal-ics="${esc(
              id
          )}"><i class="bi bi-download" aria-hidden="true"></i>ICS</button>
        </div>
        <p class="akt-drawer-cal__hint muted">Nutzt Klassen-Teams / Verteiler und Begleitung. ICS für Outlook, Apple und Google.</p>
        ${renderInviteTargetList(targets)}
      </details>
    </section>`;
}

/** Kompakte Klassenzeile für den Drawer-Kopf (viele Klassen → Kurzform). */
function drawerTeilnahmeSub(it, labels) {
    const codes = klasseCodesOf(it);
    const full = teilnahmeKurz(it, labels.klasse);
    if (codes.length <= 3) return full;
    const first = codes
        .slice(0, 2)
        .map((c) => labels.klasse[c] || c)
        .join(', ');
    return `${first} u. ${codes.length - 2} weitere Klassen`;
}

function renderStatusBlock(it, state) {
    const st = String(it.status || '').toLowerCase();
    const editable = canEditItem(it, state);
    if (st === 'genehmigt') {
        return `<section class="akt-drawer__section akt-drawer__section--status">
      <h4>Status</h4>
      <p class="akt-drawer__status-msg">${
          scheduleStatusOf(it) === 'pending'
              ? 'Genehmigt – organisatorische Eintragung noch ausständig.'
              : 'Aktuell ist nichts von dir erforderlich.'
      }</p>
    </section>
    ${renderCalendarBlock(it, state)}`;
    }
    if (st === 'beantragt' || st === 'kv_stellungnahme') {
        return `<section class="akt-drawer__section">
      <h4>Status</h4>
      ${statusBadge(st)}
      <p class="akt-drawer__lead"><strong>Zur Genehmigung eingereicht</strong><br>Der Antrag wird geprüft.${
          canDecide(state) && st === 'beantragt' ? '' : ' Aktuell ist keine Aktion erforderlich.'
      }</p>
    </section>`;
    }
    if (st === 'nachforderung') {
        return `<section class="akt-drawer__section akt-drawer__section--action">
      <h4>Status</h4>
      ${statusBadge(st)}
      <p class="akt-drawer__lead"><strong>Ergänzung erforderlich</strong></p>
      ${
          it.rueckfrageText
              ? `<p>Folgende Information wird noch benötigt:<br><em>${esc(it.rueckfrageText)}</em></p>`
              : '<p>Bitte ergänze die fehlenden Angaben und reiche erneut ein.</p>'
      }
      ${
          editable
              ? `<button type="button" class="btn akt-btn-primary" data-akt-edit="${esc(it.itemId)}">Angaben ergänzen</button>`
              : ''
      }
    </section>`;
    }
    if (st === 'abgelehnt') {
        return `<section class="akt-drawer__section akt-drawer__section--action">
      <h4>Status</h4>
      ${statusBadge(st)}
      <p class="akt-drawer__lead"><strong>Nicht genehmigt</strong></p>
      <div class="akt-drawer__reason">
        <p class="akt-drawer__reason-label">Warum wurde der Antrag abgelehnt?</p>
        <p>${esc(it.ablehnungsGrund || 'Keine Begründung hinterlegt.')}</p>
      </div>
      ${
          editable
              ? `<button type="button" class="btn akt-btn-primary" data-akt-edit="${esc(it.itemId)}">Antrag bearbeiten</button>`
              : ''
      }
    </section>`;
    }
    if (st === 'entwurf') {
        return `<section class="akt-drawer__section akt-drawer__section--action">
      <h4>Status</h4>
      ${statusBadge(st)}
      <p class="akt-drawer__lead">Dieser Entwurf ist noch nicht eingereicht.</p>
      ${
          editable
              ? `<button type="button" class="btn akt-btn-primary" data-akt-edit="${esc(it.itemId)}">Weiter bearbeiten</button>`
              : ''
      }
    </section>`;
    }
    if (st === 'abgesagt') {
        return `<section class="akt-drawer__section">
      <h4>Status</h4>
      ${statusBadge(st)}
      <p class="akt-drawer__lead">Diese Aktivität wurde abgesagt.</p>
    </section>`;
    }
    return `<section class="akt-drawer__section"><h4>Status</h4>${statusBadge(st)}</section>`;
}

/**
 * @param {object} state
 */
export function renderActivityDrawer(state) {
    const id = state.detailId;
    if (!id) return '';
    const it = (state.items || []).find((x) => x.aktivitaetId === id || x.itemId === id);
    if (!it) return '';
    const labels = labelMaps(state.stammdaten);
    const st = String(it.status || '').toLowerCase();
    const editable = canEditItem(it, state);
    const klassen = klasseCodesOf(it);
    const schuelerN = klassen.reduce(
        (n, c) => n + (studentCountForKlasse(state.stammdaten, c) || 0),
        0
    );
    const zeit =
        formatUeZeitraum(it) ||
        [it.startZeit, it.endZeit].filter(Boolean).join('–') ||
        '';
    const begleitungCodes = String(it.begleitungCodes || it.begleitung || '')
        .split(/[;,]/)
        .map((c) => c.trim())
        .filter(Boolean);
    const verlauf = parseVerlauf(it.aenderungsLog);
    const personaEmail = effectiveAccountEmail(state) || state.accountEmail || '';
    const teacherCode = effectiveTeacherCode(state);
    const teacher =
        teacherCode
            ? { code: teacherCode }
            : matchTeacherByEmail((state.stammdaten && state.stammdaten.teachers) || [], personaEmail);
    const who = {
        email: personaEmail,
        teacherCode: (teacher && teacher.code) || teacherCode || ''
    };
    const rels = resolveUserRelations(it, who);
    const relSummary = formatRelationSummary(rels);
    const kvCodes = hasAktKvCapability(state)
        ? kvKlasseCodesForEmail(state.stammdaten, personaEmail)
        : [];
    const touchedKv = kvCodes.filter((c) => itemTouchesKvKlassen(it, [c]));
    const showClassCtx = touchedKv.length > 0;
    const primaryKv = touchedKv[0] || '';
    const fachNotiz = sanitizeFachNotiz(it.notiz);
    const sjStart = schoolYearStartFromDate(it.startdatum || new Date());
    const sjLabel = schoolYearLabel(sjStart);

    const primaryActions = [];
    const secondaryActions = [];
    const moreActions = [];
    if (canDecide(state) && st === 'beantragt') {
        primaryActions.push(
            `<button type="button" class="btn akt-btn-primary" data-akt-approve="${esc(it.itemId)}">Genehmigen</button>`
        );
        secondaryActions.push(
            `<button type="button" class="btn" data-akt-nachforderung="${esc(it.itemId)}">Zur Ergänzung</button>`
        );
        moreActions.push(
            `<button type="button" class="akt-more-menu__item akt-btn-danger" data-akt-reject="${esc(it.itemId)}">Ablehnen</button>`
        );
    }
    if (canMarkUntis(it, state) && !it.untisErledigt) {
        primaryActions.unshift(
            `<button type="button" class="btn akt-btn-primary" data-akt-untis="${esc(
                it.itemId
            )}">Als in Untis verarbeitet markieren</button>`
        );
    }
    if (editable && st !== 'nachforderung' && st !== 'abgelehnt' && st !== 'entwurf') {
        secondaryActions.push(
            `<button type="button" class="btn" data-akt-edit="${esc(it.itemId)}">Aktivität bearbeiten</button>`
        );
    } else if (editable && (st === 'nachforderung' || st === 'abgelehnt' || st === 'entwurf')) {
        /* edit already in status block */
    } else if (editable) {
        secondaryActions.push(
            `<button type="button" class="btn" data-akt-edit="${esc(it.itemId)}">Aktivität bearbeiten</button>`
        );
    }
    if (editable) {
        if (st === 'genehmigt' || st === 'beantragt' || st === 'nachforderung' || st === 'kv_stellungnahme') {
            moreActions.push(
                `<button type="button" class="akt-more-menu__item" data-akt-cancel="${esc(it.itemId)}">Als abgesagt markieren</button>`
            );
        }
        if (st === 'entwurf' || st === 'abgelehnt') {
            moreActions.push(
                `<button type="button" class="akt-more-menu__item akt-btn-danger" data-akt-delete="${esc(it.itemId)}">Löschen</button>`
            );
        }
    }
    const moreMenu = moreActions.length
        ? `<details class="akt-more-menu">
        <summary class="btn akt-btn-ghost" aria-label="Weitere Aktionen">…</summary>
        <div class="akt-more-menu__panel">${moreActions.join('')}</div>
      </details>`
        : '';

    return `
    <div class="akt-drawer-root" role="dialog" aria-modal="true" aria-label="Aktivitätsdetails">
      <button type="button" class="akt-drawer__backdrop" id="aktDetailClose" aria-label="Schließen"></button>
      <aside class="akt-drawer">
        <header class="akt-drawer__head">
          <div>
            <h2 class="akt-drawer__title">${esc(it.titel || 'Aktivität')}</h2>
            <p class="akt-drawer__sub" title="${esc(teilnahmeKurz(it, labels.klasse))}">${esc(
                drawerTeilnahmeSub(it, labels)
            )}${schuelerN ? ` · ${schuelerN} Schüler:innen` : ''}</p>
            <div class="akt-drawer__badges">${statusBadge(it.status)}${scheduleBadge(it)}</div>
          </div>
          <button type="button" class="btn akt-drawer__close" id="aktDetailCloseBtn" aria-label="Schließen"><i class="bi bi-x-lg"></i></button>
        </header>

        <div class="akt-drawer__core">
          <p class="akt-drawer__when"><i class="bi bi-calendar3" aria-hidden="true"></i> ${esc(
              formatReadableDateRange(it.startdatum, it.enddatum)
          )}</p>
          ${zeit ? `<p class="akt-drawer__when"><i class="bi bi-clock" aria-hidden="true"></i> ${esc(zeit)}</p>` : ''}
          ${
              it.ort
                  ? `<p class="akt-drawer__when"><i class="bi bi-geo-alt" aria-hidden="true"></i> ${esc(it.ort)}</p>`
                  : '<p class="akt-drawer__when muted">Ort nicht erfasst</p>'
          }
          <p class="akt-drawer__meta muted">Schuljahr ${esc(sjLabel)}</p>
        </div>

        <div class="akt-drawer__body">
          ${
              rels.length
                  ? `<section class="akt-drawer__section">
            <h4>Deine Beteiligung</h4>
            <p class="akt-drawer__lead">${esc(relSummary)}</p>
          </section>`
                  : ''
          }

          ${
              showClassCtx
                  ? `<section class="akt-drawer__section">
            <h4>Deine Klasse</h4>
            <p><strong>${esc(labels.klasse[primaryKv] || primaryKv)}</strong></p>
            <p>${esc(teilnahmeModusLabel(it.teilnahmeModus))} · ${esc(teilnahmeKurz(it, labels.klasse))}</p>
            ${
                schuelerN
                    ? `<p class="muted">Betroffen: ${schuelerN} Schüler:innen</p>`
                    : ''
            }
            ${zeit ? `<p class="muted">Unterricht: ${esc(zeit)}</p>` : ''}
            ${
                it.nichtTeilnehmende
                    ? `<p class="muted">Nehmen nicht teil: ${esc(it.nichtTeilnehmende)}</p>`
                    : ''
            }
          </section>`
                  : ''
          }

          ${renderStatusBlock(it, state)}

          <section class="akt-drawer__section">
            <h4>Antrag</h4>
            <p>${esc(typLabel(it.typ))}${
                it.beantragtAm ? ` · eingereicht ${esc(formatDeDate(it.beantragtAm))}` : ''
            }</p>
            <div class="akt-drawer__people">
              <div class="akt-drawer__people-group">
                <span class="akt-drawer__people-label">Antrag gestellt von</span>
                <div class="akt-user-tiles">${renderPersonTile(
                    state,
                    {
                        email: it.beantragtVon,
                        code: it.beantragtVon === it.lehrerEmail ? it.lehrerCode : '',
                        name: ''
                    },
                    'Antragsteller nicht erfasst'
                )}</div>
              </div>
              <div class="akt-drawer__people-group">
                <span class="akt-drawer__people-label">Verantwortliche Lehrkraft</span>
                <div class="akt-user-tiles">${renderPersonTile(
                    state,
                    {
                        code: it.lehrerCode,
                        email: it.lehrerEmail
                    },
                    'Verantwortliche Lehrkraft nicht erfasst'
                )}</div>
              </div>
            </div>
          </section>

          ${
              it.kvStellungnahme || canKvStellungnahme(it, state)
                  ? `<section class="akt-drawer__section${
                        canKvStellungnahme(it, state) ? ' akt-drawer__section--action' : ''
                    }">
            <h4>KV-Stellungnahme</h4>
            ${
                canKvStellungnahme(it, state)
                    ? `<p class="akt-drawer__lead">Deine Stellungnahme ist erforderlich.</p>
            <button type="button" class="btn akt-btn-primary" data-akt-kv="${esc(
                it.itemId
            )}">Stellungnahme abgeben</button>`
                    : `<p><strong>Status</strong><br>Vorhanden</p>
            <p>${esc(it.kvStellungnahme)}</p>
            ${
                it.kvStellungnahmeVon || it.kvStellungnahmeAm
                    ? `<p class="muted">${esc(
                          [it.kvStellungnahmeVon, it.kvStellungnahmeAm && formatDeDate(it.kvStellungnahmeAm)]
                              .filter(Boolean)
                              .join(' · ')
                      )}</p>`
                    : ''
            }`
            }
          </section>`
                  : ''
          }

          <section class="akt-drawer__section">
            <h4>Teilnahme</h4>
            <p>${esc(teilnahmeModusLabel(it.teilnahmeModus))} · ${esc(teilnahmeKurz(it, labels.klasse))}</p>
            ${it.gruppenText ? `<p class="muted">${esc(it.gruppenText)}</p>` : ''}
            ${it.teilnehmendeText ? `<p class="muted">Teilnehmend: ${esc(it.teilnehmendeText)}</p>` : ''}
            ${
                it.nichtTeilnehmende
                    ? `<p class="muted">Nehmen nicht teil: ${esc(it.nichtTeilnehmende)}</p>`
                    : ''
            }
            ${schuelerN ? `<p class="muted">${schuelerN} Schüler:innen</p>` : ''}
          </section>

          <section class="akt-drawer__section">
            <h4>Organisation</h4>
            <div class="akt-drawer__people-group">
              <span class="akt-drawer__people-label">Begleitung</span>
              ${
                  begleitungCodes.length
                      ? `<div class="akt-user-tiles">${begleitungCodes
                            .map((c) =>
                                renderPersonTile(state, { code: c }, 'Begleitperson nicht erfasst')
                            )
                            .join('')}</div>`
                      : '<span class="akt-missing">Keine Begleitperson eingetragen</span>'
              }
            </div>
            ${it.verkehrsmittel ? `<p><strong>Verkehrsmittel</strong> · ${esc(it.verkehrsmittel)}</p>` : ''}
            ${
                it.kostenHinweis
                    ? `<p><strong>Kosten</strong> · ${esc(it.kostenHinweis)}${(() => {
                          const amount = parseEuroFromText(it.kostenHinweis);
                          const level = Number.isFinite(amount)
                              ? classifyKostenAmpel(it.typ, amount, state.rules)
                              : 'unknown';
                          const label = kostenAmpelLabel(level);
                          const lim = kostenLimitForTyp(it.typ, state.rules);
                          if (!label) return '';
                          return ` <span class="akt-kosten-ampel akt-kosten-ampel--${level}" title="${esc(
                              lim
                                  ? `Schwelle ${lim.typ}: bis ${lim.okBis} € / ${lim.warnBis} €`
                                  : ''
                          )}">${esc(label)}</span>`;
                      })()}</p>`
                    : ''
            }
            ${it.externeVortragende ? `<p><strong>Externe</strong> · ${esc(it.externeVortragende)}</p>` : ''}
            ${it.ablauf ? `<p><strong>Ablauf</strong><br>${esc(it.ablauf)}</p>` : ''}
            ${
                it.unterrichtsbezug
                    ? `<p><strong>Unterrichtsbezug</strong><br>${esc(it.unterrichtsbezug)}</p>`
                    : ''
            }
            ${fachNotiz ? `<p><strong>Hinweise</strong><br>${esc(fachNotiz)}</p>` : ''}
          </section>

          ${
              verlauf.length
                  ? `<section class="akt-drawer__section akt-drawer__section--verlauf">
            <h4>Verlauf</h4>
            <ul class="akt-drawer__timeline">
              ${verlauf
                  .map(
                      (v) =>
                          `<li class="akt-drawer__timeline-item">
                ${v.when ? `<time class="akt-drawer__twhen">${esc(v.when)}</time>` : ''}
                <div class="akt-drawer__ttext">${esc(v.text)}</div>
              </li>`
                  )
                  .join('')}
            </ul>
          </section>`
                  : ''
          }

          ${
              canDecide(state) && st === 'beantragt'
                  ? `<section class="akt-drawer__section akt-drawer__section--action">
            <h4>Entscheidung</h4>
            <p class="akt-drawer__lead">Antrag prüfen und entscheiden.</p>
          </section>`
                  : ''
          }
        </div>

        <footer class="akt-drawer__foot">
          ${primaryActions.join('')}
          ${secondaryActions.join('')}
          ${moreMenu}
        </footer>
      </aside>
    </div>`;
}

/**
 * Success-State nach Einreichung.
 * @param {object} state
 */
export function renderSubmitSuccess(state) {
    const s = state.submitSuccess;
    if (!s) return '';
    return `
    <section class="akt-success" role="status">
      <div class="akt-success__card">
        <p class="akt-success__eyebrow">Antrag eingereicht ✓</p>
        <h2>${esc(s.titel || 'Aktivität')} wurde zur Genehmigung eingereicht.</h2>
        <p class="akt-status akt-status--info"><span class="akt-status__dot" aria-hidden="true"></span>Eingereicht</p>
        <p class="akt-success__hint">Du musst aktuell nichts weiter tun. Änderungen oder Rückfragen siehst du unter Meine Aktivitäten.</p>
        <div class="akt-success__actions">
          <button type="button" class="btn akt-btn-primary" data-akt-success-open="${esc(s.id || '')}">Aktivität anzeigen</button>
          <button type="button" class="btn" data-akt-success-home>Zur Übersicht</button>
        </div>
      </div>
    </section>`;
}
