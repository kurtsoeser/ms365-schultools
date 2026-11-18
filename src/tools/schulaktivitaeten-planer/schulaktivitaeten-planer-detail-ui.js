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
 * @param {object} state
 * @param {{ code?: string, email?: string, name?: string }} person
 * @param {string} missingLabel
 */
function renderPersonLine(state, person, missingLabel) {
    const p = resolvePerson(state, person);
    if (!p.name && !p.email) {
        return `<span class="akt-missing">${esc(missingLabel || 'Nicht erfasst')}</span>`;
    }
    if (p.name) {
        return `<span class="akt-person"><strong>${esc(p.name)}</strong>${
            p.email && p.email.toLowerCase() !== p.name.toLowerCase()
                ? `<span class="akt-person__email muted">${esc(p.email)}</span>`
                : ''
        }</span>`;
    }
    return `<span class="akt-person">${esc(p.email)}</span>`;
}

function parseVerlauf(aenderungsLog) {
    const raw = String(aenderungsLog || '').trim();
    if (!raw) return [];
    return raw
        .split(/\n+/)
        .map((line) => line.trim())
        .filter(Boolean)
        .map((line) => {
            const m = /^(\d{4}-\d{2}-\d{2}|\d{1,2}\.\d{1,2}\.\d{2,4})?\s*[·.]?\s*(.*)$/.exec(line);
            if (m && m[1]) {
                return { when: m[1], text: m[2] || line };
            }
            return { when: '', text: line };
        })
        .slice(-8)
        .reverse();
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
    const targetList = targets.length
        ? `<ul class="akt-cal-panel__targets">${targets
              .map(
                  (t) =>
                      `<li class="${t.ok ? 'is-ok' : 'is-warn'}"><strong>${esc(t.label)}</strong> · ${esc(
                          t.detail
                      )}</li>`
              )
              .join('')}</ul>`
        : `<p class="muted">Keine Klassen oder Begleitung hinterlegt – Einladung nur manuell möglich.</p>`;
    return `<section class="akt-drawer__section akt-drawer__section--action akt-cal-panel">
      <h4>Kalender &amp; Einladung</h4>
      <p class="akt-drawer__lead">Termin in Outlook übernehmen oder Klasse / Begleitung per Microsoft&nbsp;365 einladen.</p>
      ${
          linked
              ? '<p class="akt-cal-panel__linked"><i class="bi bi-check2-circle" aria-hidden="true"></i> Bereits mit deinem Outlook-Kalender verknüpft.</p>'
              : ''
      }
      <div class="akt-cal-panel__actions">
        <button type="button" class="btn akt-btn-primary" data-akt-cal-mine="${esc(
            id
        )}"><i class="bi bi-calendar-plus" aria-hidden="true"></i>In meinen Outlook-Kalender</button>
        <button type="button" class="btn" data-akt-cal-invite="${esc(
            id
        )}" ${okN ? '' : 'disabled title="Keine einladbaren Adressen gefunden"'}><i class="bi bi-envelope-plus" aria-hidden="true"></i>Termin anlegen &amp; einladen${
            okN ? ` (${okN})` : ''
        }</button>
        <button type="button" class="btn akt-btn-ghost" data-akt-cal-ics="${esc(
            id
        )}"><i class="bi bi-download" aria-hidden="true"></i>ICS herunterladen</button>
      </div>
      <p class="akt-cal-panel__hint muted">Einladung nutzt die verknüpften Klassen-Teams / Verteiler und E-Mails der Begleitung. ICS funktioniert in Outlook, Apple und Google.</p>
      ${targetList}
    </section>`;
}

function renderStatusBlock(it, state) {
    const st = String(it.status || '').toLowerCase();
    const editable = canEditItem(it, state);
    const sched = scheduleBadge(it);
    if (st === 'genehmigt') {
        return `<section class="akt-drawer__section">
      <h4>Status</h4>
      <div class="akt-drawer__badges">${statusBadge(st)}${sched}</div>
      <p class="akt-drawer__lead">${
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
            <p class="akt-drawer__sub">${esc(teilnahmeKurz(it, labels.klasse))}${
                schuelerN ? ` · ${schuelerN} Schüler:innen` : ''
            }</p>
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
            <p>${esc(typLabel(it.typ))}</p>
            ${
                it.beantragtAm
                    ? `<p><strong>Eingereicht am</strong><br>${esc(formatDeDate(it.beantragtAm))}</p>`
                    : ''
            }
            <p><strong>Antrag gestellt von</strong><br>${renderPersonLine(
                state,
                {
                    email: it.beantragtVon,
                    code: it.beantragtVon === it.lehrerEmail ? it.lehrerCode : '',
                    name: ''
                },
                'Antragsteller nicht erfasst'
            )}</p>
            <p><strong>Verantwortliche Lehrkraft</strong><br>${renderPersonLine(
                state,
                {
                    code: it.lehrerCode,
                    email: it.lehrerEmail
                },
                'Verantwortliche Lehrkraft nicht erfasst'
            )}</p>
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
            <p><strong>Begleitung</strong><br>${
                begleitungCodes.length
                    ? begleitungCodes
                          .map((c) =>
                              renderPersonLine(state, { code: c }, 'Begleitperson nicht erfasst')
                          )
                          .join('<br>')
                    : '<span class="akt-missing">Keine Begleitperson eingetragen</span>'
            }</p>
            ${it.verkehrsmittel ? `<p><strong>Verkehrsmittel</strong><br>${esc(it.verkehrsmittel)}</p>` : ''}
            ${it.kostenHinweis ? `<p><strong>Kosten</strong><br>${esc(it.kostenHinweis)}</p>` : ''}
            ${it.externeVortragende ? `<p><strong>Externe</strong><br>${esc(it.externeVortragende)}</p>` : ''}
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
                  ? `<section class="akt-drawer__section">
            <h4>Verlauf</h4>
            <ul class="akt-drawer__timeline">
              ${verlauf
                  .map(
                      (v) =>
                          `<li>${v.when ? `<span class="akt-drawer__twhen">${esc(v.when)}</span>` : ''}<span>${esc(
                              v.text
                          )}</span></li>`
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
