/**
 * Guided Wizard für sRDP/sRP-Anmeldung (profilgesteuert).
 */
import { emptyForm, classGroupLinksFromAppData } from './srdp-anmeldung-state.js';
import { getProfile } from './srdp-anmeldung-profiles.js';
import {
    buildPruefplanDetail,
    describeExamSlots,
    isSeminarWahlfach,
    parseTeacherList,
    resolveStudentAssignedKlasse
} from './srdp-anmeldung-logic.js';

function esc(s) {
    return String(s ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function optionsHtml(list, selected) {
    const sel = String(selected || '');
    return (Array.isArray(list) ? list : [])
        .map((v) => {
            const val = typeof v === 'object' ? String(v.label || v.code || '') : String(v);
            return `<option value="${esc(val)}"${val === sel ? ' selected' : ''}>${esc(val)}</option>`;
        })
        .join('');
}

function planDateHtml(row) {
    if (row && row.datum) return `<time datetime="${esc(row.datumIso)}">${esc(row.datum)}</time>`;
    if (row && row.datumHinweis) return `<time class="is-school">${esc(row.datumHinweis)}</time>`;
    return '';
}

function planSplitHtml(item, profileId, terminJahr, bfkDatum) {
    const plan = buildPruefplanDetail(item, profileId, { terminJahr, bfkDatum });
    const col = (title, rows, mod) => {
        const body = (rows || [])
            .map((r) => {
                if (!r.part) return '';
                return `<li><strong>${esc(r.part)}</strong>${planDateHtml(r)}${
                    r.detail ? `<span>${esc(r.detail)}</span>` : ''
                }</li>`;
            })
            .join('');
        return `<section class="srdp-plan-col srdp-plan-col--${mod}">
          <h4>${esc(title)}</h4>
          <ul>${body || '<li><span>Keine Teile</span></li>'}</ul>
        </section>`;
    };
    if (!plan.schriftlich.length && !plan.muendlich.length) return '';
    return `<div class="srdp-plan-split">
      ${col('Schriftlich', plan.schriftlich, 'schriftlich')}
      ${col('Mündlich', plan.muendlich, 'muendlich')}
    </div>`;
}

/**
 * @param {string[]} lehrer
 * @param {string} selectedCsv
 * @param {string} fieldName
 * @param {string} idPrefix
 */
function teacherMultiHtml(lehrer, selectedCsv, fieldName, idPrefix) {
    const selected = new Set(parseTeacherList(selectedCsv).map((s) => s.toLowerCase()));
    const list = Array.isArray(lehrer) ? lehrer : [];
    if (!list.length) {
        return `<p class="srdp-field__hint">Keine Lehrkräfte in den Stammdaten.</p>
      <input type="hidden" id="${esc(idPrefix)}" name="${esc(fieldName)}" value="${esc(selectedCsv || '')}">`;
    }
    const boxes = list
        .map((name, i) => {
            const id = idPrefix + '_' + i;
            const checked = selected.has(String(name).toLowerCase()) ? ' checked' : '';
            return `<label class="srdp-teacher-chip" for="${esc(id)}">
        <input type="checkbox" id="${esc(id)}" data-srdp-teacher-field="${esc(fieldName)}" value="${esc(name)}"${checked}>
        <span>${esc(name)}</span>
      </label>`;
        })
        .join('');
    return `<div class="srdp-teacher-multi" data-srdp-teacher-group="${esc(fieldName)}" role="group" aria-label="${esc(
        fieldName
    )}">
      ${boxes}
      <input type="hidden" id="${esc(idPrefix)}" name="${esc(fieldName)}" value="${esc(selectedCsv || '')}">
    </div>
    <p class="srdp-field__hint">Mehrfachauswahl möglich (Prüfungskommission).</p>`;
}

function subjectOptionsForSlot(slot, o) {
    if (!slot || !slot.subjectField) return [];
    if (slot.subjectField === 'lfs') return o.lfs || [];
    if (slot.subjectField === 'wahlfachMuendlich') return o.wahlfaecher || [];
    if (slot.subjectField === 'schwerpunktFach') return o.schwerpunkte || [];
    if (slot.subjectField === 'fachkolloquium') return o.fachkolloquien || o.wahlfaecher || [];
    if (slot.subjectField === 'fachtheorie') return o.fachtheorien || ['Didaktik', 'Pädagogik'];
    if (slot.subjectField === 'klausurKombi') {
        return [
            'LFS + AM (BW-RW mündlich)',
            'LFS + BW-RW (AM mündlich)',
            'AM + BW-RW (LFS mündlich)',
            'alle vier (Variante 2)'
        ];
    }
    if (slot.subjectField === 'muendlichExtra') return ['Angewandte Mathematik', 'Englisch'];
    if (slot.subjectField === 'vierteKlausur') {
        return o.vierteKlausuren || ['Weitere lebende Fremdsprache', 'Latein', 'Darstellende Geometrie'];
    }
    if (
        slot.subjectField === 'muendlich1' ||
        slot.subjectField === 'muendlich2' ||
        slot.subjectField === 'muendlich3'
    ) {
        return o.muendlichFaecher || o.wahlfaecher || [];
    }
    return [];
}

function subjectFieldId(field) {
    const map = {
        lfs: 'srdpLfs',
        wahlfachMuendlich: 'srdpWahlfach',
        schwerpunktFach: 'srdpSchwerpunkt',
        fachkolloquium: 'srdpFachkolloquium',
        fachtheorie: 'srdpFachtheorie',
        klausurKombi: 'srdpKlausurKombi',
        muendlichExtra: 'srdpMuendlichExtra',
        vierteKlausur: 'srdpVierteKlausur',
        muendlich1: 'srdpMuendlich1',
        muendlich2: 'srdpMuendlich2',
        muendlich3: 'srdpMuendlich3'
    };
    return map[field] || 'srdpSubj_' + field;
}

function teacherFieldId(field) {
    const map = {
        lehrerLFS: 'srdpLehrerLfs',
        lehrerBKO: 'srdpLehrerBko',
        lehrerAM: 'srdpLehrerAm',
        lehrerWahlfach: 'srdpLehrerWahl',
        lehrerSchwerpunkt: 'srdpLehrerSchwerpunkt',
        lehrerFachkolloquium: 'srdpLehrerFachkolloquium',
        lehrerFachtheorie: 'srdpLehrerFachtheorie',
        lehrerMuendlichExtra: 'srdpLehrerMuendlichExtra',
        lehrerMuendlich1: 'srdpLehrerMuendlich1',
        lehrerMuendlich2: 'srdpLehrerMuendlich2',
        lehrerMuendlich3: 'srdpLehrerMuendlich3'
    };
    return map[field] || 'srdpLehrer_' + field;
}

const STEPS = [
    { id: 1, label: 'Person & Klasse' },
    { id: 2, label: 'Variante & Arbeit' },
    { id: 3, label: 'Prüfungen' },
    { id: 4, label: 'Bestätigen' }
];

/**
 * @param {object} state
 */
export function renderAntragWizard(state) {
    const f = state.form || emptyForm();
    const o = state.options || {};
    const profile = getProfile(o.profileId);
    const step = Math.min(4, Math.max(1, Number(f.wizardStep) || 1));
    const showSeminar = isSeminarWahlfach(f.wahlfachMuendlich);
    const slots = describeExamSlots(f.variante, o.profileId);
    const isAhs = profile.id === 'ahs';
    const arbeitLabel = profile.arbeitLabel || 'Diplomarbeit';
    const errors = state.formErrors || [];
    const accountEmail = String((state && (state.demoPersonaEmail || state.accountEmail)) || '')
        .trim()
        .toLowerCase();
    let loggedInHint = false;
    try {
        loggedInHint =
            typeof window !== 'undefined' &&
            typeof window.ms365AuthIsLoggedIn === 'function' &&
            window.ms365AuthIsLoggedIn();
    } catch {
        loggedInHint = false;
    }
    const displayName = [f.vorname, f.nachname].filter(Boolean).join(' ').trim();
    const nameKnown = !!(f.nachname && f.vorname);
    const needsLogin = !accountEmail && !loggedInHint;
    const abschlussKlassen = (Array.isArray(o.klassen) ? o.klassen : [])
        .map((k) => String(k || '').trim())
        .filter(Boolean);
    const ownKlasse = resolveStudentAssignedKlasse({
        email: accountEmail,
        students: (state.stammdaten && state.stammdaten.students) || [],
        classes: (state.stammdaten && state.stammdaten.classes) || [],
        classLinks: classGroupLinksFromAppData(),
        memberGroups: state.memberGroups || [],
        preferredCodes: abschlussKlassen
    });
    const classChoices = ownKlasse ? [ownKlasse] : abschlussKlassen;
    const selectedKlasse = ownKlasse || f.klasse || '';
    const abaHits = Array.isArray(state.abschlussarbeitHits) ? state.abschlussarbeitHits : [];

    const stepsHtml = STEPS.map(
        (s) =>
            `<li class="srdp-wizard__step${s.id === step ? ' is-active' : ''}${
                s.id < step ? ' is-done' : ''
            }"><span>${s.id}</span>${esc(s.label)}</li>`
    ).join('');

    const variantCards = (o.variants || profile.variants || [])
        .map((v, idx) => {
            const label = v.label || 'Variante ' + (v.key || idx + 1);
            const active = f.variante === label;
            const mod = 'v' + String(v.key || idx + 1);
            return `<label class="srdp-variant-card srdp-variant-card--${esc(mod)}${
                active ? ' is-active' : ''
            }">
        <input type="radio" name="variante" value="${esc(label)}" ${active ? 'checked' : ''} data-srdp-variante-card>
        <span class="srdp-variant-card__title">${esc(label)}</span>
        <span class="srdp-variant-card__block"><strong>schriftlich</strong> ${(v.schriftlich || [])
            .map((p) => esc(p))
            .join(' · ')}</span>
        <span class="srdp-variant-card__block"><strong>mündlich</strong> ${(v.muendlich || [])
            .map((p) => esc(p))
            .join(' · ')}</span>
      </label>`;
        })
        .join('');

    const arbeitBlock = isAhs
        ? `<div class="srdp-field srdp-field--full">
            <label class="srdp-check">
              <input type="checkbox" id="srdpHatAba" name="hatABA"${f.hatABA ? ' checked' : ''}>
              <span>Mit abschließender Arbeit (ABA / VWA)</span>
            </label>
          </div>
          <div class="srdp-field srdp-field--full">
            <label for="srdpDA">${esc(arbeitLabel)} – Titel</label>
            <textarea id="srdpDA" name="titelABA">${esc(f.titelABA || f.titelDiplomarbeit)}</textarea>
          </div>
          <div class="srdp-field srdp-field--full">
            <label for="srdpBetreuung">BetreuungslehrerIn ABA</label>
            <select id="srdpBetreuung" name="betreuungslehrerABA">
              <option value="">– wählen –</option>
              ${optionsHtml(o.lehrer, f.betreuungslehrerABA || f.betreuungslehrerDA)}
            </select>
          </div>`
        : `<div class="srdp-field srdp-field--full">
            <label for="srdpDA">Titel ${esc(arbeitLabel)}</label>
            <textarea id="srdpDA" name="titelDiplomarbeit">${esc(f.titelDiplomarbeit)}</textarea>
          </div>
          <div class="srdp-field srdp-field--full">
            <label for="srdpBetreuung">BetreuungslehrerIn ${esc(arbeitLabel)}</label>
            <select id="srdpBetreuung" name="betreuungslehrerDA">
              <option value="">– wählen –</option>
              ${optionsHtml(o.lehrer, f.betreuungslehrerDA)}
            </select>
          </div>`;

    const abaHint = f.abaPrefill
        ? `<div class="srdp-alert srdp-alert--ok srdp-field--full" role="status">Titel und Betreuung aus ABA-Portal vorausgefüllt – bitte prüfen.</div>`
        : '';

    const abaChoice =
        abaHits.length > 1
            ? `<div class="srdp-field srdp-field--full">
          <label for="srdpAbaPick">Mehrere Abschlussarbeiten gefunden</label>
          <select id="srdpAbaPick" data-srdp-aba-pick>
            <option value="">– bitte wählen –</option>
            ${abaHits
                .map(
                    (h, i) =>
                        `<option value="${i}">${esc(h.titel || '')} · ${esc(h.betreuer || '')}</option>`
                )
                .join('')}
          </select>
        </div>`
            : '';

    /** already rendered subject/teacher fields */
    const seenSubject = new Set();
    const seenTeacher = new Set();
    const examBlocks = [];

    if (slots.needsLfsCode) {
        seenSubject.add('lfs');
        examBlocks.push(`<div class="srdp-exam-card srdp-field--full">
        <h4>${esc(o.lfsFieldLabel || profile.lfsFieldLabel || 'LFS')}</h4>
        <p class="srdp-field__hint">Sprache wählen${
            slots.schriftlich.some((s) => s.key === 'lfs')
                ? ' (schriftlich zentral – kein Prüfer nötig)'
                : ' (mündlich – Prüfer unten)'
        }.</p>
        <div class="srdp-field">
          <label for="srdpLfs">${esc(o.lfsFieldLabel || 'LFS')}</label>
          <select id="srdpLfs" name="lfs">
            <option value="">– wählen –</option>
            ${optionsHtml(o.lfs, f.lfs)}
          </select>
        </div>
      </div>`);
    }

    if (slots.schriftlich.length) {
        examBlocks.push(`<div class="srdp-exam-card srdp-exam-card--written srdp-field--full">
        <h4>Schriftlich</h4>
        <ul class="srdp-exam-list">${slots.schriftlich
            .map(
                (s) =>
                    `<li><strong>${esc(s.part)}</strong> <span class="srdp-pill">zentral / schulisch – kein Prüfer</span></li>`
            )
            .join('')}</ul>
      </div>`);
    }

    slots.muendlich.forEach((slot) => {
        const parts = [];
        if (slot.subjectField && slot.subjectField !== 'lfs' && !seenSubject.has(slot.subjectField)) {
            seenSubject.add(slot.subjectField);
            const sid = subjectFieldId(slot.subjectField);
            const opts = subjectOptionsForSlot(slot, o);
            parts.push(`<div class="srdp-field">
          <label for="${esc(sid)}">${esc(slot.subjectLabel || slot.part)}</label>
          <select id="${esc(sid)}" name="${esc(slot.subjectField)}"${
                slot.subjectField === 'wahlfachMuendlich' ? ' data-srdp-wahlfach' : ''
            }>
            <option value="">– wählen –</option>
            ${optionsHtml(opts, f[slot.subjectField])}
          </select>
        </div>`);
        }
        if (slot.teacherField && slot.needsTeacher && !seenTeacher.has(slot.teacherField)) {
            seenTeacher.add(slot.teacherField);
            const tid = teacherFieldId(slot.teacherField);
            parts.push(`<div class="srdp-field srdp-field--full">
          <span class="srdp-field__label">${esc(slot.teacherLabel || 'LehrerIn')}</span>
          ${teacherMultiHtml(o.lehrer, f[slot.teacherField], slot.teacherField, tid)}
        </div>`);
        }
        if (!parts.length && !slot.teacherField) {
            parts.push(`<p class="srdp-field__hint">${esc(slot.part)} – keine Eingabe nötig.</p>`);
        }
        examBlocks.push(`<div class="srdp-exam-card srdp-exam-card--oral srdp-field--full">
        <h4>Mündlich: ${esc(slot.part)}</h4>
        ${parts.join('')}
      </div>`);
    });

    if (slots.muendlich.some((s) => s.key === 'wahlfach') || f.wahlfachMuendlich) {
        examBlocks.push(`<div class="srdp-field srdp-field--full"${showSeminar ? '' : ' hidden'} id="srdpSeminarWrap">
        <label for="srdpSeminar">Seminar</label>
        ${
            (o.seminare || []).length
                ? `<select id="srdpSeminar" name="seminar">
            <option value="">– wählen –</option>
            ${optionsHtml(o.seminare, f.seminar)}
          </select>`
                : '<p class="srdp-field__hint">In der Liste sind noch keine Seminar-Bezeichnungen hinterlegt.</p>'
        }
        <label for="srdpSeminarText">Seminar-Bezeichnung</label>
        <input type="text" id="srdpSeminarText" name="seminarText" value="${esc(
            f.seminar || ''
        )}" placeholder="Bezeichnung, falls sie nicht in der Liste steht" autocomplete="off">
        <p class="srdp-field__hint">Pflicht, wenn das Wahlfach mit „Seminar“ beginnt.</p>
      </div>`);
    }

    return `<section class="srdp-panel srdp-panel--calm srdp-panel--wizard">
      <h3>Anmeldung Haupttermin ${esc(o.jahr || '')}</h3>
      <p class="srdp-panel__hint">Schulform ${esc(o.profileLabel || profile.label || 'HAK')} · Schritt ${step} von 4</p>
      <ol class="srdp-wizard__steps">${stepsHtml}</ol>
      ${
          errors.length
              ? `<div class="srdp-alert srdp-alert--bad" role="alert">${errors
                    .map((e) => esc(e))
                    .join('<br>')}</div>`
              : ''
      }
      ${
          needsLogin
              ? `<div class="srdp-alert srdp-alert--bad" role="status">Bitte oben rechts mit deinem Schul-Konto anmelden – sonst kann die Anmeldung nicht gespeichert werden.</div>`
              : ''
      }
      <form id="srdpForm" class="srdp-form-grid" autocomplete="on" data-srdp-wizard-step="${step}">
        <div class="srdp-wizard__pane${step === 1 ? ' is-visible' : ''}" data-srdp-wizard-pane="1">
          ${
              nameKnown
                  ? `<div class="srdp-identity srdp-field--full">
              <span class="srdp-identity__label">Angemeldet als</span>
              <strong class="srdp-identity__name">${esc(displayName)}</strong>
              ${
                  accountEmail
                      ? `<span class="srdp-identity__mail">${esc(accountEmail)}</span>`
                      : ''
              }
              <input type="hidden" id="srdpNachname" name="nachname" value="${esc(f.nachname)}">
              <input type="hidden" id="srdpVorname" name="vorname" value="${esc(f.vorname)}">
            </div>`
                  : `<div class="srdp-field">
              <label for="srdpNachname">Nachname</label>
              <input type="text" id="srdpNachname" name="nachname" value="${esc(f.nachname)}" required>
            </div>
            <div class="srdp-field">
              <label for="srdpVorname">Vorname</label>
              <input type="text" id="srdpVorname" name="vorname" value="${esc(f.vorname)}" required>
            </div>`
          }
          <div class="srdp-field srdp-field--full">
            <label for="srdpKlasse">Klasse</label>
            ${
                classChoices.length
                    ? `<select id="srdpKlasse" name="klasse" required>
              ${ownKlasse ? '' : '<option value="">– Klasse wählen –</option>'}
              ${optionsHtml(classChoices, selectedKlasse)}
            </select>`
                    : `<p class="srdp-panel__hint">Für diesen Haupttermin ist noch keine Abschlussklasse hinterlegt.</p>
            <select id="srdpKlasse" name="klasse" required disabled>
              <option value="">– keine Klasse –</option>
            </select>`
            }
          </div>
        </div>

        <div class="srdp-wizard__pane${step === 2 ? ' is-visible' : ''}" data-srdp-wizard-pane="2">
          <div class="srdp-field srdp-field--full">
            <span class="srdp-field__label">Prüfungsvariante</span>
            <div class="srdp-variant-grid" role="radiogroup" aria-label="Variante">
              ${variantCards || '<p class="srdp-panel__hint">Keine Varianten im Profil.</p>'}
            </div>
            <input type="hidden" id="srdpVariante" name="varianteHidden" value="${esc(f.variante)}">
          </div>
          ${abaHint}
          ${abaChoice}
          ${arbeitBlock}
          ${
              f.variante
                  ? `<div class="srdp-field srdp-field--full">${planSplitHtml(
                        f,
                        o.profileId,
                        o.jahr,
                        o.bfkDatum || ''
                    )}</div>`
                  : ''
          }
        </div>

        <div class="srdp-wizard__pane${step === 3 ? ' is-visible' : ''}" data-srdp-wizard-pane="3">
          ${
              f.variante
                  ? examBlocks.join('')
                  : '<p class="srdp-panel__hint srdp-field--full">Bitte zuerst eine Variante wählen.</p>'
          }
          ${
              f.variante
                  ? `<div class="srdp-field srdp-field--full">${planSplitHtml(
                        f,
                        o.profileId,
                        o.jahr,
                        o.bfkDatum || ''
                    )}</div>`
                  : ''
          }
        </div>

        <div class="srdp-wizard__pane${step === 4 ? ' is-visible' : ''}" data-srdp-wizard-pane="4">
          <div class="srdp-field srdp-field--full">
            <p class="srdp-panel__hint">Bitte prüfe deine Angaben. Mit der Bestätigung wird die Anmeldung gespeichert.</p>
            <ul class="srdp-wizard__review">
              <li><strong>Name:</strong> ${esc(f.nachname)}, ${esc(f.vorname)}</li>
              <li><strong>Klasse:</strong> ${esc(f.klasse)}</li>
              <li><strong>Variante:</strong> ${esc(f.variante)}</li>
              <li><strong>${esc(arbeitLabel)}:</strong> ${esc(
                  isAhs ? f.titelABA || f.titelDiplomarbeit : f.titelDiplomarbeit
              )}</li>
            </ul>
            ${planSplitHtml(f, o.profileId, o.jahr, o.bfkDatum || '')}
          </div>
          <div class="srdp-field srdp-field--full">
            <label class="srdp-check">
              <input type="checkbox" id="srdpBestaetigung" name="bestaetigung"${
                  f.bestaetigung ? ' checked' : ''
              }>
              <span>Ich bestätige die Angaben meiner ${esc(profile.examShort || 'sRDP')}-Anmeldung.</span>
            </label>
          </div>
        </div>
      </form>
      <div class="srdp-actions">
        ${
            step > 1
                ? `<button type="button" class="btn" data-srdp-wizard-prev><i class="bi bi-arrow-left"></i>Zurück</button>`
                : ''
        }
        ${
            step < 4
                ? `<button type="button" class="btn btn-success" data-srdp-wizard-next${
                      needsLogin ? ' disabled' : ''
                  }>Weiter<i class="bi bi-arrow-right"></i></button>`
                : `<button type="button" class="btn btn-success" data-srdp-submit ${
                      state.saving || needsLogin ? 'disabled' : ''
                  }>
            <i class="bi bi-check2-circle"></i>${state.saving ? 'Speichern…' : 'Anmeldung absenden'}
          </button>`
        }
      </div>
    </section>`;
}
