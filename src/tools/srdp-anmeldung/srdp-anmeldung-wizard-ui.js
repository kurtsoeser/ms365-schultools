/**
 * Guided Wizard für sRDP-Anmeldung (Schüler Create).
 */
import { emptyForm, classGroupLinksFromAppData } from './srdp-anmeldung-state.js';
import {
    buildPruefplanKurz,
    buildPruefplanDetail,
    isSeminarWahlfach,
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

const STEPS = [
    { id: 1, label: 'Person & Klasse' },
    { id: 2, label: 'Variante & DA' },
    { id: 3, label: 'Prüfungen' },
    { id: 4, label: 'Bestätigen' }
];

/**
 * @param {object} state
 */
export function renderAntragWizard(state) {
    const f = state.form || emptyForm();
    const o = state.options || {};
    const step = Math.min(4, Math.max(1, Number(f.wizardStep) || 1));
    const showSeminar = isSeminarWahlfach(f.wahlfachMuendlich);
    const livePlan = f.variante
        ? buildPruefplanDetail(f, o.profileId, { terminJahr: o.jahr, bfkDatum: o.bfkDatum || '' })
              .schriftlich.map(
                  (row) =>
                      row.part + (row.datum ? ' · ' + row.datum : row.datumHinweis ? ' · ' + row.datumHinweis : '')
              )
              .join(', ')
        : '';
    const errors = state.formErrors || [];
    const accountEmail = String(
        (state && (state.demoPersonaEmail || state.accountEmail)) || ''
    )
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
    // Konto-E-Mail ist maßgeblich; kurz nach MSAL-Init kann sie noch fehlen – dann nicht hart sperren,
    // wenn die Auth-UI bereits „angemeldet“ meldet (Race beim Seitenstart).
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

    const stepsHtml = STEPS.map(
        (s) =>
            `<li class="srdp-wizard__step${s.id === step ? ' is-active' : ''}${
                s.id < step ? ' is-done' : ''
            }"><span>${s.id}</span>${esc(s.label)}</li>`
    ).join('');

    return `<section class="srdp-panel srdp-panel--calm srdp-panel--wizard">
      <h3>Anmeldung Haupttermin ${esc(o.jahr || '')}</h3>
      <p class="srdp-panel__hint">Schulform ${esc(o.profileLabel || 'HAK')} · Schritt ${step} von 4</p>
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
            <label for="srdpVariante">Variante</label>
            <select id="srdpVariante" name="variante" required>
              <option value="">– wählen –</option>
              ${optionsHtml(
                  (o.variants || []).map((v) => v.label),
                  f.variante
              )}
            </select>
          </div>
          <div class="srdp-field srdp-field--full">
            <label for="srdpDA">Titel Diplomarbeit</label>
            <textarea id="srdpDA" name="titelDiplomarbeit">${esc(f.titelDiplomarbeit)}</textarea>
          </div>
          <div class="srdp-field srdp-field--full">
            <label for="srdpBetreuung">BetreuungslehrerIn Diplomarbeit</label>
            <select id="srdpBetreuung" name="betreuungslehrerDA">
              <option value="">– wählen –</option>
              ${optionsHtml(o.lehrer, f.betreuungslehrerDA)}
            </select>
          </div>
          ${
              livePlan
                  ? `<div class="srdp-live-plan srdp-field--full" id="srdpLivePlan"><strong>Schriftlich:</strong> ${esc(
                        livePlan
                    )}</div>`
                  : '<div class="srdp-live-plan srdp-field--full" id="srdpLivePlan" hidden></div>'
          }
        </div>

        <div class="srdp-wizard__pane${step === 3 ? ' is-visible' : ''}" data-srdp-wizard-pane="3">
          <div class="srdp-field-pair srdp-field--full">
            <div class="srdp-field">
              <label for="srdpLfs">${esc(o.lfsFieldLabel || 'LFS')}</label>
              <select id="srdpLfs" name="lfs">
                <option value="">– wählen –</option>
                ${optionsHtml(o.lfs, f.lfs)}
              </select>
            </div>
            <div class="srdp-field">
              <label for="srdpLehrerLfs">LehrerIn ${esc(o.lfsFieldLabel || 'LFS')}</label>
              <select id="srdpLehrerLfs" name="lehrerLFS">
                <option value="">– wählen –</option>
                ${optionsHtml(o.lehrer, f.lehrerLFS)}
              </select>
            </div>
          </div>

          <div class="srdp-field srdp-field--full">
            <label for="srdpLehrerBko">LehrerIn BKO mündlich</label>
            <select id="srdpLehrerBko" name="lehrerBKO">
              <option value="">– wählen –</option>
              ${optionsHtml(o.lehrer, f.lehrerBKO)}
            </select>
            <p class="srdp-field__hint">Fach BKO ist durch die Variante vorgegeben – hier nur die Prüfer:in wählen.</p>
          </div>

          <div class="srdp-field-pair srdp-field--full">
            <div class="srdp-field">
              <label for="srdpWahlfach">Wahlfach mündlich</label>
              <select id="srdpWahlfach" name="wahlfachMuendlich">
                <option value="">– wählen –</option>
                ${optionsHtml(o.wahlfaecher, f.wahlfachMuendlich)}
              </select>
            </div>
            <div class="srdp-field">
              <label for="srdpLehrerWahl">LehrerIn Wahlfach</label>
              <select id="srdpLehrerWahl" name="lehrerWahlfach">
                <option value="">– wählen –</option>
                ${optionsHtml(o.lehrer, f.lehrerWahlfach)}
              </select>
            </div>
          </div>

          <div class="srdp-field srdp-field--full"${showSeminar ? '' : ' hidden'} id="srdpSeminarWrap">
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
            <p class="srdp-field__hint">Pflicht, wenn das Wahlfach mit „Seminar“ beginnt. Liste oder eigene Bezeichnung.</p>
          </div>
          ${f.variante ? `<div class="srdp-field srdp-field--full">${planSplitHtml(f, o.profileId, o.jahr, o.bfkDatum || '')}</div>` : ''}
        </div>

        <div class="srdp-wizard__pane${step === 4 ? ' is-visible' : ''}" data-srdp-wizard-pane="4">
          <div class="srdp-field srdp-field--full">
            <p class="srdp-panel__hint">Bitte prüfe deine Angaben. Mit der Bestätigung wird die Anmeldung gespeichert.</p>
            <ul class="srdp-wizard__review">
              <li><strong>Name:</strong> ${esc(f.nachname)}, ${esc(f.vorname)}</li>
              <li><strong>Klasse:</strong> ${esc(f.klasse)}</li>
              <li><strong>Variante:</strong> ${esc(f.variante)}</li>
            </ul>
            ${planSplitHtml(f, o.profileId, o.jahr, o.bfkDatum || '')}
          </div>
          <div class="srdp-field srdp-field--full">
            <label class="srdp-check">
              <input type="checkbox" id="srdpBestaetigung" name="bestaetigung"${
                  f.bestaetigung ? ' checked' : ''
              }>
              <span>Ich bestätige die Angaben meiner sRDP-Anmeldung.</span>
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
