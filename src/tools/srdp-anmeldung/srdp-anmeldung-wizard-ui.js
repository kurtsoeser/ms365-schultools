/**
 * Guided Wizard für sRDP-Anmeldung (Schüler Create).
 */
import { emptyForm } from './srdp-anmeldung-state.js';
import { buildPruefplanKurz, isSeminarWahlfach } from './srdp-anmeldung-logic.js';

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
    const livePlan = f.variante ? buildPruefplanKurz(f.variante, o.profileId) : '';
    const errors = state.formErrors || [];

    const stepsHtml = STEPS.map(
        (s) =>
            `<li class="srdp-wizard__step${s.id === step ? ' is-active' : ''}${
                s.id < step ? ' is-done' : ''
            }"><span>${s.id}</span>${esc(s.label)}</li>`
    ).join('');

    return `<section class="srdp-panel srdp-panel--calm">
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
      <form id="srdpForm" class="srdp-form-grid" autocomplete="on" data-srdp-wizard-step="${step}">
        <div class="srdp-wizard__pane${step === 1 ? ' is-visible' : ''}" data-srdp-wizard-pane="1">
          <div class="srdp-field">
            <label for="srdpKlasse">Klasse</label>
            <select id="srdpKlasse" name="klasse" required>
              <option value="">– wählen –</option>
              ${optionsHtml(o.klassen, f.klasse)}
            </select>
          </div>
          <div class="srdp-field">
            <label for="srdpNachname">Nachname</label>
            <input type="text" id="srdpNachname" name="nachname" value="${esc(f.nachname)}" required>
          </div>
          <div class="srdp-field">
            <label for="srdpVorname">Vorname</label>
            <input type="text" id="srdpVorname" name="vorname" value="${esc(f.vorname)}" required>
          </div>
        </div>

        <div class="srdp-wizard__pane${step === 2 ? ' is-visible' : ''}" data-srdp-wizard-pane="2">
          <div class="srdp-field">
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
          <div class="srdp-field">
            <label for="srdpBetreuung">BetreuungslehrerIn Diplomarbeit</label>
            <select id="srdpBetreuung" name="betreuungslehrerDA">
              <option value="">– wählen –</option>
              ${optionsHtml(o.lehrer, f.betreuungslehrerDA)}
            </select>
          </div>
          ${
              livePlan
                  ? `<div class="srdp-live-plan srdp-field--full" id="srdpLivePlan"><strong>Prüfplan:</strong> ${esc(
                        livePlan
                    )}</div>`
                  : '<div class="srdp-live-plan srdp-field--full" id="srdpLivePlan" hidden></div>'
          }
        </div>

        <div class="srdp-wizard__pane${step === 3 ? ' is-visible' : ''}" data-srdp-wizard-pane="3">
          <div class="srdp-field">
            <label for="srdpLfs">LFS</label>
            <select id="srdpLfs" name="lfs">
              <option value="">– wählen –</option>
              ${optionsHtml(o.lfs, f.lfs)}
            </select>
          </div>
          <div class="srdp-field">
            <label for="srdpLehrerLfs">LehrerIn LFS</label>
            <select id="srdpLehrerLfs" name="lehrerLFS">
              <option value="">– wählen –</option>
              ${optionsHtml(o.lehrer, f.lehrerLFS)}
            </select>
          </div>
          <div class="srdp-field">
            <label for="srdpLehrerBko">LehrerIn BKO mündlich</label>
            <select id="srdpLehrerBko" name="lehrerBKO">
              <option value="">– wählen –</option>
              ${optionsHtml(o.lehrer, f.lehrerBKO)}
            </select>
          </div>
          <div class="srdp-field">
            <label for="srdpWahlfach">Wahlfach mündlich</label>
            <select id="srdpWahlfach" name="wahlfachMuendlich">
              <option value="">– wählen –</option>
              ${optionsHtml(o.wahlfaecher, f.wahlfachMuendlich)}
            </select>
          </div>
          <div class="srdp-field"${showSeminar ? '' : ' hidden'} id="srdpSeminarWrap">
            <label for="srdpSeminar">Seminar</label>
            <select id="srdpSeminar" name="seminar">
              <option value="">– wählen –</option>
              ${optionsHtml(o.seminare, f.seminar)}
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

        <div class="srdp-wizard__pane${step === 4 ? ' is-visible' : ''}" data-srdp-wizard-pane="4">
          <div class="srdp-field srdp-field--full">
            <p class="srdp-panel__hint">Bitte prüfe deine Angaben. Mit der Bestätigung wird die Anmeldung gespeichert.</p>
            <ul class="srdp-wizard__review">
              <li><strong>Name:</strong> ${esc(f.nachname)}, ${esc(f.vorname)}</li>
              <li><strong>Klasse:</strong> ${esc(f.klasse)}</li>
              <li><strong>Variante:</strong> ${esc(f.variante)}</li>
              <li><strong>Diplomarbeit:</strong> ${esc(f.titelDiplomarbeit || '–')}</li>
              <li><strong>LFS:</strong> ${esc(f.lfs || '–')} · ${esc(f.lehrerLFS || '–')}</li>
              <li><strong>Wahlfach:</strong> ${esc(f.wahlfachMuendlich || '–')}${
                  f.seminar ? ' · Seminar ' + esc(f.seminar) : ''
              }</li>
            </ul>
            ${
                livePlan
                    ? `<div class="srdp-live-plan"><strong>Prüfplan:</strong> ${esc(livePlan)}</div>`
                    : ''
            }
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
                ? `<button type="button" class="btn btn-success" data-srdp-wizard-next>Weiter<i class="bi bi-arrow-right"></i></button>`
                : `<button type="button" class="btn btn-success" data-srdp-submit ${
                      state.saving ? 'disabled' : ''
                  }>
            <i class="bi bi-check2-circle"></i>${state.saving ? 'Speichern…' : 'Anmeldung absenden'}
          </button>`
        }
      </div>
    </section>`;
}
