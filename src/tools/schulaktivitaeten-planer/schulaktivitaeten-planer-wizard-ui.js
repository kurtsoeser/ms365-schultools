/**
 * Guided Flow: Neue Aktivität (4 Schritte) – Detail-Refinement.
 */
import {
    validateAktivitaet,
    DEFAULT_RULES,
    joinCodeList,
    klasseCodesOf,
    vorlaufStatusForDraft,
    mapValidationIssueToStep,
    parseEuroFromText,
    kostenLimitForTyp,
    classifyKostenAmpel,
    kostenAmpelLabel
} from './schulaktivitaeten-planer-logic.js';
import {
    ueChoices,
    formatUeZeitraum,
    timesFromUnterrichtseinheiten
} from './schulaktivitaeten-planer-zeitraster.js';
import {
    labelMaps,
    emptyForm,
    teilnahmeModusChoices,
    teilnahmeKurz,
    teilnahmeModusLabel,
    studentCountForKlasse,
    typLabel
} from './schulaktivitaeten-planer-state.js';
import { formatReadableDateRange } from './schulaktivitaeten-planer-detail-ui.js';

function esc(s) {
    return String(s ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function optionList(items, selected, emptyLabel) {
    const opts = [`<option value="">${esc(emptyLabel)}</option>`];
    (items || []).forEach((it) => {
        const v = it.code || it.value;
        const label = it.name || it.label || v;
        opts.push(
            `<option value="${esc(v)}"${String(selected) === String(v) ? ' selected' : ''}>${esc(label)}</option>`
        );
    });
    return opts.join('');
}

function typChoices() {
    return [
        { code: 'Exkursion', name: 'Exkursion' },
        { code: 'Schulaktivitaet', name: 'Schulaktivität' },
        { code: 'Veranstaltung', name: 'Veranstaltung' },
        { code: 'Klassenabwesenheit', name: 'Klassenabwesenheit' },
        { code: 'Sonstiges', name: 'Sonstiges' }
    ];
}

/**
 * Orientierung aus Kostenschwellen – keine Zahlung, keine Sperre.
 * @param {{ typ?: string, kostenHinweis?: string }} f
 * @param {object|null|undefined} rules
 */
function renderKostenAmpelHint(f, rules) {
    const limit = kostenLimitForTyp(f && f.typ, rules);
    const amount = parseEuroFromText(f && f.kostenHinweis);
    const level = Number.isFinite(amount)
        ? classifyKostenAmpel(f && f.typ, amount, rules)
        : 'unknown';
    const label = kostenAmpelLabel(level);
    if (!limit) {
        return `<span class="akt-kosten-hint muted">Wähle zuerst eine Antrags-Art – dann erscheint die Kostenschwelle.</span>`;
    }
    const scale = `<span class="akt-kosten-scale" aria-hidden="true">
      <span class="akt-kosten-scale__ok" style="flex:${Math.max(limit.okBis, 1)}"></span>
      <span class="akt-kosten-scale__warn" style="flex:${Math.max(limit.warnBis - limit.okBis, 1)}"></span>
      <span class="akt-kosten-scale__teuer" style="flex:${Math.max(limit.warnBis * 0.35, 1)}"></span>
    </span>
    <span class="akt-kosten-scale__labels muted">
      <span>bis ${esc(limit.okBis)} €</span>
      <span>bis ${esc(limit.warnBis)} €</span>
      <span>darüber</span>
    </span>`;
    if (label) {
        return `<span class="akt-kosten-hint">
      <span class="akt-kosten-ampel akt-kosten-ampel--${level}">${esc(label)}</span>
      ${scale}
    </span>`;
    }
    return `<span class="akt-kosten-hint">
      <span class="muted">Schwelle ${esc(limit.typ)}: Orientierung für die Direktion.</span>
      ${scale}
    </span>`;
}

function fieldHintsFromPreview(preview) {
    /** @type {Record<string, { error?: string, warn?: string }>} */
    const map = {};
    const put = (field, kind, msg) => {
        if (!map[field]) map[field] = {};
        if (!map[field][kind]) map[field][kind] = msg;
    };
    (preview.errors || []).forEach((raw) => {
        const m = String(raw || '');
        const low = m.toLowerCase();
        if (low.includes('titel')) put('titel', 'error', m);
        else if (low.includes('klasse')) put('klasseCodes', 'error', m);
        else if (low.includes('art der') || low.includes('typ')) put('typ', 'error', m);
        else if (
            low.includes('startdatum') ||
            low.includes('enddatum') ||
            low.includes('datum') ||
            low.includes('beginn') ||
            low.includes('frist') ||
            low.includes('vorlauf')
        )
            put('startdatum', 'error', m);
        else if (low.includes('gruppe')) put('gruppenText', 'error', m);
        else if (low.includes('teilnehmend')) put('teilnehmendeText', 'error', m);
        else if (low.includes('unterrichtseinheit') || low.includes('zeit')) put('startUe', 'error', m);
        else put('_form', 'error', m);
    });
    (preview.warnings || []).forEach((raw) => {
        const m = String(raw || '');
        const low = m.toLowerCase();
        if (low.includes('ort')) put('ort', 'warn', m);
        else if (low.includes('überschneid') || low.includes('ueberschneid')) put('startdatum', 'warn', m);
        else if (low.includes('frist') || low.includes('vorlauf')) put('startdatum', 'warn', m);
        else if (low.includes('wann findet') || low.includes('zeit')) put('startUe', 'warn', m);
        else put('_form', 'warn', m);
    });
    return map;
}

function shouldShowFieldHint(state, field) {
    if (state.formAttempted) return true;
    return !!(state.formTouched && state.formTouched[field]);
}

function fieldMsg(state, hints, name) {
    if (!shouldShowFieldHint(state, name)) return '';
    const h = hints && hints[name];
    if (!h) return '';
    if (h.error) return `<p class="akt-field-msg akt-field-msg--error" role="alert">${esc(h.error)}</p>`;
    if (h.warn) return `<p class="akt-field-msg akt-field-msg--warn">${esc(h.warn)}</p>`;
    return '';
}

function fieldMsgHtml(state, hints, name) {
    return fieldMsg(state, hints, name);
}

/**
 * Hinweise im Wizard aktualisieren – ohne komplettes Re-Render (kein Einfrieren beim Tippen).
 * @param {object} state
 */
export function refreshAntragWizardHints(state) {
    const form = typeof document !== 'undefined' ? document.getElementById('aktForm') : null;
    if (!form || !state || state.view !== 'antrag') return;
    const f = state.form || emptyForm();
    const rules = state.rules || DEFAULT_RULES;
    const preview = validateAktivitaet({
        draft: f,
        existing: state.items || [],
        rules: {
            minVorlaufTage: rules.minVorlaufTage,
            maxGleichzeitigProKlasse: rules.maxGleichzeitigProKlasse,
            vorlaufModus: rules.vorlaufModus,
            genehmigungPfad: rules.genehmigungPfad
        }
    });
    const hints = fieldHintsFromPreview(preview);
    const fields = [
        'titel',
        'ort',
        'typ',
        'startdatum',
        'klasseCodes',
        'startUe',
        'gruppenText',
        'teilnehmendeText'
    ];
    fields.forEach((name) => {
        const control = form.querySelector('[name="' + name + '"]');
        if (!control) return;
        const host =
            control.closest('.akt-field') ||
            control.closest('.akt-date-block') ||
            control.parentElement;
        if (!host) return;
        host.querySelectorAll(':scope > .akt-field-msg').forEach((n) => n.remove());
        // auch nach dem Control (z. B. in Labels)
        let sib = control.nextElementSibling;
        while (sib && sib.classList && sib.classList.contains('akt-field-msg')) {
            const rm = sib;
            sib = sib.nextElementSibling;
            rm.remove();
        }
        const html = fieldMsgHtml(state, hints, name);
        if (!html) return;
        const wrap = document.createElement('div');
        wrap.innerHTML = html;
        const node = wrap.firstElementChild;
        if (!node) return;
        if (control.parentElement === host) {
            control.insertAdjacentElement('afterend', node);
        } else {
            host.appendChild(node);
        }
    });

    const dateBlock = form.querySelector('.akt-date-block');
    if (dateBlock) {
        dateBlock.querySelectorAll('.akt-inline-ok, .akt-inline-warn').forEach((n) => n.remove());
        const vorlaufHtml = renderVorlaufHint(state, f, rules);
        if (vorlaufHtml) {
            const check = dateBlock.querySelector('.akt-check');
            const wrap = document.createElement('div');
            wrap.innerHTML = vorlaufHtml;
            const node = wrap.firstElementChild;
            if (node) {
                if (check && check.parentNode === dateBlock) {
                    check.insertAdjacentElement('afterend', node);
                } else {
                    dateBlock.appendChild(node);
                }
            }
        }
    }
}

function hiddenCodes(name, raw) {
    return String(raw || '')
        .split(/[;,]/)
        .map((c) => c.trim())
        .filter(Boolean)
        .map((c) => `<input type="hidden" name="${esc(name)}" value="${esc(c)}">`)
        .join('');
}

function renderVorlaufHint(state, f, rules) {
    const st = vorlaufStatusForDraft(f, rules);
    if (st.kind === 'none') return '';
    if (st.kind === 'ok') {
        return `<p class="akt-inline-ok"><i class="bi bi-check-circle-fill" aria-hidden="true"></i> ${esc(st.message)}</p>`;
    }
    return `<div class="akt-inline-warn" role="status">
      <p><strong>${esc(st.message)}</strong></p>
      <details class="akt-inline-help"><summary>Was kann ich tun?</summary>
        <p>Verschiebe den Termin, oder speichere vorerst als Entwurf und kläre eine Ausnahme mit der Direktion.</p>
      </details>
    </div>`;
}

function renderBegleitungChips(state, sd) {
    const codes = String(state.form.begleitungCodes || state.form.begleitung || '')
        .split(/[;,]/)
        .map((c) => c.trim())
        .filter(Boolean);
    const labels = labelMaps(sd).lehrer;
    const teacherOpts = (sd.teachers || []).filter((t) => {
        if (!t) return false;
        const code = String(t.code || t.email || '').trim();
        return code && !codes.includes(code);
    });
    const chips = codes
        .map(
            (c) => `<span class="akt-person-chip">
        <span>${esc(labels[c] || c)}</span>
        <button type="button" class="akt-person-chip__x" data-akt-remove-begleitung="${esc(c)}" aria-label="Entfernen">×</button>
      </span>`
        )
        .join('');
    const liveErr =
        state.stammdatenLive && state.stammdatenLive.error
            ? String(state.stammdatenLive.error)
            : '';
    const emptyTeachersHint =
        !(sd.teachers || []).length
            ? `<p class="akt-field-msg akt-field-msg--warn" role="status">Keine Lehrkräfte geladen${
                  liveErr ? ': ' + esc(liveErr) : ''
              }. Bitte „Klassen neu laden“ – die Liste „Lehrerinnen“ auf der Tool-Site muss für Einreicher lesbar sein (IT: Systemadministration → Nur Berechtigungen).</p>`
            : '';
    const picker = state.begleitungPickerOpen
        ? `<div class="akt-begleitung-picker">
        <label class="akt-sr-only" for="aktBegleitungPick">Begleitperson suchen</label>
        <select id="aktBegleitungPick" data-akt-begleitung-pick>
          <option value="">${teacherOpts.length ? 'Person suchen …' : 'Keine Lehrkräfte verfügbar'}</option>
          ${teacherOpts
              .map((t) => {
                  const code = String(t.code || t.email || '').trim();
                  return `<option value="${esc(code)}">${esc(t.name || code)}</option>`;
              })
              .join('')}
        </select>
      </div>`
        : '';
    return `
      <div class="akt-begleitung">
        <div class="akt-person-chips">${chips || '<span class="muted">Noch keine Begleitung</span>'}</div>
        ${emptyTeachersHint}
        ${picker}
        <button type="button" class="btn akt-btn-ghost" data-akt-toggle-begleitung>
          <i class="bi bi-plus-lg" aria-hidden="true"></i>Begleitperson hinzufügen
        </button>
        ${hiddenCodes('begleitungCodes', codes.join(';'))}
      </div>`;
}

function renderStepper(step) {
    const steps = [
        { n: 1, label: 'Aktivität' },
        { n: 2, label: 'Teilnahme' },
        { n: 3, label: 'Organisation' },
        { n: 4, label: 'Prüfen & Einreichen' }
    ];
    return `<ol class="akt-steps" aria-label="Schritte">
      ${steps
          .map((s) => {
              const done = s.n < step;
              const cur = s.n === step;
              const cls = done ? ' is-done' : cur ? ' is-active' : ' is-future';
              return `<li class="akt-steps__item${cls}">
          <button type="button" class="akt-steps__btn" data-akt-goto-step="${s.n}" aria-current="${cur ? 'step' : 'false'}">
            <span class="akt-steps__n">${done ? '<i class="bi bi-check-lg" aria-hidden="true"></i>' : s.n}</span>
            <span class="akt-steps__label">${esc(s.label)}</span>
          </button>
        </li>`;
          })
          .join('')}
    </ol>`;
}

function stammdatenKlassenEmptyHint(state) {
    const classes = (state.stammdaten && state.stammdaten.classes) || [];
    if (classes.length) return '';
    const live = state.stammdatenLive || {};
    const bits = [];
    const err = String(live.error || '');
    const siteHint = String(live.siteUrl || state.siteUrl || '').trim();
    const is403 = /403|Leserecht|kein Zugriff/i.test(err);
    if (is403) {
        bits.push(err);
        bits.push(
            'Die Listen sind befüllt – dieser Account darf sie aber nicht lesen. IT: unter Listen „Nur Berechtigungen“ erneut mit Einreicher-Gruppe/Person ausführen.'
        );
    } else if (live.source === 'sharepoint' || live.source === 'sharepoint-spo' || live.source === 'local-fallback') {
        bits.push(
            live.classCount === 0
                ? err || 'Die SharePoint-Klassenliste liefert 0 Zeilen für diesen Account.'
                : 'Keine Klassen im Planer geladen.'
        );
    } else if (err) {
        bits.push('Klassen konnten nicht geladen werden: ' + err);
    } else if (!siteHint) {
        bits.push(
            'In diesem Browser fehlt die Schul-Site-URL (liegt oft nur beim IT-PC). IT: Admin → Schule „Tool-Site-URL“ setzen, dann „Klassen neu laden“.'
        );
    } else {
        bits.push(
            'Noch keine Klassen geladen (Anmeldung/Site). Bitte „Klassen neu laden“.'
        );
    }
    if (siteHint) bits.push('Site: ' + siteHint);
    return `<div class="akt-inline-warn" role="status">
      <strong>Keine Klassen zur Auswahl.</strong>
      <p class="muted" style="margin:6px 0 0">${esc(bits.join(' '))}</p>
      <p style="margin:8px 0 0;display:flex;flex-wrap:wrap;gap:8px">
        <button type="button" class="btn btn-sm" data-akt-reload-stammdaten><i class="bi bi-arrow-clockwise"></i>Klassen neu laden</button>
        <a class="btn btn-sm" href="sharepoint-liste-schulaktivitaeten.html"><i class="bi bi-shield-lock"></i>Rechte / Listen (IT)</a>
      </p>
    </div>`;
}

function renderKlassenBlock(state, helpers, f, sd, hints, labels, modus, primaryKlasse, schuelerN) {
    const renderMultiPick = helpers.renderMultiPick;
    const multiKlassen = modus === 'klassen' || modus === 'gruppe' || modus === 'schueler';
    const emptyHint = stammdatenKlassenEmptyHint(state);
    if (multiKlassen) {
        return `<div class="akt-field akt-field--wide">
            <span class="akt-label">${modus === 'gruppe' ? 'Beteiligte Klassen' : 'Klassen'}</span>
            ${emptyHint ||
                renderMultiPick('klasseCodes', sd.classes, f.klasseCodes || f.klasseCode, {
                    placeholder: 'Klasse suchen …',
                    emptyHint: 'Keine Klassen geladen',
                    ariaLabel: 'Klassen auswählen'
                })}
            ${fieldMsg(state, hints, 'klasseCodes')}
          </div>`;
    }
    return `<label class="akt-field">Klasse
            <select name="klasseCodes" class="akt-control akt-control--sm"${
                !(sd.classes || []).length ? ' disabled' : ''
            }>${optionList(sd.classes, primaryKlasse || f.klasseCode, '– wählen –')}</select>
            ${
                primaryKlasse && schuelerN
                    ? `<span class="akt-field-meta">${esc(labels.klasse[primaryKlasse] || primaryKlasse)} · ${schuelerN} Schüler:innen</span>`
                    : ''
            }
            ${emptyHint}
            ${fieldMsg(state, hints, 'klasseCodes')}
          </label>`;
}

/**
 * @param {object} state
 * @param {{ renderMultiPick: Function }} helpers
 */
export function renderAntragWizard(state, helpers) {
    const f = state.form || emptyForm();
    const sd = state.stammdaten || { classes: [], teachers: [], students: [] };
    const rules = state.rules || DEFAULT_RULES;
    const step = Math.min(4, Math.max(1, Number(state.formStep) || 1));
    const modus = String(f.teilnahmeModus || 'klasse').toLowerCase();
    const timeMode = state.timeMode === 'clock' ? 'clock' : 'ue';
    // Nur expliziter Wizard-Schalter – nicht aus Datumsdifferenz ableiten
    // (sonst „Mehrtägig“ + heutiges Enddatum nach Schrittwechsel)
    const multiDay = !!state.multiDay;
    const preview = validateAktivitaet({
        draft: f,
        existing: state.items,
        rules: {
            minVorlaufTage: rules.minVorlaufTage,
            maxGleichzeitigProKlasse: rules.maxGleichzeitigProKlasse,
            vorlaufModus: rules.vorlaufModus,
            genehmigungPfad: rules.genehmigungPfad
        }
    });
    const hints = fieldHintsFromPreview(preview);
    const labels = labelMaps(sd);
    const klassen = klasseCodesOf(f);
    const primaryKlasse = klassen[0] || '';
    const schuelerN = studentCountForKlasse(sd, primaryKlasse);

    const ueDerived = timesFromUnterrichtseinheiten({
        startUe: f.startUe,
        endUe: f.endUe
    });
    const uePreview =
        timeMode === 'ue' && ueDerived.ok && (ueDerived.startZeit || ueDerived.endZeit)
            ? `<p class="akt-time-derived">${esc(
                  formatUeZeitraum({ startUe: f.startUe, endUe: f.endUe }) || ''
              )} · ${esc(ueDerived.startZeit)}–${esc(ueDerived.endZeit)}</p>`
            : '';

    const klassenBlock = renderKlassenBlock(
        state,
        helpers,
        f,
        sd,
        hints,
        labels,
        modus,
        primaryKlasse,
        schuelerN
    );

    const carryBasics = `
      <input type="hidden" name="titel" value="${esc(f.titel)}">
      <input type="hidden" name="typ" value="${esc(f.typ)}">
      <input type="hidden" name="ort" value="${esc(f.ort)}">
      <input type="hidden" name="startdatum" value="${esc(f.startdatum)}">
      <input type="hidden" name="enddatum" value="${esc(f.enddatum || f.startdatum)}">
      <input type="hidden" name="lehrerCode" value="${esc(f.lehrerCode)}">
      <input type="hidden" name="lehrerEmail" value="${esc(f.lehrerEmail || state.accountEmail || '')}">
      <input type="hidden" name="beantragtVon" value="${esc(f.beantragtVon || state.accountEmail || '')}">`;

    const carryTeilnahme = `
      <input type="hidden" name="teilnahmeModus" value="${esc(modus)}">
      ${klassen.map((c) => `<input type="hidden" name="klasseCodes" value="${esc(c)}">`).join('')}
      <input type="hidden" name="gruppenText" value="${esc(f.gruppenText)}">
      <input type="hidden" name="teilnehmendeText" value="${esc(f.teilnehmendeText)}">
      <input type="hidden" name="nichtTeilnehmende" value="${esc(f.nichtTeilnehmende)}">`;

    const carryOrg = `
      <input type="hidden" name="startUe" value="${esc(f.startUe)}">
      <input type="hidden" name="endUe" value="${esc(f.endUe)}">
      <input type="hidden" name="startZeit" value="${esc(f.startZeit)}">
      <input type="hidden" name="endZeit" value="${esc(f.endZeit)}">
      ${hiddenCodes('begleitungCodes', f.begleitungCodes || f.begleitung)}
      <input type="hidden" name="externeVortragende" value="${esc(f.externeVortragende)}">
      <input type="hidden" name="verkehrsmittel" value="${esc(f.verkehrsmittel)}">
      <input type="hidden" name="ablauf" value="${esc(f.ablauf)}">
      <input type="hidden" name="kostenHinweis" value="${esc(f.kostenHinweis)}">
      <input type="hidden" name="notiz" value="${esc(f.notiz)}">
      <input type="hidden" name="unterrichtsbezug" value="${esc(f.unterrichtsbezug)}">`;

    let stepBody = '';
    if (step === 1) {
        stepBody = `
        <header class="akt-wizard-head">
          <h3>Was planst du?</h3>
          <p>Grunddaten der Aktivität – kurz und klar.</p>
        </header>
        <div class="akt-form-grid akt-form-grid--focus">
          <label class="akt-field akt-field--wide">Titel
            <input name="titel" class="akt-control" required value="${esc(f.titel)}" placeholder="z. B. Museumsbesuch Wien" data-akt-touch="titel">
            ${fieldMsg(state, hints, 'titel')}
          </label>
          <label class="akt-field">Art
            <select name="typ" class="akt-control akt-control--sm" data-akt-touch="typ">${optionList(typChoices(), f.typ, '–')}</select>
            ${fieldMsg(state, hints, 'typ')}
          </label>
          <label class="akt-field">Ort / Ziel
            <input name="ort" class="akt-control" value="${esc(f.ort)}" placeholder="Ort oder Zieladresse" data-akt-touch="ort">
            ${fieldMsg(state, hints, 'ort')}
          </label>
          <div class="akt-field akt-field--wide akt-date-block">
            <span class="akt-label">${multiDay ? 'Zeitraum' : 'Datum'}</span>
            ${
                multiDay
                    ? `<div class="akt-date-pair">
                <label class="akt-field">Startdatum
                  <input type="date" name="startdatum" class="akt-control akt-control--sm" required value="${esc(f.startdatum)}" data-akt-touch="startdatum">
                </label>
                <label class="akt-field">Enddatum
                  <input type="date" name="enddatum" class="akt-control akt-control--sm" value="${esc(f.enddatum || f.startdatum)}">
                </label>
              </div>`
                    : `<label class="akt-field">
                <span class="akt-sr-only">Datum</span>
                <input type="date" name="startdatum" class="akt-control akt-control--sm" required value="${esc(f.startdatum)}" data-akt-touch="startdatum">
                <input type="hidden" name="enddatum" value="${esc(f.startdatum)}">
              </label>`
            }
            <label class="akt-check">
              <input type="checkbox" data-akt-multiday ${multiDay ? 'checked' : ''}>
              Mehrtägige Aktivität
            </label>
            ${renderVorlaufHint(state, f, rules)}
            ${fieldMsg(state, hints, 'startdatum')}
          </div>
          <input type="hidden" name="lehrerCode" value="${esc(f.lehrerCode)}">
          <input type="hidden" name="lehrerEmail" value="${esc(f.lehrerEmail || state.accountEmail || '')}">
          <input type="hidden" name="beantragtVon" value="${esc(f.beantragtVon || state.accountEmail || '')}">
          <input type="hidden" name="teilnahmeModus" value="${esc(modus)}">
          ${klassen.map((c) => `<input type="hidden" name="klasseCodes" value="${esc(c)}">`).join('')}
          <input type="hidden" name="gruppenText" value="${esc(f.gruppenText)}">
          <input type="hidden" name="teilnehmendeText" value="${esc(f.teilnehmendeText)}">
          <input type="hidden" name="nichtTeilnehmende" value="${esc(f.nichtTeilnehmende)}">
          ${carryOrg}
        </div>`;
    } else if (step === 2) {
        const showNichtTeilnehmende = modus === 'klasse' || modus === 'klassen' || modus === 'gruppe';
        stepBody = `
        <header class="akt-wizard-head">
          <h3>Wer nimmt teil?</h3>
          <p>Klasse, Gruppe oder einzelne Schüler:innen.</p>
        </header>
        <div class="akt-form-grid akt-form-grid--focus">
          ${carryBasics}
          <label class="akt-field akt-field--wide">Teilnahme
            <select name="teilnahmeModus" class="akt-control akt-control--md">${optionList(teilnahmeModusChoices(), modus, '–')}</select>
          </label>
          ${klassenBlock}
          ${
              modus === 'gruppe'
                  ? `<label class="akt-field akt-field--wide">Unterrichtsgruppe
                      <input name="gruppenText" class="akt-control" value="${esc(f.gruppenText)}" placeholder="z. B. kath. Religion 3AK/3BK" data-akt-touch="gruppenText">
                      ${fieldMsg(state, hints, 'gruppenText')}
                    </label>`
                  : `<input type="hidden" name="gruppenText" value="${esc(f.gruppenText)}">`
          }
          ${
              modus === 'schueler'
                  ? `<label class="akt-field akt-field--wide">Teilnehmende Schüler:innen
                      <textarea name="teilnehmendeText" class="akt-control" rows="2" placeholder="Eine Person pro Zeile" data-akt-touch="teilnehmendeText">${esc(f.teilnehmendeText)}</textarea>
                      ${fieldMsg(state, hints, 'teilnehmendeText')}
                    </label>`
                  : `<input type="hidden" name="teilnehmendeText" value="${esc(f.teilnehmendeText)}">`
          }
          ${
              showNichtTeilnehmende
                  ? `<label class="akt-field akt-field--wide">Wer nimmt nicht teil?
                      <textarea name="nichtTeilnehmende" class="akt-control" rows="2" placeholder="Optional – Namen der nicht teilnehmenden Schüler:innen">${esc(f.nichtTeilnehmende)}</textarea>
                    </label>`
                  : `<input type="hidden" name="nichtTeilnehmende" value="${esc(f.nichtTeilnehmende)}">`
          }
          ${carryOrg}
        </div>`;
    } else if (step === 3) {
        stepBody = `
        <header class="akt-wizard-head">
          <h3>Wie ist die Aktivität organisiert?</h3>
          <p>Zeit, Begleitung und Organisation.</p>
        </header>
        <div class="akt-form-grid akt-form-grid--focus">
          ${carryBasics}
          ${carryTeilnahme}
          <div class="akt-field akt-field--wide akt-section-label"><span class="akt-label">Zeit</span></div>
          <div class="akt-field akt-field--wide">
            <span class="akt-label akt-label--sub">Zeitmodell</span>
            <div class="akt-seg" role="group" aria-label="Zeitmodell">
              <button type="button" class="akt-seg__btn${timeMode === 'ue' ? ' is-active' : ''}" data-akt-time-mode="ue">Unterrichtseinheiten</button>
              <button type="button" class="akt-seg__btn${timeMode === 'clock' ? ' is-active' : ''}" data-akt-time-mode="clock">Uhrzeit</button>
            </div>
          </div>
          ${
              timeMode === 'ue'
                  ? `<label class="akt-field">Von UE
                      <select name="startUe" class="akt-control akt-control--sm" data-akt-touch="startUe">${optionList(ueChoices(), f.startUe, '–')}</select>
                      ${fieldMsg(state, hints, 'startUe')}
                    </label>
                    <label class="akt-field">Bis UE
                      <select name="endUe" class="akt-control akt-control--sm">${optionList(ueChoices(), f.endUe, '–')}</select>
                    </label>
                    ${uePreview}
                    <input type="hidden" name="startZeit" value="${esc(ueDerived.startZeit || f.startZeit)}">
                    <input type="hidden" name="endZeit" value="${esc(ueDerived.endZeit || f.endZeit)}">`
                  : `<label class="akt-field">Beginn
                      <input name="startZeit" class="akt-control akt-control--sm" placeholder="07:45" value="${esc(f.startZeit)}" data-akt-touch="startUe">
                    </label>
                    <label class="akt-field">Ende
                      <input name="endZeit" class="akt-control akt-control--sm" placeholder="13:30" value="${esc(f.endZeit)}">
                    </label>
                    <input type="hidden" name="startUe" value="">
                    <input type="hidden" name="endUe" value="">`
          }
          <div class="akt-field akt-field--wide akt-section-label"><span class="akt-label">Begleitung</span></div>
          <div class="akt-field akt-field--wide">
            <span class="akt-label akt-label--sub">Wer begleitet die Klasse?</span>
            ${renderBegleitungChips(state, sd)}
          </div>
          <div class="akt-field akt-field--wide akt-section-label"><span class="akt-label">Organisation</span></div>
          <label class="akt-field">Verkehrsmittel
            <input name="verkehrsmittel" class="akt-control akt-control--md" value="${esc(f.verkehrsmittel)}">
          </label>
          <label class="akt-field">Kosten
            <input name="kostenHinweis" class="akt-control akt-control--md" value="${esc(
                f.kostenHinweis
            )}" placeholder="z. B. 12 € pro Schüler:in" data-akt-kosten-live>
            ${renderKostenAmpelHint(f, state.rules)}
          </label>
          <label class="akt-field akt-field--wide">Ablauf
            <textarea name="ablauf" class="akt-control" rows="2" placeholder="Kurzer Ablauf">${esc(f.ablauf)}</textarea>
          </label>
          <details class="akt-more${state.weitereAngabenOpen ? ' is-open' : ''}" ${state.weitereAngabenOpen ? 'open' : ''} data-akt-weitere>
            <summary>Weitere Angaben</summary>
            <div class="akt-form-grid akt-form-grid--focus" style="margin-top:10px">
              <label class="akt-field akt-field--wide">Externe Vortragende
                <input name="externeVortragende" class="akt-control" value="${esc(f.externeVortragende)}" placeholder="Name / Organisation (optional)">
              </label>
              <label class="akt-field akt-field--wide">Unterrichtsbezug
                <input name="unterrichtsbezug" class="akt-control" value="${esc(f.unterrichtsbezug)}" placeholder="Fach / Lehrplanbezug">
              </label>
              <label class="akt-field akt-field--wide">Hinweise / Begründung
                <textarea name="notiz" class="akt-control" rows="2" placeholder="Optional">${esc(f.notiz)}</textarea>
              </label>
            </div>
          </details>
          ${
              !state.weitereAngabenOpen
                  ? `<input type="hidden" name="externeVortragende" value="${esc(f.externeVortragende)}">
                     <input type="hidden" name="unterrichtsbezug" value="${esc(f.unterrichtsbezug)}">
                     <input type="hidden" name="notiz" value="${esc(f.notiz)}">`
                  : ''
          }
        </div>`;
    } else {
        const issues = (preview.errors || []).map((msg) => {
            const mapped = mapValidationIssueToStep(msg);
            return { msg, ...mapped };
        });
        const begleitungCodes = String(f.begleitungCodes || f.begleitung || '')
            .split(/[;,]/)
            .map((c) => c.trim())
            .filter(Boolean);
        const begleitungLabel = begleitungCodes.map((c) => labels.lehrer[c] || c).join(', ');
        const zeitLine =
            formatUeZeitraum(f) ||
            [f.startZeit, f.endZeit].filter(Boolean).join('–') ||
            '';
        const orgLines = [];
        if (zeitLine) orgLines.push(`<p>${esc(zeitLine)}</p>`);
        if (begleitungLabel) orgLines.push(`<p><strong>Begleitung</strong><br>${esc(begleitungLabel)}</p>`);
        if (f.verkehrsmittel) orgLines.push(`<p><strong>Verkehrsmittel</strong><br>${esc(f.verkehrsmittel)}</p>`);
        if (f.kostenHinweis) {
            const ampel = renderKostenAmpelHint(f, state.rules);
            orgLines.push(
                `<p><strong>Kosten</strong><br>${esc(f.kostenHinweis)}</p>${ampel}`
            );
        }
        if (f.ablauf) orgLines.push(`<p><strong>Ablauf</strong><br>${esc(f.ablauf)}</p>`);
        if (f.externeVortragende)
            orgLines.push(`<p><strong>Externe</strong><br>${esc(f.externeVortragende)}</p>`);
        if (f.unterrichtsbezug)
            orgLines.push(`<p><strong>Unterrichtsbezug</strong><br>${esc(f.unterrichtsbezug)}</p>`);
        if (f.notiz) orgLines.push(`<p><strong>Hinweise</strong><br>${esc(f.notiz)}</p>`);
        stepBody = `
        <header class="akt-wizard-head">
          <h3>Prüfe deine Angaben</h3>
          <p>Alles fertig? Dann kannst du einreichen.</p>
        </header>
        ${
            issues.length
                ? `<div class="akt-todo-list" role="alert">
            <p><strong>Noch ${issues.length} ${issues.length === 1 ? 'Ding' : 'Dinge'} zu erledigen</strong></p>
            <ul>
              ${issues
                  .map(
                      (it) => `<li>
                  <span>${esc(it.msg)}</span>
                  <button type="button" class="akt-text-link" data-akt-goto-step="${it.step}">${esc(it.action)}</button>
                </li>`
                  )
                  .join('')}
            </ul>
          </div>`
                : `<div class="akt-review__ok" role="status">
            <i class="bi bi-check-circle-fill" aria-hidden="true"></i>
            <div>
              <strong>Bereit zum Einreichen</strong>
              <p class="akt-review__ok-sub">Alle erforderlichen Angaben sind vollständig.</p>
            </div>
          </div>`
        }
        <div class="akt-summary">
          <section class="akt-summary__block">
            <header>
              <h4>Aktivität</h4>
              <button type="button" class="akt-text-link" data-akt-goto-step="1">Bearbeiten</button>
            </header>
            <p class="akt-summary__title">${esc(f.titel || '–')}</p>
            <p>${esc(typLabel(f.typ))} · ${esc(formatReadableDateRange(f.startdatum, f.enddatum))}</p>
            <p class="muted">${esc(f.ort || 'Ort noch offen')}</p>
          </section>
          <section class="akt-summary__block">
            <header>
              <h4>Teilnahme</h4>
              <button type="button" class="akt-text-link" data-akt-goto-step="2">Bearbeiten</button>
            </header>
            <p>${esc(teilnahmeKurz(f, labels.klasse))}${
                primaryKlasse && schuelerN ? ` · ${schuelerN} Schüler:innen` : ''
            }</p>
            <p class="muted">${esc(teilnahmeModusLabel(modus))}</p>
            ${
                f.nichtTeilnehmende
                    ? `<p class="muted">Nehmen nicht teil: ${esc(f.nichtTeilnehmende)}</p>`
                    : ''
            }
          </section>
          <section class="akt-summary__block">
            <header>
              <h4>Organisation</h4>
              <button type="button" class="akt-text-link" data-akt-goto-step="3">Bearbeiten</button>
            </header>
            ${orgLines.join('') || '<p class="muted">Noch keine Organisationsangaben</p>'}
          </section>
        </div>
        ${carryBasics}${carryTeilnahme}${carryOrg}`;
    }

    const resubmitHint =
        String(f.status || '').toLowerCase() === 'nachforderung'
            ? '<div class="akt-alert akt-alert--warn"><strong>Ergänzung erforderlich.</strong> Bitte die Angaben vervollständigen und erneut einreichen.</div>'
            : '';

    const canSubmit = preview.ok;
    const missingN = (preview.errors || []).length;

    return `
    <section class="akt-panel akt-panel--wizard">
      ${f.rueckfrageText ? `<div class="akt-alert akt-alert--warn"><strong>Rückfrage:</strong> ${esc(f.rueckfrageText)}</div>` : ''}
      ${resubmitHint}
      ${renderStepper(step)}
      <form id="aktForm" class="akt-form akt-form--wizard" autocomplete="off" data-akt-step="${step}">
        ${stepBody}
        <div class="akt-wizard-actions">
          ${
              step > 1
                  ? '<button type="button" class="btn akt-btn-tertiary" data-akt-step-prev>Zurück</button>'
                  : '<span></span>'
          }
          <div class="akt-wizard-actions__primary">
            <button type="button" class="btn akt-btn-secondary" data-akt-save-draft>Als Entwurf speichern</button>
            ${
                step < 4
                    ? '<button type="button" class="btn akt-btn-primary" data-akt-step-next>Weiter →</button>'
                    : `<div class="akt-submit-wrap">
                        <button type="submit" class="btn akt-btn-primary" ${canSubmit ? '' : 'disabled'} data-akt-submit>
                          ${
                              String(f.status || '').toLowerCase() === 'nachforderung'
                                  ? 'Erneut einreichen'
                                  : 'Zur Genehmigung einreichen'
                          }
                        </button>
                        ${
                            canSubmit
                                ? ''
                                : `<p class="akt-submit-hint">Noch ${missingN} ${
                                      missingN === 1 ? 'Angabe' : 'Angaben'
                                  } erforderlich.</p>`
                        }
                      </div>`
            }
          </div>
        </div>
      </form>
    </section>`;
}
