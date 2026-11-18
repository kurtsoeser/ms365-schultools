import vorgaben from './einschreiber-vorgaben.json';
import { createPortalApi } from './stammdaten-portal-api.js';
import { loadPublicPortalConfig } from './stammdaten-portal-config.js';
import { SECTION_LABELS, visibleFields } from './stammdaten-portal-fields.js';
import {
    createEmptyPayload,
    isParentEditLocked,
    normalizePayload,
    readFieldValue,
    validatePayloadAgainstProfile
} from './stammdaten-portal-logic.js';

/**
 * @param {HTMLElement} host
 * @param {{ enrollmentId?: string }} [opts]
 */
export async function mountParentForm(host, opts = {}) {
    if (!host) return;
    const params = new URLSearchParams(window.location.search);
    const enrollmentId = String(opts.enrollmentId || params.get('id') || '').trim();
    const config = await loadPublicPortalConfig();
    const api = createPortalApi(config);

    let state = {
        phase: enrollmentId ? 'lock' : 'missing',
        token: '',
        record: null,
        profile: config.profile,
        payload: createEmptyPayload(),
        message: '',
        error: ''
    };

    host.innerHTML = '';
    const root = document.createElement('div');
    root.className = 'ssp-page';
    host.appendChild(root);
    render();

    function render() {
        if (state.phase === 'missing') {
            root.innerHTML = `
              <div class="ssp-card">
                <h2>Link ungültig</h2>
                <p class="ssp-muted">Bitte öffnen Sie den Link aus dem Schreiben der Schule (Parameter <code>id</code> fehlt).</p>
              </div>`;
            return;
        }
        if (state.phase === 'lock') {
            root.innerHTML = `
              <div class="ssp-card ssp-lock">
                <h2>${escapeHtml(config.schoolName || 'Stammdaten')}</h2>
                <p class="ssp-muted">Bitte geben Sie den PIN aus dem Schreiben der Schule ein. Der PIN steht <strong>nicht</strong> in der Adresszeile.</p>
                <div class="ssp-field">
                  <label for="sspPin">Zugangscode (PIN)</label>
                  <input id="sspPin" type="password" inputmode="numeric" autocomplete="one-time-code" maxlength="8" />
                </div>
                <div class="ssp-actions">
                  <button type="button" class="btn" id="sspUnlockBtn">Öffnen</button>
                </div>
                ${state.error ? `<p class="ssp-error" role="alert">${escapeHtml(state.error)}</p>` : ''}
              </div>`;
            root.querySelector('#sspUnlockBtn')?.addEventListener('click', onUnlock);
            root.querySelector('#sspPin')?.addEventListener('keydown', (e) => {
                if (e.key === 'Enter') onUnlock();
            });
            return;
        }

        const fields = visibleFields(state.profile || config.profile);
        const sections = groupBySection(fields);
        const locked = isParentEditLocked(state.record?.status);
        root.innerHTML = `
          <div class="ssp-banner">Sensible Daten werden verschlüsselt über HTTPS an die Schule übertragen. Bitte Gerät nicht unbeaufsichtigt lassen.</div>
          <div class="ssp-card">
            <h2>${escapeHtml(config.schoolName || 'Schüler-Stammdaten')}</h2>
            <p class="ssp-muted">Status: <span class="ssp-badge">${escapeHtml(state.record?.status || '')}</span>
              · ${
                  locked
                      ? 'Vom Sekretariat geprüft – nur noch Ansicht.'
                      : 'Bitte Angaben prüfen und speichern.'
              }</p>
          </div>
          ${
              locked
                  ? `<div class="ssp-card ssp-locked-note" role="status">
                <p class="ssp-ok" style="margin:0">Vielen Dank. Ihre Angaben wurden von der Schule geprüft und sind abgeschlossen. Bei Korrekturen bitte die Schule kontaktieren.</p>
              </div>`
                  : ''
          }
          <div class="ssp-check">
            <input type="checkbox" id="sspSameAddr" ${state.payload.sameAddressAsStudent ? 'checked' : ''} ${locked ? 'disabled' : ''} />
            <label for="sspSameAddr">Erziehungsberechtigte haben dieselbe Wohnadresse wie das Kind</label>
          </div>
          ${sections
              .map(
                  ([section, list]) => `
            <div class="ssp-card" data-section="${escapeHtml(section)}">
              <h3>${escapeHtml(SECTION_LABELS[section] || section)}</h3>
              <div class="ssp-grid">
                ${list.map((f) => renderField(f, locked)).join('')}
              </div>
            </div>`
              )
              .join('')}
          ${
              locked
                  ? ''
                  : `<div class="ssp-actions">
            <button type="button" class="btn" id="sspSaveDraft">Zwischenspeichern</button>
            <button type="button" class="btn btn-primary" id="sspSubmit">Absenden</button>
          </div>`
          }
          ${state.error ? `<p class="ssp-error" role="alert">${escapeHtml(state.error)}</p>` : ''}
          ${state.message ? `<p class="ssp-ok" role="status">${escapeHtml(state.message)}</p>` : ''}
          ${config.privacyUrl ? `<p class="ssp-muted"><a href="${escapeHtml(config.privacyUrl)}" target="_blank" rel="noopener">Datenschutzhinweise</a></p>` : ''}`;

        if (!locked) {
            root.querySelector('#sspSameAddr')?.addEventListener('change', (e) => {
                state.payload.sameAddressAsStudent = !!e.target.checked;
                state.payload = normalizePayload(state.payload);
                render();
            });
            for (const field of fields) {
                const el = root.querySelector(`[data-field="${cssEscape(field.id)}"]`);
                el?.addEventListener('change', () => writeField(field.id, el.value));
                el?.addEventListener('input', () => writeField(field.id, el.value));
            }
            root.querySelector('#sspSaveDraft')?.addEventListener('click', () => onSave(false));
            root.querySelector('#sspSubmit')?.addEventListener('click', () => onSave(true));
        }
    }

    function renderField(field, locked = false) {
        const value = readFieldValue(state.payload, field.id);
        const req = field.required ? ' *' : '';
        const dis = locked ? 'disabled' : '';
        if (field.type === 'select') {
            const opts = vorgaben[field.optionsKey] || [];
            return `<div class="ssp-field">
              <label>${escapeHtml(field.label)}${req}</label>
              <select data-field="${escapeHtml(field.id)}" ${dis}>
                <option value="">— bitte wählen —</option>
                ${opts
                    .map(
                        (o) =>
                            `<option value="${escapeHtml(o)}" ${String(value) === String(o) ? 'selected' : ''}>${escapeHtml(o)}</option>`
                    )
                    .join('')}
              </select>
            </div>`;
        }
        const inputType = field.type === 'date' ? 'date' : field.type === 'email' ? 'email' : field.type === 'tel' ? 'tel' : 'text';
        return `<div class="ssp-field">
          <label>${escapeHtml(field.label)}${req}</label>
          <input data-field="${escapeHtml(field.id)}" type="${inputType}" value="${escapeHtml(formatInputValue(field, value))}" ${dis} />
        </div>`;
    }

    function writeField(path, value) {
        const parts = String(path).split('.');
        let cur = state.payload;
        for (let i = 0; i < parts.length - 1; i += 1) {
            const p = parts[i];
            const idx = /^\d+$/.test(p) ? Number(p) : null;
            if (idx != null) {
                if (!cur[idx]) cur[idx] = {};
                cur = cur[idx];
            } else {
                if (!cur[p]) cur[p] = {};
                cur = cur[p];
            }
        }
        const last = parts[parts.length - 1];
        cur[last] = value;
        state.payload = normalizePayload(state.payload);
    }

    async function onUnlock() {
        const pin = String(root.querySelector('#sspPin')?.value || '').trim();
        state.error = '';
        const res = await api.unlock(enrollmentId, pin);
        if (!res.ok) {
            state.error = res.message || 'Ungültiger Zugang.';
            render();
            return;
        }
        state.token = res.token;
        state.record = res.record;
        state.profile = res.profile || config.profile;
        state.payload = normalizePayload(res.record?.payload);
        state.phase = 'form';
        render();
    }

    async function onSave(submit) {
        state.error = '';
        state.message = '';
        if (isParentEditLocked(state.record?.status)) {
            state.error =
                'Dieser Datensatz wurde vom Sekretariat geprüft und kann nicht mehr geändert werden.';
            render();
            return;
        }
        const check = validatePayloadAgainstProfile(state.payload, state.profile, {
            mode: submit ? 'submit' : 'draft'
        });
        state.payload = check.payload;
        if (!check.ok) {
            state.error = check.errors.map((e) => e.message).join(' ');
            render();
            return;
        }
        const res = await api.save(state.token, state.payload, { submit });
        if (!res.ok) {
            state.error = res.message || 'Speichern fehlgeschlagen.';
            render();
            return;
        }
        state.record = res.record;
        state.message = submit
            ? 'Daten wurden abgesendet. Vielen Dank!'
            : 'Zwischenstand gespeichert. Sie können den Link später erneut öffnen.';
        render();
    }
}

function groupBySection(fields) {
    const map = new Map();
    for (const f of fields) {
        if (!map.has(f.section)) map.set(f.section, []);
        map.get(f.section).push(f);
    }
    return [...map.entries()];
}

function formatInputValue(field, value) {
    if (field.type === 'date' && value) {
        const s = String(value);
        if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0, 10);
        const m = s.match(/^(\d{2})\.(\d{2})\.(\d{4})$/);
        if (m) return `${m[3]}-${m[2]}-${m[1]}`;
    }
    return value == null ? '' : String(value);
}

function escapeHtml(value) {
    return String(value ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function cssEscape(value) {
    return String(value).replace(/"/g, '\\"');
}
