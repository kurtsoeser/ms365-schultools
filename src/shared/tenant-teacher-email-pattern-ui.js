/**
 * Stammdaten-Tab: schulweites E-Mail-Muster (Presets + Baustein-Builder).
 */
import {
    EMAIL_PATTERN_PRESETS,
    defaultEmailBuilderPattern,
    emailPatternTokenLabel,
    isBuilderPatternId,
    localPartFromEmailPatternId,
    normalizeEmailBuilderPattern,
    parseEmailPatternId,
    serializeEmailBuilderPattern
} from './person-email-pattern-logic.js';

const STORAGE_KEY = 'ms365-tenant-school-email-pattern-v1';
const STORAGE_KEY_LEGACY_TEACHERS = 'ms365-tenant-teachers-email-pattern-v1';

function normStr(v) {
    return String(v ?? '').trim();
}

function getDomainSample() {
    if (typeof window.ms365GetSchoolDomainNoAt === 'function') {
        return normStr(window.ms365GetSchoolDomainNoAt()).replace(/^@+/, '');
    }
    const dom = document.getElementById('schoolEmailDomain');
    return normStr(dom && dom.value) || 'schule.de';
}

/**
 * @param {{
 *   cardId?: string,
 *   patternInputId?: string,
 *   givenNamesSelectId?: string,
 *   builderZoneId?: string,
 *   previewId?: string,
 *   modeHintId?: string,
 *   customPanelId?: string
 * }} opts
 */
export function mountTenantSchoolEmailPatternUi(opts) {
    const o = opts || {};
    const card =
        document.getElementById(o.cardId || 'tenantSchoolEmailPatternCard') ||
        document.getElementById('tenantTeachersEmailPatternCard');
    const patternInput = document.getElementById(o.patternInputId || 'tenantSchoolEmailPattern');
    const givenSel = document.getElementById(o.givenNamesSelectId || 'tenantSchoolEmailGivenNames');
    const zone = document.getElementById(o.builderZoneId || 'tenantSchoolEmailPatternBuilder');
    const previewEl = document.getElementById(o.previewId || 'tenantSchoolEmailPatternPreview');
    const modeHint = document.getElementById(o.modeHintId || 'tenantSchoolEmailPatternModeHint');
    const customPanel = document.getElementById(o.customPanelId || 'tenantSchoolEmailPatternCustom');
    if (!patternInput || !zone || !card) return;

    let builderTokens = defaultEmailBuilderPattern();

    function firstNameMode() {
        const raw = givenSel ? String(givenSel.value || 'first') : 'first';
        const api = window.ms365PersonEmailFromName;
        if (api && typeof api.resolveFirstNameMode === 'function') {
            return api.resolveFirstNameMode(raw);
        }
        return raw === 'all' ? 'all' : 'first';
    }

    function setPatternId(id) {
        patternInput.value = String(id || 'vorname.nachname');
        try {
            localStorage.setItem(STORAGE_KEY, patternInput.value);
        } catch {
            /* ignore */
        }
        syncUiFromPatternId();
        updatePreview();
    }

    function getPatternId() {
        return String(patternInput.value || 'vorname.nachname');
    }

    function readBuilderFromDom() {
        const tokens = [];
        zone.querySelectorAll('[data-email-token-type]').forEach(function (el) {
            const type = String(el.getAttribute('data-email-token-type') || '');
            if (type === 'sep') {
                tokens.push({ type: 'sep', value: String(el.getAttribute('data-email-token-value') || '') });
            } else {
                tokens.push({ type });
            }
        });
        return normalizeEmailBuilderPattern(tokens);
    }

    function addChip(token) {
        const chip = document.createElement('span');
        chip.className = 'ts-email-pattern-chip';
        chip.setAttribute('data-email-token-type', token.type);
        if (token.type === 'sep') chip.setAttribute('data-email-token-value', String(token.value ?? ''));
        chip.draggable = true;
        const label = document.createElement('span');
        label.textContent = emailPatternTokenLabel(token);
        const x = document.createElement('button');
        x.type = 'button';
        x.className = 'ts-email-pattern-chip__remove';
        x.setAttribute('aria-label', 'Baustein entfernen');
        x.innerHTML = '<i class="bi bi-x" aria-hidden="true"></i>';
        x.addEventListener('click', function () {
            chip.remove();
            builderTokens = readBuilderFromDom();
            setPatternId(serializeEmailBuilderPattern(builderTokens));
        });
        chip.appendChild(label);
        chip.appendChild(x);
        zone.appendChild(chip);
    }

    function renderBuilder(tokens) {
        zone.replaceChildren();
        (tokens || []).forEach(addChip);
        builderTokens = normalizeEmailBuilderPattern(tokens);
    }

    function syncUiFromPatternId() {
        const id = getPatternId();
        const parsed = parseEmailPatternId(id);
        card.querySelectorAll('[data-email-preset]').forEach(function (btn) {
            const p = btn.getAttribute('data-email-preset');
            btn.classList.toggle('is-active', parsed.kind === 'preset' && parsed.id === p);
        });
        const customBtn = card.querySelector('[data-email-preset="custom"]');
        if (customBtn) customBtn.classList.toggle('is-active', parsed.kind === 'build');
        if (parsed.kind === 'build') {
            renderBuilder(parsed.tokens);
        } else if (modeHint) {
            modeHint.textContent =
                'Preset: ' + (EMAIL_PATTERN_PRESETS.find(function (x) { return x.id === parsed.id; })?.label || parsed.id);
        }
        if (customPanel) customPanel.hidden = parsed.kind !== 'build';
    }

    function personFromSemicolonLine(line, codeFirst) {
        const parts = normStr(line).split(';').map(function (p) {
            return normStr(p);
        });
        if (!parts.length) return null;
        let kuerzel = '';
        let name = '';
        if (codeFirst) {
            kuerzel = parts[0] || '';
            name = parts[1] || '';
        } else {
            name = parts[1] || parts[0] || '';
        }
        const nameBits = name.split(/\s+/).filter(Boolean);
        return {
            kuerzel: kuerzel || 'MU',
            vorname: nameBits[0] || 'Max',
            nachname: nameBits.slice(1).join(' ') || 'Mustermann'
        };
    }

    function samplePerson() {
        const sources = [
            { id: 'tenantTeachersLines', codeFirst: true },
            { id: 'tenantStudentsLines', codeFirst: false },
            { id: 'tenantAdminBundleLines', codeFirst: true },
            { id: 'tenantSgaLines', codeFirst: false }
        ];
        for (let i = 0; i < sources.length; i++) {
            const ta = document.getElementById(sources[i].id);
            if (!ta || !normStr(ta.value)) continue;
            const line = normStr(ta.value).split(/\r\n|\n|\r/).find(Boolean);
            const person = personFromSemicolonLine(line, sources[i].codeFirst);
            if (person) return person;
        }
        return { kuerzel: 'MU', vorname: 'Max', nachname: 'Mustermann' };
    }

    function updatePreview() {
        if (!previewEl) return;
        const person = samplePerson();
        const local = localPartFromEmailPatternId(getPatternId(), {
            vorname: person.vorname,
            nachname: person.nachname,
            kuerzel: person.kuerzel,
            firstNameMode: firstNameMode()
        });
        const dom = getDomainSample();
        previewEl.textContent = local ? local + '@' + dom : '— (Muster unvollständig)';
    }

    function wireDnD() {
        let dragEl = null;
        zone.addEventListener('dragstart', function (e) {
            const t = e.target && e.target.closest ? e.target.closest('.ts-email-pattern-chip') : null;
            if (!t) return;
            dragEl = t;
            t.classList.add('is-dragging');
            e.dataTransfer.effectAllowed = 'move';
        });
        zone.addEventListener('dragend', function () {
            if (dragEl) dragEl.classList.remove('is-dragging');
            dragEl = null;
            builderTokens = readBuilderFromDom();
            setPatternId(serializeEmailBuilderPattern(builderTokens));
        });
        zone.addEventListener('dragover', function (e) {
            e.preventDefault();
            const over = e.target && e.target.closest ? e.target.closest('.ts-email-pattern-chip') : null;
            if (!dragEl || !over || over === dragEl) return;
            const rect = over.getBoundingClientRect();
            if (e.clientX > rect.left + rect.width / 2) over.after(dragEl);
            else over.before(dragEl);
        });
    }

    try {
        let stored = localStorage.getItem(STORAGE_KEY);
        if (!stored) stored = localStorage.getItem(STORAGE_KEY_LEGACY_TEACHERS);
        if (stored) patternInput.value = stored;
    } catch {
        /* ignore */
    }

    EMAIL_PATTERN_PRESETS.forEach(function (p) {
        const btn = card.querySelector('[data-email-preset="' + p.id + '"]');
        if (!btn || btn.dataset.bound) return;
        btn.dataset.bound = '1';
        btn.addEventListener('click', function () {
            setPatternId(p.id);
        });
    });

    const customBtn = card.querySelector('[data-email-preset="custom"]');
    if (customBtn && !customBtn.dataset.bound) {
        customBtn.dataset.bound = '1';
        customBtn.addEventListener('click', function () {
            if (!isBuilderPatternId(getPatternId())) {
                setPatternId(serializeEmailBuilderPattern(builderTokens));
            } else {
                syncUiFromPatternId();
            }
        });
    }

    card.querySelectorAll('[data-email-add-token]').forEach(function (btn) {
        if (btn.dataset.bound) return;
        btn.dataset.bound = '1';
        btn.addEventListener('click', function () {
            const type = btn.getAttribute('data-email-add-token');
            const val = btn.getAttribute('data-email-sep-value');
            let token;
            if (type === 'sep') token = { type: 'sep', value: val != null ? val : '.' };
            else token = { type: type };
            builderTokens = readBuilderFromDom();
            builderTokens.push(token);
            renderBuilder(builderTokens);
            setPatternId(serializeEmailBuilderPattern(builderTokens));
        });
    });

    if (givenSel) givenSel.addEventListener('change', updatePreview);
    ['tenantTeachersLines', 'tenantStudentsLines', 'tenantAdminBundleLines', 'tenantSgaLines'].forEach(function (id) {
        const ta = document.getElementById(id);
        if (ta) ta.addEventListener('input', updatePreview);
    });
    const domInput = document.getElementById('schoolEmailDomain');
    if (domInput) domInput.addEventListener('input', updatePreview);
    wireDnD();
    syncUiFromPatternId();
    if (!isBuilderPatternId(getPatternId())) {
        renderBuilder(defaultEmailBuilderPattern());
    }
    updatePreview();

    return { setPatternId, getPatternId, updatePreview };
}

/** @deprecated Alias – bitte mountTenantSchoolEmailPatternUi verwenden */
export function mountTenantTeacherEmailPatternUi(opts) {
    return mountTenantSchoolEmailPatternUi(opts);
}
