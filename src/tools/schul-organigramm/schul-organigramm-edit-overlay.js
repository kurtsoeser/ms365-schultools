/**
 * Overlay close/esc + inline Entra pickers + remember role parent.
 */
import { normStr } from './schul-organigramm-edit-utils.js';
import { mountInlineEntraUserPicker } from '../../shared/entra-user-picker.js';
import { mountInlineEntraGroupPicker } from '../../shared/entra-group-picker.js';

/** @type {HTMLElement|null} */
let activeOverlay = null;
/** Zuletzt genutzter Bereich für neue Funktionen (Rechtsklick / Toolbar). */
let lastRoleParentId = 'sec_verwaltung';
const LAST_ROLE_PARENT_KEY = 'ms365-schul-org-last-role-parent-v1';

export function getActiveOverlay() {
    return activeOverlay;
}

export function setActiveOverlay(el) {
    activeOverlay = el;
}

export function readLastRoleParentId() {
    try {
        if (typeof localStorage !== 'undefined') {
            const v = normStr(localStorage.getItem(LAST_ROLE_PARENT_KEY));
            if (v) lastRoleParentId = v;
        }
    } catch {
        /* ignore */
    }
    return lastRoleParentId;
}

/**
 * @param {string} areaId
 */
export function rememberSchulOrgRoleParent(areaId) {
    const id = normStr(areaId);
    if (!id) return;
    lastRoleParentId = id;
    try {
        if (typeof localStorage !== 'undefined') {
            localStorage.setItem(LAST_ROLE_PARENT_KEY, id);
        }
    } catch {
        /* ignore */
    }
}

/**
 * @returns {string}
 */
export function getLastSchulOrgRoleParent() {
    return readLastRoleParentId();
}

export function closeOverlay() {
    document.removeEventListener('keydown', onEsc, true);
    if (activeOverlay && activeOverlay.parentNode) activeOverlay.parentNode.removeChild(activeOverlay);
    activeOverlay = null;
}

export function onEsc(ev) {
    if (ev.key === 'Escape') {
        ev.preventDefault();
        closeOverlay();
    }
}

/**
 * Entra-Personensuche im offenen Dialog einbetten (kein zweites Overlay).
 * @param {HTMLElement|null} host
 * @param {{ api: { destroy: () => void, focus: () => void }|null }} state
 * @param {{ multi?: boolean, hint?: string, excludeEmails?: string[]|Set<string>, onPick?: Function, onApply?: Function }} opts
 * @param {{ api: { destroy: () => void }|null }[]} [closeOthers]
 */
/**
 * Inline-Picker als Overlay im Dialog-Body (ohne Höhenwachstum).
 * @param {HTMLElement|null} host
 * @param {{
 *   api: { destroy: () => void }|null,
 *   _eupHost?: HTMLElement|null,
 *   _eupHome?: { parent: Node, next: Node|null }|null
 * }} state
 * @param {boolean} open
 */
export function setInlinePickerOverlay(host, state, open) {
    if (!state) return;
    const el = host || state._eupHost;
    if (!el) return;
    const box = el.closest ? el.closest('.schul-org-edit-box') : null;
    const body = box ? box.querySelector('.schul-org-edit-body') : null;
    if (open) {
        state._eupHost = el;
        if (box) box.classList.add('is-inline-picker-open');
        if (body && el.parentElement !== body) {
            state._eupHome = { parent: el.parentElement, next: el.nextSibling };
            body.appendChild(el);
        }
        return;
    }
    if (box) box.classList.remove('is-inline-picker-open');
    const home = state._eupHome;
    if (home && home.parent && el.parentElement !== home.parent) {
        try {
            home.parent.insertBefore(el, home.next);
        } catch {
            home.parent.appendChild(el);
        }
    }
    state._eupHome = null;
    state._eupHost = null;
}

export async function toggleInlineEntraUserPicker(host, state, opts, closeOthers) {
    if (!host || !state) return;
    if (state.api) {
        state.api.destroy();
        state.api = null;
        setInlinePickerOverlay(host, state, false);
        return;
    }
    (closeOthers || []).forEach(function (s) {
        if (s && s.api && s !== state) {
            const otherHost = s._eupHost || null;
            s.api.destroy();
            s.api = null;
            setInlinePickerOverlay(otherHost, s, false);
        }
    });
    const options = opts || {};
    setInlinePickerOverlay(host, state, true);
    state.api = mountInlineEntraUserPicker(host, {
        multi: !!options.multi,
        hint: options.hint,
        excludeEmails: options.excludeEmails,
        onPick: function (user) {
            state.api = null;
            setInlinePickerOverlay(host, state, false);
            if (typeof options.onPick === 'function') options.onPick(user);
        },
        onApply: function (users) {
            state.api = null;
            setInlinePickerOverlay(host, state, false);
            if (typeof options.onApply === 'function') options.onApply(users);
        },
        onCancel: function () {
            state.api = null;
            setInlinePickerOverlay(host, state, false);
        }
    });
}

/**
 * Entra-Gruppensuche im offenen Dialog einbetten (kein zweites Overlay).
 * @param {HTMLElement|null} host
 * @param {{ api: { destroy: () => void, focus: () => void }|null }} state
 * @param {{ hint?: string, onPick?: Function }} opts
 * @param {{ api: { destroy: () => void }|null }[]} [closeOthers]
 */
export async function toggleInlineEntraGroupPicker(host, state, opts, closeOthers) {
    if (!host || !state) return;
    if (state.api) {
        state.api.destroy();
        state.api = null;
        setInlinePickerOverlay(host, state, false);
        return;
    }
    (closeOthers || []).forEach(function (s) {
        if (s && s.api && s !== state) {
            const otherHost = s._eupHost || null;
            s.api.destroy();
            s.api = null;
            setInlinePickerOverlay(otherHost, s, false);
        }
    });
    const options = opts || {};
    setInlinePickerOverlay(host, state, true);
    state.api = mountInlineEntraGroupPicker(host, {
        hint: options.hint,
        onPick: function (group) {
            state.api = null;
            setInlinePickerOverlay(host, state, false);
            if (typeof options.onPick === 'function') options.onPick(group);
        },
        onCancel: function () {
            state.api = null;
            setInlinePickerOverlay(host, state, false);
        }
    });
}
