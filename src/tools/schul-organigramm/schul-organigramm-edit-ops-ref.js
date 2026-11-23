/**
 * Organigramm node label/ref patch + remove helpers.
 */
import { normStr } from './schul-organigramm-edit-utils.js';
import {
    isClassNode,
    isArgeNode,
    removeClassEntityFromSettings,
    renameClassEntityInSettings,
    removeArgeEntityFromSettings
} from './schul-organigramm-logic.js';

export function patchSchulOrganigrammRefLabel(settings, type, oldRefKey, nextLabel, nextRefKey) {
    const layout = settings.schulOrganigramm;
    if (!layout || !Array.isArray(layout.nodes)) return;
    const oldKey = normStr(oldRefKey).toLowerCase();
    const nextKey =
        nextRefKey != null && String(nextRefKey) !== '' ? normStr(nextRefKey) : '';
    // Klassen: volle Entity-Umbenennung (IDs, Spiegelbäume, Parent-Refs)
    if (type === 'class' && nextKey && nextKey.toUpperCase() !== normStr(oldRefKey).toUpperCase()) {
        const renamed = renameClassEntityInSettings(settings, oldRefKey, nextKey, {
            name: nextLabel
        });
        Object.keys(renamed.settings).forEach(function (k) {
            settings[k] = renamed.settings[k];
        });
        return;
    }
    if (type === 'class' && nextKey && nextKey.toUpperCase() === normStr(oldRefKey).toUpperCase()) {
        const renamed = renameClassEntityInSettings(settings, oldRefKey, nextKey, {
            name: nextLabel
        });
        Object.keys(renamed.settings).forEach(function (k) {
            settings[k] = renamed.settings[k];
        });
        return;
    }
    layout.nodes = layout.nodes.map(function (n) {
        const t = normStr(n.type).toLowerCase();
        const rk = normStr(n.refKind).toLowerCase();
        const matchType =
            t === type ||
            (type === 'class' && (rk === 'class' || isClassNode(n))) ||
            (type === 'arge' && (rk === 'arge' || isArgeNode(n))) ||
            (type === 'role' && (t === 'role' || rk === 'adminrole'));
        if (!matchType) return n;
        if (normStr(n.refKey).toLowerCase() !== oldKey) return n;
        const next = Object.assign({}, n, { label: nextLabel || n.label });
        if (nextKey) next.refKey = nextKey;
        return next;
    });
    settings.schulOrganigramm = layout;
}

export function removeSchulOrganigrammByRef(settings, type, refKey) {
    if (type === 'class') {
        const removed = removeClassEntityFromSettings(settings, refKey);
        Object.keys(removed.settings).forEach(function (k) {
            settings[k] = removed.settings[k];
        });
        return;
    }
    if (type === 'arge') {
        const next = removeArgeEntityFromSettings(settings, refKey);
        Object.keys(next).forEach(function (k) {
            settings[k] = next[k];
        });
        return;
    }
    const layout = settings.schulOrganigramm;
    if (!layout || !Array.isArray(layout.nodes)) return;
    const key = normStr(refKey).toLowerCase();
    settings.schulOrganigramm = {
        version: layout.version || 2,
        view: layout.view,
        removedAreaIds: layout.removedAreaIds,
        nodes: layout.nodes.filter(function (n) {
            const t = normStr(n.type).toLowerCase();
            const rk = normStr(n.refKind).toLowerCase();
            const matchType =
                t === type ||
                (type === 'role' && (t === 'role' || rk === 'adminrole'));
            if (!matchType) return true;
            return normStr(n.refKey).toLowerCase() !== key;
        })
    };
}
