/**
 * Lokales Aktionsprotokoll (kein Server). Ringpuffer in app-data-v2.setup.actionLog.
 */
(function () {
    'use strict';

    const MAX_ENTRIES = 200;

    function nowIso() {
        try {
            return new Date().toISOString();
        } catch {
            return '';
        }
    }

    function normStr(v) {
        return String(v ?? '').trim();
    }

    function normalizeEntry(raw) {
        const src = raw && typeof raw === 'object' ? raw : {};
        return {
            at: normStr(src.at) || nowIso(),
            tool: normStr(src.tool) || 'app',
            action: normStr(src.action) || 'write',
            target: normStr(src.target),
            summary: normStr(src.summary),
            result: src.result === 'error' || src.result === 'skip' ? src.result : 'ok'
        };
    }

    function readList() {
        try {
            if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.getSetup === 'function') {
                const setup = window.ms365AppDataV2.getSetup() || {};
                return Array.isArray(setup.actionLog) ? setup.actionLog.slice() : [];
            }
        } catch {
            /* ignore */
        }
        return [];
    }

    function writeList(list) {
        if (!window.ms365AppDataV2 || typeof window.ms365AppDataV2.patchSetup !== 'function') return false;
        const next = Array.isArray(list) ? list.slice() : [];
        while (next.length > MAX_ENTRIES) next.shift();
        window.ms365AppDataV2.patchSetup({ actionLog: next });
        return true;
    }

    function append(raw) {
        const entry = normalizeEntry(raw);
        const list = readList();
        list.push(entry);
        writeList(list);
        return entry;
    }

    function list(limit) {
        const all = readList();
        const n = typeof limit === 'number' && limit > 0 ? limit : all.length;
        return all.slice(Math.max(0, all.length - n)).reverse();
    }

    function clear() {
        writeList([]);
        return true;
    }

    function exportJson() {
        return JSON.stringify(
            {
                exportedAt: nowIso(),
                entries: readList()
            },
            null,
            2
        );
    }

    window.ms365ActionLog = {
        MAX_ENTRIES,
        normalizeEntry,
        append,
        list,
        clear,
        exportJson
    };
})();
