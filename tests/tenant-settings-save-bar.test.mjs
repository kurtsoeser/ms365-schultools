import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { installTenantStatusPanel } from '../src/shared/tenant-settings-ui-status.js';

function makeClassList(initial) {
    const set = new Set(initial || []);
    return {
        contains(c) {
            return set.has(c);
        },
        add(c) {
            set.add(c);
        },
        remove(c) {
            set.delete(c);
        },
        toggle(c, force) {
            if (force === true) set.add(c);
            else if (force === false) set.delete(c);
            else if (set.has(c)) set.delete(c);
            else set.add(c);
            return set.has(c);
        }
    };
}

function el(attrs) {
    const node = {
        id: attrs.id || '',
        className: attrs.className || '',
        classList: makeClassList((attrs.className || '').split(/\s+/).filter(Boolean)),
        dataset: Object.assign({}, attrs.dataset || {}),
        hidden: !!attrs.hidden,
        textContent: attrs.textContent || '',
        attrs: {},
        setAttribute(k, v) {
            this.attrs[k] = String(v);
            if (k.startsWith('data-')) this.dataset[k.slice(5)] = String(v);
        },
        removeAttribute(k) {
            delete this.attrs[k];
        },
        getAttribute(k) {
            return this.attrs[k] || null;
        }
    };
    return node;
}

describe('Stammdaten Save-Bar Dirty-State', () => {
    let dock;
    let bar;
    let hint;
    let saveBtn;
    let prevDocument;
    let prevWindow;

    beforeEach(() => {
        dock = el({ className: 'ts-admin-dock ts-register-dock' });
        bar = el({ id: 'tenantSaveBar', dataset: { dirty: '0' } });
        hint = el({ id: 'tenantSaveDirtyHint', hidden: true });
        saveBtn = el({ className: 'btn fb-btn fb-btn--ghost' });
        prevDocument = globalThis.document;
        prevWindow = globalThis.window;
        const doc = {
            querySelector(sel) {
                if (sel === '.ts-admin-dock.ts-register-dock') return dock;
                return null;
            },
            querySelectorAll(sel) {
                if (sel === '[data-tenant-action="save"]') return [saveBtn];
                return [];
            },
            getElementById(id) {
                if (id === 'tenantSaveBar') return bar;
                if (id === 'tenantSaveDirtyHint') return hint;
                return null;
            }
        };
        globalThis.document = doc;
        globalThis.window = globalThis;
    });

    afterEach(() => {
        globalThis.document = prevDocument;
        globalThis.window = prevWindow;
    });

    it('markiert Dirty und hebt Speichern hervor', () => {
        const ctx = { autoSaveNow() {} };
        installTenantStatusPanel(ctx);
        ctx.markRegisterDirty();
        expect(dock.classList.contains('is-dirty')).toBe(true);
        expect(bar.dataset.dirty).toBe('1');
        expect(saveBtn.classList.contains('fb-btn--ghost')).toBe(false);
        expect(hint.hidden).toBe(false);
        expect(hint.textContent).toMatch(/ungespeicherte/i);
    });

    it('scheduleAutoSave setzt Dirty; markRegisterClean räumt auf', () => {
        const ctx = { autoSaveNow() {} };
        installTenantStatusPanel(ctx);
        ctx.scheduleAutoSave();
        expect(ctx.registerSaveDirty).toBe(true);
        expect(ctx.listEditSincePaint).toBe(true);
        if (ctx.autoSaveTimer) {
            clearTimeout(ctx.autoSaveTimer);
            ctx.autoSaveTimer = null;
        }
        ctx.markRegisterClean();
        expect(ctx.registerSaveDirty).toBe(false);
        expect(dock.classList.contains('is-dirty')).toBe(false);
        expect(saveBtn.classList.contains('fb-btn--ghost')).toBe(true);
    });
});
