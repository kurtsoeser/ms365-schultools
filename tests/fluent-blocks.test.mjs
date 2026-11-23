import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { dismissFluentToast, showFluentToast } from '../src/shared/fluent-blocks.js';

function makeFakeDom() {
    const byId = new Map();
    const makeEl = (tag) => {
        const listeners = new Map();
        const children = [];
        const el = {
            tagName: tag.toUpperCase(),
            id: '',
            className: '',
            type: '',
            textContent: '',
            attrs: {},
            children,
            setAttribute(k, v) {
                this.attrs[k] = v;
            },
            appendChild(child) {
                children.push(child);
                if (child.id) byId.set(child.id, child);
                return child;
            },
            remove() {
                if (this.id) byId.delete(this.id);
            },
            addEventListener(type, fn) {
                listeners.set(type, fn);
            },
            click() {
                listeners.get('click')?.();
            },
            querySelector(sel) {
                const want = sel.toUpperCase();
                return children.find((c) => c.tagName === want) || null;
            },
            get fullText() {
                return [this.textContent, ...children.map((c) => c.fullText ?? c.textContent)].join(' ');
            }
        };
        return new Proxy(el, {
            set(target, prop, value) {
                if (prop === 'id' && target.id && target.id !== value) byId.delete(target.id);
                target[prop] = value;
                if (prop === 'id' && value) byId.set(value, target);
                return true;
            }
        });
    };
    const body = makeEl('body');
    return {
        body,
        createElement: makeEl,
        getElementById: (id) => byId.get(id) || null
    };
}

describe('fluent-blocks toast', () => {
    const prevDoc = globalThis.document;
    beforeEach(() => {
        globalThis.document = makeFakeDom();
    });
    afterEach(() => {
        dismissFluentToast();
        if (prevDoc === undefined) delete globalThis.document;
        else globalThis.document = prevDoc;
    });

    it('zeigt Meldung mit Rückgängig und räumt beim Klick auf', () => {
        const onUndo = vi.fn();
        const el = showFluentToast({
            message: 'Freistellung genehmigt.',
            actionLabel: 'Rückgängig',
            onAction: onUndo,
            durationMs: 5000
        });
        expect(el).toBeTruthy();
        expect(document.getElementById('ms365FluentToast')?.fullText).toContain('Freistellung genehmigt.');
        const btn = el.querySelector('button');
        expect(btn?.textContent).toBe('Rückgängig');
        btn.click();
        expect(onUndo).toHaveBeenCalledTimes(1);
        expect(document.getElementById('ms365FluentToast')).toBeNull();
    });
});
