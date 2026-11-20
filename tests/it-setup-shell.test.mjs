import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import {
    clampItSetupStep,
    formatItSetupPhaseHint,
    renderItSetupStepperHtml,
    renderItSetupSummaryListHtml,
    createItSetupWizardController
} from '../src/shared/it-setup-shell.js';

function memoryStorage() {
    /** @type {Record<string, string>} */
    const map = {};
    return {
        getItem(k) {
            return Object.prototype.hasOwnProperty.call(map, k) ? map[k] : null;
        },
        setItem(k, v) {
            map[k] = String(v);
        },
        removeItem(k) {
            delete map[k];
        }
    };
}

describe('it-setup-shell', () => {
    it('clampItSetupStep', () => {
        expect(clampItSetupStep(0, 4)).toBe(1);
        expect(clampItSetupStep(2, 4)).toBe(2);
        expect(clampItSetupStep(99, 4)).toBe(4);
        expect(clampItSetupStep('3', 5)).toBe(3);
    });

    it('formatItSetupPhaseHint', () => {
        expect(formatItSetupPhaseHint(2, 4, ['Website', 'Genehmigung', 'Einreicher', 'Paket'])).toBe(
            'Schritt 2 von 4: Genehmigung'
        );
    });

    it('renderItSetupStepperHtml enthält Nav-Klassen', () => {
        const html = renderItSetupStepperHtml({
            id: 'demoGlance',
            stepAttr: 'data-demo-step',
            cols: 4,
            steps: [
                { n: 1, label: 'Website', icon: 'bi-link-45deg' },
                { n: 2, label: 'Rechte', icon: 'bi-shield-lock' }
            ]
        });
        expect(html).toContain('it-setup-glance--nav');
        expect(html).toContain('it-setup-glance--cols-4');
        expect(html).toContain('data-demo-step="1"');
        expect(html).toContain('1 · Website');
        // Standard: zweizeilig wie sRDP (kein Value-Slot ohne st.value)
        expect(html).not.toContain('tm-glance__value');
    });

    it('renderItSetupSummaryListHtml', () => {
        const html = renderItSetupSummaryListHtml([
            { label: 'Website', value: 'https://schule.sharepoint.com' },
            'Paket bereit'
        ]);
        expect(html).toContain('it-setup-summary-list');
        expect(html).toContain('<strong>Website:</strong>');
        expect(html).toContain('Paket bereit');
    });
});

describe('createItSetupWizardController', () => {
    const prev = globalThis.localStorage;
    const prevDoc = globalThis.document;

    beforeEach(() => {
        globalThis.localStorage = memoryStorage();
        globalThis.document = {
            getElementById() {
                return null;
            },
            querySelector() {
                return null;
            }
        };
    });

    afterEach(() => {
        if (prev === undefined) delete globalThis.localStorage;
        else globalThis.localStorage = prev;
        if (prevDoc === undefined) delete globalThis.document;
        else globalThis.document = prevDoc;
    });

    it('speichert und lädt Schritt', () => {
        const wiz = createItSetupWizardController({
            stepCount: 4,
            labels: ['A', 'B', 'C', 'D'],
            storageKey: 'test-it-setup-step',
            glanceSelector: '#x',
            stepAttr: 'data-x',
            panelIdPrefix: 'pane'
        });
        expect(wiz.loadStep()).toBe(1);
        wiz.saveStep(3);
        expect(wiz.loadStep()).toBe(3);
        expect(wiz.clampStep(99)).toBe(4);
    });
});
