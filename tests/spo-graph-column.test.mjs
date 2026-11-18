import { describe, expect, it } from 'vitest';
import { toGraphColumnBody } from '../src/shared/spo-graph-column.js';

describe('toGraphColumnBody', () => {
    it('kappt Single-Line maxLength auf 255', () => {
        const body = toGraphColumnBody({
            name: 'AnhangUrl',
            displayName: 'URL',
            text: { allowMultipleLines: false, maxLength: 500 }
        });
        expect(body.text.allowMultipleLines).toBe(false);
        expect(body.text.maxLength).toBe(255);
        expect(body.text.textType).toBe('plain');
    });

    it('setzt Mehrzeiler ohne maxLength und mit textType', () => {
        const body = toGraphColumnBody({
            name: 'InhaltHtml',
            displayName: 'Inhalt',
            text: { allowMultipleLines: true }
        });
        expect(body.text).toEqual({
            allowMultipleLines: true,
            textType: 'plain'
        });
        expect(body.text.maxLength).toBeUndefined();
    });

    it('serialisiert Choice und Boolean sauber', () => {
        expect(
            toGraphColumnBody({
                name: 'Status',
                choice: { allowTextEntry: false, choices: ['a', 'b'] }
            }).choice
        ).toEqual({ allowTextEntry: false, choices: ['a', 'b'] });
        expect(
            toGraphColumnBody({
                name: 'Flag',
                boolean: { anything: true }
            }).boolean
        ).toEqual({});
    });
});
