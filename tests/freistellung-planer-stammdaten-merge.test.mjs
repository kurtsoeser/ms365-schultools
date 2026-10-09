import { describe, it, expect } from 'vitest';
import { mergeSchoolClassRows } from '../src/tools/freistellung-planer/freistellung-planer-state.js';

describe('mergeSchoolClassRows', () => {
    it('behält headEmail aus Schulregister wenn Kanon nur Code hat', () => {
        const merged = mergeSchoolClassRows(
            [{ code: '1A', headName: 'Brian May', headEmail: 'brian.may@kurtrocks.com' }],
            [{ code: '1A', name: 'DEMO Klasse 1A', year: '2032' }]
        );
        expect(merged).toHaveLength(1);
        expect(merged[0].headEmail).toBe('brian.may@kurtrocks.com');
        expect(merged[0].name).toContain('DEMO');
    });
});
