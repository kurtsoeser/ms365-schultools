import { describe, expect, it } from 'vitest';
import { readSchoolItAccessDeniedParams } from '../src/shared/school-it-access-denied-ui.js';

describe('school-it-access-denied-ui', () => {
    it('liest access=denied&reason=school-it', () => {
        expect(readSchoolItAccessDeniedParams('?access=denied&reason=school-it')).toEqual({
            access: 'denied',
            reason: 'school-it'
        });
    });

    it('ignoriert andere Query-Parameter', () => {
        expect(readSchoolItAccessDeniedParams('?foo=1')).toBeNull();
        expect(readSchoolItAccessDeniedParams('?access=ok')).toBeNull();
    });
});
