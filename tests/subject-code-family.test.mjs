import { describe, it, expect } from 'vitest';
import {
    splitSubjectBaseAndSuffix,
    groupSubjectsByBase,
    subjectHasVariantSuffix
} from '../src/shared/subject-code-family.js';

describe('subject-code-family', () => {
    it('splitSubjectBaseAndSuffix erkennt Endziffern', () => {
        expect(splitSubjectBaseAndSuffix('OMAI1')).toEqual({ base: 'OMAI', suffix: '1' });
        expect(splitSubjectBaseAndSuffix('omaik2')).toEqual({ base: 'OMAIK', suffix: '2' });
        expect(splitSubjectBaseAndSuffix('ENWS')).toEqual({ base: 'ENWS', suffix: '' });
        expect(splitSubjectBaseAndSuffix('D')).toEqual({ base: 'D', suffix: '' });
    });

    it('splitSubjectBaseAndSuffix erkennt Ü und Plus-Varianten', () => {
        expect(splitSubjectBaseAndSuffix('RWÜ')).toEqual({ base: 'RW', suffix: 'Ü' });
        expect(splitSubjectBaseAndSuffix('deü')).toEqual({ base: 'DE', suffix: 'Ü' });
        expect(splitSubjectBaseAndSuffix('M+')).toEqual({ base: 'M', suffix: '+' });
        expect(splitSubjectBaseAndSuffix('PHY+LAB')).toEqual({ base: 'PHY', suffix: '+LAB' });
    });

    it('groupSubjectsByBase fasst Ziffer, Ü und Plus zusammen', () => {
        const groups = groupSubjectsByBase(['D', 'OMAI', 'OMAI1', 'RW', 'RWÜ', 'M', 'M+']);
        const oma = groups.find((g) => g.base === 'OMAI');
        expect(oma.isFamily).toBe(true);
        expect(oma.variants).toEqual(['OMAI', 'OMAI1']);
        const rw = groups.find((g) => g.base === 'RW');
        expect(rw.variants).toEqual(['RW', 'RWÜ']);
        const m = groups.find((g) => g.base === 'M');
        expect(m.variants).toEqual(['M', 'M+']);
        expect(subjectHasVariantSuffix('RWÜ')).toBe(true);
        expect(subjectHasVariantSuffix('M+')).toBe(true);
        expect(subjectHasVariantSuffix('RW')).toBe(false);
    });
});
