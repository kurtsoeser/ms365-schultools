import { describe, it, expect } from 'vitest';
import { emptyForm, prefillLehrerFormIdentity } from '../src/tools/lehrer-freistellung-planer/lfr-state.js';

describe('prefillLehrerFormIdentity', () => {
    it('füllt Lehrkraft und E-Mail aus dem angemeldeten Konto', () => {
        const state = {
            accountEmail: 'max.mueller@schule.at',
            accountName: 'Max Müller',
            stammdaten: { teachers: [] },
            form: emptyForm()
        };
        prefillLehrerFormIdentity(state);
        expect(state.form.lehrerEmail).toBe('max.mueller@schule.at');
        expect(state.form.lehrerName).toBe('Max Müller');
    });

    it('nutzt Stammdaten-Namen wenn Konto keinen Anzeigenamen hat', () => {
        const state = {
            accountEmail: 'anna@schule.at',
            accountName: '',
            stammdaten: {
                teachers: [{ email: 'anna@schule.at', name: 'Mag. Anna Mayer' }]
            },
            form: emptyForm()
        };
        prefillLehrerFormIdentity(state);
        expect(state.form.lehrerEmail).toBe('anna@schule.at');
        expect(state.form.lehrerName).toBe('Mag. Anna Mayer');
    });

    it('überschreibt bestehende Werte nur mit force', () => {
        const state = {
            accountEmail: 'neu@schule.at',
            accountName: 'Neu',
            stammdaten: { teachers: [] },
            form: emptyForm({ lehrerEmail: 'alt@schule.at', lehrerName: 'Alt' })
        };
        prefillLehrerFormIdentity(state);
        expect(state.form.lehrerEmail).toBe('alt@schule.at');
        expect(state.form.lehrerName).toBe('Alt');
        prefillLehrerFormIdentity(state, { force: true });
        expect(state.form.lehrerEmail).toBe('neu@schule.at');
        expect(state.form.lehrerName).toBe('Neu');
    });
});
