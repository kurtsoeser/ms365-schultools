import { describe, expect, it } from 'vitest';
import {
    SALES_BUNDLES,
    salesBundleById
} from '../src/shared/sales-bundles-catalog.js';
import {
    SALES_INQUIRY_MAIL,
    buildSalesInquiryMailto,
    expandPackageDeps,
    recommendSalesOffer
} from '../src/shared/sales-offer-inquiry.js';

describe('sales-offer-inquiry', () => {
    it('erkennt exaktes Bundle Schulalltag', () => {
        const pkgs = salesBundleById('schulalltag').packageIds;
        const rec = recommendSalesOffer(pkgs);
        expect(rec.mode).toBe('bundle');
        expect(rec.bundle?.id).toBe('schulalltag');
        expect(rec.offerEuro).toBe(890);
    });

    it('erweitert Abhängigkeiten (schulapps → intranet)', () => {
        const ids = expandPackageDeps(['schulapps']);
        expect(ids).toContain('schulapps');
        expect(ids).toContain('intranet');
    });

    it('markiert Bundle + Extra als Individuell mit Tipp', () => {
        const pkgs = [...salesBundleById('starter-it').packageIds, 'kommunikation'];
        const rec = recommendSalesOffer(pkgs);
        expect(rec.mode).toBe('custom');
        expect(rec.nearest?.id).toBe('starter-it');
        expect(rec.tip).toMatch(/Starter IT/i);
        expect(rec.tip).toMatch(/Kommunikation/i);
    });

    it('empfiehlt nahes Bundle wenn wenig fehlt', () => {
        const pkgs = salesBundleById('schulalltag').packageIds.filter(
            (id) => id !== 'freistellung'
        );
        const rec = recommendSalesOffer(pkgs);
        expect(rec.mode).toBe('custom');
        expect(rec.nearest?.id).toBe('schulalltag');
        expect(rec.tip).toMatch(/Freistellung/i);
    });

    it('baut mailto an kontakt@ms365.schule mit Auswahl', () => {
        const pkgs = salesBundleById('starter-it').packageIds;
        const mail = buildSalesInquiryMailto({
            name: 'Maria Muster',
            school: 'HTL Demo',
            email: 'maria@schule.at',
            message: 'Bitte um Rückruf',
            packageIds: pkgs,
            addonIds: ['addon-setup']
        });
        expect(mail.to).toBe(SALES_INQUIRY_MAIL);
        expect(mail.href).toMatch(/^mailto:kontakt@ms365\.schule\?/);
        expect(mail.subject).toMatch(/Starter IT/);
        expect(mail.body).toContain('HTL Demo');
        expect(mail.body).toContain('Basis (Core)');
        expect(mail.body).toContain('Setup / Onboarding');
        expect(mail.body).toContain('Bitte um Rückruf');
    });

    it('drei Verkaufsbündel bleiben Referenz', () => {
        expect(SALES_BUNDLES).toHaveLength(3);
    });
});
