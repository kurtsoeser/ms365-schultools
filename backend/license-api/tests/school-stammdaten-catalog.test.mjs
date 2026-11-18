import assert from 'node:assert/strict';
import test from 'node:test';
import { createRequire } from 'node:module';

const require = createRequire(import.meta.url);
const {
    SCHOOL_STAMMDATEN_FIELDS,
    isRealTenantGuid,
    isPendingTenantId,
    makePendingTenantId,
    normalizeTenantIdInput,
    assertTenantIdAllowed
} = require('../src/lib/school-stammdaten-catalog.js');

test('Stammdaten-Katalog enthält Adresse und Rollen', () => {
    const names = SCHOOL_STAMMDATEN_FIELDS.map((f) => f.name);
    assert.ok(names.includes('Standort'));
    assert.ok(names.includes('Strasse'));
    assert.ok(names.includes('Direktion'));
    assert.ok(names.includes('Schulnummer'));
    assert.ok(names.includes('ItM365Admin'));
    assert.ok(names.includes('OffenePunkte'));
});

test('Tenant GUID und PENDING erkennen', () => {
    assert.equal(isRealTenantGuid('11111111-1111-4111-8111-111111111111'), true);
    assert.equal(isRealTenantGuid('PENDING-hak-linz'), false);
    assert.equal(isPendingTenantId(''), true);
    assert.equal(isPendingTenantId('OFFEN – nicht dokumentiert'), true);
    assert.equal(isPendingTenantId('PENDING-hak-linz'), true);
    assert.equal(isPendingTenantId('11111111-1111-4111-8111-111111111111'), false);
});

test('leere Tenant-ID wird zu PENDING-Slug', () => {
    assert.equal(makePendingTenantId('HAK.HAS Linz'), 'PENDING-hak-has-linz');
    assert.equal(normalizeTenantIdInput('', 'Perg'), 'PENDING-perg');
    assert.equal(normalizeTenantIdInput('OFFEN – x', 'Gmunden'), 'PENDING-gmunden');
    assert.equal(
        assertTenantIdAllowed('', 'HAK Linz'),
        'PENDING-hak-linz'
    );
    assert.equal(
        assertTenantIdAllowed('11111111-1111-4111-8111-111111111111', 'X'),
        '11111111-1111-4111-8111-111111111111'
    );
    assert.notEqual(
        makePendingTenantId('Bundeshandelsakademie und Bundeshandelsschule Perg'),
        makePendingTenantId('Bundeshandelsakademie und Bundeshandelsschule Gmunden')
    );
    assert.match(makePendingTenantId('BHAK Schärding'), /^PENDING-bhak-schaerding$/i);
});

test('ungültige Tenant-ID wird abgelehnt', () => {
    assert.throws(() => assertTenantIdAllowed('nicht-gueltig', 'Test'));
});
