import assert from 'node:assert/strict';
import test from 'node:test';
import { createRequire } from 'node:module';

const require = createRequire(import.meta.url);
const {
    mapListItem,
    fieldsFromBody,
    isExtraColumnName,
    sanitizeInternalName,
    buildColumnDefinition
} = require('../src/lib/sharepoint-license.js');

test('isExtraColumnName filters core and system fields', () => {
    assert.equal(isExtraColumnName('Title'), false);
    assert.equal(isExtraColumnName('TenantId'), false);
    assert.equal(isExtraColumnName('Author'), false);
    assert.equal(isExtraColumnName('Schulnummer'), true);
});

test('sanitizeInternalName removes diacritics and spaces', () => {
    assert.equal(sanitizeInternalName('Schulnummer'), 'Schulnummer');
    assert.equal(sanitizeInternalName('Bezirk Süd'), 'BezirkSud');
    assert.equal(sanitizeInternalName('123x'), 'X123x');
});

test('buildColumnDefinition supports text and choice', () => {
    const text = buildColumnDefinition({ displayName: 'Partner', type: 'text' });
    assert.equal(text.name, 'Partner');
    assert.ok(text.text);

    const choice = buildColumnDefinition({
        displayName: 'Stufe',
        type: 'choice',
        choices: 'A\nB'
    });
    assert.deepEqual(choice.choice.choices, ['A', 'B']);
});

test('mapListItem maps extra fields', () => {
    const item = {
        id: '9',
        fields: {
            Title: 'Testschule',
            TenantId: 'abc',
            Status: 'active',
            Schulnummer: '1234',
            Notes: 'hi'
        }
    };
    const mapped = mapListItem(item, [{ name: 'Schulnummer', type: 'text' }]);
    assert.equal(mapped.schoolName, 'Testschule');
    assert.equal(mapped.extra.Schulnummer, '1234');
    assert.equal(mapped.notes, 'hi');
});

test('fieldsFromBody writes extra values', () => {
    const fields = fieldsFromBody(
        {
            schoolName: 'A',
            tenantId: 't1',
            status: 'active',
            extra: { Schulnummer: '99', Flag: true }
        },
        {
            partial: false,
            extraColumns: [
                { name: 'Schulnummer', type: 'text' },
                { name: 'Flag', type: 'boolean' }
            ]
        }
    );
    assert.equal(fields.Title, 'A');
    assert.equal(fields.Schulnummer, '99');
    assert.equal(fields.Flag, true);
});
