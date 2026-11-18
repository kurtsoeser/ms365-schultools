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
    assert.equal(isExtraColumnName('ToolSiteUrl'), false);
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
            Notes: 'hi',
            ToolSiteUrl: 'https://schule.sharepoint.com/sites/MS365-Schultools'
        }
    };
    const mapped = mapListItem(item, [{ name: 'Schulnummer', type: 'text' }]);
    assert.equal(mapped.schoolName, 'Testschule');
    assert.equal(mapped.extra.Schulnummer, '1234');
    assert.equal(mapped.notes, 'hi');
    assert.equal(
        mapped.toolSiteUrl,
        'https://schule.sharepoint.com/sites/MS365-Schultools'
    );
});

test('fieldsFromBody writes ToolSiteUrl', () => {
    const fields = fieldsFromBody(
        {
            schoolName: 'A',
            tenantId: '11111111-1111-4111-8111-111111111111',
            status: 'active',
            toolSiteUrl: 'https://schule.sharepoint.com/sites/X/'
        },
        { partial: false }
    );
    assert.equal(fields.ToolSiteUrl, 'https://schule.sharepoint.com/sites/X/');
});

test('pickToolSiteUrlFromFields liest Text und Hyperlink', () => {
    const { pickToolSiteUrlFromFields, findToolSiteUrlColumnMeta, toolSiteUrlFieldValue } =
        require('../src/lib/sharepoint-license.js');
    assert.equal(
        pickToolSiteUrlFromFields({ ToolSiteUrl: 'https://a.sharepoint.com/sites/S' }),
        'https://a.sharepoint.com/sites/S'
    );
    assert.equal(
        pickToolSiteUrlFromFields({
            Tool_x002d_Site_x002d_URL: { Url: 'https://b.sharepoint.com/sites/T', Description: 'T' }
        }),
        'https://b.sharepoint.com/sites/T'
    );
    const meta = findToolSiteUrlColumnMeta([
        { name: 'Title', displayName: 'Title' },
        {
            name: 'Tool_x002d_Site_x002d_URL',
            displayName: 'Tool-Site-URL (Schul-SharePoint)',
            hyperlinkOrPicture: {}
        }
    ]);
    assert.equal(meta.name, 'Tool_x002d_Site_x002d_URL');
    assert.equal(meta.kind, 'hyperlink');
    assert.deepEqual(toolSiteUrlFieldValue('https://x/', 'hyperlink'), {
        Url: 'https://x/',
        Description: 'https://x/'
    });
});

test('fieldsFromBody schreibt Hyperlink-Spalte mit alternativem Namen', () => {
    const fields = fieldsFromBody(
        {
            schoolName: 'A',
            tenantId: '11111111-1111-4111-8111-111111111111',
            status: 'active',
            toolSiteUrl: 'https://schule.sharepoint.com/sites/Y'
        },
        {
            partial: true,
            toolSiteUrlFieldName: 'Tool_x002d_Site_x002d_URL',
            toolSiteUrlFieldKind: 'hyperlink'
        }
    );
    assert.deepEqual(fields.Tool_x002d_Site_x002d_URL, {
        Url: 'https://schule.sharepoint.com/sites/Y',
        Description: 'https://schule.sharepoint.com/sites/Y'
    });
    assert.equal(fields.ToolSiteUrl, undefined);
});

test('fieldsFromBody writes extra values', () => {
    const fields = fieldsFromBody(
        {
            schoolName: 'A',
            tenantId: '11111111-1111-4111-8111-111111111111',
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
    assert.equal(fields.TenantId, '11111111-1111-4111-8111-111111111111');
    assert.equal(fields.Schulnummer, '99');
    assert.equal(fields.Flag, true);
});

test('fieldsFromBody erlaubt PENDING-Tenant', () => {
    const fields = fieldsFromBody(
        {
            schoolName: 'HAK Linz',
            tenantId: '',
            status: 'trial',
            extra: { Ort: 'Linz' }
        },
        {
            partial: false,
            extraColumns: [{ name: 'Ort', type: 'text' }]
        }
    );
    assert.equal(fields.TenantId, 'PENDING-hak-linz');
    assert.equal(fields.Ort, 'Linz');
});
