/**
 * Schreibt Datenschutz-Texte aus src/shared/privacy-copy.js in:
 *   - hilfe.html (Kern + Kurztexte)
 *   - landing/datenschutz.html (gemeinsamer Kern)
 *   - landing/index.html (Teaser-Sektion)
 *   - README.md (Kurzabsatz)
 *
 * Neu: node scripts/sync-privacy-copy.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { PRIVACY_COPY } from '../src/shared/privacy-copy.js';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = path.join(__dirname, '..');

function replaceMarked(html, id, body, endIndent = '') {
    const start = `<!-- privacy-copy:${id}:start -->`;
    const end = `<!-- privacy-copy:${id}:end -->`;
    const i0 = html.indexOf(start);
    const i1 = html.indexOf(end);
    if (i0 < 0 || i1 < 0 || i1 < i0) {
        throw new Error(`Marker privacy-copy:${id} fehlt oder ist kaputt`);
    }
    const afterEnd = html.slice(i1 + end.length);
    return (
        html.slice(0, i0 + start.length) +
        '\n' +
        body +
        '\n' +
        endIndent +
        end +
        afterEnd
    );
}

function p(html) {
    return `                        <p>\n                            ${html}\n                        </p>`;
}

function pLanding(html) {
    return `        <p>\n          ${html}\n        </p>`;
}

function ulHelp(items) {
    const lis = items.map((item) => `                            <li>${item}</li>`).join('\n');
    return `                        <ul>\n${lis}\n                        </ul>`;
}

function ulLanding(items, className = 'legal-list') {
    const lis = items.map((item) => `          <li>\n            ${item}\n          </li>`).join('\n');
    return `        <ul class="${className}">\n${lis}\n        </ul>`;
}

function buildHelpCore() {
    const c = PRIVACY_COPY.core;
    const parts = [];

    parts.push(`                        <h2 id="datenschutz">Datenschutz</h2>`);
    for (const para of c.intro.paragraphs) parts.push(p(para));
    parts.push(
        `                        <p class="help-lead" style="margin-top:8px;">Stand: ${PRIVACY_COPY.standLabel}. Quelle: <code>src/shared/privacy-copy.js</code> (gemeinsam mit ms365.schule).</p>`
    );

    parts.push(`                        <h3>${c.schoolResponsibility.titleHelp}</h3>`);
    for (const para of c.schoolResponsibility.paragraphs) parts.push(p(para));

    parts.push(`                        <h3>${c.noProductDb.titleHelp}</h3>`);
    for (const para of c.noProductDb.paragraphs) parts.push(p(para));

    parts.push(`                        <h3>${c.localData.titleHelp}</h3>`);
    parts.push(p(c.localData.lead));
    parts.push(ulHelp(c.localData.list));
    parts.push(p(c.localData.after));

    parts.push(`                        <h3>${c.backup.titleHelp}</h3>`);
    for (const para of c.backup.paragraphs) {
        // Hilfe: Link zur Stammdaten-Übergabe im ersten SharePoint-Absatz ergänzen
        const withLink = para.includes('Nach Einrichtung der IT-Sicherungsbibliothek schreibt')
            ? para.replace(
                  'Nach Einrichtung der IT-Sicherungsbibliothek schreibt',
                  'Nach Einrichtung der IT-Sicherungsbibliothek (<a href="#tool-stammdaten-uebergabe">Stammdaten-Übergabe</a>) schreibt'
              )
            : para;
        parts.push(p(withLink));
    }
    parts.push(
        p(
            '<strong>Nach SharePoint</strong> sichert, <strong>Von SharePoint</strong> ersetzt den lokalen Browser-Stand; nach der Anmeldung kann ein Auto-Sync den Stand aktuell halten.'
        )
    );

    parts.push(`                        <h3>${c.microsoft.titleHelp}</h3>`);
    for (const para of c.microsoft.paragraphs) parts.push(p(para));

    parts.push(`                        <h3>${c.operatorServices.titleHelp}</h3>`);
    parts.push(p(c.operatorServices.lead));
    parts.push(ulHelp(c.operatorServices.list));
    parts.push(p(c.operatorServices.after));

    parts.push(`                        <h3>${c.retention.titleHelp}</h3>`);
    parts.push(ulHelp(c.retention.list));

    parts.push(`                        <h3>${c.rightsShared.titleHelp}</h3>`);
    for (const para of c.rightsShared.paragraphs) parts.push(p(para));

    parts.push(`                        <h3>${c.tips.titleHelp}</h3>`);
    parts.push(ulHelp(c.tips.list));

    parts.push(`                        <div class="ok">`);
    parts.push(`                            ${c.summary.paragraphs[0]}`);
    parts.push(`                        </div>`);
    parts.push(`                        <p class="help-open" style="margin-top:14px;">`);
    parts.push(
        `                            <a href="https://ms365.schule/datenschutz.html" target="_blank" rel="noopener noreferrer">→ Vollständige Datenschutzerklärung auf ms365.schule</a>`
    );
    parts.push(`                            ·`);
    parts.push(
        `                            <a href="https://ms365.schule/impressum.html" target="_blank" rel="noopener noreferrer">Impressum</a>`
    );
    parts.push(`                            ·`);
    parts.push(
        `                            <a href="https://ms365.schule/nutzungsbedingungen.html" target="_blank" rel="noopener noreferrer">Nutzungsbedingungen</a>`
    );
    parts.push(`                        </p>`);

    return parts.join('\n');
}

function buildLandingCore() {
    const c = PRIVACY_COPY.core;
    const parts = [];

    // Abschnitte 2–3 und 5–12 (1=Controller, 4=Website bleiben landingspezifisch außerhalb)
    parts.push(`      <section class="legal-block" aria-labelledby="school-controller">`);
    parts.push(`        <h2 id="school-controller">2. ${c.schoolResponsibility.titleLanding}</h2>`);
    for (const para of c.schoolResponsibility.paragraphs) parts.push(pLanding(para));
    parts.push(`      </section>`);
    parts.push('');

    parts.push(`      <section class="legal-block" aria-labelledby="no-backend">`);
    parts.push(`        <h2 id="no-backend">3. ${c.noProductDb.titleLanding}</h2>`);
    for (const para of c.noProductDb.paragraphs) parts.push(pLanding(para));
    parts.push(`      </section>`);
    parts.push('');

    parts.push(`      <!-- privacy-copy:landing-website bleibt manuell (Abschnitt 4) -->`);

    return parts.join('\n');
}

function buildLandingCoreRest() {
    const c = PRIVACY_COPY.core;
    const parts = [];

    parts.push(`      <section class="legal-block" aria-labelledby="local-data">`);
    parts.push(`        <h2 id="local-data">5. ${c.localData.titleLanding}</h2>`);
    parts.push(pLanding(c.localData.lead));
    parts.push(ulLanding(c.localData.list));
    parts.push(pLanding(c.localData.after));
    parts.push(`      </section>`);
    parts.push('');

    parts.push(`      <section class="legal-block" aria-labelledby="backup">`);
    parts.push(`        <h2 id="backup">6. ${c.backup.titleLanding}</h2>`);
    for (const para of c.backup.paragraphs) parts.push(pLanding(para));
    parts.push(`      </section>`);
    parts.push('');

    parts.push(`      <section class="legal-block" aria-labelledby="microsoft">`);
    parts.push(`        <h2 id="microsoft">7. ${c.microsoft.titleLanding}</h2>`);
    for (const para of c.microsoft.paragraphs) parts.push(pLanding(para));
    parts.push(`        <p class="legal-note">`);
    parts.push(
        `          Weitere Informationen zu Microsoft-Diensten:`
    );
    parts.push(
        `          <a href="https://privacy.microsoft.com/de-de/privacystatement" target="_blank" rel="noopener noreferrer">Microsoft-Datenschutzbestimmungen</a>.`
    );
    parts.push(`        </p>`);
    parts.push(`      </section>`);
    parts.push('');

    parts.push(`      <section class="legal-block" aria-labelledby="license">`);
    parts.push(`        <h2 id="license">8. ${c.operatorServices.titleLanding}</h2>`);
    parts.push(pLanding(c.operatorServices.lead));
    parts.push(ulLanding(c.operatorServices.list));
    parts.push(pLanding(c.operatorServices.after));
    parts.push(`      </section>`);
    parts.push('');

    parts.push(`      <section class="legal-block" aria-labelledby="retention">`);
    parts.push(`        <h2 id="retention">9. ${c.retention.titleLanding}</h2>`);
    parts.push(ulLanding(c.retention.list));
    parts.push(`      </section>`);
    parts.push('');

    parts.push(`      <section class="legal-block" aria-labelledby="rights">`);
    parts.push(`        <h2 id="rights">10. ${c.rightsShared.titleLanding}</h2>`);
    for (const para of c.rightsShared.paragraphs) parts.push(pLanding(para));
    parts.push(`        <p>`);
    parts.push(
        `          Gegenüber dem Betreiber dieser Website/App können Sie sich in Bezug auf die unter`
    );
    parts.push(
        `          Abschnitt&nbsp;1 genannten Verarbeitungen an`
    );
    parts.push(
        `          <a href="mailto:kontakt@ms365.schule">kontakt@ms365.schule</a>`
    );
    parts.push(
        `          wenden. Sie können sich außerdem bei der Österreichischen Datenschutzbehörde beschweren:`
    );
    parts.push(
        `          <a href="https://www.dsb.gv.at/" target="_blank" rel="noopener noreferrer">www.dsb.gv.at</a>.`
    );
    parts.push(`        </p>`);
    parts.push(`      </section>`);
    parts.push('');

    parts.push(`      <section class="legal-block" aria-labelledby="tips">`);
    parts.push(`        <h2 id="tips">11. ${c.tips.titleLanding}</h2>`);
    parts.push(ulLanding(c.tips.list));
    parts.push(`      </section>`);
    parts.push('');

    parts.push(`      <section class="legal-block" aria-labelledby="summary">`);
    parts.push(`        <h2 id="summary">12. ${c.summary.titleLanding}</h2>`);
    for (const para of c.summary.paragraphs) parts.push(pLanding(para));
    parts.push(`      </section>`);

    return parts.join('\n');
}

function buildLandingTeaser() {
    const t = PRIVACY_COPY.landingTeaser;
    const points = t.points
        .map(
            (pt, i) =>
                `          <li class="reveal" style="--d:${i}">${pt}</li>`
        )
        .join('\n');
    return [
        `        <p class="eyebrow">${t.eyebrow}</p>`,
        `        <h2 id="privacy-title">${t.title}</h2>`,
        `        <p class="section-lead">${t.lead}</p>`,
        `        <ul class="privacy-points">`,
        points,
        `        </ul>`,
        `        <p class="privacy-more">`,
        `          <a href="datenschutz.html">Datenschutzerklärung</a>`,
        `          ·`,
        `          <a href="https://app.ms365.schule/hilfe.html#datenschutz">App-Hilfe · Datenschutz</a>`,
        `        </p>`
    ].join('\n');
}

function patchReadme(md) {
    const start = '<!-- privacy-copy:readme:start -->';
    const end = '<!-- privacy-copy:readme:end -->';
    const body = PRIVACY_COPY.readmeBlurb;
    if (md.includes(start) && md.includes(end)) {
        return replaceMarked(md, 'readme', body);
    }
    // Ersten Datenschutz-Absatz ersetzen, wenn noch keine Marker
    const needle =
        /Es werden \*\*keine Stammdaten an einen App-Server\*\*[\s\S]*?hilfe\.html#tool-stammdaten-uebergabe\)\./;
    if (needle.test(md)) {
        return md.replace(
            needle,
            `${start}\n${body}\n${end}`
        );
    }
    throw new Error('README: Datenschutz-Absatz nicht gefunden');
}

function main() {
    const helpPath = path.join(root, 'hilfe.html');
    let help = fs.readFileSync(helpPath, 'utf8');

    // Marker einsetzen, falls noch nicht vorhanden (einmalig beim Umbau)
    if (!help.includes('<!-- privacy-copy:help-core:start -->')) {
        const re =
            /(<article class="help-article"[^>]*data-search="datenschutz[^"]*"[^>]*>)([\s\S]*?)(<\/article>\s*<\/section>\s*<\/main>)/;
        if (!re.test(help)) throw new Error('hilfe.html: Datenschutz-Artikel nicht gefunden');
        help = help.replace(re, (_, open, _old, close) => {
            return (
                open +
                '\n                        <!-- privacy-copy:help-core:start -->\n' +
                'PLACEHOLDER\n' +
                '                        <!-- privacy-copy:help-core:end -->\n' +
                '                    ' +
                close
            );
        });
    }

    if (!help.includes('<!-- privacy-copy:help-schnellstart-lead:start -->')) {
        const re =
            /(<h2 id="schnellstart">Schnellstart<\/h2>\s*)<p class="help-lead">[\s\S]*?<\/p>/;
        if (!re.test(help)) throw new Error('hilfe.html: Schnellstart-Lead nicht gefunden');
        help = help.replace(
            re,
            `$1<!-- privacy-copy:help-schnellstart-lead:start -->\n                        <p class="help-lead">PLACEHOLDER</p>\n                        <!-- privacy-copy:help-schnellstart-lead:end -->`
        );
    }

    if (!help.includes('<!-- privacy-copy:help-datenhaltung:start -->')) {
        const re =
            /(<div class="hint" id="datenspeicher">)([\s\S]*?)(<\/div>\s*<div class="warn">)/;
        if (!re.test(help)) throw new Error('hilfe.html: #datenspeicher nicht gefunden');
        help = help.replace(
            re,
            `$1\n                            <!-- privacy-copy:help-datenhaltung:start -->\n                            PLACEHOLDER\n                            <!-- privacy-copy:help-datenhaltung:end -->\n                        $3`
        );
    }

    if (!help.includes('<!-- privacy-copy:help-glossary-stammdaten:start -->')) {
        const re = /(<dt>Stammdaten<\/dt>\s*<dd>)([\s\S]*?)(<\/dd>)/;
        if (!re.test(help)) throw new Error('hilfe.html: Glossar Stammdaten nicht gefunden');
        help = help.replace(
            re,
            `$1<!-- privacy-copy:help-glossary-stammdaten:start -->PLACEHOLDER<!-- privacy-copy:help-glossary-stammdaten:end -->$3`
        );
    }

    // data-search um Betreiber-Dienste erweitern
    help = help.replace(
        /data-search="datenschutz[^"]*"/,
        'data-search="datenschutz dsgvo localstorage msal graph microsoft betreiber lizenz-api minimalprinzip schul-tenant sharepoint sicherungsbibliothek backup it-bibliothek keine produkt-datenbank"'
    );

    help = replaceMarked(
        help,
        'help-schnellstart-lead',
        `                        <p class="help-lead">${PRIVACY_COPY.helpSchnellstartLead}</p>`,
        '                        '
    );
    help = replaceMarked(
        help,
        'help-datenhaltung',
        `                            ${PRIVACY_COPY.helpDatenhaltungHint}`,
        '                            '
    );
    help = replaceMarked(
        help,
        'help-glossary-stammdaten',
        PRIVACY_COPY.helpGlossaryStammdaten,
        ''
    );
    help = replaceMarked(help, 'help-core', buildHelpCore(), '                        ');
    fs.writeFileSync(helpPath, help, 'utf8');

    // Landing datenschutz.html
    const landPath = path.join(root, 'landing', 'datenschutz.html');
    let land = fs.readFileSync(landPath, 'utf8');

    land = land.replace(
        /<meta name="description" content="[^"]*">/,
        `<meta name="description" content="${PRIVACY_COPY.metaDescription}">`
    );

    if (!land.includes('<!-- privacy-copy:landing-core:start -->')) {
        const re =
            /(<section class="legal-block" aria-labelledby="school-controller">)([\s\S]*?)(<section class="legal-block" aria-labelledby="website">)/;
        if (!re.test(land)) throw new Error('landing/datenschutz.html: Abschnitte 2–3 nicht gefunden');
        land = land.replace(
            re,
            `<!-- privacy-copy:landing-core:start -->\nPLACEHOLDER\n      <!-- privacy-copy:landing-core:end -->\n\n      $3`
        );
    }

    if (!land.includes('<!-- privacy-copy:landing-core-rest:start -->')) {
        const re =
            /(<section class="legal-block" aria-labelledby="local-data">)([\s\S]*?)(<p class="legal-back">)/;
        if (!re.test(land)) throw new Error('landing/datenschutz.html: Abschnitte 5–12 nicht gefunden');
        land = land.replace(
            re,
            `<!-- privacy-copy:landing-core-rest:start -->\nPLACEHOLDER\n      <!-- privacy-copy:landing-core-rest:end -->\n\n      $3`
        );
    }

    // Intro / Stand
    if (!land.includes('<!-- privacy-copy:landing-intro:start -->')) {
        const re =
            /(<p class="legal-lead">)([\s\S]*?)(<\/p>\s*<p class="legal-note">[\s\S]*?Stand:[\s\S]*?<\/p>)/;
        if (!re.test(land)) throw new Error('landing/datenschutz.html: Intro nicht gefunden');
        land = land.replace(
            re,
            `<!-- privacy-copy:landing-intro:start -->\nPLACEHOLDER\n      <!-- privacy-copy:landing-intro:end -->`
        );
    }

    const introBody = [
        `      <p class="legal-lead">`,
        `        ${PRIVACY_COPY.core.intro.paragraphs[0]}`,
        `      </p>`,
        `      <p class="legal-note">`,
        `        Stand: ${PRIVACY_COPY.standLabel}. Bei wesentlichen Änderungen der App wird diese Seite angepasst.`,
        `        Gemeinsame Textquelle mit der App-Hilfe: <code>src/shared/privacy-copy.js</code>.`,
        `      </p>`
    ].join('\n');

    land = replaceMarked(land, 'landing-intro', introBody, '      ');
    land = replaceMarked(land, 'landing-core', buildLandingCore(), '      ');
    land = replaceMarked(land, 'landing-core-rest', buildLandingCoreRest(), '      ');
    fs.writeFileSync(landPath, land, 'utf8');

    // Landing index teaser
    const idxPath = path.join(root, 'landing', 'index.html');
    let idx = fs.readFileSync(idxPath, 'utf8');
    if (!idx.includes('<!-- privacy-copy:landing-teaser:start -->')) {
        const re =
            /(<section class="section section-privacy"[^>]*>\s*<div class="section-inner narrow reveal">)([\s\S]*?)(<\/div>\s*<\/section>\s*\n\s*<section class="section section-bookings")/;
        if (!re.test(idx)) throw new Error('landing/index.html: Privacy-Sektion nicht gefunden');
        idx = idx.replace(
            re,
            `$1\n        <!-- privacy-copy:landing-teaser:start -->\nPLACEHOLDER\n        <!-- privacy-copy:landing-teaser:end -->\n      $3`
        );
    }
    idx = replaceMarked(idx, 'landing-teaser', buildLandingTeaser(), '        ');
    fs.writeFileSync(idxPath, idx, 'utf8');

    // README
    const readmePath = path.join(root, 'README.md');
    let readme = fs.readFileSync(readmePath, 'utf8');
    readme = patchReadme(readme);
    fs.writeFileSync(readmePath, readme, 'utf8');

    console.log('synced privacy-copy → hilfe.html, landing/datenschutz.html, landing/index.html, README.md');
}

main();
