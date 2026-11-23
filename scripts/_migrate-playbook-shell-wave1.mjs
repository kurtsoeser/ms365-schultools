/**
 * Wave 1: Strip inline cp-* styles, inject Fluent/playbook-shell CSS,
 * add pb-page / pb-step classes, flatten tm-panel where easy.
 */
import fs from 'node:fs';
import path from 'node:path';

const root = path.resolve('tools');
const files = [
    'playbook-daten-import-verknuepfen.html',
    'playbook-schuljahresstart.html',
    'playbook-intranet.html',
    'playbook-schularbeiten-planer.html',
    'playbook-freistellungen.html',
    'playbook-kursteams.html',
    'playbook-eltern.html',
    'playbook-elternsprechtag.html',
    'cleanup-playbook.html',
    'playbook-schulaktivitaeten.html',
    'playbook-projektwochen.html',
    'playbook-einsammler.html',
    'playbook-elternbrief.html',
    'playbook-krankmeldung.html',
    'playbook-srdp.html',
    'playbook-wahlen.html'
];

const SHELL_LINKS = `    <link rel="stylesheet" href="../src/shared/design-tokens.css?v=2">
    <link rel="stylesheet" href="../src/shared/shell-fluent.css?v=10">
    <link rel="stylesheet" href="../src/shared/fluent-blocks.css?v=8">
    <link rel="stylesheet" href="../src/shared/tool-modern.css">
    <link rel="stylesheet" href="../src/shared/playbook-shell.css?v=1">
    <link rel="stylesheet" href="../vendor/bootstrap-icons/bootstrap-icons.min.css" vite-ignore>`;

function stripInlineStyleBlock(html) {
    return html.replace(/<style>[\s\S]*?<\/style>\s*/i, '');
}

function ensureShellLinks(html) {
    if (html.includes('playbook-shell.css')) {
        return html;
    }
    // Replace legacy app.css + tool-modern + bootstrap trio if present
    const legacy = /<link rel="stylesheet" href="\.\.\/app\.css[^>]*>\s*<link rel="stylesheet" href="\.\.\/src\/shared\/tool-modern\.css"[^>]*>\s*<link rel="stylesheet" href="\.\.\/vendor\/bootstrap-icons\/bootstrap-icons\.min\.css"[^>]*>/i;
    if (legacy.test(html)) {
        return html.replace(
            legacy,
            `<link rel="stylesheet" href="../app.css?v=14">\n${SHELL_LINKS}`
        );
    }
    // After title or first stylesheet
    if (html.includes('app.css')) {
        return html.replace(
            /(<link rel="stylesheet" href="\.\.\/app\.css[^>]*>)/i,
            `$1\n${SHELL_LINKS}`
        );
    }
    return html.replace('</head>', `${SHELL_LINKS}\n</head>`);
}

function upgradeBodyAndContainer(html) {
    html = html.replace(/<body([^>]*)>/i, (m, attrs) => {
        if (/app-shell-chrome/.test(attrs)) return m;
        const a = String(attrs || '');
        if (/class=/.test(a)) {
            return `<body${a.replace(/class="([^"]*)"/, 'class="$1 app-shell-chrome"')}>`;
        }
        return `<body class="app-shell-chrome"${a}>`;
    });
    html = html.replace(
        /class="container page-card tm-page"/g,
        'class="container page-card tm-page pb-page"'
    );
    html = html.replace(
        /class="container page-card"/g,
        'class="container page-card tm-page pb-page"'
    );
    return html;
}

function upgradeStepMarkup(html) {
    html = html.replace(/class="cp-phases"/g, 'class="pb-phases"');
    html = html.replace(/class="cp-phases-lead"/g, 'class="pb-lead"');
    html = html.replace(/class="cp-phase-head"/g, 'class="pb-phase__head"');
    html = html.replace(/class="cp-steps"/g, 'class="pb-steps"');
    html = html.replace(/class="cp-steps" id="/g, 'class="pb-steps" id="');

    // articles
    html = html.replace(/<article class="cp-step([^"]*)"/g, '<article class="pb-step cp-step$1"');
    html = html.replace(/class="cp-step__n"/g, 'class="pb-step__n"');
    html = html.replace(/class="cp-step__actions"/g, 'class="pb-step__actions"');
    html = html.replace(/class="cp-check"/g, 'class="pb-check"');
    html = html.replace(/class="cp-pill /g, 'class="pb-pill ');
    html = html.replace(/class="cp-pill"/g, 'class="pb-pill"');
    html = html.replace(/cp-pill--/g, 'pb-pill--');
    html = html.replace(/class="cp-step-gate"/g, 'class="pb-step__gate"');
    html = html.replace(/class="cp-step-ready"/g, 'class="pb-step__ready"');
    html = html.replace(/class="cp-step--optional"/g, 'class="pb-step cp-step pb-step--optional cp-step--optional"');

    // Fix double pb-step if optional already had article class rewrite
    html = html.replace(/class="pb-step cp-step pb-step cp-step /g, 'class="pb-step cp-step ');

    // Flatten tm-actions to pb-actions
    html = html.replace(
        /class="tm-actions" style="margin-top:16px;"/g,
        'class="pb-actions tm-actions"'
    );
    html = html.replace(/class="tm-actions"/g, 'class="pb-actions tm-actions"');

    // Primary CTAs: btn-primary stays; wrap bare single .btn after step body is fine
    // Progress class
    html = html.replace(
        /id="cpProgress" class="muted" style="margin:0;"/g,
        'id="cpProgress" class="pb-progress"'
    );
    html = html.replace(/id="cpProgress" class="muted"/g, 'id="cpProgress" class="pb-progress"');

    return html;
}

function upgradeStubBanner(html, file) {
    if (!/Stub/.test(html) && !file.includes('playbook-schulaktivitaeten') && !/Wave-0-Stub/.test(html)) {
        return html;
    }
    if (html.includes('pb-stub-banner')) return html;
    return html.replace(
        /(<div class="content">\s*<section[^>]*>)/,
        `$1\n                <p class="pb-stub-banner">Wave-0/1-Stub – Inhalt folgt in Wave 3. Shell bereits Fluent.</p>`
    );
}

for (const file of files) {
    const p = path.join(root, file);
    let html = fs.readFileSync(p, 'utf8');
    html = stripInlineStyleBlock(html);
    html = ensureShellLinks(html);
    html = upgradeBodyAndContainer(html);
    html = upgradeStepMarkup(html);
    html = upgradeStubBanner(html, file);
    // Deduplicate shell links if run twice
    const seen = new Set();
    html = html.replace(
        /[ \t]*<link rel="stylesheet" href="\.\.\/src\/shared\/(?:design-tokens|shell-fluent|fluent-blocks|tool-modern|playbook-shell)\.css[^>]*>\n?/g,
        (line) => {
            const key = line.trim();
            if (seen.has(key)) return '';
            seen.add(key);
            return line.startsWith(' ') ? line : '    ' + line.trim() + '\n';
        }
    );
    // Ensure bootstrap icons once after shell
    if ((html.match(/bootstrap-icons\.min\.css/g) || []).length > 1) {
        let first = true;
        html = html.replace(
            /[ \t]*<link rel="stylesheet" href="\.\.\/vendor\/bootstrap-icons\/bootstrap-icons\.min\.css"[^>]*>\n?/g,
            (line) => {
                if (!first) return '';
                first = false;
                return line;
            }
        );
    }
    fs.writeFileSync(p, html);
    console.log('migrated', file);
}
