/**
 * Schreibt landing/site-build.json und aktualisiert den Footer auf allen Landing-HTML-Seiten.
 * Usage: node scripts/write-landing-build-info.mjs
 */
import fs from 'node:fs';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(__dirname, '..');
const landingDir = path.join(root, 'landing');

function readScreenCacheV() {
  const indexPath = path.join(landingDir, 'index.html');
  const html = fs.readFileSync(indexPath, 'utf8');
  const m = html.match(/screens\/[a-z0-9-]+\.png\?v=([^"'&\s]+)/i);
  return m ? m[1] : '';
}

function readGitShortSha() {
  try {
    return execFileSync('git', ['rev-parse', '--short', 'HEAD'], {
      cwd: root,
      encoding: 'utf8'
    }).trim();
  } catch {
    return '';
  }
}

function formatDeVienna(iso) {
  return new Intl.DateTimeFormat('de-AT', {
    dateStyle: 'medium',
    timeStyle: 'short',
    timeZone: 'Europe/Vienna'
  }).format(new Date(iso));
}

function buildFooterBlock(info) {
  const label = formatDeVienna(info.publishedAt);
  const screens = info.screensCache
    ? ` <span class="footer-updated__screens">· Screens v${info.screensCache}</span>`
    : '';
  const sha = info.gitSha ? ` <span class="footer-updated__sha">· Build ${info.gitSha}</span>` : '';
  return (
    `<p class="footer-updated" id="landing-last-updated">` +
    `<span class="footer-updated__label">Homepage zuletzt aktualisiert:</span> ` +
    `<time datetime="${info.publishedAt}">${label}</time>${screens}${sha}</p>`
  );
}

const FOOTER_RE =
  /<p class="footer-updated" id="landing-last-updated"[^>]*>[\s\S]*?<\/p>\s*/;

function patchMetaBuild(html, info) {
  const compact = [info.publishedAt, info.screensCache || '', info.gitSha || ''].join('|');
  if (html.includes('name="ms365-landing-build"')) {
    return html.replace(
      /<meta name="ms365-landing-build" content="[^"]*">/,
      `<meta name="ms365-landing-build" content="${compact}">`
    );
  }
  return html.replace(
    /(<meta property="og:url"[^>]*>\s*)/i,
    `$1  <meta name="ms365-landing-build" content="${compact}">\n  `
  );
}

function patchLandingHtml(filePath, footerHtml, info) {
  let html = fs.readFileSync(filePath, 'utf8');
  html = patchMetaBuild(html, info);
  if (FOOTER_RE.test(html)) {
    html = html.replace(FOOTER_RE, footerHtml + '\n      ');
  } else {
    html = html.replace(
      /(<p class="footer-note">)/,
      footerHtml + '\n      $1'
    );
  }
  fs.writeFileSync(filePath, html);
}

function main() {
  const publishedAt = new Date().toISOString();
  const info = {
    publishedAt,
    screensCache: readScreenCacheV(),
    gitSha: readGitShortSha()
  };

  fs.writeFileSync(
    path.join(landingDir, 'site-build.json'),
    JSON.stringify(info, null, 2) + '\n',
    'utf8'
  );

  const footerHtml = buildFooterBlock(info);
  for (const name of fs.readdirSync(landingDir)) {
    if (!name.endsWith('.html')) continue;
    patchLandingHtml(path.join(landingDir, name), footerHtml, info);
  }

  console.log('landing build info', info.publishedAt, 'screens', info.screensCache || '—', info.gitSha || '');
}

main();
