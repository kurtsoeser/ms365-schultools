/**
 * Vollständiger Veröffentlichungs-Vorbereitungslauf (lokal, vor git push).
 *
 *   npm run release:prep
 *   npm run release:prep -- --no-screens    # ohne Playwright-Screenshots
 *   npm run release:prep -- --no-notes      # ohne release-notes.json Sync
 *
 * Push auf main triggert CI + Deploy (FTP app.ms365.schule, GitHub Pages dist).
 */
import { spawn } from 'node:child_process';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(__dirname, '..');
const SHOT_PORT = Number(process.env.MS365_RELEASE_DEV_PORT || 5199);
const SHOT_BASE = `http://localhost:${SHOT_PORT}`;

const args = process.argv.slice(2);
const skipScreens = args.includes('--no-screens');
const skipNotes = args.includes('--no-notes');

function run(cmd, cmdArgs, env = {}) {
  return new Promise((resolve, reject) => {
    const child = spawn(cmd, cmdArgs, {
      cwd: root,
      stdio: 'inherit',
      shell: true,
      env: { ...process.env, ...env }
    });
    child.on('exit', (code) => {
      if (code === 0) resolve();
      else reject(new Error(`${cmd} ${cmdArgs.join(' ')} exited with ${code}`));
    });
  });
}

async function waitForDevServer(ms = 90000) {
  const start = Date.now();
  while (Date.now() - start < ms) {
    try {
      const res = await fetch(`${SHOT_BASE}/index.html`, { method: 'HEAD' });
      if (res.ok) return;
    } catch {
      /* retry */
    }
    await new Promise((r) => setTimeout(r, 400));
  }
  throw new Error(`Dev-Server nicht erreichbar unter ${SHOT_BASE}`);
}

async function captureScreens() {
  let vite;
  try {
    vite = spawn('npm', ['run', 'dev', '--', '--port', String(SHOT_PORT), '--strictPort'], {
      cwd: root,
      stdio: 'ignore',
      shell: true
    });
    await waitForDevServer();
    await run('node', ['scripts/capture-landing-screens.mjs'], {
      MS365_SHOT_BASE: SHOT_BASE
    });
    await run('node', ['scripts/bump-landing-screen-cache.mjs']);
  } finally {
    if (vite && !vite.killed) {
      vite.kill('SIGTERM');
    }
  }
}

async function main() {
  console.log('→ check:secrets');
  await run('npm', ['run', 'check:secrets']);

  if (!skipScreens) {
    console.log('→ landing screenshots (Playwright + Vite auf Port ' + SHOT_PORT + ')');
    await captureScreens();
  } else {
    console.log('→ screenshots übersprungen (--no-screens)');
  }

  if (!skipNotes) {
    console.log('→ release-notes.json aus Commits');
    await run('node', ['scripts/sync-release-notes-from-commits.mjs']).catch((e) => {
      console.warn('release-notes sync übersprungen:', e.message);
    });
  }

  console.log('→ test');
  await run('npm', ['test']);
  console.log('→ lint');
  await run('npm', ['run', 'lint']);
  console.log('→ build');
  await run('npm', ['run', 'build']);

  console.log('');
  console.log('Fertig. Nächste Schritte:');
  console.log('  git add -A && git commit -m "Release: …" && git push origin main');
  console.log('(Push startet CI und Deploy-Workflows auf GitHub.)');
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
