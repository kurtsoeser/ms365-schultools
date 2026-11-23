(() => {
  const meta = document.querySelector('meta[name="ms365-landing-build"]');
  const line = document.getElementById('trust-build-line');
  if (!meta || !line) return;
  const raw = String(meta.getAttribute('content') || '').trim();
  if (!raw) return;
  const parts = raw.split('|');
  const iso = parts[0] || '';
  const screen = parts[1] || '';
  const sha = parts[2] || '';
  let when = iso;
  try {
    if (iso) {
      when = new Date(iso).toLocaleString('de-AT', {
        dateStyle: 'medium',
        timeStyle: 'short'
      });
    }
  } catch {
    /* ignore */
  }
  line.textContent =
    'Build ' +
    (sha || 'unbekannt') +
    (screen ? ' · Screens ' + screen : '') +
    (when ? ' · ' + when : '') +
    ' (Meta ms365-landing-build).';
})();
