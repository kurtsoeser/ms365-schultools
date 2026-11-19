/*
 * Seiten-Skript für tools/abgleichen.html (klassisches Skript, kein Modul).
 * Stand früher inline in der Seite; ausgelagert, damit die
 * Content-Security-Policy ohne 'unsafe-inline' auskommt.
 */
// Eigenes Modul "Abgleichen": wir nutzen die bestehende Oberfläche/Logik,
// starten aber gezielt im Abgleich-Modus.
try { sessionStorage.setItem('ms365-schulstruktur-sync-ui-mode-v1', 'match'); } catch { /* ignore */ }
try { window.location.replace('schulstruktur-sync.html?mode=match'); } catch { /* ignore */ }
