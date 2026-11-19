/*
 * Seiten-Skript für ms365-schooltool.html (klassisches Skript, kein Modul).
 * Stand früher inline in der Seite; ausgelagert, damit die
 * Content-Security-Policy ohne 'unsafe-inline' auskommt.
 */
(async function () {
  // Wichtig: Diese Seite ist die Redirect-URI. Bevor wir weiterleiten,
  // muss MSAL die Redirect-Antwort verarbeiten (handleRedirectPromise),
  // sonst „verliert“ die App die Anmeldung und startet erneut.

  function resolveClientId() {
    let cfg = window.MS365_MSAL_CONFIG || {};
    let id = String(cfg.clientId || '').trim();
    if (!id) {
      const meta = document.querySelector('meta[name="ms365-graph-client-id"]');
      const fromMeta = meta && meta.getAttribute('content') ? meta.getAttribute('content').trim() : '';
      if (fromMeta) id = fromMeta;
    }
    return id;
  }

  async function tryHandleMsalRedirect() {
    const clientId = resolveClientId();
    if (!clientId) return;
    try {
      // Pfad relativ zu dieser Datei (src/pages/), nicht zur HTML-Seite.
      const loader = await import('../shared/msal-loader.js');
      const msalMod = await loader.loadMsalBrowser();
      const PublicClientApplication =
        msalMod.PublicClientApplication || (msalMod.default && msalMod.default.PublicClientApplication);
      if (!PublicClientApplication) return;
      const cfg = window.MS365_MSAL_CONFIG || {};
      const pca = new PublicClientApplication({
        auth: {
          clientId: clientId,
          authority: cfg.authority || 'https://login.microsoftonline.com/organizations',
          redirectUri: window.location.href.split('#')[0]
        },
        // Muss zu msal-auth-ui.js passen, damit der Token nach dem Redirect
        // auf den Tool-Seiten (anderer Tab/Reload möglich) wiedergefunden wird.
        cache: { cacheLocation: 'localStorage', storeAuthStateInCookie: true }
      });
      await pca.initialize();
      await pca.handleRedirectPromise();
    } catch {
      // Wenn MSAL hier scheitert, versuchen wir trotzdem die Navigation.
    }
  }

  await tryHandleMsalRedirect();

  // Nach MSAL Redirect wieder zur ursprünglich gewünschten Tool-Seite zurück.
  try {
    const back = sessionStorage.getItem('ms365-post-login-url');
    if (back && typeof back === 'string' && back.trim()) {
      sessionStorage.removeItem('ms365-post-login-url');
      window.location.replace(back);
      return;
    }
  } catch {
    // ignore
  }

  const params = new URLSearchParams(window.location.search);
  const mode = String(params.get('mode') || '').toLowerCase().trim();

  const map = {
    kursteams: 'tools/kursteams.html',
    kursteam: 'tools/kursteams.html',
    webuntis: 'tools/kursteams.html',
    unterrichtsteams: 'tools/unterrichtsteams-katalog.html',
    'unterrichtsteams-katalog': 'tools/unterrichtsteams-katalog.html',
    'kursteam-templates': 'tools/kursteam-templates.html',
    'kursteam-vorlagen': 'tools/kursteam-templates.html',
    vorlagen: 'tools/kursteam-templates.html',
    jahrgang: 'tools/jahrgangsgruppen.html',
    jahrgangsgruppen: 'tools/jahrgangsgruppen.html',
    'klassen-merge': 'tools/klassen-merge.html',
    klassenmerge: 'tools/klassen-merge.html',
    zusammenlegen: 'tools/klassen-merge.html',
    diplomarbeiten: 'tools/diplomarbeiten.html',
    diplom: 'tools/diplomarbeiten.html',
    dipl: 'tools/diplomarbeiten.html',
    'teams-antrag': 'tools/teams-antrag.html',
    teamsantrag: 'tools/teams-antrag.html',
    spielwiesen: 'tools/spielwiesen.html',
    spielwiese: 'tools/spielwiesen.html',
    demo: 'tools/spielwiesen.html',
    klassenchats: 'tools/klassenchats.html',
    klassenchat: 'tools/klassenchats.html',
    'klassen-chats': 'tools/klassenchats.html',
    'datei-migration': 'tools/datei-migration.html',
    dateimigration: 'tools/datei-migration.html',
    migration: 'tools/datei-migration.html',
    'nextcloud-teams-migrate': 'tools/nextcloud-teams-migrate.html',
    nextcloud: 'tools/nextcloud-teams-migrate.html',
    'nextcloud-migration': 'tools/nextcloud-teams-migrate.html',
    'power-automate': 'tools/power-automate-rezepte.html',
    'power-automate-rezepte': 'tools/power-automate-rezepte.html',
    flow: 'tools/power-automate-rezepte.html',
    arge: 'tools/arge-fachgruppen.html',
    'arge-fachgruppen': 'tools/arge-fachgruppen.html',
    fachgruppen: 'tools/arge-fachgruppen.html',
    fachgruppe: 'tools/arge-fachgruppen.html',
    schulstruktur: 'tools/schulstruktur-sync.html',
    schulstrukturync: 'tools/schulstruktur-sync.html',
    schulstruktursync: 'tools/schulstruktur-sync.html',
    schulstrukturtool: 'tools/schulstruktur-sync.html',
    'schulstruktur-sync': 'tools/schulstruktur-sync.html',
    gruppenerstellung: 'tools/gruppenerstellung.html',
    grouppolicy: 'tools/gruppenerstellung.html',
    'gaeste-verwalten': 'tools/gaeste-verwalten.html',
    gaesteverwalten: 'tools/gaeste-verwalten.html',
    gastverwaltung: 'tools/gaeste-verwalten.html',
    'gast-verwaltung': 'tools/gaeste-verwalten.html',
    'gast-einlader': 'tools/gaeste-verwalten.html#tab=einlader',
    gasteinlader: 'tools/gaeste-verwalten.html#tab=einlader',
    gasteinladen: 'tools/gaeste-verwalten.html#tab=einlader',
    'gaeste-einladen': 'tools/gaeste-verwalten.html#tab=einlader',
    gastinviter: 'tools/gaeste-verwalten.html#tab=einlader',
    guestinviter: 'tools/gaeste-verwalten.html#tab=einlader',
    gruppen: 'tools/schulstruktur-sync.html',
    gruppenuebersicht: 'tools/schulstruktur-sync.html',
    'gruppen-uebersicht': 'tools/schulstruktur-sync.html',
    personen: 'tools/personen-verwaltung.html',
    personenverwaltung: 'tools/personen-verwaltung.html',
    'personen-verwaltung': 'tools/personen-verwaltung.html',
    klassen: 'tenant.html#classes',
    klassenumbenennen: 'tenant.html#classes',
    'klassen-umbenennen': 'tenant.html#classes',
    'schueler-lehrer-gruppen': 'tools/schueler-lehrer-gruppen.html',
    schuelerlehrer: 'tools/schueler-lehrer-gruppen.html',
    slg: 'tools/schueler-lehrer-gruppen.html',
    'schueler-sammelgruppe': 'tools/schueler-sammelgruppe.html',
    schuelersammelgruppe: 'tools/schueler-sammelgruppe.html',
    'lehrer-sammelgruppe': 'tools/lehrer-sammelgruppe.html',
    lehrersammelgruppe: 'tools/lehrer-sammelgruppe.html',
    verwaltung: 'tools/verwaltung.html',
    'verwaltung-gruppe': 'tools/verwaltung.html',
    schulverwaltung: 'tools/verwaltung.html',
    klassenvorstaende: 'tools/klassenvorstaende.html',
    kv: 'tools/klassenvorstaende.html',
    klassenvorstand: 'tools/klassenvorstaende.html',
    teamsarchiv: 'tools/teams-archiv.html',
    'teams-archiv': 'tools/teams-archiv.html',
    leeregruppen: 'tools/leere-gruppen-report.html',
    'leere-gruppen': 'tools/leere-gruppen-report.html',
    leeregruppenreport: 'tools/leere-gruppen-report.html',
    cleanup: 'tools/cleanup-playbook.html',
    'cleanup-playbook': 'tools/cleanup-playbook.html',
    aufräumen: 'tools/cleanup-playbook.html',
    'stammdaten-uebergabe': 'tools/stammdaten-uebergabe.html',
    stammdatenuebergabe: 'tools/stammdaten-uebergabe.html',
    uebergabe: 'tools/stammdaten-uebergabe.html',
    namenskonvention: 'tools/namenskonvention-audit.html',
    'namenskonvention-audit': 'tools/namenskonvention-audit.html',
    naming: 'tools/namenskonvention-audit.html',
    gast: 'tools/gaeste-verwalten.html',
    gaeste: 'tools/gaeste-verwalten.html',
    gastzugaenge: 'tools/gaeste-verwalten.html#tab=teams',
    'gast-zugaenge': 'tools/gaeste-verwalten.html#tab=teams',
    externe: 'tools/gaeste-verwalten.html#tab=teams',
    organisation: 'tools/organisations-assistent.html',
    organisationsassistent: 'tools/organisations-assistent.html',
    'organisations-assistent': 'tools/organisations-assistent.html',
    schuljahr: 'tools/organisations-assistent.html',
    schuljahreswechsel: 'tools/organisations-assistent.html',
    sharepoint: 'tools/sharepoint-mandant-website.html',
    'sharepoint-website': 'tools/sharepoint-mandant-website.html',
    'sharepoint-teilen': 'tools/sharepoint-mandant-teilen.html',
    'sharepoint-intranet': 'tools/sharepoint-intranet-hub.html',
    intranet: 'tools/sharepoint-intranet-hub.html',
    hub: 'tools/sharepoint-intranet-hub.html',
    lehrerliste: 'tools/sharepoint-liste-lehrer.html',
    'sharepoint-lehrer': 'tools/sharepoint-liste-lehrer.html',
    'sharepoint-liste-lehrer': 'tools/sharepoint-liste-lehrer.html',
    'stammdaten-listen': 'tools/sharepoint-liste-stammdaten.html',
    'sharepoint-stammdaten': 'tools/sharepoint-liste-stammdaten.html',
    'sharepoint-liste-stammdaten': 'tools/sharepoint-liste-stammdaten.html',
    schuelerliste: 'tools/sharepoint-liste-stammdaten.html',
    faecherliste: 'tools/sharepoint-liste-stammdaten.html',
    schultermine: 'tools/sharepoint-liste-schultermine.html',
    'sharepoint-schultermine': 'tools/sharepoint-liste-schultermine.html',
    'sharepoint-liste-schultermine': 'tools/sharepoint-liste-schultermine.html',
    schularbeiten: 'tools/schularbeiten-planer.html',
    'schularbeiten-planer': 'tools/schularbeiten-planer.html',
    planer: 'tools/schularbeiten-planer.html',
    'schularbeiten-listen': 'tools/sharepoint-liste-schularbeiten.html',
    'sharepoint-schularbeiten': 'tools/sharepoint-liste-schularbeiten.html',
    'sharepoint-liste-schularbeiten': 'tools/sharepoint-liste-schularbeiten.html',
    projektwochen: 'tools/projektwochen.html',
    'projektwochen-planer': 'tools/projektwochen.html',
    angebote: 'tools/projektwochen.html',
    'projektwochen-listen': 'tools/sharepoint-liste-projektwochen.html',
    'sharepoint-projektwochen': 'tools/sharepoint-liste-projektwochen.html',
    'sharepoint-liste-projektwochen': 'tools/sharepoint-liste-projektwochen.html',
    'pa-projektwochen-mail': 'tools/pa-projektwochen-mail.html',
    'projektwochen-mail': 'tools/pa-projektwochen-mail.html',
    srdp: 'tools/srdp-anmeldung.html',
    'srdp-anmeldung': 'tools/srdp-anmeldung.html',
    'sharepoint-liste-srdp': 'tools/sharepoint-liste-srdp.html',
    'sharepoint-srdp': 'tools/sharepoint-liste-srdp.html',
    'srdp-mail': 'tools/srdp-mail-setup.html',
    'srdp-mail-setup': 'tools/srdp-mail-setup.html',
    schulgraph: 'tools/schulgraph.html',
    'schuldaten-karte': 'tools/schulgraph.html',
    datenlandkarte: 'tools/datenlandkarte.html',
    'daten-landkarte': 'tools/datenlandkarte.html',
    karte: 'tools/schulgraph.html'
  };

  window.location.replace(map[mode] || 'index.html');
})();
