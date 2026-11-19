/*
 * Seiten-Skript für tools/freistellung-planer.html (klassisches Skript, kein Modul).
 * Stand früher inline in der Seite; ausgelagert, damit die
 * Content-Security-Policy ohne 'unsafe-inline' auskommt.
 */
// Schüler: nur User.Read beim Login – keine IT-Admin-Genehmigung (Graph-IT-Scopes kommen ggf. still nach).
window.MS365_AUTH_LOGIN_SCOPES = ['https://graph.microsoft.com/User.Read'];
