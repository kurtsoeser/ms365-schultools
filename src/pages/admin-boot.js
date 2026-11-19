/*
 * Seiten-Skript für admin.html (klassisches Skript, kein Modul).
 * Stand früher inline in der Seite; ausgelagert, damit die
 * Content-Security-Policy ohne 'unsafe-inline' auskommt.
 */
(function(){try{var d=document.documentElement;var t=localStorage.getItem('ms365-theme-v1');if(t!=='dark'&&t!=='light'){t='light';}d.setAttribute('data-theme',t);d.style.colorScheme=t;var b=localStorage.getItem('ms365-brand-v1');if(b!=='classic'&&b!=='teal'&&b!=='wine'){b='teal';}d.setAttribute('data-brand',b);d.removeAttribute('data-ms365-admin-boot');}catch(e){}})();
