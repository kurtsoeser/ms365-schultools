/*
 * Seiten-Skript für index.html (klassisches Skript, kein Modul).
 * Stand früher inline in der Seite; ausgelagert, damit die
 * Content-Security-Policy ohne 'unsafe-inline' auskommt.
 */
(function(){try{var d=document.documentElement;var t=localStorage.getItem('ms365-theme-v1');if(t!=='dark'&&t!=='light'){t='light';}d.setAttribute('data-theme',t);d.style.colorScheme=t;var b=localStorage.getItem('ms365-brand-v1');if(b!=='classic'&&b!=='teal'&&b!=='wine'){b='teal';}d.setAttribute('data-brand',b);var sec=localStorage.getItem('ms365-dash-main-section-v1');if(sec!=='tasks'&&sec!=='catalog'&&sec!=='playbooks'){sec=localStorage.getItem('ms365-dash-it-focus-v1')==='0'?'catalog':'tasks';}d.setAttribute('data-dash-main-section',sec);/* Portal-first: kein IT-Flash vor Persona-Resolve */d.setAttribute('data-dash-app-only','');}catch(e){}})();
