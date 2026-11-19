/*
 * Ersatz für Inline-Handler (onclick="…"), die eine Content-Security-Policy
 * ohne 'unsafe-inline' blockiert.
 *
 *   <button data-ms365-call="goToStep" data-ms365-arg="2">   ruft window.goToStep(2)
 *   <button data-ms365-call="resetApp">                      ruft window.resetApp()
 *   <form data-ms365-no-submit>                              unterdrückt das Absenden
 *
 * Aufrufbar sind nur Funktionsnamen, die beim Laden im statischen HTML der Seite
 * stehen. Nachträglich eingefügtes Markup (z. B. über innerHTML) kann damit keine
 * weiteren globalen Funktionen auslösen.
 *
 * Klassisches Skript, mit defer einbinden.
 */
(function () {
    'use strict';

    if (window.__ms365DeclarativeActions) return;
    window.__ms365DeclarativeActions = true;

    var allowed = Object.create(null);

    function collectAllowed() {
        var nodes = document.querySelectorAll('[data-ms365-call]');
        for (var i = 0; i < nodes.length; i++) {
            var name = nodes[i].getAttribute('data-ms365-call');
            if (name) allowed[name] = true;
        }
    }

    function onClick(ev) {
        var target = ev.target;
        var el = target && typeof target.closest === 'function' ? target.closest('[data-ms365-call]') : null;
        if (!el || el.disabled) return;
        var name = el.getAttribute('data-ms365-call');
        if (!name || allowed[name] !== true) return;
        var fn = window[name];
        if (typeof fn !== 'function') {
            console.warn('declarative-actions: window.' + name + ' ist (noch) nicht verfügbar.');
            return;
        }
        var raw = el.getAttribute('data-ms365-arg');
        if (raw === null) {
            fn();
            return;
        }
        var arg;
        try {
            arg = JSON.parse(raw);
        } catch (e) {
            arg = raw;
        }
        fn(arg);
    }

    function onSubmit(ev) {
        var form = ev.target;
        if (form && typeof form.hasAttribute === 'function' && form.hasAttribute('data-ms365-no-submit')) {
            ev.preventDefault();
        }
    }

    function init() {
        collectAllowed();
        // Capture-Phase: läuft wie ein Inline-Handler vor anderen Listenern und
        // unabhängig davon, ob ein Vorfahre die Weitergabe stoppt.
        document.addEventListener('click', onClick, true);
        document.addEventListener('submit', onSubmit, true);
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
})();
