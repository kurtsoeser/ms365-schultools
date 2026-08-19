import { fixToolsAnchors } from './app-paths.js';

function boot() {
    fixToolsAnchors(document);
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', boot);
} else {
    boot();
}

export { boot };
