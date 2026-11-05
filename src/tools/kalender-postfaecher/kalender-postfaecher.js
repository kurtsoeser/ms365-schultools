import { createKalenderPostfaecherUi } from './kalender-postfaecher-ui.js';

function boot() {
    createKalenderPostfaecherUi();
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', boot);
} else {
    boot();
}
