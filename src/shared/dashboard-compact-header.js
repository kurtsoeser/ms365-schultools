/**
 * Dashboard (index): Sticky-Kompakt-Header.
 */

/** @param {number} scrollY */
export function dashboardHeroScrollPhase(scrollY) {
    if (!Number.isFinite(scrollY) || scrollY <= 0) return 'top';
    if (scrollY > 72) return 'scrolled';
    return 'transition';
}

/**
 * @param {number} scrollY
 * @param {{ maxShift?: number, fadeDistance?: number }} [opts]
 */
export function dashboardHeroParallaxStyle(scrollY, opts) {
    const maxShift = (opts && opts.maxShift) || 100;
    const fadeDistance = (opts && opts.fadeDistance) || 200;
    const y = Math.max(0, scrollY);
    const shift = Math.min(y * 0.4, maxShift);
    const opacity = Math.max(0, 1 - y / fadeDistance);
    return {
        transform: 'translate3d(0, -' + shift + 'px, 0)',
        opacity: String(opacity)
    };
}

export function mountDashboardCompactHeader() {
    const root = document.documentElement;
    if (!document.getElementById('dashCompactHeader')) return;

    let frame = 0;
    function apply() {
        frame = 0;
        const y = window.scrollY || document.documentElement.scrollTop || 0;
        root.classList.toggle('dash-header-scrolled', y > 12);
    }

    function onScroll() {
        if (frame) return;
        frame = window.requestAnimationFrame(apply);
    }

    window.addEventListener('scroll', onScroll, { passive: true });
    apply();

    window.addEventListener('ms365-auth-widget-ready', function () {
        const slot = document.getElementById('adminAppTopActions');
        const widget = document.getElementById('ms365AuthWidget');
        if (slot && widget && widget.parentElement === slot) {
            widget.style.position = '';
            widget.style.top = '';
            widget.style.right = '';
            widget.style.zIndex = '';
        }
    });
}

if (typeof document !== 'undefined') {
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', mountDashboardCompactHeader);
    } else {
        mountDashboardCompactHeader();
    }
}
