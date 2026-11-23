/**
 * Dashboard: Suche mittig in den Sticky-Header (Admin-Center-Stil).
 */
export function placeDashboardSearchInHeader() {
    const slot = document.getElementById('dashHeaderSearchSlot');
    const search = document.querySelector('.dashboard-search--tasks');
    if (!slot || !search) return;

    let center = slot.closest('.dash-compact-header__center');
    if (!center) {
        const bar = document.querySelector('.dash-compact-header__bar');
        if (bar) {
            center = document.createElement('div');
            center.className = 'dash-compact-header__center';
            const actions = document.getElementById('adminAppTopActions');
            if (actions && actions.parentElement === bar) bar.insertBefore(center, actions);
            else bar.appendChild(center);
            center.appendChild(slot);
        }
    }

    if (slot.contains(search)) {
        search.classList.add('dash-compact-header__search', 'dashboard-tasks-head-search');
        return;
    }
    search.classList.add('dash-compact-header__search', 'dashboard-tasks-head-search');
    slot.appendChild(search);
}

if (typeof document !== 'undefined') {
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', placeDashboardSearchInHeader);
    } else {
        placeDashboardSearchInHeader();
    }
}
