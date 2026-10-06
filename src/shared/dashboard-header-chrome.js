/**
 * Dashboard: Suche in den Sticky-Header (Zone C).
 */
export function placeDashboardSearchInHeader() {
    const slot = document.getElementById('dashHeaderSearchSlot');
    const search = document.querySelector('.dashboard-search--tasks');
    if (!slot || !search || slot.contains(search)) return;
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
