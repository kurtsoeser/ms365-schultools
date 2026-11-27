/**
 * Eigener Outlook-Kalender am Sprechtag. Kein Bookings.Read.All.
 */

import { fetchAllPages, getGraphToken } from '../../shared/graph-client.js';
import { dayBounds } from './elternsprechtag-logic.js';

const SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Calendars.ReadWrite'
];

const SELECT = [
    'id',
    'subject',
    'start',
    'end',
    'isAllDay',
    'isCancelled',
    'location',
    'bodyPreview',
    'webLink',
    'organizer',
    'attendees',
    'isOnlineMeeting',
    'onlineMeeting',
    'categories'
].join(',');

/**
 * @param {string} isoDate
 */
export async function loadMyDayEvents(isoDate) {
    const bounds = dayBounds(isoDate);
    if (!bounds) throw new Error('Bitte ein Datum wählen.');
    const start = encodeURIComponent(bounds.start.toISOString());
    const end = encodeURIComponent(bounds.end.toISOString());
    const path =
        '/me/calendarView?startDateTime=' +
        start +
        '&endDateTime=' +
        end +
        '&$select=' +
        SELECT +
        '&$orderby=start/dateTime&$top=100';
    const token = await getGraphToken(SCOPES);
    const { items } = await fetchAllPages(token, path, {
        maxItems: 200,
        maxPages: 5,
        extraHeaders: { Prefer: 'outlook.timezone="Europe/Vienna"' }
    });
    return items || [];
}
