/**
 * Area editor Microsoft 365 wiring entry.
 */
import { wireAreaEditorM365Cards } from './schul-organigramm-edit-area-m365-cards-ui.js';
import { wireAreaEditorM365People } from './schul-organigramm-edit-area-m365-people-ui.js';
import { wireAreaEditorM365Listeners } from './schul-organigramm-edit-area-m365-listeners-ui.js';

/**
 * @param {Record<string, any>} ae
 */
export function wireAreaEditorM365(ae) {
    wireAreaEditorM365Cards(ae);
    wireAreaEditorM365People(ae);
    wireAreaEditorM365Listeners(ae);
}
