/**
 * Area M365 listeners entry.
 */
import { wireAreaEditorM365Group } from './schul-organigramm-edit-area-m365-group-ui.js';
import { wireAreaEditorM365Mail } from './schul-organigramm-edit-area-m365-mail-ui.js';

/**
 * @param {Record<string, any>} ae
 */
export function wireAreaEditorM365Listeners(ae) {
    wireAreaEditorM365Group(ae);
    wireAreaEditorM365Mail(ae);
}
