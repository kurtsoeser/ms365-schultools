/**
 * Area editor orchestration for Schul-Organigramm.
 */
import { createAreaCreateUi } from './schul-organigramm-edit-area-create-ui.js';
import { mountAreaEditorShell } from './schul-organigramm-edit-area-shell-ui.js';
import { wireAreaEditorM365 } from './schul-organigramm-edit-area-m365-ui.js';
import { wireAreaEditorPeople } from './schul-organigramm-edit-area-people-ui.js';

/**
 * @param {Record<string, any>} ctx
 */
export function createAreaUi(ctx) {
    const created = createAreaCreateUi(ctx);
    const createAudienceSection = created.createAudienceSection;
    const createAreaAt = created.createAreaAt;
    const softToast = created.softToast;
    const promptFreeAreaName = created.promptFreeAreaName;
    const persistOrganigrammLayout = created.persistOrganigrammLayout;
    const defaultCanvasSlotForNewArea = created.defaultCanvasSlotForNewArea;
    const removeAudienceSection = created.removeAudienceSection;
    const removeArea = created.removeArea;

    // make layout helpers available to shell/m365 via ctx
    ctx.persistOrganigrammLayout = persistOrganigrammLayout;

    /**
     * @param {string|object} areaIdOrNode
     */
    function openAreaEditor(areaIdOrNode) {
        const ae = mountAreaEditorShell(ctx, areaIdOrNode, { removeArea: removeArea });
        if (!ae) return;
        wireAreaEditorM365(ae);
        wireAreaEditorPeople(ae);
    }

    return {
        createAudienceSection: createAudienceSection,
        createAreaAt: createAreaAt,
        softToast: softToast,
        promptFreeAreaName: promptFreeAreaName,
        persistOrganigrammLayout: persistOrganigrammLayout,
        defaultCanvasSlotForNewArea: defaultCanvasSlotForNewArea,
        removeAudienceSection: removeAudienceSection,
        removeArea: removeArea,
        openAreaEditor: openAreaEditor
    };
}
