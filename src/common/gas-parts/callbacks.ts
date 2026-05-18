/**
 * Callback to request authorization for the current file.
 */
export function onAskPermission(): GoogleAppsScript.Card_Service.EditorFileScopeActionResponse {
    return CardService.newEditorFileScopeActionResponseBuilder().requestFileScopeForActiveDocument().build();
}

/**
 * Pops the Card Stack to the root.
 */
export function onPopCardStack(): GoogleAppsScript.Card_Service.ActionResponse {
    return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().popToRoot()).build();
}
