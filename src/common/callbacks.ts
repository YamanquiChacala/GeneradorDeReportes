/**
 * Pops the Card Stack to the root.
 */
export function onPopCardStack(): GoogleAppsScript.Card_Service.ActionResponse {
    return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().popToRoot()).build();
}
