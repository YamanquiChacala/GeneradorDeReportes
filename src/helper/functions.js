/**
 * Pops the Card Stack to the root.
 * @returns {GoogleAppsScript.Card_Service.ActionResponse}
 */
function onPopCardStack() {
    return CardService.newActionResponseBuilder().setNavigation(CardService.newNavigation().popToRoot()).build();
}