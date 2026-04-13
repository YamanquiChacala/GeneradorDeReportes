export function buildCard(e: GoogleAppsScript.Addons.EventObject): GoogleAppsScript.Card_Service.Card {
    const action = CardService.newAction().setFunctionName(onClick.name);
    const button = CardService.newTextButton().setText("Press").setOnClickAction(action);
    const section = CardService.newCardSection().addWidget(button);
    return CardService.newCardBuilder().addSection(section).build();
}

export function onClick(e: GoogleAppsScript.Addons.EventObject): GoogleAppsScript.Card_Service.ActionResponse {
    return CardService.newActionResponseBuilder()
        .setNotification(CardService.newNotification()
            .setText("Button clicked"))
        .build();
}