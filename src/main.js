
/**
 * 
 * @param {GoogleAppsScript.Addons.EventObject} e
 * @returns {GoogleAppsScript.Card_Service.Card} 
 */
function buildDriveCard(e) {
    console.log(e);

    const setPropAction = CardService.newAction().setFunctionName(onSetPropClick.name);
    const getPropAction = CardService.newAction().setFunctionName(onGetPropClick.name);

    const setPropButton = CardService.newTextButton()
        .setText("Add Identifier")
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setOnClickAction(setPropAction);
    const getPropButton = CardService.newTextButton()
        .setText("Get Identifier")
        .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
        .setOnClickAction(getPropAction);

    const buttonSection = CardService.newCardSection()
        .addWidget(setPropButton)
        .addWidget(getPropButton);

    const card = CardService.newCardBuilder()
        .setHeader(CardParts.header("File Properties", "", "spiral"))
        .addSection(buttonSection);

    return card.build();
}


/**
 * 
 * @param {GoogleAppsScript.Addons.EventObject} e
 * @returns {GoogleAppsScript.Card_Service.Card} 
 */
function buildSheetsCard(e) {
    console.log(e);

    const card = CardService.newCardBuilder()
        .setHeader(CardParts.header("Sheets Cardd"));

    return card.build();
}


/**
 * 
 * @param {GoogleAppsScript.Addons.EventObject} e 
 * @returns {GoogleAppsScript.Card_Service.ActionResponse}
 */
function onSetPropClick(e) {

    let message = "No item selected";

    if (e.drive?.selectedItems?.length) {
        const itemId = e.drive.selectedItems[0].id;

        FileValidation.setFileType(itemId, FileType.INIT);

        message = "Property Set";
    }

    return CardService.newActionResponseBuilder()
        .setNotification(
            CardService.newNotification()
                .setText(message)
        )
        .build();
}

/**
 * 
 * @param {GoogleAppsScript.Addons.EventObject} e 
 * @returns {GoogleAppsScript.Card_Service.ActionResponse}
 */
function onGetPropClick(e) {

    let message = "No item selected";

    if (e.drive?.selectedItems?.length) {
        const itemId = e.drive.selectedItems[0].id;

        const type = FileValidation.getFileType(itemId);

        message = type ?? "No file type"
    }

    return CardService.newActionResponseBuilder()
        .setNotification(
            CardService.newNotification()
                .setText(message)
        )
        .build();
}

