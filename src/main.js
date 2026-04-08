
/**
 * 
 * @param {GoogleAppsScript.Addons.EventObject} e
 * @returns {GoogleAppsScript.Card_Service.Card} 
 */
function buildDriveCard(e) {

    if (e.drive?.selectedItems?.length) {
        const selectedFile = e.drive.selectedItems[0];
        const selectedFileFileType = FileValidation.getFileType(selectedFile.id);
        switch (selectedFileFileType) {
            case FileType.INIT:
            // TODO: return Card "Ready to initialize?"
            case FileType.REPORT:
            // TODO: return Card "Open in Sheets to edit or generate all reports"
        }

        /** @type {GoogleAppsScript.Drive.Folder} */
        let selectedFolder;
        if (selectedFile.mimeType === "application/vnd.google-apps.folder") {
            selectedFolder = DriveApp.getFolderById(selectedFile.id);
        } else {
            const parents = DriveApp.getFileById(selectedFile.id).getParents();
            if (parents.hasNext()) {
                selectedFolder = parents.next();
            } else {
                return buildWrongSelectionCard();
            }
        }
        // TODO: return Card "Create Initialization file here?"
    }

    return buildWrongSelectionCard();
}


/**
 * 
 * @param {GoogleAppsScript.Addons.EventObject} e
 * @returns {GoogleAppsScript.Card_Service.Card} 
 */
function buildSheetsCard(e) {
    if (!e.sheets?.addonHasFileScopePermission) {
        return buildRequestAuthorizationCard()
    }

    const sheetId = e.sheets.id;
    const sheetFileType = FileValidation.getFileType(sheetId);

    switch (sheetFileType) {
        case FileType.INIT:
        // TODO: Return Card "Init options"
        case FileType.REPORT:
        // TODO: Return Card "Report options"
    }

    return buildWrongSelectionCard();
}

/**
 * A card indicating the user has selected an invalid file or folder.
 * Reminds them of the specific files/folders the add-on can manage.
 * 
 * @returns {GoogleAppsScript.Card_Service.Card}
 */
function buildWrongSelectionCard() {
    const warning = CardService.newDecoratedText()
        .setStartIcon(CardParts.icon({ name: Icon.folder_question, color: "orange", height: 64 }))
        .setText("Selección no válida");

    const explanation = CardService.newTextParagraph()
        .setText("El archivo o carpeta seleccionado no es compatible.<br/><br/>Por favor, selecciona una de las siguientes opciones para continuar:<br/><br/>• 📁 Una <b>carpeta vacía</b> (para crear la configuración inicial)<br/><br/>• 📋 Un archivo de <b>Registro inicial de grupos</b><br/><br/>• 📊 Un archivo de <b>Calificaciones, asistencias y reportes</b>");

    const mainSection = CardService.newCardSection()
        .addWidget(warning)
        .addWidget(explanation);

    return CardService.newCardBuilder()
        .setHeader(CardParts.header({ title: "Montessori Chacala", subtitle: "Archivo no reconocido", icon: "school" }))
        .addSection(mainSection)
        .build();
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




