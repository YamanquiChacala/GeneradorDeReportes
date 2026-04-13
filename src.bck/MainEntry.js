
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
            case FileType.setup:
            // TODO: return Card "Ready to initialize?"
            case FileType.report:
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
        return InitializationCards.buildCreateInitializationFileCard(selectedFolder.getId())
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

    if (!sheetId) return buildWrongSelectionCard();

    const sheetFileType = FileValidation.getFileType(sheetId);

    switch (sheetFileType) {
        case FileType.setup:
            return InitializationCards.buildInitializationFileEditCard(sheetId)
        case FileType.report:
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
        .setStartIcon(CardParts.icon({ iconName: Icon.folder_question, color: "orange", height: 64 }))
        .setText("Selección no válida");

    const explanation = CardService.newTextParagraph()
        .setText("El archivo o carpeta seleccionado no es compatible.<br/><br/>Por favor, selecciona una de las siguientes opciones para continuar:<br/><br/>• 📁 Una <b>carpeta vacía</b> (para crear la configuración inicial)<br/><br/>• 📋 Un archivo de <b>Registro inicial de grupos</b><br/><br/>• 📊 Un archivo de <b>Calificaciones, asistencias y reportes</b>");

    const mainSection = CardService.newCardSection()
        .addWidget(warning)
        .addWidget(explanation);

    return CardService.newCardBuilder()
        .setHeader(CardParts.headerImage({ title: "Montessori Chacala", subtitle: "Archivo no reconocido", image: "school" }))
        .addSection(mainSection)
        .build();
}




