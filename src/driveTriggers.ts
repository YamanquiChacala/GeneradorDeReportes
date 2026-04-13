import { getFileType } from "./common/fileValidation";
import { FileType } from "./common/enums";

/**
 * 
 */
export function buildDriveCard(e: GoogleAppsScript.Addons.EventObject): GoogleAppsScript.Card_Service.Card {

    if (e.drive?.selectedItems?.length) {
        const selectedFile = e.drive.selectedItems[0];
        const selectedFileFileType = getFileType(selectedFile.id);
        switch (selectedFileFileType) {
            case FileType.SETUP:
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
        // TODO: return Card
        // return InitializationCards.buildCreateInitializationFileCard(selectedFolder.getId())
    }

    return buildWrongSelectionCard();
}


/**
 * A card indicating the user has selected an invalid file or folder.
 * Reminds them of the specific files/folders the add-on can manage.
 */
function buildWrongSelectionCard(): GoogleAppsScript.Card_Service.Card {

    const explanation = CardService.newTextParagraph()
        .setText("El archivo o carpeta seleccionado no es compatible.<br/><br/>Por favor, selecciona una de las siguientes opciones para continuar:<br/><br/>• 📁 Una <b>carpeta vacía</b> (para crear la configuración inicial)<br/><br/>• 📋 Un archivo de <b>Registro inicial de grupos</b><br/><br/>• 📊 Un archivo de <b>Calificaciones, asistencias y reportes</b>");

    const mainSection = CardService.newCardSection()
        .addWidget(explanation);

    return CardService.newCardBuilder()
        .addSection(mainSection)
        .build();
}