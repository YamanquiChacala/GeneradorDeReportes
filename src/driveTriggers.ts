import { FileType } from "./common/enums";
import { getFileType } from "./common/fileValidation";
import { buildWrongSelectionCard } from "./common/premadeCards";

/**
 *
 */
export function buildDriveCard(e: GoogleAppsScript.Addons.EventObject): GoogleAppsScript.Card_Service.Card {
    const selectedFile = e.drive?.selectedItems?.[0];
    if (selectedFile) {
        const selectedFileFileType = getFileType(selectedFile.id);
        switch (selectedFileFileType) {
            case FileType.SETUP:
            // TODO: return Card "Ready to initialize?"
            case FileType.REPORT:
            // TODO: return Card "Open in Sheets to edit or generate all reports"
        }

        let selectedFolder: GoogleAppsScript.Drive.Folder;
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
