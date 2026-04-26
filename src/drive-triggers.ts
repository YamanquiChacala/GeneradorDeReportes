import { FileType } from "./common/enums";
import { buildWrongSelectionCard } from "./common/premade-cards";
import { getFileType } from "./common/utils/file-validation";
import { buildCreateSetupFileCard } from "./setup/cards";

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
        return buildCreateSetupFileCard(selectedFolder.getId());
    }

    return buildWrongSelectionCard();
}
