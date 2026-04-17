import { FileType } from "./common/enums";
import { buildRequestAuthorizationCard, buildWrongSelectionCard } from "./common/premadeCards";
import { getFileType } from "./common/utils/fileValidation";
import { buildInitializationFileEditCard } from "./setup/cards";

/**
 *
 */
export function buildSheetsCard(e: GoogleAppsScript.Addons.EventObject): GoogleAppsScript.Card_Service.Card {
    if (!e.sheets?.addonHasFileScopePermission) {
        return buildRequestAuthorizationCard();
    }

    const sheetId = e?.sheets?.id;

    if (!sheetId) return buildWrongSelectionCard();

    const sheetFileType = getFileType(sheetId);

    switch (sheetFileType) {
        case FileType.SETUP:
            return buildInitializationFileEditCard(sheetId);
        case FileType.REPORT:
        // TODO: Return Card "Report options"
    }

    return buildWrongSelectionCard();
}
