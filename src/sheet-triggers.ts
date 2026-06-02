import { buildRequestAuthorizationCard, buildWrongSelectionCard, FileType, getFileType } from "./common/gas-parts";
import { buildEditSetupFileCard } from "./setup/cards";

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
            return buildEditSetupFileCard(sheetId);
        case FileType.REPORT:
        // TODO: Return Card "Report options"
    }

    return buildWrongSelectionCard();
}
