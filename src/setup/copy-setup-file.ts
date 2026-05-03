import { FileType } from "../common/enums";
import { SetupSheetSchema } from "../common/sheet-schema";
import { key as FILE_VALIDATION_KEY } from "../common/utils/file-validation";

/**
 * Copies the SetupFile `fileId` changing the group name and saves it into `folderId`.
 */
export function copySetupFile(setupFileId: string, folderId: string | undefined, groupName: string) {
    const filename = `__Registro ${groupName}`;

    const fileRequest: GoogleAppsScript.Drive_v3.Drive.V3.Schema.File = {
        name: filename,
        appProperties: {
            [FILE_VALIDATION_KEY]: FileType.SETUP,
        },
    };
    if (folderId) {
        fileRequest.parents = [folderId];
    }

    const newFile = Drive?.Files.copy(fileRequest, setupFileId, {
        supportsAllDrives: true,
    });

    const newFileId = newFile?.id;

    if (!newFileId) throw new Error("Error al crear copia del Regirsto Inicial");

    Sheets?.Spreadsheets.Values.update({ values: [[groupName]] }, newFileId, SetupSheetSchema.sheets.groupData.ranges.groupName, { valueInputOption: "USER_ENTERED" });
}
