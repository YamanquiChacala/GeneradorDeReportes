import { FILE_VALIDATION_KEY, FileType, SetupSheetSchema } from "../../common/gas-parts";
import { ValueInputOption } from "../../common/gas-utils/types";
import { buildSetupFileUpdatePayload, SETUP_FILE_PREFIX, type SetupFileData } from "../../common/setup-utils";
import { generateCalendar } from "../generate-calendar";

/**
 * Creates a new Group Initialization file with the given data.
 */
export function createSetupFile(initData: SetupFileData) {
    // ========== Create File ============
    const fileName = `${SETUP_FILE_PREFIX}${initData.groupName}`;

    const newFile = Drive?.Files.copy(
        {
            name: fileName,
            parents: [initData.folderId],
            appProperties: {
                [FILE_VALIDATION_KEY]: FileType.SETUP,
            },
        },
        SetupSheetSchema.templateId,
        { supportsAllDrives: true },
    );

    const newFileId = newFile?.id;
    if (!newFileId) throw new Error("Error al crear copia del Registro Inicial");

    // ======== Update namedRanges ==========

    // Get the pure payload
    const updateData = buildSetupFileUpdatePayload(initData);

    Sheets?.Spreadsheets.Values.batchUpdate(
        {
            valueInputOption: ValueInputOption.USER_ENTERED,
            data: updateData,
        },
        newFileId,
    );

    // ========= Create Calendar ==========
    generateCalendar(newFileId);
}
