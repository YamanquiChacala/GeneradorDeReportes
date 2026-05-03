import { FileType } from "../common/enums";
import { SetupSheetSchema } from "../common/sheet-schema";
import { key as FILE_VALIDATION_KEY } from "../common/utils/file-validation";
import { defineRangesDataConfig, type MappedInput } from "../common/utils/gas-types";
import { generateCalendar } from "./generate-calendar";

const SetupFileDataConfig = defineRangesDataConfig({
    groupName: { range: SetupSheetSchema.sheets.groupData.ranges.groupName, type: "string" },
    attendancePerClass: { range: SetupSheetSchema.sheets.groupData.ranges.attendancePerClass, type: "boolean" },
    averagePerField: { range: SetupSheetSchema.sheets.groupData.ranges.averagePerField, type: "boolean" },
    dateStart: { range: SetupSheetSchema.sheets.groupData.ranges.dateStart, type: "date" },
    dateEndTrimester1: { range: SetupSheetSchema.sheets.groupData.ranges.dateTrim1, type: "date" },
    dateEndTrimester2: { range: SetupSheetSchema.sheets.groupData.ranges.dateTrim2, type: "date" },
    dateEnd: { range: SetupSheetSchema.sheets.groupData.ranges.dateEnd, type: "date" },
} as const);

export type SetupFileData = { folderId: string } & {
    [K in keyof typeof SetupFileDataConfig]: MappedInput<(typeof SetupFileDataConfig)[K]["type"]>;
};

/**
 * Creates a new Group Initialization file with the given data.
 */
export function createSetupFile(initData: SetupFileData) {
    // ========== Create File ============

    const fileName = `__Registro Inicial - ${initData.groupName}`;

    const newFile = Drive?.Files.copy(
        {
            name: fileName,
            parents: [initData.folderId],
            appProperties: {
                [FILE_VALIDATION_KEY]: FileType.SETUP,
            },
        },
        SetupSheetSchema.templateId,
        {
            supportsAllDrives: true,
        },
    );

    const newFileId = newFile?.id;

    if (!newFileId) throw new Error("Error al crear copia del Regirsto Inicial");

    // ======== Update namedRanges ==========

    const updateData: GoogleAppsScript.Sheets.Schema.ValueRange[] = [];

    for (const [key, config] of Object.entries(SetupFileDataConfig)) {
        const rawValue = initData[key as keyof typeof SetupFileDataConfig];

        if (rawValue == null) continue;

        let cellValue = rawValue;

        if (config.type === "date") {
            const date = new Date(rawValue as number);
            cellValue = Utilities.formatDate(date, "UTC", "yyyy-MM-dd");
        }
        updateData.push({
            range: config.range,
            values: [[cellValue]],
        });
    }

    Sheets?.Spreadsheets.Values.batchUpdate(
        {
            valueInputOption: "USER_ENTERED",
            data: updateData,
        },
        newFileId,
    );

    // ========= Create Calendar ==========
    generateCalendar(newFileId);
}
