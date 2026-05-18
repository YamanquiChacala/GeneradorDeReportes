import { FileType, SETUP_FILE_PREFIX, ValueInputOption } from "../../common/constants";
import { FILE_VALIDATION_KEY, SetupSheetSchema } from "../../common/gas-parts";
import type { ExtractRangeNames } from "../../common/gas-utils";
import { generateCalendar } from "../generate-calendar";

type RangeName = ExtractRangeNames<typeof SetupSheetSchema>;

export interface SetupFileData {
    folderId: string;
    groupName: string;
    attendancePerClass: boolean;
    averagePerField: boolean;
    dateStart: number;
    dateEndTrimester1: number;
    dateEndTrimester2: number;
    dateEnd: number;
}

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
        {
            supportsAllDrives: true,
        },
    );

    const newFileId = newFile?.id;

    if (!newFileId) throw new Error("Error al crear copia del Regirsto Inicial");

    // ======== Update namedRanges ==========

    const ranges = SetupSheetSchema.sheets.groupData.ranges;

    const updateData: GoogleAppsScript.Sheets.Schema.ValueRange[] = [];

    const simpleCopy: Array<{ key: keyof SetupFileData; rangeName: RangeName }> = [
        { key: "groupName", rangeName: ranges.groupName },
        { key: "attendancePerClass", rangeName: ranges.attendancePerClass },
        { key: "averagePerField", rangeName: ranges.averagePerField },
    ];

    for (const { key, rangeName } of simpleCopy) {
        const rawValue = initData[key];
        updateData.push({
            range: rangeName,
            values: [[rawValue]],
        });
    }

    const dates = [initData.dateStart, initData.dateEndTrimester1, initData.dateEndTrimester2, initData.dateEnd].map((ms) => [
        Utilities.formatDate(new Date(ms), "UTC", "yyyy-MM-dd"),
    ]);

    updateData.push({
        range: ranges.dates,
        values: dates,
    });

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
