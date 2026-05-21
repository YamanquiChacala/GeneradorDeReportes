import { SetupSheetSchema } from "../gas-parts";
import type { ExtractRangeNames } from "../gas-utils";
import type { SetupFileData } from ".";

type RangeName = ExtractRangeNames<typeof SetupSheetSchema>;

export function buildSetupFileUpdatePayload(initData: SetupFileData): GoogleAppsScript.Sheets.Schema.ValueRange[] {
    const ranges = SetupSheetSchema.sheets.groupData.ranges;

    const updateData: GoogleAppsScript.Sheets.Schema.ValueRange[] = [];

    // Simple mappings
    const simpleCopy: Array<{ key: keyof SetupFileData; rangeName: RangeName }> = [
        { key: "groupName", rangeName: ranges.groupName },
        { key: "attendancePerClass", rangeName: ranges.attendancePerClass },
        { key: "averagePerField", rangeName: ranges.averagePerField },
    ];

    for (const { key, rangeName } of simpleCopy) {
        updateData.push({
            range: rangeName,
            values: [[initData[key]]],
        });
    }

    // Date mappings (Using pure JS instead of GAS Utilities)
    const dates = [initData.dateStart, initData.dateEndTrimester1, initData.dateEndTrimester2, initData.dateEnd].map((ms) => {
        // toISOString() returns "YYYY-MM-DDTHH:mm:ss.sssZ"
        // Splitting by "T" and taking the first part gives us "YYYY-MM-DD" natively in UTC.
        return [new Date(ms).toISOString().split("T")[0]];
    });

    updateData.push({
        range: ranges.dates,
        values: dates,
    });

    return updateData;
}
