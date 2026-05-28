import type { ReportSheetSchema } from "../../common/gas-parts";
import type { ParsedSpreadsheet } from "../../common/gas-utils";
import type { ReportPersistentData } from "../../common/report-utils";

export function createStudentSheets(
    parsedReport: ParsedSpreadsheet<typeof ReportSheetSchema>,
    persistentData: ReportPersistentData,
): GoogleAppsScript.Sheets.Schema.Request[] {
    return [];
}
