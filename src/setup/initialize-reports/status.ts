import { ReportSheetSchema } from "../../common/gas-parts";
import { buildTransferRequests, buildUpdateSheetPropertiesRequest, createRequiredGetter, type ParsedSpreadsheet } from "../../common/gas-utils";
import type { ReportPersistentData } from "../../common/report-utils";

/**
 * Fills in the Status sheet with the subjects, students and the formulas to put ✔️ or ❌ everywhere.
 */
export function prepareStatusSheet(
    parsedReport: ParsedSpreadsheet<typeof ReportSheetSchema>,
    persistentData: ReportPersistentData,
): GoogleAppsScript.Sheets.Schema.Request[] {
    const getSheet = createRequiredGetter(parsedReport.mappedSheets, "hoja de reporte");

    const statusSheetId = getSheet(ReportSheetSchema.sheets.status.sheetName).properties?.sheetId ?? 0;

    // Rezise the sheet
    const width = 4 + Math.max(9, 4 * persistentData.subjects.length);
    const propertiesRequest = buildUpdateSheetPropertiesRequest({ sheetId: statusSheetId, columnCount: width, hidden: false, frozenColumnCount: 4, index: 1 });

    // Fill in the data
    const colOffset = 0;
    const { requests, colOffset } = buildTransferRequests({});

    // Fill in the data
    return [propertiesRequest];
}

/**
 * Fill the General Data section of the sheet
 */
