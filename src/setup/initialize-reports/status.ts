import { ReportSheetSchema } from "../../common/gas-parts";
import {
    buildTransferRequests,
    buildUnmergeCellsRequest,
    buildUpdateSheetPropertiesRequest,
    createRange,
    createRequiredGetter,
    type ParsedSpreadsheet,
} from "../../common/gas-utils";
import type { ReportPersistentData } from "../../common/report-utils";

/**
 * Fills in the Status sheet with the subjects, students and the formulas to put ✔️ or ❌ everywhere.
 */
export function prepareStatusSheet(
    parsedReport: ParsedSpreadsheet<typeof ReportSheetSchema>,
    persistentData: ReportPersistentData,
): GoogleAppsScript.Sheets.Schema.Request[] {
    // Prepare the sheet
    const prepareRequests = prepareSheet(parsedReport, persistentData);

    // Fill in the data
    const colOffset = 0;

    // Fill in the data
    return [...prepareRequests];
}

/**
 * Resizes, sorts and gets the sheet ready to input information.
 */
function prepareSheet(parsedReport: ParsedSpreadsheet<typeof ReportSheetSchema>, persistentData: ReportPersistentData): GoogleAppsScript.Sheets.Schema.Request[] {
    const getSheet = createRequiredGetter(parsedReport.mappedSheets, "hoja de reporte");
    const statusSheetId = getSheet(ReportSheetSchema.sheets.status.sheetName).properties?.sheetId ?? 0;

    const frozenCols = 4;

    // Calculate width
    const width = frozenCols + Math.max(9, 4 * persistentData.subjects.length);

    // Set the properties
    const propertiesRequest = buildUpdateSheetPropertiesRequest({ sheetId: statusSheetId, columnCount: width, hidden: false, frozenColumnCount: frozenCols, index: 1 });

    // Remove merged cells to avoid problems with data input
    const unfrozenRange = createRange(statusSheetId, 0, frozenCols);
    const unmergeRequest = buildUnmergeCellsRequest(unfrozenRange);

    return [propertiesRequest, unmergeRequest];
}

/**
 * Fill the General Data section of the sheet
 */
function fillGeneralData(parsedReport: ParsedSpreadsheet<typeof ReportSheetSchema>, persistentData: ReportPersistentData): GoogleAppsScript.Sheets.Schema.Request[] {
    const getMappedRange = createRequiredGetter(parsedReport.mappedRanges, "rango de reporte");

    return [];
}
