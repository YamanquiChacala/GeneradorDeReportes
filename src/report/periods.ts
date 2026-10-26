import { ReportSheetSchema } from "../common/gas-parts";
import {
    buildCopyPasteRequest,
    buildFieldsMask,
    buildTransferRequests,
    buildUpdateCellsRequest,
    buildUpdatePersistenMetadata,
    createRequiredGetter,
    offsetGridRange,
    type ParsedSpreadsheet,
    PasteType,
} from "../common/gas-utils";
import { generatePeriodString, type Period, type ReportPersistentData, TRIMESTER_NAMES } from "../common/report-utils";
import { loadReport } from "./load";

/**
 * Changes the active reporting period for a report file and optionally resets student comments.
 *
 * @param reportFileId - The spreadsheet file ID to update.
 * @param newPeriod - The period to apply to all student sheets.
 * @param resetComments - Whether to clear the copied comments and ability values before updating the period.
 */
export function switchPeriod(reportFileId: string, newPeriod: Period, resetComments: boolean) {
    const { parsedReport, persistentData } = loadReport(reportFileId);

    if (persistentData.currentPeriod === newPeriod && !resetComments) return; // Nothing to change

    // Update Persistent Data
    const updatePersistentDataRequests = updatePersistentData(parsedReport, persistentData, newPeriod);

    // Update each student's data
    const updateStudentDataRequests = updateStudentsPeriodData(parsedReport, persistentData);

    // Reset abilities and comments
    const resetStudentsCommentsRequests = resetComments ? resetStudentsComments(parsedReport) : [];

    // Update unprotected ranges

    // Hide calendar columns

    const apiRequests: GoogleAppsScript.Sheets.Schema.Request[] = [...updatePersistentDataRequests, ...updateStudentDataRequests, ...resetStudentsCommentsRequests];

    Sheets?.Spreadsheets.batchUpdate({ requests: apiRequests }, reportFileId);
}

/**
 * Updates the stored report metadata to reflect the selected period and returns the corresponding spreadsheet request.
 *
 * @param persistentData - The in-memory metadata for the report.
 * @param newPeriod - The new active period to store.
 * @returns The request used to persist the updated period metadata.
 */
function updatePersistentData(
    parsedReport: ParsedSpreadsheet<typeof ReportSheetSchema>,
    persistentData: ReportPersistentData,
    newPeriod: Period,
): GoogleAppsScript.Sheets.Schema.Request[] {
    const getMappedRange = createRequiredGetter(parsedReport.mappedRanges, "rango de reporte");

    // --------------
    // Mutate local variable

    // Swap grade protections.
    const trimesters = persistentData.protectedSections.trimesters;
    [trimesters[persistentData.currentPeriod], trimesters[newPeriod]] = [trimesters[newPeriod], trimesters[persistentData.currentPeriod]];

    persistentData.currentPeriod = newPeriod;

    // --------------
    // Mutate metadata
    const metadataRequest = buildUpdatePersistenMetadata(persistentData);

    // --------------
    // Mutate sheet
    const currentPeriodData: GoogleAppsScript.Sheets.Schema.CellData[][] = [[{ userEnteredValue: { numberValue: newPeriod } }]];

    const transferResult = buildTransferRequests({
        destination: getMappedRange(ReportSheetSchema.sheets.persistentData.ranges.currentPeriod),
        data: currentPeriodData,
        fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.CellData>("userEnteredValue.numberValue"),
    });

    const protectedPeriodsData: GoogleAppsScript.Sheets.Schema.CellData[][] = [
        [{ userEnteredValue: { boolValue: trimesters[0] } }],
        [{ userEnteredValue: { boolValue: trimesters[1] } }],
        [{ userEnteredValue: { boolValue: trimesters[2] } }],
    ];

    const 

    return [metadataRequest, ...transferResult.requests];
}

/**
 * Builds the spreadsheet update requests needed to write the selected trimester and period text into each student sheet.
 *
 * @param parsedReport - Parsed report metadata and mapped ranges.
 * @param persistentData - Stored report data that provides the current period context.
 * @returns The request list used to update the period labels in every student sheet.
 */
function updateStudentsPeriodData(
    parsedReport: ParsedSpreadsheet<typeof ReportSheetSchema>,
    persistentData: ReportPersistentData,
): GoogleAppsScript.Sheets.Schema.Request[] {
    const getMappedRange = createRequiredGetter(parsedReport.mappedRanges, "rango de reporte");

    const requests: GoogleAppsScript.Sheets.Schema.Request[] = [];

    const periodString = generatePeriodString(persistentData, persistentData.currentPeriod);

    const data: GoogleAppsScript.Sheets.Schema.CellData[][] = [
        [{ userEnteredValue: { stringValue: TRIMESTER_NAMES[persistentData.currentPeriod] } }],
        [{ userEnteredValue: { stringValue: periodString } }],
    ];

    const fields = buildFieldsMask<GoogleAppsScript.Sheets.Schema.CellData>("userEnteredValue.stringValue");

    const origin = getMappedRange(ReportSheetSchema.sheets.studentTemplate.ranges.generalInfo).namedRange.range;
    const dataRange = offsetGridRange({ origin, rowOffset: 3, height: 2, width: 1 });

    for (const sheet of parsedReport.extraSheets) {
        const sheetId = sheet.properties?.sheetId ?? 0;
        const destination = offsetGridRange({ origin: dataRange, sheetId });
        const request = buildUpdateCellsRequest({ destination, data, fields });
        if (request) requests.push(request);
    }

    return requests;
}

/**
 * Creates the copy requests needed to restore the default ability and comments ranges for each student sheet.
 *
 * @param parsedReport - Parsed report metadata and mapped ranges.
 * @returns The list of paste requests that reset student notes and abilities for all sheets.
 */
function resetStudentsComments(parsedReport: ParsedSpreadsheet<typeof ReportSheetSchema>): GoogleAppsScript.Sheets.Schema.Request[] {
    const getMappedRange = createRequiredGetter(parsedReport.mappedRanges, "rango de reporte");

    const requests: GoogleAppsScript.Sheets.Schema.Request[] = [];

    const abilitiesSource = getMappedRange(ReportSheetSchema.sheets.studentTemplate.ranges.unprotectedAbilities).namedRange.range;
    const commentsSource = getMappedRange(ReportSheetSchema.sheets.studentTemplate.ranges.unprotectedComments).namedRange.range;

    for (const sheet of parsedReport.extraSheets) {
        const sheetId = sheet.properties?.sheetId ?? 0;
        const abilitiesDestination = offsetGridRange({ origin: abilitiesSource, sheetId });
        const commentsDestination = offsetGridRange({ origin: commentsSource, sheetId });
        requests.push(buildCopyPasteRequest(abilitiesSource, abilitiesDestination, PasteType.PASTE_NORMAL));
        requests.push(buildCopyPasteRequest(commentsSource, commentsDestination, PasteType.PASTE_NORMAL));
    }

    return requests;
}
