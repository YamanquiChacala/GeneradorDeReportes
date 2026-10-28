import { ReportSheetSchema } from "../common/gas-parts";
import { buildFieldsMask, buildTransferRequests, buildUpdatePersistenMetadata, createRequiredGetter, type ParsedSpreadsheet } from "../common/gas-utils";
import type { ReportPersistentData } from "../common/report-utils";
import { updateStudentsUnprotectedRanges } from "./helpers";
import { loadReport } from "./load";

export function updateProtectedSections(reportFileId: string, abilities: boolean, comments: boolean, grades: [boolean, boolean, boolean]) {
    const { parsedReport, persistentData } = loadReport(reportFileId);

    if (
        persistentData.protectedSections.abilities === abilities &&
        persistentData.protectedSections.comments === comments &&
        persistentData.protectedSections.trimesters[0] === grades[0] &&
        persistentData.protectedSections.trimesters[1] === grades[1] &&
        persistentData.protectedSections.trimesters[2] === grades[2]
    )
        return; // No changes

    const userEmail = Session.getActiveUser().getEmail();

    // Update Persistent Data
    const updatePersistentDataRequests = updatePersistentData(parsedReport, persistentData, abilities, comments, grades);

    // Update unprotected ranges of students
    const studentUnprotectedRangesRequests = updateStudentsUnprotectedRanges(parsedReport, persistentData, userEmail);

    const apiRequests: GoogleAppsScript.Sheets.Schema.Request[] = [...updatePersistentDataRequests, ...studentUnprotectedRangesRequests];

    Sheets?.Spreadsheets.batchUpdate({ requests: apiRequests }, reportFileId);
}

function updatePersistentData(
    parsedReport: ParsedSpreadsheet<typeof ReportSheetSchema>,
    persistentData: ReportPersistentData,
    abilities: boolean,
    comments: boolean,
    grades: [boolean, boolean, boolean],
): GoogleAppsScript.Sheets.Schema.Request[] {
    const getMappedRange = createRequiredGetter(parsedReport.mappedRanges, "rango de reporte");

    // --------------
    // Mutate local variable

    persistentData.protectedSections.abilities = abilities;
    persistentData.protectedSections.comments = comments;
    persistentData.protectedSections.trimesters = grades;

    // --------------
    // Mutate metadata

    const metadataRequest = buildUpdatePersistenMetadata(persistentData);

    // --------------
    // Mutate sheet

    const protectAbilitiesData: GoogleAppsScript.Sheets.Schema.CellData[][] = [[{ userEnteredValue: { boolValue: abilities } }]];
    const protectCommentsData: GoogleAppsScript.Sheets.Schema.CellData[][] = [[{ userEnteredValue: { boolValue: comments } }]];
    const protectGradesData: GoogleAppsScript.Sheets.Schema.CellData[][] = [
        [{ userEnteredValue: { boolValue: grades[0] } }],
        [{ userEnteredValue: { boolValue: grades[1] } }],
        [{ userEnteredValue: { boolValue: grades[2] } }],
    ];

    const protectAbilitiesTransferResult = buildTransferRequests({
        destination: getMappedRange(ReportSheetSchema.sheets.persistentData.ranges.protectSkills),
        data: protectAbilitiesData,
        fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.CellData>("userEnteredValue.boolValue"),
    });
    const protectCommentsTransferResult = buildTransferRequests({
        destination: getMappedRange(ReportSheetSchema.sheets.persistentData.ranges.protectComments),
        data: protectCommentsData,
        fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.CellData>("userEnteredValue.boolValue"),
    });
    const protectGradesTransferResult = buildTransferRequests({
        destination: getMappedRange(ReportSheetSchema.sheets.persistentData.ranges.protectPeriods),
        data: protectGradesData,
        fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.CellData>("userEnteredValue.boolValue"),
    });

    return [metadataRequest, ...protectAbilitiesTransferResult.requests, ...protectCommentsTransferResult.requests, ...protectGradesTransferResult.requests];
}
