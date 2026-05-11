import { FileType } from "../../common/enums";
import { ReportSheetSchema, SetupSheetSchema } from "../../common/sheet-schema";
import { key as FILE_VALIDATION_KEY } from "../../common/utils/file-validation";
import { buildFieldsMask } from "../../common/utils/gas-types";
import { type ExtractRangeNames, MappedNamedRange, parseSpreadsheet } from "../../common/utils/mapped-name-range";
import { sanitizeFileName } from "../../common/utils/text-utils";
import { createAttendanceSheet } from "./attendance";
import { fillPersistentData } from "./persistent-data";
import { prepareStudentTemplate } from "./student-template";

/**
 * Initializes a Report spreadsheet in the same folder as the setupFile, with the information from the Setup file.
 */
export function initializeReport(setupFileId: string, parentId: string) {
    // Fetch and parse setup file data

    const setupFieldsMask = buildFieldsMask<GoogleAppsScript.Sheets.Schema.Spreadsheet>(
        "sheets.properties.sheetId",
        "sheets.properties.title",
        "sheets.properties.gridProperties.rowCount",
        "sheets.properties.gridProperties.columnCount",
        "sheets.data.rowData.values.effectiveValue",
        "sheets.data.rowData.values.effectiveFormat.backgroundColor",
        "namedRanges",
    );
    const reportFieldsMask = buildFieldsMask<GoogleAppsScript.Sheets.Schema.Spreadsheet>(
        "sheets.properties.sheetId",
        "sheets.properties.title",
        "sheets.properties.gridProperties.rowCount",
        "sheets.properties.gridProperties.columnCount",
        "sheets.data.rowData.values.effectiveValue",
        "namedRanges",
    );
    const SetupSpreadsheet = Sheets?.Spreadsheets.get(setupFileId, { fields: setupFieldsMask });
    const { namedRanges: setupRanges } = parseSpreadsheet(SetupSpreadsheet, SetupSheetSchema);

    // Create report file.

    const reportFileId = createReportFile(parentId, setupRanges);

    // Fetch and parse report file data

    const reportSpreadsheet = Sheets?.Spreadsheets.get(reportFileId, { fields: reportFieldsMask });
    const parsedReportSheet = parseSpreadsheet(reportSpreadsheet, ReportSheetSchema);

    // Fill Report Persistent Data

    const { data: persistentData, requests: persistentDataRequests } = fillPersistentData(setupRanges, parsedReportSheet.namedRanges);

    // Create Attendance sheet

    const attendanceRequests = createAttendanceSheet(parsedReportSheet, persistentData);

    // Prepare Student template sheet

    const studentTemplateSetup = prepareStudentTemplate(parsedReportSheet, persistentData);

    // ============ Batch Changes ==============

    const apiRequests: GoogleAppsScript.Sheets.Schema.Request[] = [...persistentDataRequests, ...attendanceRequests, ...studentTemplateSetup];

    // ============ Execute Batch update ===========

    Sheets?.Spreadsheets.batchUpdate({ requests: apiRequests }, reportFileId);
}

/**
 * Creates an empty report file in the given folder, with the name from the setup file.
 */
function createReportFile(parentId: string, setupFileRanges: Partial<Record<ExtractRangeNames<typeof SetupSheetSchema>, MappedNamedRange>>): string {
    const groupNameRange = setupFileRanges[SetupSheetSchema.sheets.groupData.ranges.groupName];
    const groupName = MappedNamedRange.getCellText({ mappedRange: groupNameRange });

    if (!groupName) throw new Error("Falta nombre del grupo.");

    const fileName = sanitizeFileName(groupName);

    const reportFile = Drive?.Files.copy(
        {
            name: fileName,
            parents: [parentId],
            appProperties: {
                [FILE_VALIDATION_KEY]: FileType.REPORT,
            },
        },
        ReportSheetSchema.templateId,
        {
            supportsAllDrives: true,
        },
    );

    const reportFileId = reportFile?.id;
    if (!reportFileId) throw new Error("Error copiando template de calificaciones en Drive.");

    return reportFileId;
}
