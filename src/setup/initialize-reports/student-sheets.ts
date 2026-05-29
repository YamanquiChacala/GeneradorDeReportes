import { ReportSheetSchema } from "../../common/gas-parts";
import type { ExtractRangeNames, ParsedSpreadsheet } from "../../common/gas-utils";
import type { ReportPersistentData } from "../../common/report-utils";

type RangeName = ExtractRangeNames<typeof ReportSheetSchema>;

/**
 * Creates a new sheet for every student.
 */
export function createStudentSheets(
    parsedReport: ParsedSpreadsheet<typeof ReportSheetSchema>,
    persistentData: ReportPersistentData,
): GoogleAppsScript.Sheets.Schema.Request[] {
    const { requests: cleanTemplateRequest, studentNamedRanges } = cleanTemplateSheet(parsedReport);

    return [...cleanTemplateRequest];
}

/**
 * Prepares the template so it can be copied for each student.
 */
function cleanTemplateSheet(parsedReport: ParsedSpreadsheet<typeof ReportSheetSchema>): {
    requests: GoogleAppsScript.Sheets.Schema.Request[];
    studentNamedRanges: GoogleAppsScript.Sheets.Schema.NamedRange[];
} {
    const requests: GoogleAppsScript.Sheets.Schema.Request[] = [];
    const studentNamedRanges = parsedReport.mappedSheetNamedRanges[ReportSheetSchema.sheets.studentTemplate.sheetName] ?? [];

    for (const namedRange of studentNamedRanges) requests.push({ deleteNamedRange: { namedRangeId: namedRange.namedRangeId } });

    return { requests, studentNamedRanges };
}
