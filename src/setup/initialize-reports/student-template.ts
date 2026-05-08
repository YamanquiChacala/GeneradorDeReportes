import { Dimension } from "../../common/gas-enums";
import { ReportSheetSchema } from "../../common/sheet-schema";
import type { ExtractRangeNames, MappedNamedRange, ParsedSpreadsheet } from "../../common/utils/mapped-name-range";
import type { ReportPersistentData } from "./persistent-data";

/**
 * Prepares the student template sheet, filling in subjects, formulas, etc, so it can be duplicated for each student.
 */
export function prepareStudentTemplate(
    parsedReport: ParsedSpreadsheet<typeof ReportSheetSchema>,
    persistentData: ReportPersistentData,
): GoogleAppsScript.Sheets.Schema.Request[] {
    const requests: GoogleAppsScript.Sheets.Schema.Request[] = [];

    const studentTemplateSheetId = parsedReport.sheets[ReportSheetSchema.sheets.studentTemplate.sheetName]?.properties?.sheetId;

    if (!studentTemplateSheetId) throw new Error("No se encuentran template de estudiantes.");

    // Adapt the sheet's ranges

    const adaptSheetRangesRequests = adaptSizeAndRanges(studentTemplateSheetId, parsedReport.namedRanges, persistentData);

    // Build requests

    requests.push(...adaptSheetRangesRequests);

    return requests;
}

/**
 * Adapts the size of the template sheet, and updates the named ranges
 */
function adaptSizeAndRanges(
    sheetId: number,
    namedRanges: Partial<Record<ExtractRangeNames<typeof ReportSheetSchema>, MappedNamedRange>>,
    persistentData: ReportPersistentData,
): GoogleAppsScript.Sheets.Schema.Request[] {
    const requests: GoogleAppsScript.Sheets.Schema.Request[] = [];

    let rowOffset = 0;

    if (persistentData.configData.attendancePerClass) {
        const generalAbsencesRange = namedRanges[ReportSheetSchema.sheets.studentTemplate.ranges.generalAbsences]?.range;
        rowOffset -= (generalAbsencesRange?.startRowIndex ?? 0) - (generalAbsencesRange?.startRowIndex ?? 0);
        requests.push({
            deleteDimension: {
                range: {
                    sheetId,
                    dimension: Dimension.ROWS,
                    startIndex: generalAbsencesRange?.startRowIndex,
                    endIndex: generalAbsencesRange?.endRowIndex,
                },
            },
        });
    }

    return requests;
}
