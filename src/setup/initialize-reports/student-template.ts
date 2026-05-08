import type { ReportSheetSchema } from "../../common/sheet-schema";
import type { ExtractRangeNames, MappedNamedRange } from "../../common/utils/mapped-name-range";
import type { ReportPersistentData } from "./persistent-data";

/**
 * Prepares the student template sheet, filling in subjects, formulas, etc, so it can be duplicated for each student.
 */
export function prepareStudentTemplate(
    reportRanges: Partial<Record<ExtractRangeNames<typeof ReportSheetSchema>, MappedNamedRange>>,
    persistentData: ReportPersistentData,
): GoogleAppsScript.Sheets.Schema.Request[] {
    const requests: GoogleAppsScript.Sheets.Schema.Request[] = []


    return requests;
}


/**
 * Adapts the size of the template sheet, and updates the named ranges
 */
function adaptSizeAndRanges(namedRanges: Partial<Record<ExtractRangeNames<typeof ReportSheetSchema>, MappedNamedRange>>, persistentData: ReportPersistentData): GoogleAppsScript.Sheets.Schema.Request[] {
    const requests: GoogleAppsScript.Sheets.Schema.Request[] = []


    return requests;
}