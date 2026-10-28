import { ReportSheetSchema } from "../common/gas-parts";
import { buildProtectExtraSheetRequests, createRequiredGetter, type ParsedSpreadsheet } from "../common/gas-utils";
import type { ReportPersistentData } from "../common/report-utils";

export function updateStudentsUnprotectedRanges(
    parsedReport: ParsedSpreadsheet<typeof ReportSheetSchema>,
    persistentData: ReportPersistentData,
    userEmail: string,
): GoogleAppsScript.Sheets.Schema.Request[] {
    const getMappedRange = createRequiredGetter(parsedReport.mappedRanges, "rango de reporte");

    const unprotectedRanges: GoogleAppsScript.Sheets.Schema.GridRange[] = [];

    if (!persistentData.protectedSections.abilities) {
        unprotectedRanges.push(getMappedRange(ReportSheetSchema.sheets.studentTemplate.ranges.unprotectedAbilities).namedRange.range);
    }
    if (!persistentData.protectedSections.comments) {
        unprotectedRanges.push(getMappedRange(ReportSheetSchema.sheets.studentTemplate.ranges.unprotectedComments).namedRange.range);
    }
    if (!persistentData.protectedSections.trimesters[0]) {
        unprotectedRanges.push(getMappedRange(ReportSheetSchema.sheets.studentTemplate.ranges.unprotectedTrim1).namedRange.range);
    }
    if (!persistentData.protectedSections.trimesters[1]) {
        unprotectedRanges.push(getMappedRange(ReportSheetSchema.sheets.studentTemplate.ranges.unprotectedTrim2).namedRange.range);
    }
    if (!persistentData.protectedSections.trimesters[2]) {
        unprotectedRanges.push(getMappedRange(ReportSheetSchema.sheets.studentTemplate.ranges.unprotectedTrim3).namedRange.range);
    }

    return buildProtectExtraSheetRequests(parsedReport, userEmail, unprotectedRanges);
}
