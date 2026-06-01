import { ReportSheetSchema } from "../../common/gas-parts";
import {
    buildFieldsMask,
    buildUpdateCellsRequest,
    buildUpdateSheetPropertiesRequest,
    changeGridRangeSheet,
    createRequiredGetter,
    type ExtractRangeNames,
    type MappedNamedRange,
    offsetGridRange,
    type ParsedSpreadsheet,
} from "../../common/gas-utils";
import type { ReportPersistentData } from "../../common/report-utils";
import { getRandomId } from "../../common/setup-utils";

type RangeName = ExtractRangeNames<typeof ReportSheetSchema>;

/**
 * Creates a new sheet for every student.
 */
export function createStudentSheets(
    parsedReport: ParsedSpreadsheet<typeof ReportSheetSchema>,
    persistentData: ReportPersistentData,
): GoogleAppsScript.Sheets.Schema.Request[] {
    const getMappedSheet = createRequiredGetter(parsedReport.mappedSheets, "hoja de reportes");

    const studentTemplateSheet = getMappedSheet(ReportSheetSchema.sheets.studentTemplate.sheetName);
    const studentTemplateId = studentTemplateSheet.properties?.sheetId ?? 0;

    // Remove named ranges so they don't get copied on each new sheet.
    const { requests: cleanTemplateRequest, studentNamedRanges } = cleanTemplateSheet(parsedReport);

    const newSheetsRequests = buildSheetsRequests(studentTemplateId, parsedReport.mappedRanges, persistentData);

    // Return the deleted named ranges to the template.
    const resetTemplateRequests = resetTemplateSheet(studentNamedRanges);

    // Hide the template
    const studentTemplateSheetId = getMappedSheet(ReportSheetSchema.sheets.studentTemplate.sheetName).properties?.sheetId ?? 0;
    const propertiesRequest = buildUpdateSheetPropertiesRequest({
        sheetId: studentTemplateSheetId,
        hidden: true,
    });

    return [...cleanTemplateRequest, ...newSheetsRequests, ...resetTemplateRequests, propertiesRequest];
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

/**
 * Return named ranges to the template
 */
function resetTemplateSheet(namedRanges: GoogleAppsScript.Sheets.Schema.NamedRange[]): GoogleAppsScript.Sheets.Schema.Request[] {
    const requests: GoogleAppsScript.Sheets.Schema.Request[] = [];
    for (const namedRange of namedRanges) requests.push({ addNamedRange: { namedRange: namedRange } });
    return requests;
}

/**
 * Creates a new sheet for each student.
 */
function buildSheetsRequests(
    templateId: number,
    mappedRanges: Partial<Record<RangeName, MappedNamedRange>>,
    persistenData: ReportPersistentData,
): GoogleAppsScript.Sheets.Schema.Request[] {
    const rangeNames = ReportSheetSchema.sheets.studentTemplate.ranges;
    const getMappedRange = createRequiredGetter(mappedRanges, "rango de reporte");

    const requests: GoogleAppsScript.Sheets.Schema.Request[] = [];

    const fields = buildFieldsMask<GoogleAppsScript.Sheets.Schema.CellData>("userEnteredValue");

    const baseUnprotectedRanges = [
        getMappedRange(rangeNames.unprotectedAbilities).namedRange.range,
        getMappedRange(rangeNames.unprotectedComments).namedRange.range,
        getMappedRange(rangeNames.unprotectedTrim1).namedRange.range,
    ];

    let insertSheetIndex: number = 0;
    for (const studentRow of persistenData.students) {
        if (studentRow.type === "student") {
            // Create a random id for the sheet.
            const newSheetId = getRandomId();

            insertSheetIndex++;

            // Copy the template sheet
            requests.push({
                duplicateSheet: {
                    insertSheetIndex,
                    newSheetId,
                    sourceSheetId: templateId,
                    newSheetName: studentRow.sheetName,
                },
            });

            // Add the information to the sheet
            const simpleCopyOps: Array<{ rangeName: RangeName; value: string }> = [
                { rangeName: rangeNames.firstName, value: studentRow.firstName },
                { rangeName: rangeNames.lastName, value: studentRow.lastName },
            ];
            for (const op of simpleCopyOps) {
                const mappedRange = getMappedRange(op.rangeName);
                const destination = changeGridRangeSheet(mappedRange.namedRange.range, newSheetId);
                const data: GoogleAppsScript.Sheets.Schema.CellData[][] = [[{ userEnteredValue: { stringValue: op.value } }]];
                const updateRequest = buildUpdateCellsRequest({ destination, data, fields });
                if (updateRequest) requests.push(updateRequest);
            }

            const infoMappedRange = getMappedRange(rangeNames.generalInfo);
            const infoOriginalRange = offsetGridRange({ origin: infoMappedRange.namedRange.range, height: 3 });
            const infoDestinationRange = changeGridRangeSheet(infoOriginalRange, newSheetId);
            const data: GoogleAppsScript.Sheets.Schema.CellData[][] = [
                [{ userEnteredValue: { stringValue: studentRow.curp } }],
                [{ userEnteredValue: { stringValue: studentRow.grade } }],
                [{ userEnteredValue: { stringValue: studentRow.level } }],
            ];
            const infoRequest = buildUpdateCellsRequest({ destination: infoDestinationRange, data, fields });
            if (infoRequest) requests.push(infoRequest);

            // Protect the sheet
            const unprotectedRanges = baseUnprotectedRanges.map((range) => changeGridRangeSheet(range, newSheetId));
            requests.push({
                addProtectedRange: {
                    protectedRange: {
                        range: { sheetId: newSheetId },
                        description: studentRow.sheetName,
                        warningOnly: false,
                        unprotectedRanges,
                    },
                },
            });
        }
    }

    return requests;
}
