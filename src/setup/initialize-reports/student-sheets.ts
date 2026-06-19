import { ReportSheetSchema } from "../../common/gas-parts";
import {
    addNewSheet,
    buildFieldsMask,
    buildProtectExtraSheetRequests,
    buildUpdateCellsRequest,
    createRequiredGetter,
    type ExtractRangeNames,
    offsetGridRange,
    type ParsedSpreadsheet,
} from "../../common/gas-utils";
import { type ReportPersistentData, type Student, StudentRowType } from "../../common/report-utils";
import { zip } from "../../common/utils";

type RangeName = ExtractRangeNames<typeof ReportSheetSchema>;

/**
 * Creates a new sheet for every student.
 */
export function createStudentSheets(
    parsedReport: ParsedSpreadsheet<typeof ReportSheetSchema>,
    persistentData: ReportPersistentData,
): GoogleAppsScript.Sheets.Schema.Request[] {
    const rangeNames = ReportSheetSchema.sheets.studentTemplate.ranges;
    const getMappedRange = createRequiredGetter(parsedReport.mappedRanges, "rango de reporte");

    const realStudents = persistentData.students.filter((studentRow) => studentRow.type === StudentRowType.STUDENT);

    const { requests: createSheetsRequests, newSheetIds } = addNewSheet({
        parsedData: parsedReport,
        sourceSheetTitle: ReportSheetSchema.sheets.studentTemplate.sheetName,
        insertSheetIndex: 1,
        multipleSheetNames: realStudents.map((student) => student.sheetName),
    });

    const fillInDataRequests: GoogleAppsScript.Sheets.Schema.Request[] = [];
    for (const [sheetId, student] of zip(newSheetIds, realStudents)) {
        fillInDataRequests.push(...fillStudentData(parsedReport, sheetId, student));
    }

    const baseUnprotectedRanges = [
        getMappedRange(rangeNames.unprotectedAbilities).namedRange.range,
        getMappedRange(rangeNames.unprotectedComments).namedRange.range,
        getMappedRange(rangeNames.unprotectedTrim1).namedRange.range,
    ];

    const protectSheetsRequests = buildProtectExtraSheetRequests(parsedReport, baseUnprotectedRanges);

    return [...createSheetsRequests, ...fillInDataRequests, ...protectSheetsRequests];
}

/**
 * Fills a student sheet with the data from the student
 */
function fillStudentData(parsedReport: ParsedSpreadsheet<typeof ReportSheetSchema>, sheetId: number, student: Student): GoogleAppsScript.Sheets.Schema.Request[] {
    const requests: GoogleAppsScript.Sheets.Schema.Request[] = [];

    const getMappedRange = createRequiredGetter(parsedReport.mappedRanges, "rango de reporte");
    const rangeNames = ReportSheetSchema.sheets.studentTemplate.ranges;

    const fields = buildFieldsMask<GoogleAppsScript.Sheets.Schema.CellData>("userEnteredValue");

    const simpleCopyOps: Array<{ rangeName: RangeName; value: string }> = [
        { rangeName: rangeNames.firstName, value: student.firstName },
        { rangeName: rangeNames.lastName, value: student.lastName },
    ];
    for (const op of simpleCopyOps) {
        const mappedRange = getMappedRange(op.rangeName);
        const destination = offsetGridRange({ origin: mappedRange.namedRange.range, sheetId });
        const data: GoogleAppsScript.Sheets.Schema.CellData[][] = [[{ userEnteredValue: { stringValue: op.value } }]];
        const updateRequest = buildUpdateCellsRequest({ destination, data, fields });
        if (updateRequest) requests.push(updateRequest);
    }

    const infoMappedRange = getMappedRange(rangeNames.generalInfo);
    const infoDestinationRange = offsetGridRange({ origin: infoMappedRange.namedRange.range, height: 3, sheetId });
    const data: GoogleAppsScript.Sheets.Schema.CellData[][] = [
        [{ userEnteredValue: { stringValue: student.curp } }],
        [{ userEnteredValue: { stringValue: student.grade } }],
        [{ userEnteredValue: { stringValue: student.level } }],
    ];
    const infoRequest = buildUpdateCellsRequest({ destination: infoDestinationRange, data, fields });
    if (infoRequest) requests.push(infoRequest);

    return requests;
}
