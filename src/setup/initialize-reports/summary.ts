import { ReportSheetSchema } from "../../common/gas-parts";
import {
    buildFieldsMask,
    buildMergeCellsRequest,
    buildTransferRequests,
    buildUnmergeCellsRequest,
    buildUpdateColumnWidthRequests,
    buildUpdateSheetPropertiesRequest,
    createRange,
    createRequiredGetter,
    type MappedNamedRange,
    type ParsedSpreadsheet,
    RangeBehavior,
} from "../../common/gas-utils";
import type { Range, ReportPersistentData } from "../../common/report-utils";
import { buildSummaryHeadersData, buildSummaryStudentData, getSummaryColumnWidths } from "../../common/setup-utils";

export function prepareSummarySheet(
    parsedReport: ParsedSpreadsheet<typeof ReportSheetSchema>,
    persistentData: ReportPersistentData,
): GoogleAppsScript.Sheets.Schema.Request[] {
    // Prepare the sheet (resize and properties)
    const { requests: prepareSheetRequests, mergeColumns } = prepareSheet(parsedReport, persistentData);

    // Fill header
    const headerRequests = addHeaders(parsedReport, persistentData, mergeColumns);

    // Fill cntents
    const contentRequests = addContent(parsedReport, persistentData);

    // TODO: Fill periods
    return [...prepareSheetRequests, ...headerRequests, ...contentRequests];
}

/**
 * Resizes, sorts and gets the sheet ready to input information.
 */
function prepareSheet(
    parsedReport: ParsedSpreadsheet<typeof ReportSheetSchema>,
    persistentData: ReportPersistentData,
): { requests: GoogleAppsScript.Sheets.Schema.Request[]; mergeColumns: Range[] } {
    const getSheet = createRequiredGetter(parsedReport.mappedSheets, "hoja de reporte");
    const summarySheetId = getSheet(ReportSheetSchema.sheets.summary.sheetName).properties?.sheetId ?? 0;

    const frozenCols = 3;

    const { columWidths, mergeRanges } = getSummaryColumnWidths(
        persistentData.configData.attendancePerClass,
        persistentData.configData.averagePerField,
        persistentData.subjects.length,
        persistentData.academicFields.length,
    );

    const totalCols = frozenCols + columWidths.length;

    // Set the properties
    const propertiesRequest = buildUpdateSheetPropertiesRequest({
        sheetId: summarySheetId,
        hidden: false,
        columnCount: totalCols,
        frozenColumnCount: frozenCols,
        index: 1,
    });

    // Remove merged cells to avoid problems with data input
    const unfrozenRange = createRange(summarySheetId, 0, frozenCols);
    const unmergeRequest = buildUnmergeCellsRequest(unfrozenRange);

    const adjustColWidthRequests = buildUpdateColumnWidthRequests(summarySheetId, frozenCols, columWidths);

    return { requests: [propertiesRequest, unmergeRequest, ...adjustColWidthRequests], mergeColumns: mergeRanges };
}

/**
 * Add the labels for the header of the sheet
 */
function addHeaders(
    parsedReport: ParsedSpreadsheet<typeof ReportSheetSchema>,
    persistentData: ReportPersistentData,
    mergeColumns: Range[],
): GoogleAppsScript.Sheets.Schema.Request[] {
    const getMappedRange = createRequiredGetter(parsedReport.mappedRanges, "rango de reporte");
    const statusHeadersRange = getMappedRange(ReportSheetSchema.sheets.summary.ranges.labels);

    const frozenCols = 3;

    const subjectNames = persistentData.subjects.map((x) => x.subject);
    const fieldNames = persistentData.academicFields.map((x) => x.name);

    const data = buildSummaryHeadersData(persistentData.configData.attendancePerClass, persistentData.configData.averagePerField, subjectNames, fieldNames);

    const transferResult = buildTransferRequests({
        destination: statusHeadersRange,
        data,
        fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.CellData>("userEnteredValue", "userEnteredFormat.borders"),
        rowBehavior: RangeBehavior.INSERT_DELETE,
        colBehavior: RangeBehavior.MODIFY_RANGE,
    });

    const mergeRequests: GoogleAppsScript.Sheets.Schema.Request[] = [];

    for (const merge of mergeColumns) {
        const mergeRange = createRange(statusHeadersRange.sheet.properties?.sheetId ?? 0, 0, frozenCols + merge.start, 1, merge.end - merge.start);
        mergeRequests.push(buildMergeCellsRequest(mergeRange));
    }

    return [...transferResult.requests, ...mergeRequests];
}

/**
 * Fills in the content of the sheet
 */

function addContent(parsedReport: ParsedSpreadsheet<typeof ReportSheetSchema>, persistentData: ReportPersistentData): GoogleAppsScript.Sheets.Schema.Request[] {
    const getMappedRange = createRequiredGetter(parsedReport.mappedRanges, "rango de reporte'");

    const fields = persistentData.academicFields.map((field) => field.subjects);
    const attendanceSheetName = ReportSheetSchema.sheets.attendance.sheetName;

    const commentsRange = getMappedRange(ReportSheetSchema.sheets.studentTemplate.ranges.comments);

    const periodBundle: Array<{
        periodMappedRange: MappedNamedRange;
        period: 0 | 1 | 2;
        subjectsRange: MappedNamedRange;
        fieldsRange: MappedNamedRange;
        averagesRange: MappedNamedRange;
    }> = [
        {
            periodMappedRange: getMappedRange(ReportSheetSchema.sheets.summary.ranges.trim1),
            period: 0,
            subjectsRange: getMappedRange(ReportSheetSchema.sheets.studentTemplate.ranges.trim1Subjects),
            fieldsRange: getMappedRange(ReportSheetSchema.sheets.studentTemplate.ranges.trim1Fields),
            averagesRange: getMappedRange(ReportSheetSchema.sheets.studentTemplate.ranges.trim1Totals),
        },
        {
            periodMappedRange: getMappedRange(ReportSheetSchema.sheets.summary.ranges.trim2),
            period: 1,
            subjectsRange: getMappedRange(ReportSheetSchema.sheets.studentTemplate.ranges.trim2Subjects),
            fieldsRange: getMappedRange(ReportSheetSchema.sheets.studentTemplate.ranges.trim2Fields),
            averagesRange: getMappedRange(ReportSheetSchema.sheets.studentTemplate.ranges.trim2Totals),
        },
        {
            periodMappedRange: getMappedRange(ReportSheetSchema.sheets.summary.ranges.trim3),
            period: 2,
            subjectsRange: getMappedRange(ReportSheetSchema.sheets.studentTemplate.ranges.trim3Subjects),
            fieldsRange: getMappedRange(ReportSheetSchema.sheets.studentTemplate.ranges.trim3Fields),
            averagesRange: getMappedRange(ReportSheetSchema.sheets.studentTemplate.ranges.trim3Totals),
        },
    ];

    const requests: GoogleAppsScript.Sheets.Schema.Request[] = [];

    let rowOffset = 0;

    for (const bundle of periodBundle) {
        const periodData = buildSummaryStudentData({
            mappedRange: bundle.periodMappedRange,
            rowOffset,
            attendancePerClass: persistentData.configData.attendancePerClass,
            averagePerField: persistentData.configData.averagePerField,
            subjects: persistentData.subjects.length,
            fields,
            students: persistentData.students,
            period: bundle.period,
            attendanceSheetName,
            commentsRange,
            subjectsRange: bundle.subjectsRange,
            fieldsRange: bundle.fieldsRange,
            averagesRange: bundle.averagesRange,
        });

        const periodDataResponse = buildTransferRequests({
            destination: bundle.periodMappedRange,
            data: periodData,
            fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.CellData>("userEnteredValue"),
            rowBehavior: RangeBehavior.INSERT_DELETE,
            colBehavior: RangeBehavior.MODIFY_RANGE,
            rowOffset,
        });

        rowOffset = periodDataResponse.rowOffset;

        requests.push(...periodDataResponse.requests);
    }

    return requests;
}
