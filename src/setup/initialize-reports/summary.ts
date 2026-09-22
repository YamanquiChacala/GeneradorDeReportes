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
    type ParsedSpreadsheet,
    RangeBehavior,
} from "../../common/gas-utils";
import type { Range, ReportPersistentData } from "../../common/report-utils";
import { buildSummaryHeadersData, buildSummaryStudentData, getSummaryColumnWidths, type PeriodRanges } from "../../common/setup-utils";

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

    const commentRange = getMappedRange(ReportSheetSchema.sheets.studentTemplate.ranges.comments);
    const subjectRanges: PeriodRanges = [
        getMappedRange(ReportSheetSchema.sheets.studentTemplate.ranges.trim1Subjects),
        getMappedRange(ReportSheetSchema.sheets.studentTemplate.ranges.trim2Subjects),
        getMappedRange(ReportSheetSchema.sheets.studentTemplate.ranges.trim3Subjects),
    ];
    const fieldRanges: PeriodRanges = [
        getMappedRange(ReportSheetSchema.sheets.studentTemplate.ranges.trim1Fields),
        getMappedRange(ReportSheetSchema.sheets.studentTemplate.ranges.trim2Fields),
        getMappedRange(ReportSheetSchema.sheets.studentTemplate.ranges.trim3Fields),
    ];
    const averageRanges: PeriodRanges = [
        getMappedRange(ReportSheetSchema.sheets.studentTemplate.ranges.trim1Totals),
        getMappedRange(ReportSheetSchema.sheets.studentTemplate.ranges.trim2Totals),
        getMappedRange(ReportSheetSchema.sheets.studentTemplate.ranges.trim3Totals),
    ];

    const trim1SummaryRange = getMappedRange(ReportSheetSchema.sheets.summary.ranges.trim1);

    let rowOffset = 0;

    const trim1SummaryData = buildSummaryStudentData({
        attendancePerClass: persistentData.configData.attendancePerClass,
        averagePerField: persistentData.configData.averagePerField,
        subjects: persistentData.subjects.length,
        fields,
        students: persistentData.students,
        period: 0,
        attendanceSheetName,
        commentsRange: commentRange,
        subjectsRange: subjectRanges,
        fieldsRange: fieldRanges,
        averagesRange: averageRanges,
    });

    const transferResult = buildTransferRequests({
        destination: trim1SummaryRange,
        data: trim1SummaryData,
        fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.CellData>("userEnteredValue"),
        rowBehavior: RangeBehavior.INSERT_DELETE,
        colBehavior: RangeBehavior.MODIFY_RANGE,
        rowOffset,
    });

    rowOffset = transferResult.rowOffset;

    return transferResult.requests;
}
