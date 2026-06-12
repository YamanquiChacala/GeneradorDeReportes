import { ReportSheetSchema } from "../../common/gas-parts";
import type { ExtractRangeNames, NestedSheetSchema, ParsedSpreadsheet } from "../../common/gas-utils";
import {
    buildAddBandingRequest,
    buildCopyPasteRequest,
    buildFieldsMask,
    buildMergeCellsRequest,
    buildUpdateCellsRequest,
    buildUpdateSheetPropertiesRequest,
    createBanding,
    createRange,
    createRequiredGetter,
    createSingleCellRange,
    getCellDataArray,
    type MappedNamedRange,
    offsetGridRange,
} from "../../common/gas-utils";
import { MergeType, PasteType } from "../../common/gas-utils/api-types";
import type { ReportPersistentData } from "../../common/report-utils";
import {
    calculateAttendanceGridSize,
    calculateCalendarHeaders,
    calculatePerClassLayout,
    generateStudentGrid,
    type Range,
    type Trimesters,
} from "../../common/setup-utils";
import { getDistinctHues, getRandomId, getUpperBoundIndex } from "../../common/utils";

type RangeName = ExtractRangeNames<typeof ReportSheetSchema>;

/**
 * Generates batch update `addNamedRange` request.
 */
function buildAddNamedRangeRequest<T extends NestedSheetSchema>(
    name: ExtractRangeNames<T>,
    range: GoogleAppsScript.Sheets.Schema.GridRange,
    namedRangeId?: string,
): GoogleAppsScript.Sheets.Schema.Request {
    return {
        addNamedRange: {
            namedRange: {
                namedRangeId,
                name,
                range,
            },
        },
    };
}

/**
 * Initializes an attendance sheet with the persistent data.
 */
export function createAttendanceSheet(
    parsedReport: ParsedSpreadsheet<typeof ReportSheetSchema>,
    persistentData: ReportPersistentData,
): GoogleAppsScript.Sheets.Schema.Request[] {
    const requests: GoogleAppsScript.Sheets.Schema.Request[] = [];

    const frozenArea = getFrozenArea(parsedReport.mappedRanges);
    const trimesters = calculateTrimesters(persistentData, frozenArea);

    // Copy Attendance template
    const attendanceSheetId = getRandomId(new Set());

    const copyTemplateRequests = copyAttendanceTemplate(parsedReport, persistentData, attendanceSheetId);

    // Add Attendance dates
    const datesRequests = addDateHeaders(attendanceSheetId, parsedReport.mappedRanges, persistentData.calendar);

    // Add student list(s)
    const { requests: studentListRequest, formatRanges, writableRanges } = addStudentLists(attendanceSheetId, parsedReport.mappedRanges, persistentData, trimesters);

    // Put format in the main edit area.
    const formatRequests = formatMainArea(attendanceSheetId, formatRanges, parsedReport.mappedRanges, trimesters);

    // Protect sheet.
    const protectRequests = protectSheet(attendanceSheetId, writableRanges);

    // Build requests
    requests.push(...copyTemplateRequests, ...datesRequests, ...studentListRequest, ...formatRequests, ...protectRequests);

    return requests;
}

/**
 * Copies the template and adjusts the properties of the new sheet.
 */
function copyAttendanceTemplate(
    parsedReport: ParsedSpreadsheet<typeof ReportSheetSchema>,
    persistenData: ReportPersistentData,
    attendaceSheetId: number,
): GoogleAppsScript.Sheets.Schema.Request[] {
    const getMappedSheet = createRequiredGetter(parsedReport.mappedSheets, "hoja de reporte");

    const attendanceTemplateSheetId = getMappedSheet(ReportSheetSchema.sheets.attendanceTemplate.sheetName).properties?.sheetId ?? 0;

    const requests: GoogleAppsScript.Sheets.Schema.Request[] = [];

    const { frozenRows, frozenCols } = getFrozenArea(parsedReport.mappedRanges);

    const { finalRowCount, finalColumnCount } = calculateAttendanceGridSize(
        frozenRows,
        frozenCols,
        persistenData.calendar.length,
        persistenData.students.length,
        persistenData.subjects.length,
        persistenData.configData.attendancePerClass,
    );

    // Temporarly remove namedRanges so they don't get copied.
    const templateNamedRanges = parsedReport.mappedSheetNamedRanges[ReportSheetSchema.sheets.attendanceTemplate.sheetName] ?? [];
    for (const namedRange of templateNamedRanges) requests.push({ deleteNamedRange: { namedRangeId: namedRange.namedRangeId } });

    // Make the new sheet
    requests.push({
        duplicateSheet: {
            sourceSheetId: attendanceTemplateSheetId,
            newSheetId: attendaceSheetId,
            newSheetName: ReportSheetSchema.sheets.attendance.sheetName,
        },
    });

    // Adjust properties
    requests.push(
        buildUpdateSheetPropertiesRequest({
            sheetId: attendaceSheetId,
            hidden: false,
            rowCount: finalRowCount,
            columnCount: finalColumnCount,
            index: 0,
        }),
    );

    // Unmerge cells
    requests.push({
        unmergeCells: {
            range: {
                sheetId: attendaceSheetId,
                startRowIndex: 0,
                endRowIndex: 1,
                startColumnIndex: frozenCols,
            },
        },
    });
    requests.push({
        unmergeCells: {
            range: {
                sheetId: attendaceSheetId,
                startRowIndex: frozenRows,
                startColumnIndex: 0,
                endColumnIndex: frozenCols,
            },
        },
    });

    // Return the namedRanges to the template
    for (const namedRange of templateNamedRanges) requests.push({ addNamedRange: { namedRange: namedRange } });

    // Hide the template sheet
    requests.push(
        buildUpdateSheetPropertiesRequest({
            sheetId: attendanceTemplateSheetId,
            hidden: true,
            index: 1,
        }),
    );

    // Return
    return requests;
}

/**
 * Adds the date headers to the attendance.
 */
function addDateHeaders(sheetId: number, namedRanges: Partial<Record<RangeName, MappedNamedRange>>, days: number[]): GoogleAppsScript.Sheets.Schema.Request[] {
    const rangeNames = ReportSheetSchema.sheets.attendanceTemplate.ranges;
    const getMappedRange = createRequiredGetter(namedRanges, "rango de asistencia");

    const requests: GoogleAppsScript.Sheets.Schema.Request[] = [];
    if (days.length === 0) return requests;

    const { frozenRows, frozenCols } = getFrozenArea(namedRanges);

    // Helper to safely extract 12 month strings from a specific template range
    const extractMonthNames = (rangeName: keyof typeof namedRanges, colStep: number): string[] => {
        const mappedRange = getMappedRange(rangeName);
        const data = getCellDataArray(mappedRange);
        const months: string[] = [];
        for (let i = 0; i < 12; i++) {
            const cellText = data?.[0]?.[i * colStep]?.effectiveValue?.stringValue ?? "";
            months.push(cellText);
        }
        return months;
    };

    const names1 = extractMonthNames(rangeNames.monthNames1, 1);
    const names2 = extractMonthNames(rangeNames.monthNames2, 2);
    const names5 = extractMonthNames(rangeNames.monthNames5, 5);

    const dayNames = getCellDataArray(getMappedRange(rangeNames.dayNames))[0]?.map((cellData) => cellData.effectiveValue?.stringValue ?? "");
    if (!dayNames) throw new Error("Faltan los nombres de los días.");

    const { monthGroups, row1Values, row2Values, row3Values } = calculateCalendarHeaders(days, frozenCols, names1, names2, names5, dayNames);

    // Format requests to run before filling with data.
    const formatRequests: GoogleAppsScript.Sheets.Schema.Request[] = [];

    // Format row 2
    const row2FormatSource = offsetGridRange({
        origin: namedRanges[rangeNames.dayNames]?.namedRange.range ?? {},
        height: 1,
        width: 1,
    });
    const row2FormatDest = createRange(sheetId, 1, frozenCols, 1, -1);
    formatRequests.push(buildCopyPasteRequest(row2FormatSource, row2FormatDest, PasteType.PASTE_FORMAT));

    for (const group of monthGroups) {
        let origin = getMappedRange(rangeNames.monthNames5).namedRange.range;
        let width = 5;

        if (group.template === 1) {
            origin = getMappedRange(rangeNames.monthNames1).namedRange.range;
            width = 1;
        } else if (group.template === 2) {
            origin = getMappedRange(rangeNames.monthNames2).namedRange.range;
            width = 2;
        }

        const colOffset = group.month * width;

        const source = offsetGridRange({ origin, colOffset, width, height: 1 });
        const dest = createSingleCellRange(sheetId, 0, group.startCol);

        formatRequests.push(buildCopyPasteRequest(source, dest, PasteType.PASTE_NORMAL));

        if (group.count !== 5 && group.count > 2) {
            const mergeRange = createRange(sheetId, 0, group.startCol, 1, group.count);
            formatRequests.push(buildMergeCellsRequest(mergeRange, MergeType.MERGE_ALL));
        }
    }

    const row1Data = row1Values.map((val): GoogleAppsScript.Sheets.Schema.CellData => (val != null ? { userEnteredValue: { stringValue: val } } : {}));
    const row2Data = row2Values.map((val): GoogleAppsScript.Sheets.Schema.CellData => ({ userEnteredValue: { stringValue: val } }));
    const row3Data = row3Values.map((val): GoogleAppsScript.Sheets.Schema.CellData => ({ userEnteredValue: { numberValue: val } }));

    const dataDestination = createRange(sheetId, 0, frozenCols, frozenRows, days.length);
    const bulkDataTransfer = buildUpdateCellsRequest({
        destination: dataDestination,
        data: [row1Data, row2Data, row3Data],
        fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.CellData>("userEnteredValue"),
    });

    requests.push(...formatRequests);
    if (bulkDataTransfer) requests.push(bulkDataTransfer);

    return requests;
}

interface AddStrudentListResponse {
    requests: GoogleAppsScript.Sheets.Schema.Request[];
    formatRanges: FormatRanges;
    writableRanges: GoogleAppsScript.Sheets.Schema.GridRange[];
}

interface FormatRanges {
    trim1: GoogleAppsScript.Sheets.Schema.GridRange[];
    trim2: GoogleAppsScript.Sheets.Schema.GridRange[];
    trim3: GoogleAppsScript.Sheets.Schema.GridRange[];
}

/**
 * Adds the subject and student list.
 */
function addStudentLists(
    sheetId: number,
    mappedRanges: Partial<Record<RangeName, MappedNamedRange>>,
    data: ReportPersistentData,
    trimesters: Trimesters,
): AddStrudentListResponse {
    const requests: GoogleAppsScript.Sheets.Schema.Request[] = [];
    const writableRanges: GoogleAppsScript.Sheets.Schema.GridRange[] = [];
    const formatRanges: FormatRanges = { trim1: [], trim2: [], trim3: [] };

    const getMappedRange = createRequiredGetter(mappedRanges, "rango de asistencia");

    const { frozenRows, frozenCols } = getFrozenArea(mappedRanges);
    const { trim1Range, trim2Range, trim3Range } = trimesters;

    const studentListDataRequests: GoogleAppsScript.Sheets.Schema.Request[] = [];
    const studentListFormatRequests: GoogleAppsScript.Sheets.Schema.Request[] = [];
    const namedRangesRequests: GoogleAppsScript.Sheets.Schema.Request[] = [];

    const buildGrid = (initialRow: number) => generateStudentGrid(data.students, initialRow, trimesters);

    if (data.configData.attendancePerClass) {
        const subjectTitleFormatOrigin = getMappedRange(ReportSheetSchema.sheets.attendanceTemplate.ranges.spaceSubjectRows).namedRange.range;
        const studentRowFormatOrigin = getMappedRange(ReportSheetSchema.sheets.attendanceTemplate.ranges.attendanceStudentRow).namedRange.range;

        const hues = getDistinctHues(data.subjects.length, 0.4);
        const subjectStudentListData: GoogleAppsScript.Sheets.Schema.CellData[][] = [];

        const layouts = calculatePerClassLayout(data.subjects.length, data.students.length, frozenRows);

        for (const layout of layouts) {
            const weightedSubject = data.subjects[layout.subjectIndex];

            // Build data
            subjectStudentListData.push([], [{ userEnteredValue: { stringValue: weightedSubject?.subject } }]); // Space and subject name
            subjectStudentListData.push(...buildGrid(layout.studentStartRow));

            // Format Title
            const subjectTitleFormatDestination = createRange(sheetId, layout.titleFormatStartRow, 0, 2, frozenCols);
            studentListFormatRequests.push(buildCopyPasteRequest(subjectTitleFormatOrigin, subjectTitleFormatDestination, PasteType.PASTE_FORMAT));

            // Format Students
            const studentRowFormatDestination = createRange(sheetId, layout.studentStartRow, data.students.length, frozenCols);
            studentListFormatRequests.push(buildCopyPasteRequest(studentRowFormatOrigin, studentRowFormatDestination, PasteType.PASTE_FORMAT));

            // Banding
            const hue = hues[layout.subjectIndex] ?? 0;
            const studentListBandingDestination = createRange(sheetId, layout.bandingStartRow, 0, layout.bandingNumRows, -1);
            studentListFormatRequests.push(buildAddBandingRequest(studentListBandingDestination, createBanding(hue, true)));

            // Named ranges
            const strIndex = String(layout.subjectIndex).padStart(2, "0");
            if (trim1Range.start !== -1) {
                const rangeName = `${ReportSheetSchema.sheets.attendance.dynamicRanges.unprotectTrim1}_Mat${strIndex}`;
                const range = createRange(sheetId, layout.studentStartRow, trim1Range.start, data.students.length, trim1Range.end - trim1Range.start + 1);
                namedRangesRequests.push(buildAddNamedRangeRequest(rangeName, range));
                formatRanges.trim1.push(range);
                writableRanges.push(range);
            }
            if (trim2Range.start !== -1) {
                const rangeName = `${ReportSheetSchema.sheets.attendance.dynamicRanges.unprotectTrim2}_Mat${strIndex}`;
                const range = createRange(sheetId, layout.studentStartRow, trim2Range.start, data.students.length, trim2Range.end - trim2Range.start + 1);
                namedRangesRequests.push(buildAddNamedRangeRequest(rangeName, range));
                formatRanges.trim2.push(range);
            }
            if (trim3Range.start !== -1) {
                const rangeName = `${ReportSheetSchema.sheets.attendance.dynamicRanges.unprotectTrim3}_Mat${strIndex}`;
                const range = createRange(sheetId, layout.studentStartRow, trim3Range.start, data.students.length, trim3Range.end - trim3Range.start + 1);
                namedRangesRequests.push(buildAddNamedRangeRequest(rangeName, range));
                formatRanges.trim3.push(range);
            }
        }

        // Data transfer
        const space = data.students.length + 2;
        const subjectStudentListDataRange = createRange(sheetId, frozenRows, 0, space * data.subjects.length, frozenCols);

        const subjectStudentListDataTransferRequest = buildUpdateCellsRequest({
            destination: subjectStudentListDataRange,
            data: subjectStudentListData,
            fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.CellData>("userEnteredValue"),
        });
        if (subjectStudentListDataTransferRequest) studentListDataRequests.push(subjectStudentListDataTransferRequest);
    } else {
        const studentStartRow = frozenRows + 1;
        const studentListDataRange = createRange(sheetId, studentStartRow, 0, data.students.length, frozenCols);

        // Named Ranges
        if (trim1Range.start !== -1) {
            const range = createRange(sheetId, frozenRows + 1, trim1Range.start, data.students.length, trim1Range.end - trim1Range.start + 1);
            namedRangesRequests.push(buildAddNamedRangeRequest(`${ReportSheetSchema.sheets.attendance.dynamicRanges.unprotectTrim1}_General`, range));
            formatRanges.trim1.push(range);
            writableRanges.push(range);
        }
        if (trim2Range.start !== -1) {
            const range = createRange(sheetId, frozenRows + 1, trim2Range.start, data.students.length, trim2Range.end - trim2Range.start + 1);
            namedRangesRequests.push(buildAddNamedRangeRequest(`${ReportSheetSchema.sheets.attendance.dynamicRanges.unprotectTrim2}_General`, range));
            formatRanges.trim2.push(range);
        }
        if (trim3Range.start !== -1) {
            const range = createRange(sheetId, frozenRows + 1, trim3Range.start, data.students.length, trim3Range.end - trim3Range.start + 1);
            namedRangesRequests.push(buildAddNamedRangeRequest(`${ReportSheetSchema.sheets.attendance.dynamicRanges.unprotectTrim3}_General`, range));
            formatRanges.trim3.push(range);
        }

        const subjectStudentListDataTransferRequest = buildUpdateCellsRequest({
            destination: studentListDataRange,
            data: buildGrid(studentStartRow),
            fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.CellData>("userEnteredValue"),
        });
        if (subjectStudentListDataTransferRequest) studentListDataRequests.push(subjectStudentListDataTransferRequest);

        const studentRow = getMappedRange(ReportSheetSchema.sheets.attendanceTemplate.ranges.attendanceStudentRow).namedRange.range;
        studentListFormatRequests.push(buildCopyPasteRequest(studentRow, studentListDataRange, PasteType.PASTE_FORMAT));

        // Banding
        studentListFormatRequests.push(buildAddBandingRequest(studentListDataRange, createBanding(0.4)));
    }

    requests.push(...studentListDataRequests, ...studentListFormatRequests, ...namedRangesRequests);
    return { requests, formatRanges, writableRanges };
}

/**
 * Copies the trimester format on the edit area.
 */
function formatMainArea(
    sheetId: number,
    formatRanges: FormatRanges,
    mappedRanges: Partial<Record<RangeName, MappedNamedRange>>,
    trimesters: Trimesters,
): GoogleAppsScript.Sheets.Schema.Request[] {
    const requests: GoogleAppsScript.Sheets.Schema.Request[] = [];

    const getMappedRange = createRequiredGetter(mappedRanges, "rango de asistencia");

    const { frozenRows, frozenCols } = getFrozenArea(mappedRanges);

    // Clean up format
    const noFormatSource = getMappedRange(ReportSheetSchema.sheets.attendanceTemplate.ranges.formatAttendanceCell).namedRange.range;
    const noFormatDest = createRange(sheetId, frozenRows, frozenCols);
    requests.push(buildCopyPasteRequest(noFormatSource, noFormatDest, PasteType.PASTE_NORMAL));

    const trimesterConfig = [
        { schemaRange: ReportSheetSchema.sheets.attendanceTemplate.ranges.formatPeriod1, targetRanges: formatRanges.trim1, trimStart: trimesters.trim1Range.start },
        { schemaRange: ReportSheetSchema.sheets.attendanceTemplate.ranges.formatPeriod2, targetRanges: formatRanges.trim2, trimStart: trimesters.trim2Range.start },
        { schemaRange: ReportSheetSchema.sheets.attendanceTemplate.ranges.formatPeriod3, targetRanges: formatRanges.trim3, trimStart: trimesters.trim3Range.start },
    ];

    for (const config of trimesterConfig) {
        if (config.trimStart !== 1) {
            const formatSource = getMappedRange(config.schemaRange).namedRange.range;
            for (const range of config.targetRanges) {
                requests.push(buildCopyPasteRequest(formatSource, range, PasteType.PASTE_FORMAT));
            }
        }
    }

    return requests;
}

/**
 * Protect the final Sheet.
 */
function protectSheet(sheetId: number, unprotectedRanges: GoogleAppsScript.Sheets.Schema.GridRange[]): GoogleAppsScript.Sheets.Schema.Request[] {
    const requests: GoogleAppsScript.Sheets.Schema.Request[] = [];

    requests.push({
        addProtectedRange: {
            protectedRange: {
                range: { sheetId },
                description: ReportSheetSchema.sheets.attendance.sheetName,
                warningOnly: false,
                unprotectedRanges,
            },
        },
    });

    return requests;
}

interface FrozenArea {
    rows: number;
    cols: number;
}

/**
 * Get the start of the editable area.
 */
function getFrozenArea(mappedRanges: Partial<Record<RangeName, MappedNamedRange>>): FrozenArea {
    const getRange = createRequiredGetter(mappedRanges, "rango de reporte");

    const mappedRange = getRange(ReportSheetSchema.sheets.attendanceTemplate.ranges.frozenArea);

    return {
        rows: mappedRange.namedRange.range.startRowIndex ?? 0,
        cols: mappedRange.namedRange.range.startColumnIndex ?? 0,
    };
}

/**
 * Calculate trimester column indexes.
 */
function calculateTrimesters(data: ReportPersistentData, frozenArea: FrozenArea): Trimesters {
    const frozenCols = frozenArea.cols;
    const format = (start: number, end: number): Range => (start <= end ? { start: start + frozenCols, end: end + frozenCols } : { start: -1, end: -1 });

    return {
        trim1Range: format(getUpperBoundIndex(data.calendar, data.configData.dates[0] - 1), getUpperBoundIndex(data.calendar, data.configData.dates[1]) - 1),
        trim2Range: format(getUpperBoundIndex(data.calendar, data.configData.dates[1]), getUpperBoundIndex(data.calendar, data.configData.dates[2]) - 1),
        trim3Range: format(getUpperBoundIndex(data.calendar, data.configData.dates[2]), getUpperBoundIndex(data.calendar, data.configData.dates[3]) - 1),
    };
}
