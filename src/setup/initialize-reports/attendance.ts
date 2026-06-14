import { ReportSheetSchema } from "../../common/gas-parts";
import type { ExtractRangeNames, ParsedSpreadsheet } from "../../common/gas-utils";
import {
    addNewNamedRange,
    addNewSheet,
    buildAddBandingRequest,
    buildCopyPasteRequest,
    buildFieldsMask,
    buildMergeCellsRequest,
    buildProtectSheetRequest,
    buildUnmergeCellsRequest,
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
import {
    calculateAssistanceTrimRanges as calculateAttendanceTrimRanges,
    type FrozenArea,
    type ReportPersistentData,
    type TrimesterRanges,
} from "../../common/report-utils";
import { calculateAttendanceGridSize, calculateCalendarHeaders, calculatePerClassLayout, generateStudentGrid, TemplateSize } from "../../common/setup-utils";
import { getDistinctHues } from "../../common/utils";

type RangeName = ExtractRangeNames<typeof ReportSheetSchema>;

/**
 * Initializes an attendance sheet with the persistent data.
 */
export function createAttendanceSheet(
    parsedReport: ParsedSpreadsheet<typeof ReportSheetSchema>,
    persistentData: ReportPersistentData,
): GoogleAppsScript.Sheets.Schema.Request[] {
    const requests: GoogleAppsScript.Sheets.Schema.Request[] = [];

    const frozenArea = getFrozenArea(parsedReport.mappedRanges);
    const trimesters = calculateAttendanceTrimRanges(persistentData, frozenArea);

    // Copy Attendance template
    const { requests: copyTemplateRequests, attendanceSheetId } = copyAttendanceTemplate(parsedReport, persistentData, frozenArea);

    // Add Attendance dates
    const datesRequests = addDateHeaders(attendanceSheetId, parsedReport.mappedRanges, persistentData.calendar, frozenArea);

    // Add student list(s)
    const { requests: studentListRequest, formatRanges, writableRanges } = addStudentLists(attendanceSheetId, parsedReport, persistentData, trimesters, frozenArea);

    // Put format in the main edit area.
    const formatRequests = formatMainArea(attendanceSheetId, formatRanges, parsedReport.mappedRanges, trimesters, frozenArea);

    // Protect sheet
    const protectRequest = buildProtectSheetRequest(parsedReport, ReportSheetSchema.sheets.attendance.sheetName, writableRanges);

    // Build requests
    requests.push(...copyTemplateRequests, ...datesRequests, ...studentListRequest, ...formatRequests, protectRequest);

    return requests;
}

/**
 * Copies the template and adjusts the properties of the new sheet.
 */
function copyAttendanceTemplate(
    parsedReport: ParsedSpreadsheet<typeof ReportSheetSchema>,
    persistenData: ReportPersistentData,
    frozenArea: FrozenArea,
): { requests: GoogleAppsScript.Sheets.Schema.Request[]; attendanceSheetId: number } {
    const getMappedSheet = createRequiredGetter(parsedReport.mappedSheets, "hoja de reporte");
    const attendanceTemplateSheetId = getMappedSheet(ReportSheetSchema.sheets.attendanceTemplate.sheetName).properties?.sheetId ?? 0;

    const { finalRowCount, finalColumnCount } = calculateAttendanceGridSize(
        frozenArea,
        persistenData.calendar.length,
        persistenData.students.length,
        persistenData.subjects.length,
        persistenData.configData.attendancePerClass,
    );

    const { requests, newSheetIds } = addNewSheet({
        parsedData: parsedReport,
        sourceSheetTitle: ReportSheetSchema.sheets.attendanceTemplate.sheetName,
        schemaSheetName: ReportSheetSchema.sheets.attendance.sheetName,
        insertSheetIndex: 0,
    });

    const attendanceSheetId = newSheetIds[0] ?? 0;

    // Adjust properties
    requests.push(
        buildUpdateSheetPropertiesRequest({
            sheetId: attendanceSheetId,
            hidden: false,
            rowCount: finalRowCount,
            columnCount: finalColumnCount,
            index: 0,
        }),
    );

    // Unmerge cells
    requests.push(
        buildUnmergeCellsRequest({
            sheetId: attendanceSheetId,
            startRowIndex: 0,
            endRowIndex: 1,
            startColumnIndex: frozenArea.cols,
        }),
    );

    requests.push(
        buildUnmergeCellsRequest({
            sheetId: attendanceSheetId,
            startRowIndex: frozenArea.rows,
            startColumnIndex: 0,
            endColumnIndex: frozenArea.cols,
        }),
    );

    // Hide the template sheet
    requests.push(
        buildUpdateSheetPropertiesRequest({
            sheetId: attendanceTemplateSheetId,
            hidden: true,
            index: 1,
        }),
    );

    // Return
    return { requests, attendanceSheetId };
}

/**
 * Adds the date headers to the attendance.
 */
function addDateHeaders(
    sheetId: number,
    namedRanges: Partial<Record<RangeName, MappedNamedRange>>,
    days: number[],
    frozenArea: FrozenArea,
): GoogleAppsScript.Sheets.Schema.Request[] {
    const rangeNames = ReportSheetSchema.sheets.attendanceTemplate.ranges;
    const getMappedRange = createRequiredGetter(namedRanges, "rango de asistencia");

    const requests: GoogleAppsScript.Sheets.Schema.Request[] = [];
    if (days.length === 0) return requests;

    // Helper to safely extract 12 month strings from a specific template range
    const extractMonthNames = (rangeName: keyof typeof namedRanges, colStep: number): string[] => {
        const mappedRange = getMappedRange(rangeName);
        const data = getCellDataArray(mappedRange);
        const months: string[] = [];
        for (let i = 0; i < 12; i++) {
            const cellText = data[0]?.[i * colStep]?.effectiveValue?.stringValue ?? "";
            months.push(cellText);
        }
        return months;
    };

    const names1 = extractMonthNames(rangeNames.monthNames1, 1);
    const names2 = extractMonthNames(rangeNames.monthNames2, 2);
    const names5 = extractMonthNames(rangeNames.monthNames5, 5);

    const dayNames = getCellDataArray(getMappedRange(rangeNames.dayNames))[0]?.map((cellData) => cellData.effectiveValue?.stringValue ?? "");
    if (!dayNames) throw new Error("Faltan los nombres de los días.");

    const { monthGroups, row1Values, row2Values, row3Values } = calculateCalendarHeaders(days, frozenArea.cols, names1, names2, names5, dayNames);

    // Format requests to run before filling with data.
    const formatRequests: GoogleAppsScript.Sheets.Schema.Request[] = [];

    // Format row 2
    const row2FormatSource = offsetGridRange({
        origin: namedRanges[rangeNames.dayNames]?.namedRange.range ?? {},
        height: 1,
        width: 1,
    });
    const row2FormatDest = createRange(sheetId, 1, frozenArea.cols, 1, -1);
    formatRequests.push(buildCopyPasteRequest(row2FormatSource, row2FormatDest, PasteType.PASTE_FORMAT));

    for (const group of monthGroups) {
        let origin = getMappedRange(rangeNames.monthNames5).namedRange.range;
        let width = 5;

        if (group.template === TemplateSize.SMALL) {
            origin = getMappedRange(rangeNames.monthNames1).namedRange.range;
            width = 1;
        } else if (group.template === TemplateSize.MEDIUM) {
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

    const dataDestination = createRange(sheetId, 0, frozenArea.cols, frozenArea.rows, days.length);
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
    parsedReport: ParsedSpreadsheet<typeof ReportSheetSchema>,
    data: ReportPersistentData,
    trimesters: TrimesterRanges,
    frozenArea: FrozenArea,
): AddStrudentListResponse {
    const requests: GoogleAppsScript.Sheets.Schema.Request[] = [];
    const formatRanges: FormatRanges = { trim1: [], trim2: [], trim3: [] };
    const writableRanges: GoogleAppsScript.Sheets.Schema.GridRange[] = [];

    const getMappedRange = createRequiredGetter(parsedReport.mappedRanges, "rango de asistencia");

    const studentListDataRequests: GoogleAppsScript.Sheets.Schema.Request[] = [];
    const studentListFormatRequests: GoogleAppsScript.Sheets.Schema.Request[] = [];
    const namedRangesRequests: GoogleAppsScript.Sheets.Schema.Request[] = [];

    const buildGrid = (initialRow: number) => generateStudentGrid(data.students, initialRow, trimesters);

    if (data.configData.attendancePerClass) {
        const subjectTitleFormatOrigin = getMappedRange(ReportSheetSchema.sheets.attendanceTemplate.ranges.spaceSubjectRows).namedRange.range;
        const studentRowFormatOrigin = getMappedRange(ReportSheetSchema.sheets.attendanceTemplate.ranges.attendanceStudentRow).namedRange.range;

        const hues = getDistinctHues(data.subjects.length, 0.4);
        const subjectStudentListData: GoogleAppsScript.Sheets.Schema.CellData[][] = [];

        const layouts = calculatePerClassLayout(data.subjects.length, data.students.length, frozenArea.rows);

        for (const layout of layouts) {
            const weightedSubject = data.subjects[layout.subjectIndex];

            // Build data
            subjectStudentListData.push([], [{ userEnteredValue: { stringValue: weightedSubject?.subject } }]); // Space and subject name
            subjectStudentListData.push(...buildGrid(layout.studentStartRow));

            // Format Title
            const subjectTitleFormatDestination = createRange(sheetId, layout.titleFormatStartRow, 0, 2, frozenArea.cols);
            studentListFormatRequests.push(buildCopyPasteRequest(subjectTitleFormatOrigin, subjectTitleFormatDestination, PasteType.PASTE_FORMAT));

            // Format Students
            const studentRowFormatDestination = createRange(sheetId, layout.studentStartRow, data.students.length, frozenArea.cols);
            studentListFormatRequests.push(buildCopyPasteRequest(studentRowFormatOrigin, studentRowFormatDestination, PasteType.PASTE_FORMAT));

            // Banding
            const hue = hues[layout.subjectIndex] ?? 0;
            const studentListBandingDestination = createRange(sheetId, layout.bandingStartRow, 0, layout.bandingNumRows, -1);
            studentListFormatRequests.push(buildAddBandingRequest(studentListBandingDestination, createBanding(hue, true)));

            // Borders
            // TODO: Add `innerVertical` borders to colums 3 - frozen

            // Named ranges
            const strIndex = String(layout.subjectIndex).padStart(2, "0");
            if (trimesters.trim1.start !== -1) {
                const rangeName = `${ReportSheetSchema.sheets.attendance.dynamicRanges.unprotectTrim1}_Mat${strIndex}`;
                const gridRange = createRange(
                    sheetId,
                    layout.studentStartRow,
                    trimesters.trim1.start,
                    data.students.length,
                    trimesters.trim1.end - trimesters.trim1.start + 1,
                );
                namedRangesRequests.push(
                    addNewNamedRange({
                        parsedData: parsedReport,
                        sheetTitle: ReportSheetSchema.sheets.attendance.sheetName,
                        gridRange,
                        rangeName,
                        dynamicRangeKey: ReportSheetSchema.sheets.attendance.dynamicRanges.unprotectTrim1,
                    }),
                );
                formatRanges.trim1.push(gridRange);
                writableRanges.push(gridRange);
            }
            if (trimesters.trim2.start !== -1) {
                const rangeName = `${ReportSheetSchema.sheets.attendance.dynamicRanges.unprotectTrim2}_Mat${strIndex}`;
                const gridRange = createRange(
                    sheetId,
                    layout.studentStartRow,
                    trimesters.trim2.start,
                    data.students.length,
                    trimesters.trim2.end - trimesters.trim2.start + 1,
                );
                namedRangesRequests.push(
                    addNewNamedRange({
                        parsedData: parsedReport,
                        sheetTitle: ReportSheetSchema.sheets.attendance.sheetName,
                        gridRange,
                        rangeName,
                        dynamicRangeKey: ReportSheetSchema.sheets.attendance.dynamicRanges.unprotectTrim2,
                    }),
                );
                formatRanges.trim2.push(gridRange);
            }
            if (trimesters.trim3.start !== -1) {
                const rangeName = `${ReportSheetSchema.sheets.attendance.dynamicRanges.unprotectTrim3}_Mat${strIndex}`;
                const gridRange = createRange(
                    sheetId,
                    layout.studentStartRow,
                    trimesters.trim3.start,
                    data.students.length,
                    trimesters.trim3.end - trimesters.trim3.start + 1,
                );
                namedRangesRequests.push(
                    addNewNamedRange({
                        parsedData: parsedReport,
                        sheetTitle: ReportSheetSchema.sheets.attendance.sheetName,
                        gridRange,
                        rangeName,
                        dynamicRangeKey: ReportSheetSchema.sheets.attendance.dynamicRanges.unprotectTrim3,
                    }),
                );
                formatRanges.trim3.push(gridRange);
            }
        }

        // Data transfer
        const space = data.students.length + 2;
        const subjectStudentListDataRange = createRange(sheetId, frozenArea.rows, 0, space * data.subjects.length, frozenArea.cols);

        const subjectStudentListDataTransferRequest = buildUpdateCellsRequest({
            destination: subjectStudentListDataRange,
            data: subjectStudentListData,
            fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.CellData>("userEnteredValue"),
        });
        if (subjectStudentListDataTransferRequest) studentListDataRequests.push(subjectStudentListDataTransferRequest);
    } else {
        const studentStartRow = frozenArea.rows + 1;
        const studentListDataRange = createRange(sheetId, studentStartRow, 0, data.students.length, frozenArea.cols);

        // Named Ranges
        if (trimesters.trim1.start !== -1) {
            const gridRange = createRange(sheetId, frozenArea.rows + 1, trimesters.trim1.start, data.students.length, trimesters.trim1.end - trimesters.trim1.start + 1);
            const rangeName = `${ReportSheetSchema.sheets.attendance.dynamicRanges.unprotectTrim1}_General`;
            namedRangesRequests.push(
                addNewNamedRange({
                    parsedData: parsedReport,
                    sheetTitle: ReportSheetSchema.sheets.attendance.sheetName,
                    gridRange,
                    rangeName,
                    dynamicRangeKey: ReportSheetSchema.sheets.attendance.dynamicRanges.unprotectTrim1,
                }),
            );
            formatRanges.trim1.push(gridRange);
            writableRanges.push(gridRange);
        }
        if (trimesters.trim2.start !== -1) {
            const gridRange = createRange(sheetId, frozenArea.rows + 1, trimesters.trim2.start, data.students.length, trimesters.trim2.end - trimesters.trim2.start + 1);
            const rangeName = `${ReportSheetSchema.sheets.attendance.dynamicRanges.unprotectTrim2}_General`;
            namedRangesRequests.push(
                addNewNamedRange({
                    parsedData: parsedReport,
                    sheetTitle: ReportSheetSchema.sheets.attendance.sheetName,
                    gridRange,
                    rangeName,
                    dynamicRangeKey: ReportSheetSchema.sheets.attendance.dynamicRanges.unprotectTrim2,
                }),
            );
            formatRanges.trim2.push(gridRange);
        }
        if (trimesters.trim3.start !== -1) {
            const gridRange = createRange(sheetId, frozenArea.rows + 1, trimesters.trim3.start, data.students.length, trimesters.trim3.end - trimesters.trim3.start + 1);
            const rangeName = `${ReportSheetSchema.sheets.attendance.dynamicRanges.unprotectTrim3}_General`;
            namedRangesRequests.push(
                addNewNamedRange({
                    parsedData: parsedReport,
                    sheetTitle: ReportSheetSchema.sheets.attendance.sheetName,
                    gridRange,
                    rangeName,
                    dynamicRangeKey: ReportSheetSchema.sheets.attendance.dynamicRanges.unprotectTrim3,
                }),
            );
            formatRanges.trim3.push(gridRange);
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
    trimesters: TrimesterRanges,
    frozenArea: FrozenArea,
): GoogleAppsScript.Sheets.Schema.Request[] {
    const requests: GoogleAppsScript.Sheets.Schema.Request[] = [];

    const getMappedRange = createRequiredGetter(mappedRanges, "rango de asistencia");

    // Add basic format (validation and conditional format)
    const baseFormatSource = getMappedRange(ReportSheetSchema.sheets.attendanceTemplate.ranges.formatAttendanceCell).namedRange.range;
    const baseFormatDest = createRange(sheetId, frozenArea.rows, frozenArea.cols);
    requests.push(buildCopyPasteRequest(baseFormatSource, baseFormatDest, PasteType.PASTE_NORMAL));

    const trimesterConfig: Array<{ schemaRange: RangeName; targetRanges: GoogleAppsScript.Sheets.Schema.GridRange[]; trimStart: number }> = [
        { schemaRange: ReportSheetSchema.sheets.attendanceTemplate.ranges.formatPeriod1, targetRanges: formatRanges.trim1, trimStart: trimesters.trim1.start },
        { schemaRange: ReportSheetSchema.sheets.attendanceTemplate.ranges.formatPeriod2, targetRanges: formatRanges.trim2, trimStart: trimesters.trim2.start },
        { schemaRange: ReportSheetSchema.sheets.attendanceTemplate.ranges.formatPeriod3, targetRanges: formatRanges.trim3, trimStart: trimesters.trim3.start },
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
