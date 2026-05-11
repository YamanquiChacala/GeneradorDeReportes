import { BASE_ASSISTANCE_PROTECTED_RANGE } from "../../common/constants";
import { MergeType, PasteType } from "../../common/gas-enums";
import { ReportSheetSchema } from "../../common/sheet-schema";
import { getDistinctHues } from "../../common/utils/color-utils";
import { buildFieldsMask } from "../../common/utils/gas-types";
import {
    buildCopyPasteRequest,
    buildTransferRequests,
    createBanding,
    createRange,
    createSingleCellRange,
    getColumnLetter,
    offsetGridRange,
} from "../../common/utils/gas-utils";
import { type ExtractRangeNames, MappedNamedRange, type ParsedSpreadsheet } from "../../common/utils/mapped-name-range";
import type { ReportPersistentData } from "../../common/utils/report-utils";

type RangeName = ExtractRangeNames<typeof ReportSheetSchema>;

/**
 * Initializes an attendance sheet with the persistent data.
 */
export function createAttendanceSheet(
    parsedReport: ParsedSpreadsheet<typeof ReportSheetSchema>,
    persistentData: ReportPersistentData,
): GoogleAppsScript.Sheets.Schema.Request[] {
    const requests: GoogleAppsScript.Sheets.Schema.Request[] = [];

    // Copy Attendance template

    const attendanceSheetId = Math.floor(Math.random() * (2 ** 31 - 1));

    const { frozenCols } = getFrozenRowCols(parsedReport.namedRanges);
    const trimesters = calculateTrimesters(persistentData, frozenCols);

    const copyTemplateRequests = copyAttendanceTemplate(parsedReport, persistentData, attendanceSheetId);

    // Add Attendance dates
    const datesRequests = addDateHeaders(attendanceSheetId, parsedReport.namedRanges, persistentData.calendar);

    // Add student list(s)

    const { requests: studentListRequest, writableRanges } = addStudentLists(attendanceSheetId, parsedReport.namedRanges, persistentData, trimesters);

    // Put format in the main edit area.

    const formatRequests = formatMainArea(attendanceSheetId, writableRanges, parsedReport.namedRanges, trimesters);

    // Protect sheet.

    const protectRequests = protectSheet(attendanceSheetId, writableRanges);

    // Build requests

    requests.push(...copyTemplateRequests, ...datesRequests, ...studentListRequest, ...formatRequests, ...protectRequests);

    return requests;
}

/**
 * Get the start of the editable area.
 */
function getFrozenRowCols(namedRanges: Partial<Record<RangeName, MappedNamedRange>>): { frozenRows: number; frozenCols: number } {
    const range = namedRanges[ReportSheetSchema.sheets.attendanceTemplate.ranges.frozenArea]?.range;

    if (!range) throw new Error("Falta limite de formato.");

    return {
        frozenRows: range.startRowIndex ?? 0,
        frozenCols: range.startColumnIndex ?? 0,
    };
}

/**
 * Copies the template and adjusts the properties of the new sheet.
 */
function copyAttendanceTemplate(
    parsedReport: ParsedSpreadsheet<typeof ReportSheetSchema>,
    persistenData: ReportPersistentData,
    attendaceSheetId: number,
): GoogleAppsScript.Sheets.Schema.Request[] {
    const requests: GoogleAppsScript.Sheets.Schema.Request[] = [];

    const { frozenRows, frozenCols } = getFrozenRowCols(parsedReport.namedRanges);

    const finalColumnCount = frozenCols + persistenData.calendar.length;
    let finalRowCount = frozenRows;
    if (persistenData.configData.attendancePerClass) {
        finalRowCount += (2 + persistenData.students.length) * persistenData.subjects.length;
    } else {
        finalRowCount += 1 + persistenData.students.length;
    }

    // Temporarly remove namedRanges so they don't get copied.
    const templateNamedRanges = parsedReport.sheetNamedRanges[ReportSheetSchema.sheets.attendanceTemplate.sheetName] ?? [];
    for (const namedRange of templateNamedRanges) requests.push({ deleteNamedRange: { namedRangeId: namedRange.namedRangeId } });

    // Make the new sheet
    requests.push({
        duplicateSheet: {
            sourceSheetId: parsedReport.sheets[ReportSheetSchema.sheets.attendanceTemplate.sheetName]?.properties?.sheetId,
            newSheetId: attendaceSheetId,
            newSheetName: ReportSheetSchema.sheets.attendance.sheetName,
        },
    });

    // Adjust properties
    requests.push({
        updateSheetProperties: {
            properties: {
                sheetId: attendaceSheetId,
                hidden: false,
                gridProperties: { columnCount: finalColumnCount, rowCount: finalRowCount },
            },
            fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.SheetProperties>("hidden", "gridProperties.columnCount", "gridProperties.rowCount", "index"),
        },
    });

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

    // Return
    return requests;
}

/**
 * Adds the date headers to the attendance.
 */
function addDateHeaders(sheetId: number, namedRanges: Partial<Record<RangeName, MappedNamedRange>>, days: number[]): GoogleAppsScript.Sheets.Schema.Request[] {
    const ranges = ReportSheetSchema.sheets.attendanceTemplate.ranges;

    const requests: GoogleAppsScript.Sheets.Schema.Request[] = [];
    if (days.length === 0) return requests;

    const { frozenRows, frozenCols } = getFrozenRowCols(namedRanges);

    // Helper to safely extract 12 month strings from a specific template range
    const extractMonthNames = (rangeName: keyof typeof namedRanges, colStep: number): string[] => {
        const mappedRange = namedRanges[rangeName];
        const data = MappedNamedRange.getCellDataArray(mappedRange);
        const months: string[] = [];
        for (let i = 0; i < 12; i++) {
            const cellText = data?.[0]?.[i * colStep]?.effectiveValue?.stringValue ?? "";
            months.push(cellText);
        }
        return months;
    };

    const names1 = extractMonthNames(ranges.monthNames1, 1);
    const names2 = extractMonthNames(ranges.monthNames2, 2);
    const names5 = extractMonthNames(ranges.monthNames5, 5);

    const dayNames = MappedNamedRange.getCellDataArray(namedRanges[ranges.dayNames])[0]?.map((cellData) => cellData.effectiveValue?.stringValue ?? "");

    if (!dayNames) throw new Error("Faltan los nombres de los días.");

    interface MonthGroup {
        year: number;
        month: number;
        startCol: number;
        count: number;
    }

    let currentGroup: MonthGroup | null = null;

    // Arrays to collect our batched data values
    const row1Data: GoogleAppsScript.Sheets.Schema.CellData[] = [];
    const row2Data: GoogleAppsScript.Sheets.Schema.CellData[] = [];
    const row3Data: GoogleAppsScript.Sheets.Schema.CellData[] = [];

    // Format requests to run before filling with data.
    const formatRequests: GoogleAppsScript.Sheets.Schema.Request[] = [];

    // Format row 2
    const row2FormatSource = offsetGridRange({
        origin: namedRanges[ranges.dayNames]?.range ?? {},
        height: 1,
        width: 1,
    });
    const row2FormatDest = createRange(sheetId, 1, frozenCols, 1, -1);
    const row2FormatRequest = buildCopyPasteRequest(row2FormatSource, row2FormatDest, PasteType.PASTE_FORMAT);
    if (row2FormatRequest) formatRequests.push(row2FormatRequest);

    const processMonthGroup = (group: MonthGroup) => {
        const { year, month, startCol, count } = group;
        const shortYear = String(year).slice(-2);
        const longYear = String(year);

        let origin: GoogleAppsScript.Sheets.Schema.GridRange | undefined;
        let width = 1;
        let colOffset = 0;
        let text = "";

        if (count === 1) {
            origin = namedRanges[ranges.monthNames1]?.range;
            width = 1;
            colOffset = month * 1;
            text = `${names1[month]}\n${shortYear}`;
        } else if (count >= 2 && count <= 4) {
            origin = namedRanges[ranges.monthNames2]?.range;
            width = 2;
            colOffset = month * 2;
            text = `${names2[month]}\n${longYear}`;
        } else {
            origin = namedRanges[ranges.monthNames5]?.range;
            width = 5;
            colOffset = month * 5;
            text = `${names5[month]}\n${longYear}`;
        }

        if (!origin) throw new Error("Error copiando formato de asistencias.");

        const source = offsetGridRange({ origin, colOffset, width, height: 1, rowOffset: 0 });
        const dest = createSingleCellRange(sheetId, 0, startCol);

        // Collect the template format copies for Row 1
        const copyRequest = buildCopyPasteRequest(source, dest, PasteType.PASTE_NORMAL);
        if (copyRequest) formatRequests.push(copyRequest);

        if (count !== 5 && count > 2) {
            formatRequests.push({
                mergeCells: {
                    range: {
                        sheetId,
                        startRowIndex: 0,
                        endRowIndex: 1,
                        startColumnIndex: startCol,
                        endColumnIndex: startCol + count,
                    },
                    mergeType: MergeType.MERGE_ALL,
                },
            });
        }

        // Build the text structure for Row 1
        for (let i = 0; i < count; i++) {
            if (i === 0) {
                row1Data.push({ userEnteredValue: { stringValue: text } });
            } else {
                row1Data.push({}); // Push empty cell objects for merged trail columns
            }
        }
    };

    const dayNamesOrigin = namedRanges[ranges.dayNames]?.range;

    if (!dayNamesOrigin) throw new Error("Error copiando formato de dias.");

    for (let i = 0; i < days.length; i++) {
        const date = new Date(days[i] ?? 0);
        const year = date.getUTCFullYear();
        const month = date.getUTCMonth();
        const dayOfWeek = date.getUTCDay();
        const dayOfMonth = date.getUTCDate();

        const targetCol = frozenCols + i; // Start at Col J (index 9)

        // Group tracking for Row 1
        if (!currentGroup) {
            currentGroup = { year, month, startCol: targetCol, count: 1 };
        } else if (currentGroup.year === year && currentGroup.month === month) {
            currentGroup.count++;
        } else {
            processMonthGroup(currentGroup);
            currentGroup = { year, month, startCol: targetCol, count: 1 };
        }

        // Prepare Row 2 in the bulk transfer
        row2Data.push({ userEnteredValue: { stringValue: dayNames[dayOfWeek] } });

        // Collect data for Row 3 (Day Numbers)
        row3Data.push({ userEnteredValue: { numberValue: dayOfMonth } });
    }

    // Process the final lingering month group
    if (currentGroup) processMonthGroup(currentGroup);

    const totalColumns = days.length;

    // Combine Rows 1, 2, and 3 into a single highly optimized transfer
    const bulkDataTransfer = buildTransferRequests({
        destination: {
            sheetId,
            startRowIndex: 0,
            endRowIndex: frozenRows,
            startColumnIndex: frozenCols,
            endColumnIndex: frozenCols + totalColumns,
        },
        data: [row1Data, row2Data, row3Data], // The single 3-row payload
        fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.CellData>("userEnteredValue"),
        adaptRange: false,
    });

    requests.push(...formatRequests, ...bulkDataTransfer);

    return requests;
}

/**
 * Adds the subject and student list.
 */
function addStudentLists(
    sheetId: number,
    namedRanges: Partial<Record<RangeName, MappedNamedRange>>,
    data: ReportPersistentData,
    trimesters: Trimesters,
): { requests: GoogleAppsScript.Sheets.Schema.Request[]; writableRanges: GoogleAppsScript.Sheets.Schema.GridRange[] } {
    const requests: GoogleAppsScript.Sheets.Schema.Request[] = [];
    const writableRanges: GoogleAppsScript.Sheets.Schema.GridRange[] = [];

    const { frozenRows, frozenCols } = getFrozenRowCols(namedRanges);

    const { trim1Range, trim2Range, trim3Range } = trimesters;

    const studentGroup = (initialRow: number): GoogleAppsScript.Sheets.Schema.CellData[][] => {
        const result: GoogleAppsScript.Sheets.Schema.CellData[][] = [];
        for (const [i, studentRow] of data.students.entries()) {
            if (studentRow.type === "student") {
                const { percent: trim1Percent, count: trim1Count } = createAttendaceFormulas(initialRow + i, trim1Range);
                const { percent: trim2Percent, count: trim2Count } = createAttendaceFormulas(initialRow + i, trim2Range);
                const { percent: trim3Percent, count: trim3Count } = createAttendaceFormulas(initialRow + i, trim3Range);

                const resultRow: GoogleAppsScript.Sheets.Schema.CellData[] = [
                    { userEnteredValue: { numberValue: studentRow.id } },
                    { userEnteredValue: { stringValue: studentRow.firstName } },
                    { userEnteredValue: { stringValue: studentRow.lastName } },
                    { userEnteredValue: { formulaValue: trim1Percent } },
                    { userEnteredValue: { formulaValue: trim1Count } },
                    { userEnteredValue: { formulaValue: trim2Percent } },
                    { userEnteredValue: { formulaValue: trim2Count } },
                    { userEnteredValue: { formulaValue: trim3Percent } },
                    { userEnteredValue: { formulaValue: trim3Count } },
                ];
                result.push(resultRow);
            } else {
                result.push([]);
            }
        }
        return result;
    };

    const studentListDataRequests: GoogleAppsScript.Sheets.Schema.Request[] = [];
    const studentListFormatRequests: GoogleAppsScript.Sheets.Schema.Request[] = [];
    const namedRangesRequests: GoogleAppsScript.Sheets.Schema.Request[] = [];

    if (data.configData.attendancePerClass) {
        const subjectTitleFormatOrigin = namedRanges[ReportSheetSchema.sheets.attendanceTemplate.ranges.spaceSubjectRows]?.range;
        const studentRowFormatOrigin = namedRanges[ReportSheetSchema.sheets.attendanceTemplate.ranges.attendanceStudentRow]?.range;

        const space = data.students.length + 2;

        const hues = getDistinctHues(data.subjects.length, 0.4);

        const subjectStudentListData: GoogleAppsScript.Sheets.Schema.CellData[][] = [];
        for (const [index, weightedSubject] of data.subjects.entries()) {
            subjectStudentListData.push([], [{ userEnteredValue: { stringValue: weightedSubject.subject } }]); // Space and subject name.
            subjectStudentListData.push(...studentGroup(frozenRows + 2 + index * space));

            // Copy format for the subject title
            const subjectTitleFormatDestination = createRange(sheetId, frozenRows + index * space, 0, 2, frozenCols);
            const subjectTitleFormatCopyRequest = buildCopyPasteRequest(subjectTitleFormatOrigin, subjectTitleFormatDestination, PasteType.PASTE_FORMAT);
            if (subjectTitleFormatCopyRequest) studentListFormatRequests.push(subjectTitleFormatCopyRequest);

            // Copy format for the list of students
            const studentRowFormatDestination = createRange(sheetId, frozenRows + 2 + index * space, 0, data.students.length, frozenCols);
            const studentRowFormatCopyRequest = buildCopyPasteRequest(studentRowFormatOrigin, studentRowFormatDestination, PasteType.PASTE_FORMAT);
            if (studentRowFormatCopyRequest) studentListFormatRequests.push(studentRowFormatCopyRequest);

            // Put banding on
            const hue = hues[index] ?? 0;
            studentListFormatRequests.push({
                addBanding: {
                    bandedRange: {
                        range: createRange(sheetId, frozenRows + 1 + index * space, 0, data.students.length + 1, -1),
                        rowProperties: createBanding(hue, true),
                    },
                },
            });

            // Named Ranges
            if (trim1Range.start !== -1) {
                const trim1NamedRange = createRange(
                    sheetId,
                    frozenRows + 2 + index * space,
                    trim1Range.start,
                    data.students.length,
                    trim1Range.end - trim1Range.start + 1,
                );
                namedRangesRequests.push(createAddNamedRangeRequest(trim1NamedRange, 0, `Mat_${String(index).padStart(2, "0")}`));
                writableRanges.push(trim1NamedRange);
            }
            if (trim2Range.start !== -1) {
                const trim2NamedRange = createRange(
                    sheetId,
                    frozenRows + 2 + index * space,
                    trim2Range.start,
                    data.students.length,
                    trim2Range.end - trim2Range.start + 1,
                );
                namedRangesRequests.push(createAddNamedRangeRequest(trim2NamedRange, 1, `Mat_${String(index).padStart(2, "0")}`));
            }
            if (trim3Range.start !== -1) {
                const trim3NamedRange = createRange(
                    sheetId,
                    frozenRows + 2 + index * space,
                    trim3Range.start,
                    data.students.length,
                    trim3Range.end - trim3Range.start + 1,
                );
                namedRangesRequests.push(createAddNamedRangeRequest(trim3NamedRange, 2, `Mat_${String(index).padStart(2, "0")}`));
            }
        }

        // Copy the data for all the subjects and students.
        const subjectStudentListDataRange = createRange(sheetId, frozenRows, 0, space * data.subjects.length, frozenCols);
        studentListDataRequests.push(
            ...buildTransferRequests({
                destination: subjectStudentListDataRange,
                data: subjectStudentListData,
                fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.CellData>("userEnteredValue"),
            }),
        );
    } else {
        const studentListDataRange = createRange(sheetId, frozenRows + 1, 0, data.students.length, frozenCols);

        // Named Ranges
        if (trim1Range.start !== -1) {
            const trim1NamedRange = createRange(sheetId, frozenRows + 1, trim1Range.start, data.students.length, trim1Range.end - trim1Range.start + 1);
            namedRangesRequests.push(createAddNamedRangeRequest(trim1NamedRange, 0, "General"));
            writableRanges.push(trim1NamedRange);
        }
        if (trim2Range.start !== -1) {
            const trim2NamedRange = createRange(sheetId, frozenRows + 1, trim2Range.start, data.students.length, trim2Range.end - trim2Range.start + 1);
            namedRangesRequests.push(createAddNamedRangeRequest(trim2NamedRange, 1, "General"));
        }
        if (trim3Range.start !== -1) {
            const trim3NamedRange = createRange(sheetId, frozenRows + 1, trim3Range.start, data.students.length, trim3Range.end - trim3Range.start + 1);
            namedRangesRequests.push(createAddNamedRangeRequest(trim3NamedRange, 2, "General"));
        }

        studentListDataRequests.push(
            ...buildTransferRequests({
                destination: studentListDataRange,
                data: studentGroup(frozenRows + 1),
                fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.CellData>("userEnteredValue"),
            }),
        );

        const studentRow = namedRanges[ReportSheetSchema.sheets.attendanceTemplate.ranges.attendanceStudentRow]?.range;

        const copyRequest = buildCopyPasteRequest(studentRow, studentListDataRange, PasteType.PASTE_FORMAT);
        if (copyRequest) studentListFormatRequests.push(copyRequest);

        // Banding

        studentListFormatRequests.push({
            addBanding: {
                bandedRange: {
                    range: studentListDataRange,
                    rowProperties: createBanding(0.4),
                },
            },
        });
    }
    requests.push(...studentListDataRequests, ...studentListFormatRequests, ...namedRangesRequests);

    return { requests, writableRanges };
}

/**
 * Helper function to create a named range
 */
function createAddNamedRangeRequest(range: GoogleAppsScript.Sheets.Schema.GridRange, period: 0 | 1 | 2, subject: string): GoogleAppsScript.Sheets.Schema.Request {
    return {
        addNamedRange: {
            namedRange: {
                name: `${BASE_ASSISTANCE_PROTECTED_RANGE}${period}_${subject}`,
                range,
            },
        },
    };
}

interface Range {
    start: number;
    end: number;
}

interface Trimesters {
    trim1Range: Range;
    trim2Range: Range;
    trim3Range: Range;
}

/**
 * Calculate trimester column indexes.
 */
function calculateTrimesters(data: ReportPersistentData, frozenCols: number): Trimesters {
    const upperBound = (target: number) => {
        let left = 0;
        let right = data.calendar.length;
        while (left < right) {
            const mid = Math.floor((left + right) / 2);
            // biome-ignore lint/style/noNonNullAssertion: We know we're inside the array bounds.
            if (data.calendar[mid]! > target) right = mid;
            else left = mid + 1;
        }
        return left;
    };

    const format = (start: number, end: number): Range => (start <= end ? { start: start + frozenCols, end: end + frozenCols } : { start: -1, end: -1 });

    return {
        trim1Range: format(upperBound(data.configData.dateStart - 1), upperBound(data.configData.dateTrim1) - 1),
        trim2Range: format(upperBound(data.configData.dateTrim1), upperBound(data.configData.dateTrim2) - 1),
        trim3Range: format(upperBound(data.configData.dateTrim2), upperBound(data.configData.dateEnd) - 1),
    };
}

/**
 * Define the formala for calculating the attendance percent.
 */
function createAttendaceFormulas(row: number, columms: Range): { percent: string; count: string } {
    if (columms.start === -1 || columms.start > columms.end) return { percent: '="ND"', count: '=""' };

    const str = getColumnLetter(columms.start);
    const end = getColumnLetter(columms.end);

    const range = `$${str}${row + 1}:$${end}${row + 1}`;

    const percent = `=IF(COUNTA(${range}) >= COLUMNS(${range}) / 10, (COUNTA(${range}) - SUM(${range})) / COUNTA(${range}), "ND")`;
    const count = `=IF(COUNTA(${range}) >= COLUMNS(${range}) / 10, SUM(${range}), "")`;

    return { percent, count };
}

/**
 * Copies the trimester format on the edit area.
 */
function formatMainArea(
    sheetId: number,
    editableRanges: GoogleAppsScript.Sheets.Schema.GridRange[],
    namedRanges: Partial<Record<RangeName, MappedNamedRange>>,
    trimesters: Trimesters,
): GoogleAppsScript.Sheets.Schema.Request[] {
    const requests: GoogleAppsScript.Sheets.Schema.Request[] = [];

    const { frozenRows, frozenCols } = getFrozenRowCols(namedRanges);
    const { trim1Range, trim2Range, trim3Range } = trimesters;

    const trim1FormatSource = namedRanges[ReportSheetSchema.sheets.attendanceTemplate.ranges.formatPeriod1]?.range;
    const trim2FormatSource = namedRanges[ReportSheetSchema.sheets.attendanceTemplate.ranges.formatPeriod2]?.range;
    const trim3FormatSource = namedRanges[ReportSheetSchema.sheets.attendanceTemplate.ranges.formatPeriod3]?.range;

    // Clean up format
    const noFormatSource = namedRanges[ReportSheetSchema.sheets.attendanceTemplate.ranges.formatAttendanceCell]?.range;
    const noFormatDest = createRange(sheetId, frozenRows, frozenCols);
    const noFormatCopyRequest = buildCopyPasteRequest(noFormatSource, noFormatDest, PasteType.PASTE_NORMAL);
    if (noFormatCopyRequest) requests.push(noFormatCopyRequest);

    // Add format to the editable ranges
    for (const fullRange of editableRanges) {
        const trim1FormatDest = offsetGridRange({ origin: fullRange, colOffset: trim1Range.start - frozenCols, width: trim1Range.end - trim1Range.start + 1 });
        const trim2FormatDest = offsetGridRange({ origin: fullRange, colOffset: trim2Range.start - frozenCols, width: trim2Range.end - trim2Range.start + 1 });
        const trim3FormatDest = offsetGridRange({ origin: fullRange, colOffset: trim3Range.start - frozenCols, width: trim3Range.end - trim3Range.start + 1 });

        const trim1CopyRequest = buildCopyPasteRequest(trim1FormatSource, trim1FormatDest, PasteType.PASTE_FORMAT);
        if (trim1Range.start !== -1 && trim1CopyRequest) requests.push(trim1CopyRequest);

        const trim2CopyRequest = buildCopyPasteRequest(trim2FormatSource, trim2FormatDest, PasteType.PASTE_FORMAT);
        if (trim2Range.start !== -1 && trim2CopyRequest) requests.push(trim2CopyRequest);

        const trim3CopyRequest = buildCopyPasteRequest(trim3FormatSource, trim3FormatDest, PasteType.PASTE_FORMAT);
        if (trim3Range.start !== -1 && trim3CopyRequest) requests.push(trim3CopyRequest);
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
