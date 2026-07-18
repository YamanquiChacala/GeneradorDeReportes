import { getA1Notation, type MappedNamedRange, offsetGridRange } from "../gas-utils";
import {
    createAttendaceFormulas,
    createSetupRowValidFormula,
    DEFAULT_COMMENT,
    DEFAULT_SEP_STRENGHT,
    DEFAULT_SEP_SUGGESTION,
    DEFAULT_SEP_WEAKNESS,
    getCharacterCountFormula,
    getSepCommentFormula,
    type StudentRow,
    StudentRowType,
    type TrimesterRanges,
} from "../report-utils";
import { TemplateSize } from "./types";

interface MonthGroupMeta {
    year: number;
    month: number;
    startCol: number;
    count: number;
    template: TemplateSize;
}

/**
 * Builds the data to form the Assistance Headers.
 */
export function calculateCalendarHeaders(
    days: number[],
    frozenCols: number,
    names1: string[],
    names2: string[],
    names5: string[],
    dayNames: string[],
): { monthGroups: MonthGroupMeta[]; row1Values: (string | null)[]; row2Values: string[]; row3Values: number[] } {
    const monthGroups: MonthGroupMeta[] = [];
    const row1Values: (string | null)[] = [];
    const row2Values: string[] = [];
    const row3Values: number[] = [];

    if (days.length === 0) {
        return { monthGroups, row1Values, row2Values, row3Values };
    }

    let currentGroup: MonthGroupMeta | null = null;

    const finalizeGroup = (group: MonthGroupMeta) => {
        const shortYear = String(group.year).slice(-2);
        const longYear = String(group.year);
        let text = "";

        if (group.count === 1) {
            group.template = TemplateSize.SMALL;
            const name = names1[group.month] ?? "";
            text = `${name}\n${shortYear}`;
        } else if (group.count >= 2 && group.count <= 4) {
            group.template = TemplateSize.MEDIUM;
            const name = names2[group.month] ?? "";
            text = `${name}\n${longYear}`;
        } else {
            group.template = TemplateSize.LARGE;
            const name = names5[group.month] ?? "";
            text = `${name}\n${longYear}`;
        }

        monthGroups.push(group);

        for (let i = 0; i < group.count; i++) {
            row1Values.push(i === 0 ? text : null);
        }
    };

    days.forEach((dayValue, i) => {
        const date = new Date(dayValue);
        const year = date.getUTCFullYear();
        const month = date.getUTCMonth();
        const dayOfWeek = date.getUTCDay();
        const dayOfMonth = date.getUTCDate();

        const targetCol = frozenCols + i;

        if (!currentGroup) {
            currentGroup = { year, month, startCol: targetCol, count: 1, template: TemplateSize.SMALL };
        } else if (currentGroup.year === year && currentGroup.month === month) {
            currentGroup.count++;
        } else {
            finalizeGroup(currentGroup);
            currentGroup = { year, month, startCol: targetCol, count: 1, template: TemplateSize.SMALL };
        }

        const dayName = dayNames[dayOfWeek] ?? "";
        row2Values.push(dayName);
        row3Values.push(dayOfMonth);
    });

    // biome-ignore lint/style/noNonNullAssertion: At lest 1 day exists, so the cycle ran at least 1.
    finalizeGroup(currentGroup!);

    return { monthGroups, row1Values, row2Values, row3Values };
}

/**
 * Builds the data for the student list in Attendance
 */
export function generateStudentGrid(students: StudentRow[], initialRow: number, trimesters: TrimesterRanges): GoogleAppsScript.Sheets.Schema.CellData[][] {
    const result: GoogleAppsScript.Sheets.Schema.CellData[][] = [];

    for (let i = 0; i < students.length; i++) {
        const studentRow = students[i];

        if (studentRow?.type === StudentRowType.STUDENT) {
            const trim1 = createAttendaceFormulas(initialRow + i, trimesters.trim1.start, trimesters.trim1.end);
            const trim2 = createAttendaceFormulas(initialRow + i, trimesters.trim2.start, trimesters.trim2.end);
            const trim3 = createAttendaceFormulas(initialRow + i, trimesters.trim3.start, trimesters.trim3.end);

            result.push([
                { userEnteredValue: { numberValue: studentRow.id } },
                { userEnteredValue: { stringValue: studentRow.firstName } },
                { userEnteredValue: { stringValue: studentRow.lastName } },
                { userEnteredValue: { formulaValue: trim1.percent } },
                { userEnteredValue: { formulaValue: trim1.count } },
                { userEnteredValue: { formulaValue: trim2.percent } },
                { userEnteredValue: { formulaValue: trim2.count } },
                { userEnteredValue: { formulaValue: trim3.percent } },
                { userEnteredValue: { formulaValue: trim3.count } },
            ]);
        } else {
            result.push([]); // Keep empty rows for separators
        }
    }

    return result;
}

interface BuildDefaultCommentForStudentTemplateDataParams {
    commentsRange: MappedNamedRange;
    fieldSubjectCounts: number[];
    subjectNames: string[];
    averagePerField: boolean;
}

/**
 * Build the default comment data for the student template
 */
export function buildDefaultCommentForStudentTemplateData({
    commentsRange,
    fieldSubjectCounts,
    subjectNames,
    averagePerField,
}: BuildDefaultCommentForStudentTemplateDataParams): { data: GoogleAppsScript.Sheets.Schema.CellData[][]; mergeRanges: GoogleAppsScript.Sheets.Schema.GridRange[] } {
    const data: GoogleAppsScript.Sheets.Schema.CellData[][] = [];
    const mergeRanges: GoogleAppsScript.Sheets.Schema.GridRange[] = [];

    let subjectIndex = 0;

    for (const subjectCount of fieldSubjectCounts) {
        const rowsToConsider = averagePerField ? subjectCount : 1;

        if (averagePerField) {
            mergeRanges.push(
                offsetGridRange({
                    origin: commentsRange.namedRange.range,
                    colOffset: 7,
                    width: 1,
                    rowOffset: subjectIndex,
                    height: rowsToConsider,
                }),
            );
            mergeRanges.push(
                offsetGridRange({
                    origin: commentsRange.namedRange.range,
                    colOffset: 8,
                    width: 2,
                    rowOffset: subjectIndex,
                    height: rowsToConsider,
                }),
            );
        }

        const commentLenght = averagePerField ? Math.min(2, rowsToConsider - 1) : 0;

        for (let i = 0; i < subjectCount; i++) {
            const subjectName = subjectNames[subjectIndex] ?? "";

            const sepCommentPartsA1 = getA1Notation({
                mappedRange: commentsRange,
                rowOffset: subjectIndex,
                colOffset: 4,
                height: rowsToConsider,
                width: 3,
                lockColumns: true,
            });

            const sepCommentA1 = getA1Notation({
                mappedRange: commentsRange,
                rowOffset: subjectIndex,
                colOffset: 8,
                height: 1,
                width: 1,
                lockColumns: true,
            });

            subjectIndex++;
            data.push([
                { userEnteredValue: { stringValue: subjectName } },
                { userEnteredValue: { stringValue: DEFAULT_COMMENT } },
                {},
                {},
                { userEnteredValue: { stringValue: DEFAULT_SEP_STRENGHT[commentLenght] } },
                { userEnteredValue: { stringValue: DEFAULT_SEP_WEAKNESS[commentLenght] } },
                { userEnteredValue: { stringValue: DEFAULT_SEP_SUGGESTION[commentLenght] } },
                { userEnteredValue: { formulaValue: getCharacterCountFormula(sepCommentA1) } },
                { userEnteredValue: { formulaValue: getSepCommentFormula(sepCommentPartsA1) } },
            ]);
        }
    }

    return { data, mergeRanges };
}

export type CellFormulaClosure = (a1Cell: string) => string;

interface StatusSectionDataParams {
    readonly statusRange: MappedNamedRange;
    readonly studentRange: MappedNamedRange;
    readonly markColOffsets: number[];
    readonly markUsesFieldValue?: boolean[];
    readonly title: string;
    readonly headers: string[];
    readonly studentRows: StudentRow[];
    readonly averagePerField?: boolean;
    readonly fieldSubjects?: number[];
    readonly statusRowOffset: number;
    readonly formulaFunction: CellFormulaClosure | CellFormulaClosure[];
}

/**
 * Builds the data for the sections of the Status sheet
 */
export function buildStatusSectionData({
    statusRange,
    studentRange,
    markColOffsets,
    markUsesFieldValue = [],
    title,
    headers,
    studentRows,
    averagePerField = false,
    fieldSubjects,
    statusRowOffset,
    formulaFunction,
}: StatusSectionDataParams): {
    data: GoogleAppsScript.Sheets.Schema.CellData[][];
    mergeRanges: GoogleAppsScript.Sheets.Schema.GridRange[];
    borderRanges: GoogleAppsScript.Sheets.Schema.GridRange[];
} {
    const data: GoogleAppsScript.Sheets.Schema.CellData[][] = [];
    const mergeRanges: GoogleAppsScript.Sheets.Schema.GridRange[] = [];
    const borderRanges: GoogleAppsScript.Sheets.Schema.GridRange[] = [];

    const marksPerItem = markColOffsets.length;
    const colsPerMark = Math.ceil(3 / marksPerItem);
    const colsPerItem = colsPerMark * marksPerItem;

    if (!Array.isArray(formulaFunction)) {
        formulaFunction = Array(marksPerItem).fill(formulaFunction);
    } else if (formulaFunction.length < marksPerItem) {
        throw new Error("Not enough formulas for Status row.");
    }

    const fieldIndices = averagePerField && fieldSubjects != null ? fieldSubjects : Array.from({ length: studentRows.length }, (_, index) => index);

    // Headers
    const header: GoogleAppsScript.Sheets.Schema.CellData[] = [{ userEnteredValue: { stringValue: title } }, {}, {}, {}];
    for (const label of headers) {
        header.push({ userEnteredValue: { stringValue: label } });
        for (let i = 1; i < colsPerItem; i++) {
            header.push({});
        }
    }
    data.push(header);

    // Student rows
    for (const [studentIndex, studentRow] of studentRows.entries()) {
        if (studentRow.type === StudentRowType.SEPARATOR) {
            data.push([]);
            continue;
        }

        const studentDataRow: GoogleAppsScript.Sheets.Schema.CellData[] = [
            { userEnteredValue: { formulaValue: createSetupRowValidFormula(statusRange, 1 + statusRowOffset + studentIndex, 4) } },
            { userEnteredValue: { numberValue: studentRow.id } },
            { userEnteredValue: { stringValue: studentRow.firstName } },
            { userEnteredValue: { stringValue: studentRow.lastName } },
        ];

        headers.forEach((_, subjectIndex) => {
            for (const [markIndex, colOffset] of markColOffsets.entries()) {
                let rowOffset = subjectIndex;
                if (averagePerField && markUsesFieldValue[markIndex]) rowOffset = fieldIndices[subjectIndex] ?? subjectIndex;

                const a1Cell = getA1Notation({
                    mappedRange: studentRange,
                    includeSheetName: true,
                    customSheetName: studentRow.sheetName,
                    rowOffset,
                    colOffset,
                    height: 1,
                    width: 1,
                    lockColumns: true,
                    lockRows: true,
                });

                // biome-ignore lint/style/noNonNullAssertion: At the start we checked it will the right lenght.
                const formula = formulaFunction[markIndex]!;
                studentDataRow.push({ userEnteredValue: { formulaValue: formula(a1Cell) } });
                for (let i = 1; i < colsPerMark; i++) {
                    studentDataRow.push({});
                }
            }
        });

        data.push(studentDataRow);
    }

    // Merges and ranges
    headers.forEach((_, subjectIndex) => {
        const subjectRange = offsetGridRange({
            origin: statusRange.namedRange.range,
            colOffset: 4 + colsPerItem * subjectIndex,
            width: colsPerItem,
            rowOffset: statusRowOffset,
            height: studentRows.length + 1,
        });
        borderRanges.push(subjectRange);

        if (colsPerMark === colsPerItem) {
            mergeRanges.push(subjectRange);
        } else {
            mergeRanges.push(
                offsetGridRange({
                    origin: statusRange.namedRange.range,
                    colOffset: 4 + colsPerItem * subjectIndex,
                    width: colsPerItem,
                    rowOffset: statusRowOffset,
                    height: 1,
                }),
            );
            if (colsPerMark > 1) {
                for (let i = 0; i < marksPerItem; i++) {
                    mergeRanges.push(
                        offsetGridRange({
                            origin: statusRange.namedRange.range,
                            colOffset: 4 + colsPerItem * subjectIndex + i * colsPerMark,
                            rowOffset: statusRowOffset + 1,
                            width: colsPerMark,
                            height: studentRows.length,
                        }),
                    );
                }
            }
        }
    });

    return { data, mergeRanges, borderRanges };
}
