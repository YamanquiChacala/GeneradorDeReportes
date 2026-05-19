import type { MappedNamedRange } from "./gas-utils";
import { getA1Notation, getColumnLetter } from "./gas-utils";
import { formatDateRange } from "./utils";

export interface ReportPersistentData {
    configData: ConfigData;
    protectedSections: ProtectedSections;
    academicFields: AcademicField[];
    subjects: WeightedSubject[];
    students: StudentRow[];
    calendar: number[];
}

export interface ConfigData {
    attendancePerClass: boolean;
    averagePerField: boolean;
    dates: [number, number, number, number];
    subjectGradingWeights: [number, number, number];
}

interface ProtectedSections {
    habilities: boolean;
    comments: boolean;
    trimesters: [boolean, boolean, boolean];
}

export interface AcademicField {
    name: string;
    color: string;
    subjects: number;
}

export interface WeightedSubject {
    subject: string;
    weight: number;
}

export type StudentRow = Student | StudentSpace;

export interface Student {
    type: "student";
    id: number;
    firstName: string;
    lastName: string;
    sheetName: string;
    sex: string;
    level: string;
    grade: string;
}

interface StudentSpace {
    type: "separator";
}

/**
 * Helper function to build the formula for shorter comments (for the SEP).
 */
export function getShortCommentFormula(source: MappedNamedRange, rowOffset: number, colOffset: number): string {
    const sourceLetter = getColumnLetter((source.range.startColumnIndex ?? 0) + colOffset);
    const sourceNumber = (source.range.startRowIndex ?? 0) + rowOffset + 1;
    return `=LET(
  raw_text, ${sourceLetter}${sourceNumber},
  lower_text, LOWER(raw_text),
  accent_map, {"á","a"; "é","e"; "í","i"; "ó","o"; "ú","u"; "ñ","n"; "ü","u"},
  cleaned_accents, REDUCE(lower_text, SEQUENCE(ROWS(accent_map)), LAMBDA(acc, i, SUBSTITUTE(acc, INDEX(accent_map, i, 1), INDEX(accent_map, i, 2)))),
  truncated, LEFT(cleaned_accents, MIN(250, IFERROR(FIND(CHAR(10), cleaned_accents) - 1, 250))),
  upper_text, UPPER(truncated),
  final_ascii, REGEXREPLACE(upper_text, "[^ -~]", ""),
  final_ascii
)`;
}

/**
 * Generate a formula for the general absences.
 */
export function createStudentAsistanceFormula(
    period: 0 | 1 | 2,
    firstNameMappedRange: MappedNamedRange,
    lastNameMappedRange: MappedNamedRange,
    attendanceSheetName: string,
    absences = false,
): string {
    const firstNameA1 = getA1Notation({ mappedRange: firstNameMappedRange, width: 1, height: 1, lockRows: true, lockColumns: true });
    const lastNameA1 = getA1Notation({ mappedRange: lastNameMappedRange, width: 1, height: 1, lockRows: true, lockColumns: true });
    const returnColumn = getColumnLetter((absences ? 4 : 3) + period * 2);
    return `=LET(
    first_names, '${attendanceSheetName}'!B:B,
    last_names, '${attendanceSheetName}'!C:C,
    return_data, '${attendanceSheetName}'!${returnColumn}:${returnColumn},
    XLOOKUP(${firstNameA1}&${lastNameA1}, ARRAYFORMULA(first_names&last_names), return_data)
)`;
}

/**
 * Formula to claculate each subject's average.
 */
export function createSubjectAverageFormula(valuesRange: MappedNamedRange, rowOffset: number, weightsRange: MappedNamedRange): string {
    const valuesA1 = getA1Notation({ mappedRange: valuesRange, rowOffset, height: 1, lockColumns: true });
    const weightsA1 = getA1Notation({ mappedRange: weightsRange, includeSheetName: true, lockRows: true, lockColumns: true });

    return `=IFERROR(
    ROUND(
        AVERAGE.WEIGHTED(
            ${valuesA1},
            TRANSPOSE(${weightsA1})
        )
    ), ""
)`;
}

/**
 * Formula to calculate the final subject's average.
 */
export function createFinalSubjectAverageFormula(trim1: MappedNamedRange, trim2: MappedNamedRange, trim3: MappedNamedRange, rowOffset: number): string {
    const colOffset = (trim3.range.endColumnIndex ?? 0) - 2;

    const trim1A1 = getA1Notation({ mappedRange: trim1, rowOffset, colOffset, height: 1, width: 1, lockColumns: true });
    const trim2A1 = getA1Notation({ mappedRange: trim2, rowOffset, colOffset, height: 1, width: 1, lockColumns: true });
    const trim3A1 = getA1Notation({ mappedRange: trim3, rowOffset, colOffset, height: 1, width: 1, lockColumns: true });

    return `=IF(
    COUNT(${trim1A1}, ${trim2A1}, ${trim3A1})=3,
    ROUND(AVERAGE(${trim1A1}, ${trim2A1}, ${trim3A1}),1),
    ""
)`;
}

/**
 * Generate the student forumala to grab the absences per subject.
 */
export function createStudentAsistancePerSubjectFormula(
    period: 0 | 1 | 2,
    subjectRange: MappedNamedRange,
    rowOffset: number,
    firstNameRange: MappedNamedRange,
    lastNameRange: MappedNamedRange,
    attendanceSheetName: string,
): string {
    const subjectA1 = getA1Notation({ mappedRange: subjectRange, width: 1, height: 1, rowOffset, lockColumns: true });
    const firstNameA1 = getA1Notation({ mappedRange: firstNameRange, width: 1, height: 1, lockRows: true, lockColumns: true });
    const lastNameA1 = getA1Notation({ mappedRange: lastNameRange, width: 1, height: 1, lockRows: true, lockColumns: true });
    const returnColum = getColumnLetter(3 + period * 2);
    return `=LET(
    start_row, MATCH(${subjectA1}, ${attendanceSheetName}!A:A, 0),
    next_subject_offset,
        IFERROR(
            MATCH(
                TRUE,
                ARRAYFORMULA(
                    ISTEXT(
                        INDEX(${attendanceSheetName}!A:A, start_row + 1):
                        INDEX(${attendanceSheetName}!A:A, ROWS(${attendanceSheetName}!A:A))
                    )
                ),
                0
            ),
            ROWS(${attendanceSheetName}!A:A) - start_row
        ),
    height, next_subject_offset - 1,
    first_names, OFFSET(${attendanceSheetName}!B$1, start_row, 0, height, 1),
    last_names, OFFSET(${attendanceSheetName}!C$1, start_row, 0, height, 1),
    return_data, OFFSET(${attendanceSheetName}!${returnColum}$1, start_row, 0, height, 2),
    XLOOKUP(${firstNameA1} & ${lastNameA1}, ARRAYFORMULA(first_names & last_names), return_data)
)`;
}

/**
 * Helper function to calculate the field value from the subjects.
 */
export function createFieldFunction(
    field: AcademicField,
    rowOffset: number,
    colOffset: number,
    subjectsOffset: number,
    subjectsRange: MappedNamedRange,
    weightsRange: MappedNamedRange,
): string {
    const valuesA1 = getA1Notation({
        mappedRange: subjectsRange,
        rowOffset: subjectsOffset + rowOffset,
        colOffset,
        width: 1,
        height: field.subjects,
        lockRows: true,
    });
    const weightsA1 = getA1Notation({
        includeSheetName: true,
        mappedRange: weightsRange,
        rowOffset: subjectsOffset,
        colOffset: 1,
        width: 1,
        height: field.subjects,
        lockRows: true,
        lockColumns: true,
    });

    return `=IFERROR(
    AVERAGE.WEIGHTED(${valuesA1}, ${weightsA1})
)`;
}

/**
 * Returns a nice string for the asked period.
 */
export function generatePeriodString(data: ReportPersistentData, period: 0 | 1 | 2): string {
    const { configData, calendar } = data;

    const nextPeriod = [1, 2, 3] as const;

    const startBoundary = configData.dates[period] - 1;
    const endBoundary = configData.dates[nextPeriod[period]];

    // Get the indices using our single binary search function
    const startIndex = getUpperBoundIndex(calendar, startBoundary);
    const endIndex = getUpperBoundIndex(calendar, endBoundary);

    // Resolve the actual days.
    const actualStart = calendar[startIndex];
    const actualEnd = calendar[endIndex - 1];

    // Exit if dates weren't found, or if the boundaries resulted in crossed dates
    if (!actualStart || !actualEnd || actualStart > actualEnd) return "No hay fechas";

    return formatDateRange(actualStart, actualEnd);
}

/**
 * Returns the index of the first element in the calendar strictly greater than the threshold.
 */
function getUpperBoundIndex(calendar: readonly number[], threshold: number): number {
    let left = 0;
    let right = calendar.length - 1;
    let bestIndex = calendar.length;

    while (left <= right) {
        const mid = Math.floor((left + right) / 2);
        const midVal = calendar[mid];

        if (midVal === undefined) {
            break;
        }

        if (midVal > threshold) {
            bestIndex = mid;
            right = mid - 1; // It's strictly greater, but look left to find an earlier one
        } else {
            left = mid + 1; // It's <= threshold, look right
        }
    }

    return bestIndex;
}
