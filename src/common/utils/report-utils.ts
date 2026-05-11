import { getA1Notation, getColumnLetter } from "./gas-utils";
import { formatDateRange } from "./text";

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
    dateStart: number;
    dateTrim1: number;
    dateTrim2: number;
    dateEnd: number;
}

interface ProtectedSections {
    data: boolean;
    habilities: boolean;
    comments: boolean;
    trim1: boolean;
    trim2: boolean;
    trim3: boolean;
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
 * Generate a formula for the general absences.
 */
export function createStudentAsistanceFormula(
    period: 0 | 1 | 2,
    firstNameRange: GoogleAppsScript.Sheets.Schema.GridRange,
    lastNameRange: GoogleAppsScript.Sheets.Schema.GridRange,
    assistanceSheetName: string,
    absences = false,
): string {
    const firstNameA1 = getA1Notation(firstNameRange, true, true, true);
    const lastNameA1 = getA1Notation(lastNameRange, true, true, true);
    const returnColumn = getColumnLetter((absences ? 4 : 3) + period * 2);
    return `=LET(
    first_names, ${assistanceSheetName}!B:B,
    last_names, ${assistanceSheetName}!C:C,
    return_data, ${assistanceSheetName}!${returnColumn}:${returnColumn},
    XLOOKUP(${firstNameA1}&${lastNameA1}, ARRAYFORMULA(first_names&last_names), return_data)
)`;
}

/**
 * Generate the student forumala to grab the absences per subject.
 */
export function createStudentAsistancePerSubjectFormula(
    period: 0 | 1 | 2,
    subjectRange: GoogleAppsScript.Sheets.Schema.GridRange,
    firstNameRange: GoogleAppsScript.Sheets.Schema.GridRange,
    lastNameRange: GoogleAppsScript.Sheets.Schema.GridRange,
    assistanceSheetName: string,
): string {
    const subjectA1 = getA1Notation(subjectRange, true, false, true);
    const firstNameA1 = getA1Notation(firstNameRange, true, true, true);
    const lastNameA1 = getA1Notation(lastNameRange, true, true, true);
    const returnColum = getColumnLetter(3 + period * 2);
    return `=LET(
    start_row, MATCH(${subjectA1}, ${assistanceSheetName}!A:A, 0),
    next_subject_offset,
        IFERROR(
            MATCH(
                TRUE,
                ARRAYFORMULA(
                    ISTEXT(
                        INDEX(${assistanceSheetName}!A:A, start_row + 1):
                        INDEX(${assistanceSheetName}!A:A, ROWS(${assistanceSheetName}!A:A))
                    )
                ),
                0
            ),
            ROWS(${assistanceSheetName}!A:A) - start_row
        ),
    height, next_subject_offset - 1,
    first_names, OFFSET(${assistanceSheetName}!B$1, start_row, 0, height, 1),
    last_names, OFFSET(${assistanceSheetName}!C$1, start_row, 0, height, 1),
    return_data, OFFSET(${assistanceSheetName}!${returnColum}$1, start_row, 0, height, 2),
    XLOOKUP(${firstNameA1} & ${lastNameA1}, ARRAYFORMULA(first_names & last_names), return_data)
)`;
}

/**
 * Returns a nice string for the asked period.
 */
export function generatePeriodString(data: ReportPersistentData, period: 0 | 1 | 2): string {
    const { configData, calendar } = data;

    let startBoundary: number;
    let endBoundary: number;

    if (period === 0) {
        startBoundary = configData.dateStart - 1;
        endBoundary = configData.dateTrim1;
    } else if (period === 1) {
        startBoundary = configData.dateTrim1;
        endBoundary = configData.dateTrim2;
    } else {
        startBoundary = configData.dateTrim2;
        endBoundary = configData.dateEnd;
    }

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
