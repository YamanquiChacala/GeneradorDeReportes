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
    if (!actualStart || !actualEnd || actualStart > actualEnd) throw new Error("Fechas de periodos erroneas.");

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
