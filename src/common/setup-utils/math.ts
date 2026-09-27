import type { AcademicField, FrozenArea, Range, WeightedSubject } from "../report-utils";
import type { SubjectBlockLayout } from "./types";

/**
 * Calculate the size of the Attendance sheet.
 */
export function calculateAttendanceGridSize(
    frozenArea: FrozenArea,
    calendarLength: number,
    studentsCount: number,
    subjectsCount: number,
    attendancePerClass: boolean,
): { finalRowCount: number; finalColumnCount: number } {
    const finalColumnCount = frozenArea.cols + calendarLength;
    let finalRowCount = frozenArea.rows;

    if (attendancePerClass) {
        finalRowCount += (3 + studentsCount) * subjectsCount; // Space + subject title + class mood
    } else {
        finalRowCount += 2 + studentsCount; // Space + class mood
    }

    return { finalRowCount, finalColumnCount };
}

/**
 * Calculates the sizes for each class block for attendance.
 */
export function calculatePerClassLayout(subjectCount: number, studentCount: number, frozenRows: number): SubjectBlockLayout[] {
    const space = studentCount + 3;
    const layouts: SubjectBlockLayout[] = [];

    for (let index = 0; index < subjectCount; index++) {
        layouts.push({
            subjectIndex: index,
            titleFormatStartRow: frozenRows + index * space,
            studentStartRow: frozenRows + 3 + index * space,
            bandingStartRow: frozenRows + 1 + index * space,
            bandingNumRows: studentCount + 2,
        });
    }

    return layouts;
}

/**
 * Generically normalizes the weights of any array of items.
 * Proportionality sums the extracted weights to 1.
 */
function normalizeItems<T>(items: readonly T[], getWeight: (item: T) => number, setWeight: (item: T, newWeight: number) => T): T[] {
    const sum = items.reduce((acc, item) => acc + Math.abs(getWeight(item)), 0);
    if (sum === 0) {
        return items.map((item) => setWeight(item, 1 / items.length));
    }

    return items.map((item) => setWeight(item, Math.abs(getWeight(item)) / sum));
}

/**
 * Normalizes exactly three trimester weights.
 */
export function normalizeTrimesterWeights(w0: number, w1: number, w2: number): [number, number, number] {
    const result = normalizeItems(
        [w0, w1, w2],
        (w) => w,
        (_, newWeight) => newWeight,
    );

    // biome-ignore lint/style/noNonNullAssertion: normalizeItems returns the same number of inputs.
    return [result[0]!, result[1]!, result[2]!];
}

/**
 * Normalizes the weights of subjects.
 * If averagePerField is true, normalizes weights in chunks based on academicFields.
 * If false, normalizes globally across all subjects.
 */
export function normalizeSubjectWeights(subjects: readonly WeightedSubject[], academicFields: readonly AcademicField[], averagePerField: boolean): WeightedSubject[] {
    if (!averagePerField) {
        // Global normalization
        return normalizeItems(
            subjects,
            (subject) => subject.weight,
            (subject, newWeight) => ({ ...subject, weight: newWeight }),
        );
    }

    // Per-field (chunked) normalization
    const normalizedSubjects: WeightedSubject[] = [];
    let subjectIndex = 0;

    for (const field of academicFields) {
        const fieldSubjectCount = field.subjects;

        // Safely extract the chunk for this field
        const chunk: WeightedSubject[] = [];
        for (let i = 0; i < fieldSubjectCount; i++) {
            const subj = subjects[subjectIndex + i];
            if (subj) {
                chunk.push(subj);
            }
        }

        // Normalize the chunk
        const normalizedChunk = normalizeItems(
            chunk,
            (subject) => subject.weight,
            (subject, newWeight) => ({ ...subject, weight: newWeight }),
        );
        normalizedSubjects.push(...normalizedChunk);

        subjectIndex += fieldSubjectCount;
    }

    // Safety net: preserve any remaining subjects that fell outside the field counts
    while (subjectIndex < subjects.length) {
        // biome-ignore lint/style/noNonNullAssertion: The while above ensures we're in bounds.
        normalizedSubjects.push(subjects[subjectIndex]!);
        subjectIndex++;
    }

    return normalizedSubjects;
}

/**
 * Calculates the widths of the columns needed for the summary sheet.
 */
export function getSummaryColumnWidths(
    attendancePerClass: boolean,
    averagePerField: boolean,
    subjects: number,
    fields: number,
): { columWidths: number[]; mergeRanges: Range[] } {
    const columWidths: number[] = [];
    const mergeRanges: Range[] = [];
    let currentIndex = 0; // Tracks our position for the merge ranges

    // Subject Column Widths must add to 120
    let subjectWidths: number[];
    if (attendancePerClass && averagePerField) {
        // 2 columns (Fal, Cal)
        subjectWidths = [60, 60];
    } else if (attendancePerClass && !averagePerField) {
        // 3 columns (Fal, Cal, SEP)
        subjectWidths = [40, 40, 40];
    } else if (!attendancePerClass && averagePerField) {
        // 1 column (Cal)
        subjectWidths = [120];
    } else {
        // 2 columns (Cal, SEP)
        subjectWidths = [60, 60];
    }

    const fieldWidths = [60, 60]; // Fields always have "Cal" and "SEP"
    const emptyColumnWidth = 40;
    const generalAttendanceWidth = 60;
    const averageWidth = 80;

    // Helper to cleanly push general attendance when appropriate
    const pushGeneralAttendance = () => {
        if (!attendancePerClass) {
            columWidths.push(generalAttendanceWidth);
            currentIndex += 1;
        }
    };

    // Helper to push subjects and their merge ranges
    const pushSubjects = () => {
        for (let i = 0; i < subjects; i++) {
            columWidths.push(...subjectWidths);
            // We only need a merge range if a subject spans more than 1 column
            if (subjectWidths.length > 1) {
                mergeRanges.push({ start: currentIndex, end: currentIndex + subjectWidths.length });
            }
            currentIndex += subjectWidths.length;
        }
    };

    // Assemble the columns
    if (averagePerField) {
        // Subjects
        pushSubjects();

        // Empty Column
        columWidths.push(emptyColumnWidth);
        currentIndex += 1;

        // General Attendance
        pushGeneralAttendance();

        // Fields
        for (let i = 0; i < fields; i++) {
            columWidths.push(...fieldWidths);
            mergeRanges.push({ start: currentIndex, end: currentIndex + fieldWidths.length });
            currentIndex += fieldWidths.length;
        }

        // Final Average
        columWidths.push(averageWidth);
    } else {
        // General Attendance
        pushGeneralAttendance();

        // Subjects
        pushSubjects();

        // Final Average
        columWidths.push(averageWidth);
    }

    return { columWidths, mergeRanges };
}
