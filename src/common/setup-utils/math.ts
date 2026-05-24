import type { AcademicField, WeightedSubject } from "../report-utils";
import type { SubjectBlockLayout } from ".";

/**
 * Generates a random 30 bit integer ID
 */
export function getRandomId(): number {
    return Math.floor(Math.random() * (2 ** 31 - 1));
}

/**
 * Calculate the size of the Attendance sheet.
 */
export function calculateAttendanceGridSize(
    frozenRows: number,
    frozenCols: number,
    calendarLength: number,
    studentsCount: number,
    subjectsCount: number,
    attendancePerClass: boolean,
): { finalRowCount: number; finalColumnCount: number } {
    const finalColumnCount = frozenCols + calendarLength;
    let finalRowCount = frozenRows;

    if (attendancePerClass) {
        finalRowCount += (2 + studentsCount) * subjectsCount;
    } else {
        finalRowCount += 1 + studentsCount;
    }

    return { finalRowCount, finalColumnCount };
}

/**
 * Calculates the sizes for each class block for attendance.
 */
export function calculatePerClassLayout(subjectCount: number, studentCount: number, frozenRows: number): SubjectBlockLayout[] {
    const space = studentCount + 2;
    const layouts: SubjectBlockLayout[] = [];

    for (let index = 0; index < subjectCount; index++) {
        layouts.push({
            subjectIndex: index,
            titleFormatStartRow: frozenRows + index * space,
            studentStartRow: frozenRows + 2 + index * space,
            bandingStartRow: frozenRows + 1 + index * space,
            bandingNumRows: studentCount + 1,
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
