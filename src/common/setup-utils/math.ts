import type { AcademicField, WeightedSubject } from "../report-utils";

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

    // Safely fallback to 0 to satisfy noUncheckedIndexedAccess
    return [result[0] ?? 0, result[1] ?? 0, result[2] ?? 0];
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
