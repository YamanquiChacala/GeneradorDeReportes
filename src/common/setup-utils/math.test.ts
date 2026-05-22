import type { AcademicField, WeightedSubject } from "../report-utils";
import { normalizeSubjectWeights, normalizeTrimesterWeights } from "./math";

describe("Setup Math Utils", () => {
    describe("normalizeTrimesterWeights", () => {
        it("normalizes positive weights correctly", () => {
            // Sum is 4
            const [w0, w1, w2] = normalizeTrimesterWeights(1, 2, 1);
            expect(w0).toBe(0.25);
            expect(w1).toBe(0.5);
            expect(w2).toBe(0.25);
        });

        it("uses absolute values for negative inputs", () => {
            // Math.abs turns this into [1, 2, 1], Sum is 4
            const [w0, w1, w2] = normalizeTrimesterWeights(-1, 2, -1);
            expect(w0).toBe(0.25);
            expect(w1).toBe(0.5);
            expect(w2).toBe(0.25);
        });

        it("defaults to equal distribution when all weights are 0", () => {
            const [w0, w1, w2] = normalizeTrimesterWeights(0, 0, 0);
            expect(w0).toBeCloseTo(1 / 3);
            expect(w1).toBeCloseTo(1 / 3);
            expect(w2).toBeCloseTo(1 / 3);
        });

        it("handles extremes with big and small numbers", () => {
            // Sum is essentially 1,000,000.000001
            const [w0, w1, w2] = normalizeTrimesterWeights(1000000, 0.000001, 0);
            expect(w0).toBeCloseTo(1); // Basically 100%
            expect(w1).toBeGreaterThan(0); // Extremely small but > 0
            expect(w2).toBe(0);
        });
    });

    describe("normalizeSubjectWeights", () => {
        // Base mock data
        const subjects: WeightedSubject[] = [
            { subject: "Math", weight: 1 },
            { subject: "Physics", weight: 3 },
            { subject: "History", weight: 2 },
            { subject: "Art", weight: 2 },
        ];

        const academicFields: AcademicField[] = [
            { name: "Sciences", color: "blue", subjects: 2 },
            { name: "Humanities", color: "red", subjects: 2 },
        ];

        it("normalizes globally when averagePerField is false", () => {
            const result = normalizeSubjectWeights(subjects, academicFields, false);
            // Global sum is 8
            expect(result).toEqual([
                { subject: "Math", weight: 0.125 }, // 1/8
                { subject: "Physics", weight: 0.375 }, // 3/8
                { subject: "History", weight: 0.25 }, // 2/8
                { subject: "Art", weight: 0.25 }, // 2/8
            ]);
        });

        it("normalizes per chunk when averagePerField is true", () => {
            const result = normalizeSubjectWeights(subjects, academicFields, true);
            // Chunk 1 (Sciences): Sum is 4 -> 1/4, 3/4
            // Chunk 2 (Humanities): Sum is 4 -> 2/4, 2/4
            expect(result).toEqual([
                { subject: "Math", weight: 0.25 },
                { subject: "Physics", weight: 0.75 },
                { subject: "History", weight: 0.5 },
                { subject: "Art", weight: 0.5 },
            ]);
        });

        it("leaves leftover subjects unmodified if there are more subjects than field slots", () => {
            const shortFields: AcademicField[] = [
                { name: "Sciences", color: "blue", subjects: 1 }, // Only accounts for 1 subject
            ];

            const result = normalizeSubjectWeights(subjects, shortFields, true);
            // Chunk 1 gets "Math". Remaining subjects fall through untouched.
            expect(result).toEqual([
                { subject: "Math", weight: 1 }, // Normalized alone
                { subject: "Physics", weight: 3 }, // Leftover
                { subject: "History", weight: 2 }, // Leftover
                { subject: "Art", weight: 2 }, // Leftover
            ]);
        });

        it("handles chunks gracefully if there are fewer subjects than expected by the fields", () => {
            const largeFields: AcademicField[] = [
                { name: "Sciences", color: "blue", subjects: 2 },
                { name: "Humanities", color: "red", subjects: 4 }, // Expects 4, but only 1 subject left
            ];
            const shortSubjects: WeightedSubject[] = [
                { subject: "Math", weight: 1 },
                { subject: "Physics", weight: 1 },
                { subject: "History", weight: 5 },
            ];

            const result = normalizeSubjectWeights(shortSubjects, largeFields, true);
            expect(result).toEqual([
                { subject: "Math", weight: 0.5 },
                { subject: "Physics", weight: 0.5 },
                { subject: "History", weight: 1 }, // Normalized alone because it was the only one left in the 4-subject chunk
            ]);
        });

        it("skips chunks correctly if an academic field has 0 subjects", () => {
            const weirdFields: AcademicField[] = [
                { name: "Empty", color: "gray", subjects: 0 },
                { name: "Sciences", color: "blue", subjects: 2 },
            ];
            const twoSubjects: WeightedSubject[] = [
                { subject: "Math", weight: 2 },
                { subject: "Physics", weight: 8 },
            ];

            const result = normalizeSubjectWeights(twoSubjects, weirdFields, true);
            expect(result).toEqual([
                { subject: "Math", weight: 0.2 },
                { subject: "Physics", weight: 0.8 },
            ]);
        });
    });
});
