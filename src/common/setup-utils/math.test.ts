import type { AcademicField, WeightedSubject } from "../report-utils";
import { calculateAttendanceGridSize, calculatePerClassLayout, getRandomId, normalizeSubjectWeights, normalizeTrimesterWeights } from "./math";

describe("Setup Utils. Math", () => {
    describe("getRandomId", () => {
        let randomSpy: jest.SpyInstance;

        beforeEach(() => {
            // Spy on Math.random to control its output for deterministic tests
            randomSpy = jest.spyOn(Math, "random");
        });

        afterEach(() => {
            randomSpy.mockRestore();
        });

        it("returns the minimum possible value when Math.random is 0", () => {
            randomSpy.mockReturnValue(0);
            expect(getRandomId()).toBe(0);
        });

        it("returns a predictable ID based on Math.random output", () => {
            randomSpy.mockReturnValue(0.5);
            // 0.5 * (2^31 - 1) = 0.5 * 2147483647 = 1073741823.5
            // Math.floor(1073741823.5) = 1073741823
            expect(getRandomId()).toBe(1073741823);
        });

        it("returns the maximum possible value when Math.random is close to 1", () => {
            randomSpy.mockReturnValue(0.9999999999);
            // Should be exactly 2^31 - 2 due to flooring
            expect(getRandomId()).toBe(2147483646);
        });
    });

    describe("calculateAttendanceGridSize", () => {
        it("calculates the correct grid size when attendancePerClass is false", () => {
            const frozenRows = 2;
            const frozenCols = 3;
            const calendarLength = 30;
            const studentsCount = 25;
            const subjectsCount = 5; // Should be ignored in this branch

            const result = calculateAttendanceGridSize(frozenRows, frozenCols, calendarLength, studentsCount, subjectsCount, false);

            expect(result).toEqual({
                finalRowCount: 2 + 1 + 25, // 28
                finalColumnCount: 3 + 30, // 33
            });
        });

        it("calculates the correct grid size when attendancePerClass is true", () => {
            const frozenRows = 2;
            const frozenCols = 3;
            const calendarLength = 30;
            const studentsCount = 20;
            const subjectsCount = 4;

            const result = calculateAttendanceGridSize(frozenRows, frozenCols, calendarLength, studentsCount, subjectsCount, true);

            expect(result).toEqual({
                finalRowCount: 2 + (2 + 20) * 4, // 2 + (22 * 4) = 90
                finalColumnCount: 3 + 30, // 33
            });
        });
    });

    describe("calculatePerClassLayout", () => {
        it("returns an empty array if there are 0 subjects", () => {
            const result = calculatePerClassLayout(0, 20, 2);
            expect(result).toEqual([]);
        });

        it("calculates correct layouts for a single subject", () => {
            const result = calculatePerClassLayout(1, 10, 2);

            // space = 10 + 2 = 12
            // index 0:
            // titleFormatStartRow = 2 + 0 = 2
            // studentStartRow = 2 + 2 + 0 = 4
            // bandingStartRow = 2 + 1 + 0 = 3
            // bandingNumRows = 10 + 1 = 11
            expect(result).toEqual([
                {
                    subjectIndex: 0,
                    titleFormatStartRow: 2,
                    studentStartRow: 4,
                    bandingStartRow: 3,
                    bandingNumRows: 11,
                },
            ]);
        });

        it("calculates correct consecutive blocks for multiple subjects", () => {
            const result = calculatePerClassLayout(2, 5, 3);

            // space = 5 + 2 = 7
            // index 0 offsets: 0
            // index 1 offsets: 7
            expect(result).toEqual([
                {
                    subjectIndex: 0,
                    titleFormatStartRow: 3, // 3 + 0
                    studentStartRow: 5, // 3 + 2 + 0
                    bandingStartRow: 4, // 3 + 1 + 0
                    bandingNumRows: 6, // 5 + 1
                },
                {
                    subjectIndex: 1,
                    titleFormatStartRow: 10, // 3 + 7
                    studentStartRow: 12, // 3 + 2 + 7
                    bandingStartRow: 11, // 3 + 1 + 7
                    bandingNumRows: 6, // 5 + 1
                },
            ]);
        });
    });

    describe("normalizeTrimesterWeights", () => {
        it("normalizes positive weights correctly", () => {
            const [w0, w1, w2] = normalizeTrimesterWeights(1, 2, 1);
            expect(w0).toBe(0.25);
            expect(w1).toBe(0.5);
            expect(w2).toBe(0.25);
        });

        it("uses absolute values for negative inputs", () => {
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
            const [w0, w1, w2] = normalizeTrimesterWeights(1000000, 0.000001, 0);
            expect(w0).toBeCloseTo(1);
            expect(w1).toBeGreaterThan(0);
            expect(w2).toBe(0);
        });
    });

    describe("normalizeSubjectWeights", () => {
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
            expect(result).toEqual([
                { subject: "Math", weight: 0.125 },
                { subject: "Physics", weight: 0.375 },
                { subject: "History", weight: 0.25 },
                { subject: "Art", weight: 0.25 },
            ]);
        });

        it("normalizes per chunk when averagePerField is true", () => {
            const result = normalizeSubjectWeights(subjects, academicFields, true);
            expect(result).toEqual([
                { subject: "Math", weight: 0.25 },
                { subject: "Physics", weight: 0.75 },
                { subject: "History", weight: 0.5 },
                { subject: "Art", weight: 0.5 },
            ]);
        });

        it("leaves leftover subjects unmodified if there are more subjects than field slots", () => {
            const shortFields: AcademicField[] = [{ name: "Sciences", color: "blue", subjects: 1 }];

            const result = normalizeSubjectWeights(subjects, shortFields, true);
            expect(result).toEqual([
                { subject: "Math", weight: 1 },
                { subject: "Physics", weight: 3 },
                { subject: "History", weight: 2 },
                { subject: "Art", weight: 2 },
            ]);
        });

        it("handles chunks gracefully if there are fewer subjects than expected by the fields", () => {
            const largeFields: AcademicField[] = [
                { name: "Sciences", color: "blue", subjects: 2 },
                { name: "Humanities", color: "red", subjects: 4 },
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
                { subject: "History", weight: 1 },
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
