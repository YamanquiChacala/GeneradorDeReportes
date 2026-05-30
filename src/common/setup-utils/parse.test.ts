import { MS_PER_DAY } from "../constants";
import { parseAcademicFieldsAndSubjects, parseCalendarDays, parseStudentList } from ".";

// Helper functions to quickly mock Google Apps Script CellData
type CellData = GoogleAppsScript.Sheets.Schema.CellData;
const str = (v: string, bgCol?: GoogleAppsScript.Sheets.Schema.Color): CellData => ({
    effectiveValue: { stringValue: v },
    effectiveFormat: bgCol ? { backgroundColor: bgCol } : undefined,
});
const num = (v: number): CellData => ({ effectiveValue: { numberValue: v } });
const bool = (v: boolean): CellData => ({ effectiveValue: { boolValue: v } });
const empty = (): CellData => ({});

// Mock Colors
const RED_BG = { red: 1, green: 0, blue: 0 };
const BLUE_BG = { red: 0, green: 0, blue: 1 };

describe("Setup Parse Utils", () => {
    describe("parseAcademicFieldsAndSubjects", () => {
        it("handles simple input correctly", () => {
            const simpleSubjectData: readonly CellData[][] = [
                [str("field1", RED_BG), num(2), str("subject1")],
                [empty(), empty(), str("subject2")],
                [str("field2", BLUE_BG), empty(), str("subject3")],
            ];

            const { academicFields, rawSubjects } = parseAcademicFieldsAndSubjects(simpleSubjectData);

            expect(academicFields).toEqual([
                expect.objectContaining({
                    name: "field1",
                    subjects: 2,
                }),
                expect.objectContaining({
                    name: "field2",
                    subjects: 1,
                }),
            ]);

            expect(academicFields).toHaveLength(2);

            expect(rawSubjects).toEqual([
                { subject: "subject1", weight: 2 },
                { subject: "subject2", weight: 1 }, // Default weight fallback
                { subject: "subject3", weight: 1 },
            ]);

            expect(rawSubjects).toHaveLength(3);
        });

        it("handles messy user input correctly: creating fields, subjects, and skipping empty/invalid data", () => {
            const setupSubjectData: readonly CellData[][] = [
                /* 0 */ [empty(), empty(), empty(), empty(), empty(), empty()],
                /* 1 */ [],
                /* 2 */ [empty(), empty(), num(4), empty(), empty(), empty()], // Ignored: No field yet
                /* 3 */ [empty(), empty(), num(2), empty(), str("ignoredSubject")], // Ignored
                /* 4 */ [str("field1", RED_BG), empty(), str("subject1"), num(5), empty(), empty()], // Start Field 1
                /* 5 */ [empty(), empty(), str("subject2")], // Belongs to Field 1, default weight 1
                /* 6 */ [empty(), empty(), num(-4), str("subject3"), empty()], // Belongs to Field 1, weight -4
                /* 7 */ [str("field2", BLUE_BG)], // Finishes Field 1, Starts Field 2
                /* 8 */ [empty(), empty(), empty(), empty(), num(6), str("subject4")], // Belongs to Field 2, weight 6
                /* 9 */ [], // Skipped
                /* 10*/ [empty(), empty(), str("subject5"), num(3)], // Belongs to Field 2, weight 3
                /* 11*/ [str("ignoredField")], // Finishes Field 2, Starts Field 3 (which will be dropped)
                /* 12*/ [],
                /* 13*/ [empty(), str(""), num(7), empty()], // Empty subject strings don't count
            ];

            const { academicFields, rawSubjects } = parseAcademicFieldsAndSubjects(setupSubjectData);

            // Because we don't have the exact colorToHex implementation in the test scope,
            // we will just assert the structure and that strings were extracted.
            expect(academicFields).toEqual([
                expect.objectContaining({
                    name: "field1",
                    subjects: 3,
                }),
                expect.objectContaining({
                    name: "field2",
                    subjects: 2,
                }),
            ]);

            // Field 3 shouldn't be included because it had 0 subjects.
            expect(academicFields.length).toBe(2);

            expect(rawSubjects).toEqual([
                { subject: "subject1", weight: 5 },
                { subject: "subject2", weight: 1 }, // Default weight fallback
                { subject: "subject3", weight: -4 },
                { subject: "subject4", weight: 6 },
                { subject: "subject5", weight: 3 },
            ]);
        });
    });

    describe("parseStudentList", () => {
        it("groups students, restarts IDs per group, handles gaps, and assigns optional fields", () => {
            const studentSetupData: readonly CellData[][] = [
                /* 0 */ [empty(), empty(), empty(), empty(), empty()],
                /* 1 */ [],
                /* 2 */ [str("first1"), str("last1"), str("male"), str("elementary"), str("1"), str("abc")], // id 1
                /* 3 */ [str("first2"), str("last2"), num(50), num(35), num(2), num(123)], // id 2
                /* 4 */ [str("first3"), str("last3"), str("male"), str("elementary"), str("3°"), str("aqui")], // id 3
                /* 5 */ [empty(), str("ignore"), str("female"), str("high school"), num(2), str("my curp")], // Gap triggers here
                /* 6 */ [str("ignore"), empty(), str("male"), str("high school"), num(2), str("nada")], // Continued gap
                /* 7 */ [num(56), str("fake"), str("male"), str("high school"), num(2), str("nada")], // Continued gap
                /* 8 */ [str("ignore"), num(35), str("male"), str("high school"), num(2), str("nada")], // Continued gap
                /* 9 */ [str("first4"), str("last4")], // Starts new group, id 1
                /* 10 */ [],
                /* 11 */ [],
            ];

            const { students, studentReportData } = parseStudentList(studentSetupData);

            expect(students).toEqual([
                {
                    type: "student",
                    id: 1,
                    firstName: "first1",
                    lastName: "last1",
                    sheetName: expect.any(String), // Relies on your sanitizeSheetName helper
                    sex: "male",
                    level: "elementary",
                    grade: "1",
                    curp: "abc",
                },
                {
                    type: "student",
                    id: 2,
                    firstName: "first2",
                    lastName: "last2",
                    sheetName: expect.any(String),
                    sex: "50",
                    level: "35",
                    grade: "2º", // Number value fallback logic kicks in
                    curp: "123",
                },
                {
                    type: "student",
                    id: 3,
                    firstName: "first3",
                    lastName: "last3",
                    sheetName: expect.any(String),
                    sex: "male",
                    level: "elementary",
                    grade: "3°",
                    curp: "aqui",
                },
                { type: "separator" }, // Separator correctly inserted between groups
                {
                    type: "student",
                    id: 1, // ID resets correctly
                    firstName: "first4",
                    lastName: "last4",
                    sheetName: expect.any(String),
                    sex: "", // Defaults for missing optionals
                    level: "",
                    grade: "",
                    curp: "",
                },
            ]);

            // Ensure no trailing separators were generated at the end due to the final empty rows
            expect(students.length).toBe(5);

            // Check studentReportData matches length and basic structure (omitting separator)
            expect(studentReportData.length).toBe(5);
        });
    });

    describe("parseCalendarDays", () => {
        const SUN_JAN_4 = Date.UTC(2026, 0, 4);
        const MON_JAN_5 = Date.UTC(2026, 0, 5);

        it("throws an error if the initial day is not a Sunday", () => {
            const emptyGrid: CellData[][] = [[], []];
            expect(() => parseCalendarDays(MON_JAN_5, emptyGrid)).toThrow("Calendario no inicia en Domingo.");
        });

        it("only records days that are checked across rows", () => {
            // Mock a small grid: 2 weeks (2 rows).
            // 7 days per week. Indexes for checks are 1, 3, 5, 7, 9, 11, 13
            const row1: CellData[] = [];
            const row2: CellData[] = [];

            for (let i = 0; i < 14; i++) {
                row1.push(empty());
                row2.push(empty());
            }

            // Check Sunday Week 1 (Index 1)
            row1[1] = bool(true);
            // Check Tuesday Week 1 (Index 5)
            row1[5] = bool(true);
            // Check Monday Week 2 (Index 3)
            row2[3] = bool(true);

            const grid = [row1, row2];

            const { days, sheetDays } = parseCalendarDays(SUN_JAN_4, grid);

            // Expect exactly 3 days to have been extracted
            expect(days).toEqual([
                SUN_JAN_4, // Sunday Week 1
                SUN_JAN_4 + 2 * MS_PER_DAY, // Tuesday Week 1
                SUN_JAN_4 + 8 * MS_PER_DAY, // Monday Week 2 (7 + 1 days)
            ]);

            // Expect sheetDays array to be populated equivalently
            expect(sheetDays.length).toBe(3);

            // Validate structural presence of userEnteredValue (value relies on getSheetsDate math)
            expect(sheetDays).toEqual([
                [{ userEnteredValue: { numberValue: expect.any(Number) } }],
                [{ userEnteredValue: { numberValue: expect.any(Number) } }],
                [{ userEnteredValue: { numberValue: expect.any(Number) } }],
            ]);
        });
    });
});
