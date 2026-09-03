import type { MappedNamedRange } from "../gas-utils";
import {
    DEFAULT_COMMENT,
    DEFAULT_SEP_STRENGHT,
    DEFAULT_SEP_SUGGESTION,
    DEFAULT_SEP_WEAKNESS,
    type StudentRow,
    StudentRowType,
    type TrimesterRanges,
} from "../report-utils";
import { buildDefaultCommentForStudentTemplateData, buildStatusSectionData, calculateCalendarHeaders, generateStudentGrid } from "./data";
import { TemplateSize } from "./types";

describe("Setup Utils. Data", () => {
    describe("calculateCalendarHeaders", () => {
        // Shared mock data for names
        const names1 = ["Jan1", "Feb1", "Mar1"];
        const names2 = ["Jan2", "Feb2", "Mar2"];
        const names5 = ["Jan5", "Feb5", "Mar5"];
        const dayNames = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];

        it("should return empty arrays when no days are provided", () => {
            const result = calculateCalendarHeaders([], 0, names1, names2, names5, dayNames);

            expect(result).toEqual({
                monthGroups: [],
                row1Values: [],
                row2Values: [],
                row3Values: [],
            });
        });

        it("should correctly format a single day (Template SMALL)", () => {
            // Date.UTC(Year, Month Index, Date) -> Jan 1, 2023 is a Sunday
            const days = [Date.UTC(2023, 0, 1)];
            const frozenCols = 2;

            const result = calculateCalendarHeaders(days, frozenCols, names1, names2, names5, dayNames);

            expect(result).toEqual({
                monthGroups: [{ year: 2023, month: 0, startCol: 2, count: 1, template: TemplateSize.SMALL }],
                row1Values: ["Jan1\n23"],
                row2Values: ["Sun"],
                row3Values: [1],
            });
        });

        it("should correctly format 2 to 4 days (Template MEDIUM)", () => {
            // Jan 1 (Sun), Jan 2 (Mon), Jan 3 (Tue)
            const days = [Date.UTC(2023, 0, 1), Date.UTC(2023, 0, 2), Date.UTC(2023, 0, 3)];

            const result = calculateCalendarHeaders(days, 0, names1, names2, names5, dayNames);

            expect(result).toEqual({
                monthGroups: [{ year: 2023, month: 0, startCol: 0, count: 3, template: TemplateSize.MEDIUM }],
                row1Values: ["Jan2\n2023", null, null],
                row2Values: ["Sun", "Mon", "Tue"],
                row3Values: [1, 2, 3],
            });
        });

        it("should correctly format 5 or more days (Template LARGE)", () => {
            // Feb 1 to Feb 5, 2023
            const days = [Date.UTC(2023, 1, 1), Date.UTC(2023, 1, 2), Date.UTC(2023, 1, 3), Date.UTC(2023, 1, 4), Date.UTC(2023, 1, 5)];

            const result = calculateCalendarHeaders(days, 0, names1, names2, names5, dayNames);

            expect(result).toEqual({
                monthGroups: [{ year: 2023, month: 1, startCol: 0, count: 5, template: TemplateSize.LARGE }],
                row1Values: ["Feb5\n2023", null, null, null, null],
                row2Values: ["Wed", "Thu", "Fri", "Sat", "Sun"],
                row3Values: [1, 2, 3, 4, 5],
            });
        });

        it("should handle month transitions seamlessly", () => {
            // Jan 31 (Tue) and Feb 1 (Wed)
            const days = [Date.UTC(2023, 0, 31), Date.UTC(2023, 1, 1)];

            const result = calculateCalendarHeaders(days, 0, names1, names2, names5, dayNames);

            expect(result).toEqual({
                monthGroups: [
                    { year: 2023, month: 0, startCol: 0, count: 1, template: TemplateSize.SMALL },
                    { year: 2023, month: 1, startCol: 1, count: 1, template: TemplateSize.SMALL },
                ],
                row1Values: ["Jan1\n23", "Feb1\n23"],
                row2Values: ["Tue", "Wed"],
                row3Values: [31, 1],
            });
        });

        it("should fallback to empty strings if names are missing", () => {
            const days1 = [Date.UTC(2023, 0, 1)];
            const days3 = [Date.UTC(2023, 0, 1), Date.UTC(2023, 0, 2), Date.UTC(2023, 0, 3)];
            const days5 = [Date.UTC(2023, 1, 1), Date.UTC(2023, 1, 2), Date.UTC(2023, 1, 3), Date.UTC(2023, 1, 4), Date.UTC(2023, 1, 5)];

            // Passing empty arrays to trigger the `?? ""` fallback
            const result1 = calculateCalendarHeaders(days1, 0, [], [], [], []);
            const result3 = calculateCalendarHeaders(days3, 0, [], [], [], []);
            const result5 = calculateCalendarHeaders(days5, 0, [], [], [], []);

            expect(result1).toEqual({
                monthGroups: [{ year: 2023, month: 0, startCol: 0, count: 1, template: TemplateSize.SMALL }],
                row1Values: ["\n23"],
                row2Values: [""],
                row3Values: [1],
            });

            expect(result3).toEqual({
                monthGroups: [{ year: 2023, month: 0, startCol: 0, count: 3, template: TemplateSize.MEDIUM }],
                row1Values: ["\n2023", null, null],
                row2Values: ["", "", ""],
                row3Values: [1, 2, 3],
            });

            expect(result5).toEqual({
                monthGroups: [{ year: 2023, month: 1, startCol: 0, count: 5, template: TemplateSize.LARGE }],
                row1Values: ["\n2023", null, null, null, null],
                row2Values: ["", "", "", "", ""],
                row3Values: [1, 2, 3, 4, 5],
            });
        });

        it("should handle cross-year transitions seamlessly", () => {
            // Dec 31, 2023 (Sun) and Jan 1, 2024 (Mon)
            const days = [Date.UTC(2023, 11, 31), Date.UTC(2024, 0, 1)];

            // Passing empty arrays to trigger the `?? ""` fallback for missing names
            const result = calculateCalendarHeaders(days, 0, [], [], [], ["Sun", "Mon"]);

            expect(result).toEqual({
                monthGroups: [
                    { year: 2023, month: 11, startCol: 0, count: 1, template: TemplateSize.SMALL },
                    { year: 2024, month: 0, startCol: 1, count: 1, template: TemplateSize.SMALL },
                ],
                row1Values: ["\n23", "\n24"],
                row2Values: ["Sun", "Mon"],
                row3Values: [31, 1],
            });
        });
    });

    describe("generateStudentGrid", () => {
        const mockStudent1: StudentRow = {
            type: StudentRowType.STUDENT,
            id: 101,
            firstName: "John",
            lastName: "Doe",
            sheetName: "JohnSheet",
            sex: "M",
            level: "1",
            grade: "A",
            curp: "123",
        };
        const mockStudent2: StudentRow = {
            type: StudentRowType.STUDENT,
            id: 102,
            firstName: "Jane",
            lastName: "Deer",
            sheetName: "JaneSheet",
            sex: "F",
            level: "2",
            grade: "B",
            curp: "456",
        };
        const mockSeparator: StudentRow = { type: StudentRowType.SEPARATOR };

        const mockTrimesters: TrimesterRanges = {
            trim1: { start: 1, end: 5 }, // B-F
            trim2: { start: 6, end: 10 }, // G-K
            trim3: { start: 11, end: 15 }, // L-P
        };

        it("should generate cell data correctly", () => {
            const resutl = generateStudentGrid([mockStudent1, mockSeparator, mockStudent2], 2 /* row 3 */, mockTrimesters);

            const expectedResult: GoogleAppsScript.Sheets.Schema.CellData[][] = [
                [
                    { userEnteredValue: { numberValue: 101 } },
                    { userEnteredValue: { stringValue: "John" } },
                    { userEnteredValue: { stringValue: "Doe" } },
                    { userEnteredValue: { formulaValue: expect.stringContaining("$B3:$F3") } },
                    { userEnteredValue: { formulaValue: expect.stringContaining("$B3:$F3") } },
                    { userEnteredValue: { formulaValue: expect.stringContaining("$G3:$K3") } },
                    { userEnteredValue: { formulaValue: expect.stringContaining("$G3:$K3") } },
                    { userEnteredValue: { formulaValue: expect.stringContaining("$L3:$P3") } },
                    { userEnteredValue: { formulaValue: expect.stringContaining("$L3:$P3") } },
                ],
                [],
                [
                    { userEnteredValue: { numberValue: 102 } },
                    { userEnteredValue: { stringValue: "Jane" } },
                    { userEnteredValue: { stringValue: "Deer" } },
                    { userEnteredValue: { formulaValue: expect.stringContaining("$B5:$F5") } },
                    { userEnteredValue: { formulaValue: expect.stringContaining("$B5:$F5") } },
                    { userEnteredValue: { formulaValue: expect.stringContaining("$G5:$K5") } },
                    { userEnteredValue: { formulaValue: expect.stringContaining("$G5:$K5") } },
                    { userEnteredValue: { formulaValue: expect.stringContaining("$L5:$P5") } },
                    { userEnteredValue: { formulaValue: expect.stringContaining("$L5:$P5") } },
                ],
            ];

            expect(resutl).toEqual(expectedResult);
        });

        it("should return an empty array when no students are provided", () => {
            const mockTrimesters: TrimesterRanges = {
                trim1: { start: 1, end: 5 },
                trim2: { start: 6, end: 10 },
                trim3: { start: 11, end: 15 },
            };

            const result = generateStudentGrid([], 2, mockTrimesters);

            expect(result).toEqual([]);
        });
    });

    describe("buildDefaultCommentForStudentTemplateData", () => {
        const mockCommentsMappedRange: MappedNamedRange = {
            namedRange: {
                namedRangeId: "comments01",
                name: "Mock Comments Range",
                range: { sheetId: 5, startRowIndex: 25, startColumnIndex: 0, endRowIndex: 6, endColumnIndex: 11 },
            },
            sheet: {},
        };

        it("should generate correct date with no fields", () => {
            const fieldSubjectCounts = [3, 1, 2];
            const subjectNames = ["sub1", "sub2", "sub3", "sub4", "sub5", "sub6", "sub7"];

            const result = buildDefaultCommentForStudentTemplateData({
                commentsRange: mockCommentsMappedRange,
                fieldSubjectCounts,
                subjectNames,
                averagePerField: false,
            });

            const expecteMergeRanges: GoogleAppsScript.Sheets.Schema.GridRange[] = [];

            const expectedData: GoogleAppsScript.Sheets.Schema.CellData[][] = [
                [
                    { userEnteredValue: { stringValue: "sub1" } },
                    { userEnteredValue: { stringValue: DEFAULT_COMMENT } },
                    {},
                    {},
                    { userEnteredValue: { stringValue: DEFAULT_SEP_STRENGHT[0] } },
                    { userEnteredValue: { stringValue: DEFAULT_SEP_WEAKNESS[0] } },
                    { userEnteredValue: { stringValue: DEFAULT_SEP_SUGGESTION[0] } },
                    { userEnteredValue: { formulaValue: expect.stringContaining("$I26") } },
                    { userEnteredValue: { formulaValue: expect.stringContaining("$E26:$G26") } },
                ],
                [
                    { userEnteredValue: { stringValue: "sub2" } },
                    { userEnteredValue: { stringValue: DEFAULT_COMMENT } },
                    {},
                    {},
                    { userEnteredValue: { stringValue: DEFAULT_SEP_STRENGHT[0] } },
                    { userEnteredValue: { stringValue: DEFAULT_SEP_WEAKNESS[0] } },
                    { userEnteredValue: { stringValue: DEFAULT_SEP_SUGGESTION[0] } },
                    { userEnteredValue: { formulaValue: expect.stringContaining("$I27") } },
                    { userEnteredValue: { formulaValue: expect.stringContaining("$E27:$G27") } },
                ],
                [
                    { userEnteredValue: { stringValue: "sub3" } },
                    { userEnteredValue: { stringValue: DEFAULT_COMMENT } },
                    {},
                    {},
                    { userEnteredValue: { stringValue: DEFAULT_SEP_STRENGHT[0] } },
                    { userEnteredValue: { stringValue: DEFAULT_SEP_WEAKNESS[0] } },
                    { userEnteredValue: { stringValue: DEFAULT_SEP_SUGGESTION[0] } },
                    { userEnteredValue: { formulaValue: expect.stringContaining("$I28") } },
                    { userEnteredValue: { formulaValue: expect.stringContaining("$E28:$G28") } },
                ],
                [
                    { userEnteredValue: { stringValue: "sub4" } },
                    { userEnteredValue: { stringValue: DEFAULT_COMMENT } },
                    {},
                    {},
                    { userEnteredValue: { stringValue: DEFAULT_SEP_STRENGHT[0] } },
                    { userEnteredValue: { stringValue: DEFAULT_SEP_WEAKNESS[0] } },
                    { userEnteredValue: { stringValue: DEFAULT_SEP_SUGGESTION[0] } },
                    { userEnteredValue: { formulaValue: expect.stringContaining("$I29") } },
                    { userEnteredValue: { formulaValue: expect.stringContaining("$E29:$G29") } },
                ],
                [
                    { userEnteredValue: { stringValue: "sub5" } },
                    { userEnteredValue: { stringValue: DEFAULT_COMMENT } },
                    {},
                    {},
                    { userEnteredValue: { stringValue: DEFAULT_SEP_STRENGHT[0] } },
                    { userEnteredValue: { stringValue: DEFAULT_SEP_WEAKNESS[0] } },
                    { userEnteredValue: { stringValue: DEFAULT_SEP_SUGGESTION[0] } },
                    { userEnteredValue: { formulaValue: expect.stringContaining("$I30") } },
                    { userEnteredValue: { formulaValue: expect.stringContaining("$E30:$G30") } },
                ],
                [
                    { userEnteredValue: { stringValue: "sub6" } },
                    { userEnteredValue: { stringValue: DEFAULT_COMMENT } },
                    {},
                    {},
                    { userEnteredValue: { stringValue: DEFAULT_SEP_STRENGHT[0] } },
                    { userEnteredValue: { stringValue: DEFAULT_SEP_WEAKNESS[0] } },
                    { userEnteredValue: { stringValue: DEFAULT_SEP_SUGGESTION[0] } },
                    { userEnteredValue: { formulaValue: expect.stringContaining("$I31") } },
                    { userEnteredValue: { formulaValue: expect.stringContaining("$E31:$G31") } },
                ],
            ];

            expect(result.mergeRanges).toEqual(expecteMergeRanges);
            expect(result.data).toEqual(expectedData);
        });

        it("should generate correct date with fields", () => {
            const fieldSubjectCounts = [3, 1, 2];
            const subjectNames = ["sub1", "sub2", "sub3", "sub4", "sub5"];

            const result = buildDefaultCommentForStudentTemplateData({
                commentsRange: mockCommentsMappedRange,
                fieldSubjectCounts,
                subjectNames,
                averagePerField: true,
            });

            const expecteMergeRanges: GoogleAppsScript.Sheets.Schema.GridRange[] = [
                { sheetId: 5, startRowIndex: 25, startColumnIndex: 7, endRowIndex: 28, endColumnIndex: 8 },
                { sheetId: 5, startRowIndex: 25, startColumnIndex: 8, endRowIndex: 28, endColumnIndex: 10 },
                { sheetId: 5, startRowIndex: 28, startColumnIndex: 7, endRowIndex: 29, endColumnIndex: 8 },
                { sheetId: 5, startRowIndex: 28, startColumnIndex: 8, endRowIndex: 29, endColumnIndex: 10 },
                { sheetId: 5, startRowIndex: 29, startColumnIndex: 7, endRowIndex: 31, endColumnIndex: 8 },
                { sheetId: 5, startRowIndex: 29, startColumnIndex: 8, endRowIndex: 31, endColumnIndex: 10 },
            ];

            const expectedData: GoogleAppsScript.Sheets.Schema.CellData[][] = [
                [
                    { userEnteredValue: { stringValue: "sub1" } },
                    { userEnteredValue: { stringValue: DEFAULT_COMMENT } },
                    {},
                    {},
                    { userEnteredValue: { stringValue: DEFAULT_SEP_STRENGHT[2] } },
                    { userEnteredValue: { stringValue: DEFAULT_SEP_WEAKNESS[2] } },
                    { userEnteredValue: { stringValue: DEFAULT_SEP_SUGGESTION[2] } },
                    { userEnteredValue: { formulaValue: expect.stringContaining("$I26") } },
                    { userEnteredValue: { formulaValue: expect.stringContaining("$E26:$G28") } },
                ],
                [
                    { userEnteredValue: { stringValue: "sub2" } },
                    { userEnteredValue: { stringValue: DEFAULT_COMMENT } },
                    {},
                    {},
                    { userEnteredValue: { stringValue: DEFAULT_SEP_STRENGHT[2] } },
                    { userEnteredValue: { stringValue: DEFAULT_SEP_WEAKNESS[2] } },
                    { userEnteredValue: { stringValue: DEFAULT_SEP_SUGGESTION[2] } },
                    { userEnteredValue: { formulaValue: expect.any(String) } },
                    { userEnteredValue: { formulaValue: expect.any(String) } },
                ],
                [
                    { userEnteredValue: { stringValue: "sub3" } },
                    { userEnteredValue: { stringValue: DEFAULT_COMMENT } },
                    {},
                    {},
                    { userEnteredValue: { stringValue: DEFAULT_SEP_STRENGHT[2] } },
                    { userEnteredValue: { stringValue: DEFAULT_SEP_WEAKNESS[2] } },
                    { userEnteredValue: { stringValue: DEFAULT_SEP_SUGGESTION[2] } },
                    { userEnteredValue: { formulaValue: expect.any(String) } },
                    { userEnteredValue: { formulaValue: expect.any(String) } },
                ],
                [
                    { userEnteredValue: { stringValue: "sub4" } },
                    { userEnteredValue: { stringValue: DEFAULT_COMMENT } },
                    {},
                    {},
                    { userEnteredValue: { stringValue: DEFAULT_SEP_STRENGHT[0] } },
                    { userEnteredValue: { stringValue: DEFAULT_SEP_WEAKNESS[0] } },
                    { userEnteredValue: { stringValue: DEFAULT_SEP_SUGGESTION[0] } },
                    { userEnteredValue: { formulaValue: expect.stringContaining("$I29") } },
                    { userEnteredValue: { formulaValue: expect.stringContaining("$E29:$G29") } },
                ],
                [
                    { userEnteredValue: { stringValue: "sub5" } },
                    { userEnteredValue: { stringValue: DEFAULT_COMMENT } },
                    {},
                    {},
                    { userEnteredValue: { stringValue: DEFAULT_SEP_STRENGHT[1] } },
                    { userEnteredValue: { stringValue: DEFAULT_SEP_WEAKNESS[1] } },
                    { userEnteredValue: { stringValue: DEFAULT_SEP_SUGGESTION[1] } },
                    { userEnteredValue: { formulaValue: expect.stringContaining("$I30") } },
                    { userEnteredValue: { formulaValue: expect.stringContaining("$E30:$G31") } },
                ],
                [
                    { userEnteredValue: { stringValue: "" } },
                    { userEnteredValue: { stringValue: DEFAULT_COMMENT } },
                    {},
                    {},
                    { userEnteredValue: { stringValue: DEFAULT_SEP_STRENGHT[1] } },
                    { userEnteredValue: { stringValue: DEFAULT_SEP_WEAKNESS[1] } },
                    { userEnteredValue: { stringValue: DEFAULT_SEP_SUGGESTION[1] } },
                    { userEnteredValue: { formulaValue: expect.any(String) } },
                    { userEnteredValue: { formulaValue: expect.any(String) } },
                ],
            ];

            expect(result.mergeRanges).toEqual(expecteMergeRanges);
            expect(result.data).toEqual(expectedData);
        });
    });

    describe("buildStatusSectionData", () => {
        // Setup dummy data
        const mockStatusRange: MappedNamedRange = {
            namedRange: {
                namedRangeId: "asdf",
                name: "Mock Status Range",
                range: { sheetId: 100, startRowIndex: 0, startColumnIndex: 0, endRowIndex: 3, endColumnIndex: 6 },
            },
            sheet: {},
        };
        const mockStudentRange: MappedNamedRange = {
            namedRange: {
                namedRangeId: "fdsa",
                name: "Mock Student Range",
                range: { sheetId: 1, startRowIndex: 25, startColumnIndex: 1, endRowIndex: 27, endColumnIndex: 4 },
            },
            sheet: {},
        };

        const mockStudent1: StudentRow = {
            type: StudentRowType.STUDENT,
            id: 101,
            firstName: "John",
            lastName: "Doe",
            sheetName: "JohnSheet",
            sex: "M",
            level: "1",
            grade: "A",
            curp: "123",
        };
        const mockStudent2: StudentRow = {
            type: StudentRowType.STUDENT,
            id: 102,
            firstName: "Jane",
            lastName: "Deer",
            sheetName: "JaneSheet",
            sex: "F",
            level: "2",
            grade: "B",
            curp: "456",
        };
        const mockSeparator: StudentRow = { type: StudentRowType.SEPARATOR };

        beforeEach(() => {
            jest.clearAllMocks();
        });

        it("shold correctly generate data with a single formula", () => {
            const title = "Status Title";

            const result = buildStatusSectionData({
                statusRange: mockStatusRange,
                studentRange: mockStudentRange,
                markColOffsets: [0],
                title,
                headers: ["Subject1", "Subject2"],
                studentRows: [mockStudent1, mockStudent2],
                statusRowOffset: 5,
                formulaFunction: (a1Cell: string) => `=CUSTOM_FORMULA(${a1Cell})`,
            });

            expect(result.data).toHaveLength(3); // Header + 2 students
            expect(result.data[0]).toHaveLength(10); // 4 frozen columns + 2 subjects * 3 columns

            const expectedData: GoogleAppsScript.Sheets.Schema.CellData[][] = [
                [
                    { userEnteredValue: { stringValue: title } },
                    {},
                    {},
                    {},
                    { userEnteredValue: { stringValue: "Subject1" } },
                    {},
                    {},
                    { userEnteredValue: { stringValue: "Subject2" } },
                    {},
                    {},
                ],
                [
                    { userEnteredValue: { formulaValue: expect.stringContaining("$E7:7") } },
                    { userEnteredValue: { numberValue: 101 } },
                    { userEnteredValue: { stringValue: "John" } },
                    { userEnteredValue: { stringValue: "Doe" } },
                    { userEnteredValue: { formulaValue: "=CUSTOM_FORMULA('JohnSheet'!$B$26)" } },
                    {},
                    {},
                    { userEnteredValue: { formulaValue: "=CUSTOM_FORMULA('JohnSheet'!$B$27)" } },
                    {},
                    {},
                ],
                [
                    { userEnteredValue: { formulaValue: expect.stringContaining("$E8:8") } },
                    { userEnteredValue: { numberValue: 102 } },
                    { userEnteredValue: { stringValue: "Jane" } },
                    { userEnteredValue: { stringValue: "Deer" } },
                    { userEnteredValue: { formulaValue: "=CUSTOM_FORMULA('JaneSheet'!$B$26)" } },
                    {},
                    {},
                    { userEnteredValue: { formulaValue: "=CUSTOM_FORMULA('JaneSheet'!$B$27)" } },
                    {},
                    {},
                ],
            ];

            const expectedMergeRanges: GoogleAppsScript.Sheets.Schema.GridRange[] = [
                { sheetId: 100, startRowIndex: 5, startColumnIndex: 4, endRowIndex: 8, endColumnIndex: 7 },
                { sheetId: 100, startRowIndex: 5, startColumnIndex: 7, endRowIndex: 8, endColumnIndex: 10 },
            ];

            const expectedBorderRanges: GoogleAppsScript.Sheets.Schema.GridRange[] = [
                { sheetId: 100, startRowIndex: 5, startColumnIndex: 4, endRowIndex: 8, endColumnIndex: 7 },
                { sheetId: 100, startRowIndex: 5, startColumnIndex: 7, endRowIndex: 8, endColumnIndex: 10 },
            ];

            expect(result.data).toEqual(expectedData);
            expect(result.mergeRanges).toEqual(expectedMergeRanges);
            expect(result.borderRanges).toEqual(expectedBorderRanges);
        });

        it("shold correctly generate data with multiple formulas", () => {
            const title = "Status Title";

            const result = buildStatusSectionData({
                statusRange: mockStatusRange,
                studentRange: mockStudentRange,
                markColOffsets: [0, 3],
                title,
                headers: ["Subject1", "Subject2"],
                studentRows: [mockSeparator, mockStudent1],
                statusRowOffset: 5,
                formulaFunction: [(a1Cell: string) => `=FORMULA1(${a1Cell})`, (a1Cell: string) => `=FORMULA2(${a1Cell})`],
            });

            expect(result.data).toHaveLength(3); // Header + 2 students
            expect(result.data[0]).toHaveLength(12); // 4 frozen columns + 2 subjects * 4 columns

            const expectedData: GoogleAppsScript.Sheets.Schema.CellData[][] = [
                [
                    { userEnteredValue: { stringValue: title } },
                    {},
                    {},
                    {},
                    { userEnteredValue: { stringValue: "Subject1" } },
                    {},
                    {},
                    {},
                    { userEnteredValue: { stringValue: "Subject2" } },
                    {},
                    {},
                    {},
                ],
                [],
                [
                    { userEnteredValue: { formulaValue: expect.stringContaining("$E8:8") } },
                    { userEnteredValue: { numberValue: 101 } },
                    { userEnteredValue: { stringValue: "John" } },
                    { userEnteredValue: { stringValue: "Doe" } },
                    { userEnteredValue: { formulaValue: "=FORMULA1('JohnSheet'!$B$26)" } },
                    {},
                    { userEnteredValue: { formulaValue: "=FORMULA2('JohnSheet'!$E$26)" } },
                    {},
                    { userEnteredValue: { formulaValue: "=FORMULA1('JohnSheet'!$B$27)" } },
                    {},
                    { userEnteredValue: { formulaValue: "=FORMULA2('JohnSheet'!$E$27)" } },
                    {},
                ],
            ];

            const expectedMergeRanges: GoogleAppsScript.Sheets.Schema.GridRange[] = [
                { sheetId: 100, startRowIndex: 5, startColumnIndex: 4, endRowIndex: 6, endColumnIndex: 8 },
                { sheetId: 100, startRowIndex: 6, startColumnIndex: 4, endRowIndex: 8, endColumnIndex: 6 },
                { sheetId: 100, startRowIndex: 6, startColumnIndex: 6, endRowIndex: 8, endColumnIndex: 8 },
                { sheetId: 100, startRowIndex: 5, startColumnIndex: 8, endRowIndex: 6, endColumnIndex: 12 },
                { sheetId: 100, startRowIndex: 6, startColumnIndex: 8, endRowIndex: 8, endColumnIndex: 10 },
                { sheetId: 100, startRowIndex: 6, startColumnIndex: 10, endRowIndex: 8, endColumnIndex: 12 },
            ];

            const expectedBorderRanges: GoogleAppsScript.Sheets.Schema.GridRange[] = [
                { sheetId: 100, startRowIndex: 5, startColumnIndex: 4, endRowIndex: 8, endColumnIndex: 8 },
                { sheetId: 100, startRowIndex: 5, startColumnIndex: 8, endRowIndex: 8, endColumnIndex: 12 },
            ];

            expect(result.data).toEqual(expectedData);
            expect(result.mergeRanges).toEqual(expectedMergeRanges);
            expect(result.borderRanges).toEqual(expectedBorderRanges);
        });

        it("should throw error if not enough formulas", () => {
            const title = "Status Title";

            expect(() =>
                buildStatusSectionData({
                    statusRange: mockStatusRange,
                    studentRange: mockStudentRange,
                    markColOffsets: [0, 3],
                    title,
                    headers: ["Subject1", "Subject2"],
                    studentRows: [mockSeparator, mockStudent1],
                    statusRowOffset: 5,
                    formulaFunction: [(a1Cell: string) => `=FORMULA1(${a1Cell})`],
                }),
            ).toThrow();
        });

        it("should handle educational fields", () => {
            const title = "Status Title";

            const result = buildStatusSectionData({
                statusRange: mockStatusRange,
                studentRange: mockStudentRange,
                markColOffsets: [0, 1, 2],
                markUsesFieldValue: [true, false, true],
                title,
                headers: ["Subject1", "Subject2", "Subject3"],
                studentRows: [mockStudent1],
                averagePerField: true,
                fieldSubjects: [0, 0],
                statusRowOffset: 5,
                formulaFunction: (a1Cell: string) => `=CUSTOM_FORMULA(${a1Cell})`,
            });

            expect(result.data).toHaveLength(2); // Header + 1 students
            expect(result.data[0]).toHaveLength(13); // 4 frozen columns + 3 subjects * 3 columns

            const expectedData: GoogleAppsScript.Sheets.Schema.CellData[][] = [
                [
                    { userEnteredValue: { stringValue: title } },
                    {},
                    {},
                    {},
                    { userEnteredValue: { stringValue: "Subject1" } },
                    {},
                    {},
                    { userEnteredValue: { stringValue: "Subject2" } },
                    {},
                    {},
                    { userEnteredValue: { stringValue: "Subject3" } },
                    {},
                    {},
                ],
                [
                    { userEnteredValue: { formulaValue: expect.stringContaining("$E7:7") } },
                    { userEnteredValue: { numberValue: 101 } },
                    { userEnteredValue: { stringValue: "John" } },
                    { userEnteredValue: { stringValue: "Doe" } },
                    { userEnteredValue: { formulaValue: "=CUSTOM_FORMULA('JohnSheet'!$B$26)" } },
                    { userEnteredValue: { formulaValue: "=CUSTOM_FORMULA('JohnSheet'!$C$26)" } },
                    { userEnteredValue: { formulaValue: "=CUSTOM_FORMULA('JohnSheet'!$D$26)" } },
                    { userEnteredValue: { formulaValue: "=CUSTOM_FORMULA('JohnSheet'!$B$26)" } },
                    { userEnteredValue: { formulaValue: "=CUSTOM_FORMULA('JohnSheet'!$C$27)" } },
                    { userEnteredValue: { formulaValue: "=CUSTOM_FORMULA('JohnSheet'!$D$26)" } },
                    { userEnteredValue: { formulaValue: "=CUSTOM_FORMULA('JohnSheet'!$B$28)" } },
                    { userEnteredValue: { formulaValue: "=CUSTOM_FORMULA('JohnSheet'!$C$28)" } },
                    { userEnteredValue: { formulaValue: "=CUSTOM_FORMULA('JohnSheet'!$D$28)" } },
                ],
            ];

            const expectedMergeRanges: GoogleAppsScript.Sheets.Schema.GridRange[] = [
                { sheetId: 100, startRowIndex: 5, startColumnIndex: 4, endRowIndex: 6, endColumnIndex: 7 },
                { sheetId: 100, startRowIndex: 5, startColumnIndex: 7, endRowIndex: 6, endColumnIndex: 10 },
                { sheetId: 100, startRowIndex: 5, startColumnIndex: 10, endRowIndex: 6, endColumnIndex: 13 },
            ];

            const expectedBorderRanges: GoogleAppsScript.Sheets.Schema.GridRange[] = [
                { sheetId: 100, startRowIndex: 5, startColumnIndex: 4, endRowIndex: 7, endColumnIndex: 7 },
                { sheetId: 100, startRowIndex: 5, startColumnIndex: 7, endRowIndex: 7, endColumnIndex: 10 },
                { sheetId: 100, startRowIndex: 5, startColumnIndex: 10, endRowIndex: 7, endColumnIndex: 13 },
            ];

            expect(result.data).toEqual(expectedData);
            expect(result.mergeRanges).toEqual(expectedMergeRanges);
            expect(result.borderRanges).toEqual(expectedBorderRanges);
        });

        it("should handle averagePerField without fieldSubjects provided gracefully", () => {
            const title = "Status Title";

            const result = buildStatusSectionData({
                statusRange: mockStatusRange,
                studentRange: mockStudentRange,
                markColOffsets: [0],
                markUsesFieldValue: [true],
                title,
                headers: ["Subject1", "Subject2"],
                studentRows: [mockStudent1],
                averagePerField: true,
                fieldSubjects: undefined, // Explicitly omitted to test fallback logic
                statusRowOffset: 5,
                formulaFunction: (a1Cell: string) => `=CUSTOM_FORMULA(${a1Cell})`,
            });

            expect(result.data).toHaveLength(2); // Header + 1 student

            const expectedData: GoogleAppsScript.Sheets.Schema.CellData[][] = [
                [
                    { userEnteredValue: { stringValue: title } },
                    {},
                    {},
                    {},
                    { userEnteredValue: { stringValue: "Subject1" } },
                    {},
                    {},
                    { userEnteredValue: { stringValue: "Subject2" } },
                    {},
                    {},
                ],
                [
                    { userEnteredValue: { formulaValue: expect.stringContaining("$E7:7") } },
                    { userEnteredValue: { numberValue: 101 } },
                    { userEnteredValue: { stringValue: "John" } },
                    { userEnteredValue: { stringValue: "Doe" } },
                    { userEnteredValue: { formulaValue: "=CUSTOM_FORMULA('JohnSheet'!$B$26)" } },
                    {},
                    {},
                    // Because fieldSubjects is missing and we have fewer students (1) than headers (2),
                    // this tests the fallback logic to ensure rowOffset resolves properly
                    { userEnteredValue: { formulaValue: "=CUSTOM_FORMULA('JohnSheet'!$B$27)" } },
                    {},
                    {},
                ],
            ];

            expect(result.data).toEqual(expectedData);
        });
    });
});
