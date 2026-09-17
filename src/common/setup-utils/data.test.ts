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
import {
    buildDefaultCommentForStudentTemplateData,
    buildStatusSectionData,
    buildSummaryHeadersData,
    buildSummaryStudentData,
    calculateCalendarHeaders,
    generateStudentGrid,
    type PeriodRanges,
} from "./data";
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

    describe("buildSummaryHeadersData", () => {
        it("should build summary headers for general attendance without fields", () => {
            const attendancePerClass = false;
            const averagePerField = false;
            const subjectsNames = ["art", "math", "biology", "english"];
            const fieldNames = ["language", "science", "humanities"];

            const data = buildSummaryHeadersData(attendancePerClass, averagePerField, subjectsNames, fieldNames);

            const expectedData: GoogleAppsScript.Sheets.Schema.CellData[][] = [
                [
                    { userEnteredValue: { stringValue: "Faltas" }, userEnteredFormat: expect.any(Object) },
                    { userEnteredValue: { stringValue: "art" } },
                    { userEnteredFormat: expect.any(Object) },
                    { userEnteredValue: { stringValue: "math" } },
                    { userEnteredFormat: expect.any(Object) },
                    { userEnteredValue: { stringValue: "biology" } },
                    { userEnteredFormat: expect.any(Object) },
                    { userEnteredValue: { stringValue: "english" } },
                    { userEnteredFormat: expect.any(Object) },
                    { userEnteredValue: { stringValue: "Promedio" } },
                ],
                [
                    { userEnteredFormat: expect.any(Object) },
                    { userEnteredValue: { stringValue: "Cal" } },
                    { userEnteredValue: { stringValue: "SEP" }, userEnteredFormat: expect.any(Object) },
                    { userEnteredValue: { stringValue: "Cal" } },
                    { userEnteredValue: { stringValue: "SEP" }, userEnteredFormat: expect.any(Object) },
                    { userEnteredValue: { stringValue: "Cal" } },
                    { userEnteredValue: { stringValue: "SEP" }, userEnteredFormat: expect.any(Object) },
                    { userEnteredValue: { stringValue: "Cal" } },
                    { userEnteredValue: { stringValue: "SEP" }, userEnteredFormat: expect.any(Object) },
                    {},
                ],
            ];

            expect(data).toEqual(expectedData);
        });

        it("should build summary headers for individual attendance without fields", () => {
            const attendancePerClass = true;
            const averagePerField = false;
            const subjectsNames = ["art", "math", "biology", "english"];
            const fieldNames = ["language", "science", "humanities"];

            const data = buildSummaryHeadersData(attendancePerClass, averagePerField, subjectsNames, fieldNames);

            const expectedData: GoogleAppsScript.Sheets.Schema.CellData[][] = [
                [
                    { userEnteredValue: { stringValue: "art" } },
                    {},
                    { userEnteredFormat: expect.any(Object) },
                    { userEnteredValue: { stringValue: "math" } },
                    {},
                    { userEnteredFormat: expect.any(Object) },
                    { userEnteredValue: { stringValue: "biology" } },
                    {},
                    { userEnteredFormat: expect.any(Object) },
                    { userEnteredValue: { stringValue: "english" } },
                    {},
                    { userEnteredFormat: expect.any(Object) },
                    { userEnteredValue: { stringValue: "Promedio" } },
                ],
                [
                    { userEnteredValue: { stringValue: "Fal" } },
                    { userEnteredValue: { stringValue: "Cal" } },
                    { userEnteredValue: { stringValue: "SEP" }, userEnteredFormat: expect.any(Object) },
                    { userEnteredValue: { stringValue: "Fal" } },
                    { userEnteredValue: { stringValue: "Cal" } },
                    { userEnteredValue: { stringValue: "SEP" }, userEnteredFormat: expect.any(Object) },
                    { userEnteredValue: { stringValue: "Fal" } },
                    { userEnteredValue: { stringValue: "Cal" } },
                    { userEnteredValue: { stringValue: "SEP" }, userEnteredFormat: expect.any(Object) },
                    { userEnteredValue: { stringValue: "Fal" } },
                    { userEnteredValue: { stringValue: "Cal" } },
                    { userEnteredValue: { stringValue: "SEP" }, userEnteredFormat: expect.any(Object) },
                    {},
                ],
            ];

            expect(data).toEqual(expectedData);
        });

        it("should build summary headers for general attendance with fields", () => {
            const attendancePerClass = false;
            const averagePerField = true;
            const subjectsNames = ["art", "math", "biology", "english"];
            const fieldNames = ["language", "science", "humanities"];

            const data = buildSummaryHeadersData(attendancePerClass, averagePerField, subjectsNames, fieldNames);

            const expectedData: GoogleAppsScript.Sheets.Schema.CellData[][] = [
                [
                    { userEnteredValue: { stringValue: "art" }, userEnteredFormat: expect.any(Object) },
                    { userEnteredValue: { stringValue: "math" }, userEnteredFormat: expect.any(Object) },
                    { userEnteredValue: { stringValue: "biology" }, userEnteredFormat: expect.any(Object) },
                    { userEnteredValue: { stringValue: "english" }, userEnteredFormat: expect.any(Object) },
                    { userEnteredFormat: expect.any(Object) },
                    { userEnteredValue: { stringValue: "Faltas" }, userEnteredFormat: expect.any(Object) },
                    { userEnteredValue: { stringValue: "language" } },
                    { userEnteredFormat: expect.any(Object) },
                    { userEnteredValue: { stringValue: "science" } },
                    { userEnteredFormat: expect.any(Object) },
                    { userEnteredValue: { stringValue: "humanities" } },
                    { userEnteredFormat: expect.any(Object) },
                    { userEnteredValue: { stringValue: "Promedio" } },
                ],
                [
                    { userEnteredValue: { stringValue: "Cal" }, userEnteredFormat: expect.any(Object) },
                    { userEnteredValue: { stringValue: "Cal" }, userEnteredFormat: expect.any(Object) },
                    { userEnteredValue: { stringValue: "Cal" }, userEnteredFormat: expect.any(Object) },
                    { userEnteredValue: { stringValue: "Cal" }, userEnteredFormat: expect.any(Object) },
                    { userEnteredFormat: expect.any(Object) },
                    { userEnteredFormat: expect.any(Object) },
                    { userEnteredValue: { stringValue: "Cal" } },
                    { userEnteredValue: { stringValue: "SEP" }, userEnteredFormat: expect.any(Object) },
                    { userEnteredValue: { stringValue: "Cal" } },
                    { userEnteredValue: { stringValue: "SEP" }, userEnteredFormat: expect.any(Object) },
                    { userEnteredValue: { stringValue: "Cal" } },
                    { userEnteredValue: { stringValue: "SEP" }, userEnteredFormat: expect.any(Object) },
                    {},
                ],
            ];

            expect(data).toEqual(expectedData);
        });

        it("should build summary headers for individual attendance with fields", () => {
            const attendancePerClass = true;
            const averagePerField = true;
            const subjectsNames = ["art", "math", "biology", "english"];
            const fieldNames = ["language", "science", "humanities"];

            const data = buildSummaryHeadersData(attendancePerClass, averagePerField, subjectsNames, fieldNames);

            const expectedData: GoogleAppsScript.Sheets.Schema.CellData[][] = [
                [
                    { userEnteredValue: { stringValue: "art" } },
                    { userEnteredFormat: expect.any(Object) },
                    { userEnteredValue: { stringValue: "math" } },
                    { userEnteredFormat: expect.any(Object) },
                    { userEnteredValue: { stringValue: "biology" } },
                    { userEnteredFormat: expect.any(Object) },
                    { userEnteredValue: { stringValue: "english" } },
                    { userEnteredFormat: expect.any(Object) },
                    { userEnteredFormat: expect.any(Object) },
                    { userEnteredValue: { stringValue: "language" } },
                    { userEnteredFormat: expect.any(Object) },
                    { userEnteredValue: { stringValue: "science" } },
                    { userEnteredFormat: expect.any(Object) },
                    { userEnteredValue: { stringValue: "humanities" } },
                    { userEnteredFormat: expect.any(Object) },
                    { userEnteredValue: { stringValue: "Promedio" } },
                ],
                [
                    { userEnteredValue: { stringValue: "Fal" } },
                    { userEnteredValue: { stringValue: "Cal" }, userEnteredFormat: expect.any(Object) },
                    { userEnteredValue: { stringValue: "Fal" } },
                    { userEnteredValue: { stringValue: "Cal" }, userEnteredFormat: expect.any(Object) },
                    { userEnteredValue: { stringValue: "Fal" } },
                    { userEnteredValue: { stringValue: "Cal" }, userEnteredFormat: expect.any(Object) },
                    { userEnteredValue: { stringValue: "Fal" } },
                    { userEnteredValue: { stringValue: "Cal" }, userEnteredFormat: expect.any(Object) },
                    { userEnteredFormat: expect.any(Object) },
                    { userEnteredValue: { stringValue: "Cal" } },
                    { userEnteredValue: { stringValue: "SEP" }, userEnteredFormat: expect.any(Object) },
                    { userEnteredValue: { stringValue: "Cal" } },
                    { userEnteredValue: { stringValue: "SEP" }, userEnteredFormat: expect.any(Object) },
                    { userEnteredValue: { stringValue: "Cal" } },
                    { userEnteredValue: { stringValue: "SEP" }, userEnteredFormat: expect.any(Object) },
                    {},
                ],
            ];

            expect(data).toEqual(expectedData);
        });
    });

    describe("buildSummaryStudentData", () => {
        const subjects = 4;
        const fields = [3, 1];
        const students: StudentRow[] = [
            {
                type: StudentRowType.STUDENT,
                id: 1,
                firstName: "Yama",
                lastName: "Nanqui",
                sheetName: "yama nanqui",
                sex: "M",
                level: "Preschool",
                grade: "3rd",
                curp: "yama1234",
            },
            {
                type: StudentRowType.SEPARATOR,
            },
            {
                type: StudentRowType.STUDENT,
                id: 3,
                firstName: "Erin",
                lastName: "Smith",
                sheetName: "erin smith",
                sex: "F",
                level: "High school",
                grade: "9th",
                curp: "erin5678",
            },
        ];
        const assistanceSheetName = "Asistencia";
        const commentsRange: MappedNamedRange = {
            sheet: {},
            namedRange: {
                namedRangeId: "coments",
                name: "comments",
                range: { sheetId: 123, startColumnIndex: 0, endColumnIndex: 10, startRowIndex: 20, endRowIndex: 24 },
            },
        };
        const subjectRanges: PeriodRanges = [
            {
                sheet: {},
                namedRange: {
                    namedRangeId: "subjects",
                    name: "subjects",
                    range: { sheetId: 123, startColumnIndex: 0, endColumnIndex: 5, startRowIndex: 30, endRowIndex: 34 },
                },
            },
            {
                sheet: {},
                namedRange: {
                    namedRangeId: "subjects",
                    name: "subjects",
                    range: { sheetId: 123, startColumnIndex: 0, endColumnIndex: 5, startRowIndex: 40, endRowIndex: 44 },
                },
            },
            {
                sheet: {},
                namedRange: {
                    namedRangeId: "subjects",
                    name: "subjects",
                    range: { sheetId: 123, startColumnIndex: 0, endColumnIndex: 6, startRowIndex: 50, endRowIndex: 54 },
                },
            },
        ];
        const fieldRanges: PeriodRanges = [
            {
                sheet: {},
                namedRange: {
                    namedRangeId: "fields",
                    name: "fields",
                    range: { sheetId: 123, startColumnIndex: 0, endColumnIndex: 5, startRowIndex: 35, endRowIndex: 37 },
                },
            },
            {
                sheet: {},
                namedRange: {
                    namedRangeId: "fields",
                    name: "fields",
                    range: { sheetId: 123, startColumnIndex: 0, endColumnIndex: 5, startRowIndex: 45, endRowIndex: 47 },
                },
            },
            {
                sheet: {},
                namedRange: {
                    namedRangeId: "fields",
                    name: "fields",
                    range: { sheetId: 123, startColumnIndex: 0, endColumnIndex: 6, startRowIndex: 55, endRowIndex: 57 },
                },
            },
        ];
        const averageRanges: PeriodRanges = [
            {
                sheet: {},
                namedRange: {
                    namedRangeId: "fields",
                    name: "fields",
                    range: { sheetId: 123, startColumnIndex: 0, endColumnIndex: 5, startRowIndex: 37, endRowIndex: 38 },
                },
            },
            {
                sheet: {},
                namedRange: {
                    namedRangeId: "fields",
                    name: "fields",
                    range: { sheetId: 123, startColumnIndex: 0, endColumnIndex: 5, startRowIndex: 47, endRowIndex: 48 },
                },
            },
            {
                sheet: {},
                namedRange: {
                    namedRangeId: "fields",
                    name: "fields",
                    range: { sheetId: 123, startColumnIndex: 0, endColumnIndex: 6, startRowIndex: 57, endRowIndex: 58 },
                },
            },
        ];
        it("should build summary student data for general attendance and simple average", () => {
            // id | Name | Last Name | Attendance | Subject N grade | Subject N Comment | ... | Average
            const result = buildSummaryStudentData(
                false,
                false,
                subjects,
                fields,
                students,
                0,
                assistanceSheetName,
                commentsRange,
                subjectRanges,
                fieldRanges,
                averageRanges,
            );

            const expectedResult: GoogleAppsScript.Sheets.Schema.CellData[][] = [
                [
                    { userEnteredValue: { numberValue: 1 } }, // id
                    { userEnteredValue: { stringValue: "Yama" } }, // Name
                    { userEnteredValue: { stringValue: "Nanqui" } }, // Last Name
                    { userEnteredValue: { formulaValue: `=FILTER('Asistencia'!$E$4:$E, 'Asistencia'!$B$4:$B = "Yama", 'Asistencia'!$C$4:$C = "Nanqui")` } }, // Attendance
                    { userEnteredValue: { formulaValue: `='yama nanqui'!$E$31` } }, // Subject 0 grade
                    { userEnteredValue: { formulaValue: `='yama nanqui'!$I$21` } }, // Subject 0 comment
                    { userEnteredValue: { formulaValue: `='yama nanqui'!$E$32` } }, // Subject 1 grade
                    { userEnteredValue: { formulaValue: `='yama nanqui'!$I$22` } }, // Subject 1 comment
                    { userEnteredValue: { formulaValue: `='yama nanqui'!$E$33` } }, // Subject 2 grade
                    { userEnteredValue: { formulaValue: `='yama nanqui'!$I$23` } }, // Subject 2 comment
                    { userEnteredValue: { formulaValue: `='yama nanqui'!$E$34` } }, // Subject 3 grade
                    { userEnteredValue: { formulaValue: `='yama nanqui'!$I$24` } }, // Subject 3 comment
                    { userEnteredValue: { formulaValue: `='yama nanqui'!$E$38` } }, // Average
                ],
                [],
                [
                    { userEnteredValue: { numberValue: 3 } }, // id
                    { userEnteredValue: { stringValue: "Erin" } }, // Name
                    { userEnteredValue: { stringValue: "Smith" } }, // Last Name
                    { userEnteredValue: { formulaValue: `=FILTER('Asistencia'!$E$4:$E, 'Asistencia'!$B$4:$B = "Erin", 'Asistencia'!$C$4:$C = "Smith")` } },
                    { userEnteredValue: { formulaValue: `='erin smith'!$E$31` } }, // Subject 0 grade
                    { userEnteredValue: { formulaValue: `='erin smith'!$I$21` } }, // Subject 0 comment
                    { userEnteredValue: { formulaValue: `='erin smith'!$E$32` } }, // Subject 1 grade
                    { userEnteredValue: { formulaValue: `='erin smith'!$I$22` } }, // Subject 1 comment
                    { userEnteredValue: { formulaValue: `='erin smith'!$E$33` } }, // Subject 2 grade
                    { userEnteredValue: { formulaValue: `='erin smith'!$I$23` } }, // Subject 2 comment
                    { userEnteredValue: { formulaValue: `='erin smith'!$E$34` } }, // Subject 3 grade
                    { userEnteredValue: { formulaValue: `='erin smith'!$I$24` } }, // Subject 3 comment
                    { userEnteredValue: { formulaValue: `='erin smith'!$E$38` } }, // Average
                ],
            ];

            expect(result).toEqual(expectedResult);
        });
        it("should build summary student data for general attendance and field average", () => {
            // id | Name | Last Name |  Subject N grade | ... | Space | Attendance | Field N grade | Field N Comment | Average
            const result = buildSummaryStudentData(
                false,
                true,
                subjects,
                fields,
                students,
                0,
                assistanceSheetName,
                commentsRange,
                subjectRanges,
                fieldRanges,
                averageRanges,
            );

            const expectedResult: GoogleAppsScript.Sheets.Schema.CellData[][] = [
                [
                    { userEnteredValue: { numberValue: 1 } }, // id
                    { userEnteredValue: { stringValue: "Yama" } }, // Name
                    { userEnteredValue: { stringValue: "Nanqui" } }, // Last Name
                    { userEnteredValue: { formulaValue: `='yama nanqui'!$E$31` } }, // Subject 0 grade
                    { userEnteredValue: { formulaValue: `='yama nanqui'!$E$32` } }, // Subject 1 grade
                    { userEnteredValue: { formulaValue: `='yama nanqui'!$E$33` } }, // Subject 2 grade
                    { userEnteredValue: { formulaValue: `='yama nanqui'!$E$34` } }, // Subject 3 grade
                    {}, // Space
                    { userEnteredValue: { formulaValue: `=FILTER('Asistencia'!$E$4:$E, 'Asistencia'!$B$4:$B = "Yama", 'Asistencia'!$C$4:$C = "Nanqui")` } }, // Attendance
                    { userEnteredValue: { formulaValue: `='yama nanqui'!$E$36` } }, // Field 0 grade
                    { userEnteredValue: { formulaValue: `='yama nanqui'!$I$21` } }, // Field 0 comment
                    { userEnteredValue: { formulaValue: `='yama nanqui'!$E$37` } }, // Field 1 grade
                    { userEnteredValue: { formulaValue: `='yama nanqui'!$I$24` } }, // Field 1 comment
                    { userEnteredValue: { formulaValue: `='yama nanqui'!$E$38` } }, // Average
                ],
                [],
                [
                    { userEnteredValue: { numberValue: 3 } }, // id
                    { userEnteredValue: { stringValue: "Erin" } }, // Name
                    { userEnteredValue: { stringValue: "Smith" } }, // Last Name
                    { userEnteredValue: { formulaValue: `='erin smith'!$E$31` } }, // Subject 0 grade
                    { userEnteredValue: { formulaValue: `='erin smith'!$E$32` } }, // Subject 1 grade
                    { userEnteredValue: { formulaValue: `='erin smith'!$E$33` } }, // Subject 2 grade
                    { userEnteredValue: { formulaValue: `='erin smith'!$E$34` } }, // Subject 3 grade
                    {}, // Space
                    { userEnteredValue: { formulaValue: `=FILTER('Asistencia'!$E$4:$E, 'Asistencia'!$B$4:$B = "Erin", 'Asistencia'!$C$4:$C = "Smith")` } }, // Attendance
                    { userEnteredValue: { formulaValue: `='erin smith'!$E$36` } }, // Field 0 grade
                    { userEnteredValue: { formulaValue: `='erin smith'!$I$21` } }, // Field 0 comment
                    { userEnteredValue: { formulaValue: `='erin smith'!$E$37` } }, // Field 1 grade
                    { userEnteredValue: { formulaValue: `='erin smith'!$I$24` } }, // Field 1 comment
                    { userEnteredValue: { formulaValue: `='erin smith'!$E$38` } }, // Average
                ],
            ];

            expect(result).toEqual(expectedResult);
        });
        it("should build summary student data for individual attendance and simple average", () => {
            // id | Name | Last Name | Subject N attendance | Subject N grade | Subject N comment | ... | Average
            const result = buildSummaryStudentData(
                true,
                false,
                subjects,
                fields,
                students,
                0,
                assistanceSheetName,
                commentsRange,
                subjectRanges,
                fieldRanges,
                averageRanges,
            );

            const expectedResult: GoogleAppsScript.Sheets.Schema.CellData[][] = [
                [
                    { userEnteredValue: { numberValue: 1 } }, // id
                    { userEnteredValue: { stringValue: "Yama" } }, // Name
                    { userEnteredValue: { stringValue: "Nanqui" } }, // Last Name
                    {}, // Subject 0 attendance
                    { userEnteredValue: { formulaValue: `='yama nanqui'!$E$31` } }, // Subject 0 grade
                    { userEnteredValue: { formulaValue: `='yama nanqui'!$I$21` } }, // Subject 0 comment
                    { userEnteredValue: { formulaValue: `='yama nanqui'!$E$32` } }, // Subject 1 grade
                    { userEnteredValue: { formulaValue: `='yama nanqui'!$I$22` } }, // Subject 1 comment
                    { userEnteredValue: { formulaValue: `='yama nanqui'!$E$33` } }, // Subject 2 grade
                    { userEnteredValue: { formulaValue: `='yama nanqui'!$I$23` } }, // Subject 2 comment
                    { userEnteredValue: { formulaValue: `='yama nanqui'!$E$34` } }, // Subject 3 grade
                    { userEnteredValue: { formulaValue: `='yama nanqui'!$I$24` } }, // Subject 3 comment
                    { userEnteredValue: { formulaValue: `='yama nanqui'!$E$38` } }, // Average
                ],
                [],
                [
                    { userEnteredValue: { numberValue: 3 } }, // id
                    { userEnteredValue: { stringValue: "Erin" } }, // Name
                    { userEnteredValue: { stringValue: "Smith" } }, // Last Name
                    { userEnteredValue: { formulaValue: `=FILTER('Asistencia'!$E$4:$E, 'Asistencia'!$B$4:$B = "Erin", 'Asistencia'!$C$4:$C = "Smith")` } },
                    { userEnteredValue: { formulaValue: `='erin smith'!$E$31` } }, // Subject 0 grade
                    { userEnteredValue: { formulaValue: `='erin smith'!$I$21` } }, // Subject 0 comment
                    { userEnteredValue: { formulaValue: `='erin smith'!$E$32` } }, // Subject 1 grade
                    { userEnteredValue: { formulaValue: `='erin smith'!$I$22` } }, // Subject 1 comment
                    { userEnteredValue: { formulaValue: `='erin smith'!$E$33` } }, // Subject 2 grade
                    { userEnteredValue: { formulaValue: `='erin smith'!$I$23` } }, // Subject 2 comment
                    { userEnteredValue: { formulaValue: `='erin smith'!$E$34` } }, // Subject 3 grade
                    { userEnteredValue: { formulaValue: `='erin smith'!$I$24` } }, // Subject 3 comment
                    { userEnteredValue: { formulaValue: `='erin smith'!$E$38` } }, // Average
                ],
            ];
        });
        it("should build summary student data for individual attendance and field average", () => {});
    });
});
