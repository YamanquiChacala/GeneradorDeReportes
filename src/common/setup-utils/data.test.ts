import { getA1Notation, type MappedNamedRange } from "../gas-utils";
import { createAttendaceFormulas, createSetupRowValidFormula, type StudentRow, StudentRowType, type TrimesterRanges } from "../report-utils";
import { buildStatusSectionData, calculateCalendarHeaders, generateStudentGrid, type StatusSectionDataParams } from "./data";
import { TemplateSize } from "./types";

// 1. Setup Typed Mocks
jest.mock("../report-utils", () => {
    const originalModule = jest.requireActual("../report-utils");
    return {
        ...originalModule,
        createAttendaceFormulas: jest.fn(),
        createSetupRowValidFormula: jest.fn(),
    };
});
jest.mock("../gas-utils", () => {
    const originalModule = jest.requireActual("../gas-utils");
    return {
        ...originalModule,
        getA1Notation: jest.fn(),
    };
});

const mockCreateAttendaceFormulas = createAttendaceFormulas as jest.MockedFunction<typeof createAttendaceFormulas>;
const mockCreateSetupRowValidFormula = createSetupRowValidFormula as jest.MockedFunction<typeof createSetupRowValidFormula>;
const mockGetA1Notation = getA1Notation as jest.MockedFunction<typeof getA1Notation>;

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
    });

    describe("generateStudentGrid", () => {
        beforeEach(() => {
            jest.clearAllMocks();
        });

        const mockTrimesters: TrimesterRanges = {
            trim1: { start: 1, end: 5 },
            trim2: { start: 6, end: 10 },
            trim3: { start: 11, end: 15 },
        };

        it("should return an empty array if no students are provided", () => {
            const result = generateStudentGrid([], 5, mockTrimesters);
            expect(result).toEqual([]);
            expect(mockCreateAttendaceFormulas).not.toHaveBeenCalled();
        });

        it("should push empty arrays for non-student rows", () => {
            // Natively satisfies the StudentSpace interface
            const students: StudentRow[] = [{ type: StudentRowType.SEPARATOR }];

            const result = generateStudentGrid(students, 5, mockTrimesters);

            expect(result).toEqual([[]]);
            expect(mockCreateAttendaceFormulas).not.toHaveBeenCalled();
        });

        it("should generate cell data correctly for student rows", () => {
            // Natively satisfies the Student interface
            const students: StudentRow[] = [
                {
                    type: StudentRowType.STUDENT,
                    id: 99,
                    firstName: "Jane",
                    lastName: "Doe",
                    sheetName: "Sheet1",
                    sex: "F",
                    level: "Elementary",
                    grade: "1st",
                    curp: "ABCD123456",
                },
            ];

            const initialRow = 2;

            // Mock the formulas returning specific strings so we can verify they map correctly
            mockCreateAttendaceFormulas
                .mockReturnValueOnce({ percent: "=P1", count: "=C1" }) // trim 1
                .mockReturnValueOnce({ percent: "=P2", count: "=C2" }) // trim 2
                .mockReturnValueOnce({ percent: "=P3", count: "=C3" }); // trim 3

            const result = generateStudentGrid(students, initialRow, mockTrimesters);

            // 1. Verify the dependency was called with the correct rows and column bounds
            expect(mockCreateAttendaceFormulas).toHaveBeenCalledTimes(3);
            expect(mockCreateAttendaceFormulas).toHaveBeenNthCalledWith(1, 2, 1, 5);
            expect(mockCreateAttendaceFormulas).toHaveBeenNthCalledWith(2, 2, 6, 10);
            expect(mockCreateAttendaceFormulas).toHaveBeenNthCalledWith(3, 2, 11, 15);

            // 2. Verify the mapping structure output
            expect(result).toEqual([
                [
                    { userEnteredValue: { numberValue: 99 } },
                    { userEnteredValue: { stringValue: "Jane" } },
                    { userEnteredValue: { stringValue: "Doe" } },
                    { userEnteredValue: { formulaValue: "=P1" } },
                    { userEnteredValue: { formulaValue: "=C1" } },
                    { userEnteredValue: { formulaValue: "=P2" } },
                    { userEnteredValue: { formulaValue: "=C2" } },
                    { userEnteredValue: { formulaValue: "=P3" } },
                    { userEnteredValue: { formulaValue: "=C3" } },
                ],
            ]);
        });
    });

    describe("buildStatusSectionData", () => {
        // Setup dummy data
        const mockStatusRange = {} as MappedNamedRange;
        const mockStudentRange = {} as MappedNamedRange;
        const title = "Status Title";
        const headers = ["HW 1", "HW 2"];
        const formulaFunction = (a1Cell: string) => `=CUSTOM_FORMULA(${a1Cell})`;

        const mockStudent: StudentRow = {
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
        const mockSeparator: StudentRow = { type: StudentRowType.SEPARATOR };

        beforeEach(() => {
            jest.clearAllMocks();
        });

        it("should correctly format the header row with colsPerItem padding", () => {
            const params: StatusSectionDataParams = {
                statusRange: mockStatusRange,
                studentRange: mockStudentRange,
                title,
                headers,
                students: [],
                colsPerItem: 3,
                rowOffset: 0,
                mergeOnlyheader: false,
                formulaFunction,
            };

            const data = buildStatusSectionData(params);

            expect(data).toHaveLength(1); // Only header
            const headerRow = data[0];

            // 4 fixed cells + 3*2 headers = 10 cells
            expect(headerRow).toHaveLength(10);

            // Title + 3 padding cells = 4 fixed cells
            expect(headerRow?.[0]?.userEnteredValue?.stringValue).toBe(title);

            // First header starts at index 4
            expect(headerRow?.[4]?.userEnteredValue?.stringValue).toBe("HW 1");
            expect(headerRow?.[5]).toEqual({}); // padding 1
            expect(headerRow?.[6]).toEqual({}); // padding 2

            // Second header starts at index 7
            expect(headerRow?.[7]?.userEnteredValue?.stringValue).toBe("HW 2");
        });

        it("should output an empty array for SEPARATOR row types", () => {
            const params: StatusSectionDataParams = {
                statusRange: mockStatusRange,
                studentRange: mockStudentRange,
                title,
                headers,
                students: [mockSeparator],
                colsPerItem: 2,
                rowOffset: 0,
                mergeOnlyheader: false,
                formulaFunction,
            };

            const data = buildStatusSectionData(params);

            expect(data).toHaveLength(2); // Header + Separator
            expect(data[0]).toHaveLength(8); // 4 fixed + 2*2 sections
            expect(data[1]).toEqual([]); // Separator row
        });

        it("should generate student formulas correctly when mergeOnlyheader is FALSE", () => {
            mockCreateSetupRowValidFormula.mockReturnValue("=SETUP_MOCK()");
            mockGetA1Notation.mockReturnValue("Sheet1!$A$1");

            const params: StatusSectionDataParams = {
                statusRange: mockStatusRange,
                studentRange: mockStudentRange,
                title,
                headers,
                students: [mockStudent],
                colsPerItem: 2,
                rowOffset: 0,
                mergeOnlyheader: false,
                formulaFunction,
            };

            const data = buildStatusSectionData(params);

            const studentRow = data[1];

            // Fixed columns
            expect(studentRow?.[0]?.userEnteredValue?.formulaValue).toBe("=SETUP_MOCK()");
            expect(studentRow?.[1]?.userEnteredValue?.numberValue).toBe(101);
            expect(studentRow?.[2]?.userEnteredValue?.stringValue).toBe("John");
            expect(studentRow?.[3]?.userEnteredValue?.stringValue).toBe("Doe");

            // Header 1 data (formula + empty padding)
            expect(studentRow?.[4]?.userEnteredValue?.formulaValue).toBe("=CUSTOM_FORMULA(Sheet1!$A$1)");
            expect(studentRow?.[5]).toEqual({});

            // Header 2 data (formula + empty padding)
            expect(studentRow?.[6]?.userEnteredValue?.formulaValue).toBe("=CUSTOM_FORMULA(Sheet1!$A$1)");
            expect(studentRow?.[7]).toEqual({});
        });

        it("should populate all columns with formulas when mergeOnlyheader is TRUE", () => {
            mockCreateSetupRowValidFormula.mockReturnValue("=SETUP_MOCK()");
            mockGetA1Notation.mockReturnValue("Sheet1!$A$1");

            const params: StatusSectionDataParams = {
                statusRange: mockStatusRange,
                studentRange: mockStudentRange,
                title,
                headers,
                students: [mockStudent],
                colsPerItem: 2,
                rowOffset: 0,
                mergeOnlyheader: true,
                formulaFunction,
            };

            const data = buildStatusSectionData(params);

            const studentRow = data[1];

            // Both cells under the header should have formulas, no empty objects
            expect(studentRow?.[4]?.userEnteredValue?.formulaValue).toBe("=CUSTOM_FORMULA(Sheet1!$A$1)");
            expect(studentRow?.[5]?.userEnteredValue?.formulaValue).toBe("=CUSTOM_FORMULA(Sheet1!$A$1)");
            expect(studentRow?.[6]?.userEnteredValue?.formulaValue).toBe("=CUSTOM_FORMULA(Sheet1!$A$1)");
            expect(studentRow?.[7]?.userEnteredValue?.formulaValue).toBe("=CUSTOM_FORMULA(Sheet1!$A$1)");
        });
    });
});
