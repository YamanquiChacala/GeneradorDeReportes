import { createAttendaceFormulas, type StudentRow, type TrimesterRanges } from "../report-utils";
import { calculateCalendarHeaders, generateStudentGrid } from "./data";
import { TemplateSize } from "./types";

// 1. Setup Typed Mocks
jest.mock("../report-utils", () => ({
    createAttendaceFormulas: jest.fn(),
}));

const mockCreateAttendaceFormulas = createAttendaceFormulas as jest.MockedFunction<typeof createAttendaceFormulas>;

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
            const students: StudentRow[] = [{ type: "separator" }];

            const result = generateStudentGrid(students, 5, mockTrimesters);

            expect(result).toEqual([[]]);
            expect(mockCreateAttendaceFormulas).not.toHaveBeenCalled();
        });

        it("should generate cell data correctly for student rows", () => {
            // Natively satisfies the Student interface
            const students: StudentRow[] = [
                {
                    type: "student",
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
});
