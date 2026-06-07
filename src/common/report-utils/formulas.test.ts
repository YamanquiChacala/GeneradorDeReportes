import { getA1Notation, getColumnLetter, type MappedNamedRange } from "../gas-utils";
import {
    createAllSubjectsAverageFormula,
    createAttendaceFormulas,
    createFieldAverageFormula,
    createFieldFormula,
    createFinalSubjectAverageFormula,
    createIndividualSubjectAverageFormula,
    createStudentGeneralAttendanceFormula,
    createStudentPerSubjectAttendanceFormula,
    getShortCommentFormula,
} from "./formulas";
import { type AcademicField, Period } from "./types";

// Mock the dependencies
jest.mock("../gas-utils", () => ({
    getA1Notation: jest.fn(),
    getColumnLetter: jest.fn(),
}));

describe("Report Utils. Formula Generators", () => {
    // Reusable dummy data
    const mockRange: MappedNamedRange = {
        namedRange: { namedRangeId: "", name: "", range: { startRowIndex: 2, endRowIndex: 5, startColumnIndex: 1, endColumnIndex: 4 } },
        sheet: { properties: { title: "TestSheet" } },
    };

    const mockRangeOpen: MappedNamedRange = {
        namedRange: { namedRangeId: "", name: "", range: {} }, // Testing the `?? 0` fallbacks
        sheet: {},
    };

    beforeEach(() => {
        jest.clearAllMocks();
    });

    describe("createAttendaceFormulas", () => {
        it("should return empty strings if startCol is -1", () => {
            const row = 0;
            const startCol = -1;
            const endCol = 5;

            const result = createAttendaceFormulas(row, startCol, endCol);

            expect(result).toEqual({ percent: '=""', count: '=""' });
            expect(getColumnLetter).not.toHaveBeenCalled();
        });

        it("should return empty strings if startCol is greater than endCol", () => {
            const row = 0;
            const startCol = 5;
            const endCol = 2;

            const result = createAttendaceFormulas(row, startCol, endCol);

            expect(result).toEqual({ percent: '=""', count: '=""' });
            expect(getColumnLetter).not.toHaveBeenCalled();
        });

        it("should calculate the correct formulas with A1 notation when valid inputs are provided", () => {
            const row = 4; // Function does row + 1, resulting in row 5
            const startCol = 1;
            const endCol = 5;

            (getColumnLetter as jest.Mock).mockReturnValueOnce("A").mockReturnValueOnce("E");

            const result = createAttendaceFormulas(row, startCol, endCol);

            expect(getColumnLetter).toHaveBeenCalledTimes(2);
            expect(getColumnLetter).toHaveBeenNthCalledWith(1, startCol);
            expect(getColumnLetter).toHaveBeenNthCalledWith(2, endCol);

            expect(result).toEqual({
                percent: '=IF(COUNTA($A5:$E5) >= COLUMNS($A5:$E5) / 10, (COUNTA($A5:$E5) - SUM($A5:$E5)) / COUNTA($A5:$E5), "")',
                count: '=IF(COUNTA($A5:$E5) >= COLUMNS($A5:$E5) / 10, SUM($A5:$E5), "")',
            });
        });
    });

    describe("getShortCommentFormula", () => {
        it("should calculate correct offsets and inject them into the LET formula", () => {
            (getColumnLetter as jest.Mock).mockReturnValue("D");

            const formula = getShortCommentFormula(mockRange, 5, 2);

            expect(getColumnLetter).toHaveBeenCalledWith(3);
            expect(formula).toContain("raw_text, D8");
            expect(formula).toContain("cleaned_accents");
            expect(formula.startsWith("=LET(")).toBe(true);
        });

        it("should safely handle ranges missing index properties by falling back to 0", () => {
            (getColumnLetter as jest.Mock).mockReturnValue("A");

            const formula = getShortCommentFormula(mockRangeOpen, 0, 0);

            expect(getColumnLetter).toHaveBeenCalledWith(0);
            expect(formula).toContain("raw_text, A1");
        });
    });

    describe("createStudentGeneralAttendanceFormula", () => {
        it("should generate attendance formula and calculate correct column offsets for Period FIRST with false absences multiplier", () => {
            (getA1Notation as jest.Mock).mockReturnValueOnce("$A$1").mockReturnValueOnce("$B$1");
            (getColumnLetter as jest.Mock).mockReturnValue("C");

            const formula = createStudentGeneralAttendanceFormula(Period.FIRST, mockRange, mockRange, "Attendance", false);

            expect(getColumnLetter).toHaveBeenCalledWith(3);
            expect(formula).toContain("return_data, 'Attendance'!C:C");
            expect(formula).toContain("XLOOKUP($A$1&$B$1");
            expect(formula).toContain("ROUND(raw_result * { 10 }");
        });

        it("should shift the return column by 1 and change multiplier if absences is true", () => {
            (getA1Notation as jest.Mock).mockReturnValue("$A$1");
            (getColumnLetter as jest.Mock).mockReturnValue("D");

            const formula = createStudentGeneralAttendanceFormula(Period.FIRST, mockRange, mockRange, "Attendance", true);

            expect(getColumnLetter).toHaveBeenCalledWith(4);
            expect(formula).toContain("return_data, 'Attendance'!D:D");
            expect(formula).toContain("ROUND(raw_result * { 1 }");
        });

        it("should shift the return column by 2 for each period", () => {
            (getA1Notation as jest.Mock).mockReturnValue("$A$1");
            (getColumnLetter as jest.Mock).mockReturnValue("G");

            createStudentGeneralAttendanceFormula(Period.THIRD, mockRange, mockRange, "Attendance", false);

            expect(getColumnLetter).toHaveBeenCalledWith(7);
        });

        it("should default absences to false if omitted, calculating the correct column offset and multiplier", () => {
            (getA1Notation as jest.Mock).mockReturnValue("$A$1");
            (getColumnLetter as jest.Mock).mockReturnValue("C");

            const formula = createStudentGeneralAttendanceFormula(Period.FIRST, mockRange, mockRange, "Attendance");

            expect(getColumnLetter).toHaveBeenCalledWith(3);
            expect(formula).toContain("return_data, 'Attendance'!C:C");
            expect(formula).toContain("ROUND(raw_result * { 10 }");
        });
    });

    describe("createIndividualSubjectAverageFormula", () => {
        it("should format the weighted average formula and pass correct constraints to A1 utility", () => {
            (getA1Notation as jest.Mock).mockReturnValueOnce("B2:D2").mockReturnValueOnce("'Weights'!$B$1:$D$1");

            const formula = createIndividualSubjectAverageFormula(mockRange, 2, mockRange);

            expect(getA1Notation).toHaveBeenNthCalledWith(1, {
                mappedRange: mockRange,
                rowOffset: 2,
                colOffset: 1,
                height: 1,
                width: 3,
                lockColumns: true,
            });

            expect(getA1Notation).toHaveBeenNthCalledWith(2, {
                mappedRange: mockRange,
                includeSheetName: true,
                lockRows: true,
                lockColumns: true,
            });

            expect(formula).toContain("AVERAGE.WEIGHTED(");
            expect(formula).toContain("B2:D2,\n            TRANSPOSE('Weights'!$B$1:$D$1)");
        });
    });

    describe("createFinalSubjectAverageFormula", () => {
        it("should extract endColumnIndex for colOffset and construct the COUNT/AVERAGE formula", () => {
            (getA1Notation as jest.Mock).mockReturnValueOnce("E2").mockReturnValueOnce("F2").mockReturnValueOnce("G2");

            const formula = createFinalSubjectAverageFormula(mockRange, mockRange, mockRange, 1);

            expect(getA1Notation).toHaveBeenCalledWith(expect.objectContaining({ colOffset: 2 }));
            expect(formula).toContain("COUNT(E2, F2, G2)=3");
            expect(formula).toContain("ROUND(AVERAGE(E2, F2, G2),1)");
        });

        it("should fallback to 0 if trim3 is missing endColumnIndex", () => {
            createFinalSubjectAverageFormula(mockRange, mockRange, mockRangeOpen, 1);
            expect(getA1Notation).toHaveBeenCalledWith(expect.objectContaining({ colOffset: -2 }));
        });
    });

    describe("createStudentPerSubjectAttendanceFormula", () => {
        it("should calculate correct return column based on period and build complex LET formula", () => {
            (getA1Notation as jest.Mock).mockReturnValueOnce("$A$5").mockReturnValueOnce("$A$1").mockReturnValueOnce("$B$1");
            (getColumnLetter as jest.Mock).mockReturnValue("E");

            const formula = createStudentPerSubjectAttendanceFormula(Period.SECOND, mockRange, 2, mockRange, mockRange, "AttData");

            expect(getColumnLetter).toHaveBeenCalledWith(5);
            expect(formula).toContain("MATCH($A$5, AttData!A:A, 0)");
            expect(formula).toContain("OFFSET(AttData!E$1");
            expect(formula).toContain("XLOOKUP($A$1 & $B$1");
        });
    });

    describe("createFieldFormula", () => {
        it("should request correct dimensions based on field.subjects and build the formula", () => {
            const mockField: AcademicField = { name: "Science", color: "#FFF", subjects: 4 };

            (getA1Notation as jest.Mock).mockReturnValueOnce("C2:C5").mockReturnValueOnce("'Weights'!$D$2:$D$5");

            const formula = createFieldFormula(mockField, 2, 1, mockRange, mockRange);

            expect(getA1Notation).toHaveBeenNthCalledWith(
                1,
                expect.objectContaining({
                    colOffset: 2,
                    rowOffset: 1,
                    height: 4,
                    lockRows: true,
                }),
            );

            expect(getA1Notation).toHaveBeenNthCalledWith(
                2,
                expect.objectContaining({
                    includeSheetName: true,
                    colOffset: 1,
                    rowOffset: 1,
                    height: 4,
                    lockRows: true,
                    lockColumns: true,
                }),
            );

            expect(formula).toContain("AVERAGE.WEIGHTED(C2:C5, 'Weights'!$D$2:$D$5)");
        });
    });

    describe("createFieldAverageFormula", () => {
        it("should request correct constraints from A1 utility and build the rounded average formula", () => {
            (getA1Notation as jest.Mock).mockReturnValueOnce("C$2:C$5");

            const formula = createFieldAverageFormula(mockRange, 3, 2);

            expect(getA1Notation).toHaveBeenCalledWith({
                mappedRange: mockRange,
                colOffset: 3,
                width: 1,
                lockRows: true,
            });

            expect(formula).toBe('=IF(COUNT(C$2:C$5) <> ROWS(C$2:C$5),"", ROUND(AVERAGE(C$2:C$5), 2))');
        });
    });

    describe("createAllSubjectsAverageFormula", () => {
        it("should request values and weights from A1 utility and build the weighted average formula", () => {
            (getA1Notation as jest.Mock).mockReturnValueOnce("E$2:E$10").mockReturnValueOnce("'Weights'!F$2:F$10");

            const formula = createAllSubjectsAverageFormula(mockRange, mockRange, 4, 1);

            expect(getA1Notation).toHaveBeenNthCalledWith(1, {
                mappedRange: mockRange,
                colOffset: 4,
                width: 1,
                lockRows: true,
            });

            expect(getA1Notation).toHaveBeenNthCalledWith(2, {
                mappedRange: mockRange,
                colOffset: 1,
                width: 1,
                includeSheetName: true,
                lockRows: true,
                lockColumns: true,
            });

            expect(formula).toBe("=IFERROR(ROUND(AVERAGE.WEIGHTED(E$2:E$10, 'Weights'!F$2:F$10), 1))");
        });
    });
});
