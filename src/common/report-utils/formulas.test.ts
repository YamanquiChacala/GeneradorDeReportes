import { getA1Notation, getColumnLetter, type MappedNamedRange } from "../gas-utils";
import type { AcademicField } from ".";
import {
    createFieldFormula,
    createFinalSubjectAverageFormula,
    createStudentAsistanceFormula,
    createStudentAsistancePerSubjectFormula,
    createSubjectAverageFormula,
    getShortCommentFormula,
} from ".";

// Mock the dependencies
jest.mock("../gas-utils", () => ({
    getA1Notation: jest.fn(),
    getColumnLetter: jest.fn(),
}));

describe("Formula Generators", () => {
    // Reusable dummy data
    const mockRange: MappedNamedRange = {
        range: { startRowIndex: 2, endRowIndex: 5, startColumnIndex: 1, endColumnIndex: 4 },
        sheet: { properties: { title: "TestSheet" } },
    };

    const mockRangeOpen: MappedNamedRange = {
        range: {}, // Testing the `?? 0` fallbacks
        sheet: {},
    };

    beforeEach(() => {
        jest.clearAllMocks();
    });

    describe("getShortCommentFormula", () => {
        it("should calculate correct offsets and inject them into the LET formula", () => {
            (getColumnLetter as jest.Mock).mockReturnValue("D");

            // startColumnIndex (1) + colOffset (2) = 3 (which maps to "D" in our mock)
            // startRowIndex (2) + rowOffset (5) + 1 = 8
            const formula = getShortCommentFormula(mockRange, 5, 2);

            expect(getColumnLetter).toHaveBeenCalledWith(3);
            expect(formula).toContain("raw_text, D8");
            expect(formula).toContain("cleaned_accents"); // Verify formula body presence
            expect(formula.startsWith("=LET(")).toBe(true);
        });

        it("should safely handle ranges missing index properties by falling back to 0", () => {
            (getColumnLetter as jest.Mock).mockReturnValue("A");

            const formula = getShortCommentFormula(mockRangeOpen, 0, 0);

            expect(getColumnLetter).toHaveBeenCalledWith(0);
            expect(formula).toContain("raw_text, A1"); // 0 + 0 + 1 = 1
        });
    });

    describe("createStudentAsistanceFormula", () => {
        it("should generate attendance formula and calculate correct column offsets for Period 0", () => {
            (getA1Notation as jest.Mock)
                .mockReturnValueOnce("$A$1") // firstName
                .mockReturnValueOnce("$B$1"); // lastName
            (getColumnLetter as jest.Mock).mockReturnValue("C"); // 3 + 0*2 = 3

            const formula = createStudentAsistanceFormula(0, mockRange, mockRange, "Attendance", false);

            expect(getColumnLetter).toHaveBeenCalledWith(3);
            expect(formula).toContain("return_data, 'Attendance'!C:C");
            expect(formula).toContain("XLOOKUP($A$1&$B$1");
        });

        it("should shift the return column by 1 if absences is true", () => {
            (getA1Notation as jest.Mock).mockReturnValue("$A$1");
            (getColumnLetter as jest.Mock).mockReturnValue("D"); // 4 + 0*2 = 4

            const formula = createStudentAsistanceFormula(0, mockRange, mockRange, "Attendance", true);

            expect(getColumnLetter).toHaveBeenCalledWith(4);
            expect(formula).toContain("return_data, 'Attendance'!D:D");
        });

        it("should shift the return column by 2 for each period", () => {
            (getA1Notation as jest.Mock).mockReturnValue("$A$1");
            (getColumnLetter as jest.Mock).mockReturnValue("G"); // 3 + 2*2 = 7

            createStudentAsistanceFormula(2, mockRange, mockRange, "Attendance", false);

            expect(getColumnLetter).toHaveBeenCalledWith(7);
        });

        it("should default absences to false if omitted, calculating the correct column offset", () => {
            (getA1Notation as jest.Mock).mockReturnValue("$A$1");
            (getColumnLetter as jest.Mock).mockReturnValue("C"); // Default false: 3 + 0*2 = 3

            // Notice: completely omitting the 5th `absences` argument here
            const formula = createStudentAsistanceFormula(0, mockRange, mockRange, "Attendance");

            // It should calculate as if absences was false
            expect(getColumnLetter).toHaveBeenCalledWith(3);
            expect(formula).toContain("return_data, 'Attendance'!C:C");
        });
    });

    describe("createSubjectAverageFormula", () => {
        it("should format the weighted average formula and pass correct constraints to A1 utility", () => {
            (getA1Notation as jest.Mock)
                .mockReturnValueOnce("B2:D2") // values
                .mockReturnValueOnce("'Weights'!$B$1:$D$1"); // weights

            const formula = createSubjectAverageFormula(mockRange, 2, mockRange);

            // Verify configuration for values
            expect(getA1Notation).toHaveBeenNthCalledWith(1, {
                mappedRange: mockRange,
                rowOffset: 2,
                colOffset: 1,
                height: 1,
                width: 3,
                lockColumns: true,
            });

            // Verify configuration for weights
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

            // mockRange.range.endColumnIndex is 4. colOffset should be 4 - 2 = 2.
            expect(getA1Notation).toHaveBeenCalledWith(expect.objectContaining({ colOffset: 2 }));

            expect(formula).toContain("COUNT(E2, F2, G2)=3");
            expect(formula).toContain("ROUND(AVERAGE(E2, F2, G2),1)");
        });

        it("should fallback to 0 if trim3 is missing endColumnIndex", () => {
            createFinalSubjectAverageFormula(mockRange, mockRange, mockRangeOpen, 1);
            // 0 - 2 = -2
            expect(getA1Notation).toHaveBeenCalledWith(expect.objectContaining({ colOffset: -2 }));
        });
    });

    describe("createStudentAsistancePerSubjectFormula", () => {
        it("should calculate correct return column based on period and build complex LET formula", () => {
            (getA1Notation as jest.Mock)
                .mockReturnValueOnce("$A$5") // subjectA1
                .mockReturnValueOnce("$A$1") // firstNameA1
                .mockReturnValueOnce("$B$1"); // lastNameA1
            (getColumnLetter as jest.Mock).mockReturnValue("E"); // 3 + 1*2 = 5

            const formula = createStudentAsistancePerSubjectFormula(1, mockRange, 2, mockRange, mockRange, "AttData");

            expect(getColumnLetter).toHaveBeenCalledWith(5);
            expect(formula).toContain("MATCH($A$5, AttData!A:A, 0)");
            expect(formula).toContain("OFFSET(AttData!E$1"); // Verifies return_data offset
            expect(formula).toContain("XLOOKUP($A$1 & $B$1");
        });
    });

    describe("createFieldFormula", () => {
        it("should request correct dimensions based on field.subjects and build the formula", () => {
            const mockField: AcademicField = { name: "Science", color: "#FFF", subjects: 4 };

            (getA1Notation as jest.Mock).mockReturnValueOnce("C2:C5").mockReturnValueOnce("'Weights'!$D$2:$D$5");

            const formula = createFieldFormula(mockField, 2, 1, mockRange, mockRange);

            // Check that it requested exactly `field.subjects` height for values
            expect(getA1Notation).toHaveBeenNthCalledWith(
                1,
                expect.objectContaining({
                    colOffset: 2,
                    rowOffset: 1,
                    height: 4,
                    lockRows: true,
                }),
            );

            // Check weights configuration
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
});
