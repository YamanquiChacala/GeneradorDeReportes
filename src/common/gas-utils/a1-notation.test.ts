import { getA1Notation, getColumnLetter } from "./a1-notation";
import type { MappedNamedRange } from "./types";

describe("Notation", () => {
    describe("getA1Notation", () => {
        type InferredA1Params = Parameters<typeof getA1Notation>[0];

        const createMockSheet = (title = "Sheet1", rowCount: number | undefined = 1000, columnCount: number | undefined = 26) =>
            ({
                properties: {
                    title,
                    gridProperties: { rowCount, columnCount },
                },
            }) as GoogleAppsScript.Sheets.Schema.Sheet;

        const createParams = (range: GoogleAppsScript.Sheets.Schema.GridRange, overrides: Partial<InferredA1Params> = {}): InferredA1Params => ({
            mappedRange: {
                namedRange: { namedRangeId: "", name: "", range },
                sheet: createMockSheet(),
            },
            ...overrides,
        });

        describe("Standard Bounded Ranges", () => {
            it("should format a single cell (A1)", () => {
                const params = createParams({ startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 1 });
                expect(getA1Notation(params)).toBe("A1");
            });

            it("should format a standard multi-cell range (A1:B2)", () => {
                const params = createParams({ startRowIndex: 0, endRowIndex: 2, startColumnIndex: 0, endColumnIndex: 2 });
                expect(getA1Notation(params)).toBe("A1:B2");
            });

            it("should handle missing sheet parameters", () => {
                const mockMappedRange: MappedNamedRange = {
                    namedRange: { namedRangeId: "", name: "", range: { endRowIndex: 1, endColumnIndex: 1 } },
                    sheet: { properties: { gridProperties: {} } },
                };
                expect(getA1Notation({ mappedRange: mockMappedRange })).toBe("A1");
            });
        });

        describe("Unbounded Dimensions (Columns & Rows)", () => {
            it("should format a single entire column (A:A)", () => {
                const params = createParams({ startColumnIndex: 0, endColumnIndex: 1 }); // Rows undefined
                expect(getA1Notation(params)).toBe("A:A");
            });

            it("should format multiple entire columns (A:C)", () => {
                const params = createParams({ startColumnIndex: 0, endColumnIndex: 3 }); // Rows undefined
                expect(getA1Notation(params)).toBe("A:C");
            });

            it("should format a single entire row (1:1)", () => {
                const params = createParams({ startRowIndex: 0, endRowIndex: 1 }); // Columns undefined
                expect(getA1Notation(params)).toBe("1:1");
            });

            it("should format multiple entire rows (1:5)", () => {
                const params = createParams({ startRowIndex: 0, endRowIndex: 5 }); // Columns undefined
                expect(getA1Notation(params)).toBe("1:5");
            });
        });

        describe("Partially Unbounded (The 'To Edge' Cases)", () => {
            it("should format an open-ended range falling back to sheet dimensions (C4:Z1000)", () => {
                // Start row 3 (Row 4), Start col 2 (Col C). Ends are undefined.
                const params = createParams({ startRowIndex: 3, startColumnIndex: 2 });
                expect(getA1Notation(params)).toBe("C4:Z1000");
            });

            it("should format an open-ended column bounded by rows (C4:Z10)", () => {
                const params = createParams({ startRowIndex: 3, endRowIndex: 10, startColumnIndex: 2 });
                expect(getA1Notation(params)).toBe("C4:Z10");
            });

            it("should format an open-ended row bounded by columns (C4:E1000)", () => {
                const params = createParams({ startRowIndex: 3, startColumnIndex: 2, endColumnIndex: 5 });
                expect(getA1Notation(params)).toBe("C4:E1000");
            });
        });

        describe("Fully Unbounded (Entire Sheet)", () => {
            it("should format a completely empty GridRange to cover the whole sheet bounds", () => {
                const params = createParams({}); // Everything undefined
                expect(getA1Notation(params)).toBe("A1:Z1000");
            });
        });

        describe("Locking References (Absolute vs Relative)", () => {
            it("should lock rows and columns when both are true ($A$1:$B$2)", () => {
                const params = createParams({ startRowIndex: 0, endRowIndex: 2, startColumnIndex: 0, endColumnIndex: 2 }, { lockRows: true, lockColumns: true });
                expect(getA1Notation(params)).toBe("$A$1:$B$2");
            });

            it("should lock only rows (A$1:B$2)", () => {
                const params = createParams({ startRowIndex: 0, endRowIndex: 2, startColumnIndex: 0, endColumnIndex: 2 }, { lockRows: true });
                expect(getA1Notation(params)).toBe("A$1:B$2");
            });

            it("should lock only columns ($A1:$B2)", () => {
                const params = createParams({ startRowIndex: 0, endRowIndex: 2, startColumnIndex: 0, endColumnIndex: 2 }, { lockColumns: true });
                expect(getA1Notation(params)).toBe("$A1:$B2");
            });

            it("should lock completely unbounded rows/columns correctly ($A:$C)", () => {
                const params = createParams({ startColumnIndex: 0, endColumnIndex: 3 }, { lockColumns: true });
                expect(getA1Notation(params)).toBe("$A:$C");
            });
        });

        describe("Sheet Name Prefixing", () => {
            it("should prepend sheet name without quotes if standard", () => {
                const params = createParams({ startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 1 }, { includeSheetName: true });
                // Even standard names get single quotes in this implementation for absolute safety
                expect(getA1Notation(params)).toBe("'Sheet1'!A1");
            });

            it("should escape single quotes inside the sheet name", () => {
                const params = createParams({ startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 1 }, { includeSheetName: true });
                // biome-ignore lint/style/noNonNullAssertion: We create the mock, we know it exist.
                params.mappedRange.sheet.properties!.title = "John's Data";

                // "John's Data" -> 'John''s Data'!A1
                expect(getA1Notation(params)).toBe("'John''s Data'!A1");
            });
        });
    });

    describe("getColumnLetter", () => {
        it("should correctly map a single letter", () => {
            expect(getColumnLetter(1 - 1)).toBe("A");
            expect(getColumnLetter(13 - 1)).toBe("M");
            expect(getColumnLetter(26 - 1)).toBe("Z");
        });

        it("should correctly map multiple letters", () => {
            expect(getColumnLetter(1 * 26 + 1 - 1)).toBe("AA");
            expect(getColumnLetter(2 * 26 + 6 - 1)).toBe("BF");
            expect(getColumnLetter(26 * 26 + 26 - 1)).toBe("ZZ");
            expect(getColumnLetter(4 * 26 * 26 + 26 + 26 - 1)).toBe("DAZ");
        });

        it("should ignore negative numbers", () => {
            expect(getColumnLetter(-1)).toBe("");
            expect(getColumnLetter(-10000)).toBe("");
        });
    });
});
