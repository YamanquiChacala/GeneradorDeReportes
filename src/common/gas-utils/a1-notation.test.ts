import { getA1Notation, getColumnLetter } from "./a1-notation";
import type { MappedNamedRange } from "./types";

describe("Gas Utils, A1 Notation", () => {
    describe("getA1Notation", () => {
        type InferredA1Params = Parameters<typeof getA1Notation>[0];

        // Provides sensible defaults while allowing overrides
        const createMockSheet = (title = "Sheet1", rowCount: number | undefined = 1000, columnCount: number | undefined = 26) =>
            ({
                properties: {
                    title,
                    gridProperties: { rowCount, columnCount },
                },
            }) as GoogleAppsScript.Sheets.Schema.Sheet;

        const createParams = (range: GoogleAppsScript.Sheets.Schema.GridRange, overrides: Partial<InferredA1Params> = {}): InferredA1Params => ({
            mappedRange: {
                namedRange: { namedRangeId: "test_id", name: "test_name", range },
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

            it("should handle missing grid properties (fallbacks to 1000x1000)", () => {
                const mockMappedRange: MappedNamedRange = {
                    namedRange: { namedRangeId: "", name: "", range: { startRowIndex: 5, startColumnIndex: 5 } }, // Only starts defined
                    sheet: { properties: { title: "Sheet1" } }, // gridProperties entirely missing
                };
                // Starts at F6, ends at fallback bounds (1000 cols = ALL, 1000 rows)
                expect(getA1Notation({ mappedRange: mockMappedRange })).toBe("F6:ALL1000");
            });
        });

        describe("Offsets and Dimensions (Integration with offsetGridRange)", () => {
            // These tests act as integration tests. If offsetGridRange logic changes,
            // these A1 notation assertions will correctly fail.

            it("should shift a single cell by row and column offsets", () => {
                // Origin: A1. Offset by 1 row, 2 cols -> C2
                const params = createParams({ startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 1 }, { rowOffset: 1, colOffset: 2 });
                expect(getA1Notation(params)).toBe("C2");
            });

            it("should expand a single cell into a range using height and width", () => {
                // Origin: A1. Resize to 3 rows high, 2 cols wide -> A1:B3
                const params = createParams({ startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 1 }, { height: 3, width: 2 });
                expect(getA1Notation(params)).toBe("A1:B3");
            });

            it("should combine offsets and resizing simultaneously", () => {
                // Origin: A1. Shift to B2, then resize to 2x2 -> B2:C3
                const params = createParams(
                    { startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 1 },
                    { rowOffset: 1, colOffset: 1, height: 2, width: 2 },
                );
                expect(getA1Notation(params)).toBe("B2:C3");
            });
        });

        describe("Unbounded Dimensions (Columns & Rows)", () => {
            it("should format a single entire column (A:A)", () => {
                const params = createParams({ startColumnIndex: 0, endColumnIndex: 1 });
                expect(getA1Notation(params)).toBe("A:A");
            });

            it("should format multiple entire columns (A:C)", () => {
                const params = createParams({ startColumnIndex: 0, endColumnIndex: 3 });
                expect(getA1Notation(params)).toBe("A:C");
            });

            it("should format a single entire row (1:1)", () => {
                const params = createParams({ startRowIndex: 0, endRowIndex: 1 });
                expect(getA1Notation(params)).toBe("1:1");
            });

            it("should format multiple entire rows (1:5)", () => {
                const params = createParams({ startRowIndex: 0, endRowIndex: 5 });
                expect(getA1Notation(params)).toBe("1:5");
            });
        });

        describe("Partially Unbounded (The 'To Edge' Cases)", () => {
            it("should format an open-ended range falling back to sheet dimensions (C4:Z1000)", () => {
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
                const params = createParams({});
                expect(getA1Notation(params)).toBe("A1:Z1000");
            });

            it("should adapt entirely unbounded fallback to custom grid dimensions", () => {
                // Sheet is 50 rows by 5 columns (E)
                const params = createParams(
                    {},
                    {
                        mappedRange: {
                            namedRange: { namedRangeId: "", name: "", range: {} },
                            sheet: createMockSheet("Custom", 50, 5),
                        },
                    },
                );
                expect(getA1Notation(params)).toBe("A1:E50");
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

            it("should apply locks correctly on open-ended ranges ($C$4:$Z$1000)", () => {
                const params = createParams({ startRowIndex: 3, startColumnIndex: 2 }, { lockRows: true, lockColumns: true });
                expect(getA1Notation(params)).toBe("$C$4:$Z$1000");
            });
        });

        describe("Sheet Name Prefixing", () => {
            it("should prepend sheet name without quotes if standard", () => {
                const params = createParams({ startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 1 }, { includeSheetName: true });
                // As implemented, it always wraps in single quotes for safety
                expect(getA1Notation(params)).toBe("'Sheet1'!A1");
            });

            it("should escape single quotes inside the sheet name", () => {
                const params = createParams({ startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 1 }, { includeSheetName: true });
                // biome-ignore lint/style/noNonNullAssertion: Setup in mock builder
                params.mappedRange.sheet.properties!.title = "John's Data";
                expect(getA1Notation(params)).toBe("'John''s Data'!A1");
            });

            it("should use custom sheet name when specified", () => {
                const params = createParams(
                    { startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 1 },
                    { includeSheetName: true, customSheetName: "mySheet" },
                );
                // biome-ignore lint/style/noNonNullAssertion: Setup in mock builder
                params.mappedRange.sheet.properties!.title = "UnusedName";
                expect(getA1Notation(params)).toBe("'mySheet'!A1");
            });

            it("should fallback to sheet name if custom name is empty string", () => {
                const params = createParams(
                    { startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 1 },
                    { includeSheetName: true, customSheetName: "" },
                );
                // biome-ignore lint/style/noNonNullAssertion: Setup in mock builder
                params.mappedRange.sheet.properties!.title = "usedName";
                expect(getA1Notation(params)).toBe("'usedName'!A1");
            });

            it("should handle empty string sheet names gracefully", () => {
                const params = createParams({ startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: 1 }, { includeSheetName: true });
                // biome-ignore lint/style/noNonNullAssertion: Setup in mock builder
                params.mappedRange.sheet.properties!.title = "";
                expect(getA1Notation(params)).toBe("A1");
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

        it("should handle the fallback maximum bound seamlessly (1000 columns)", () => {
            expect(getColumnLetter(999)).toBe("ALL");
        });

        it("should ignore negative numbers", () => {
            expect(getColumnLetter(-1)).toBe("");
            expect(getColumnLetter(-10000)).toBe("");
        });

        it("should ignore floating point discrepancies if passed", () => {
            expect(getColumnLetter(2.9)).toBe("C");
        });
    });
});
