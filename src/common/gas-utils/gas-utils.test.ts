import { PasteType } from "../constants";
import { getEpochDate, getSheetsDate } from ".";
import { getA1Notation, getColumnLetter } from "./notation";
import { changeGridRangeSheet, createSingleCellRange, offsetGridRange } from "./range";
import { buildCopyPasteRequest } from "./request";

describe("googleAPI Utilities", () => {
    describe("Notation", () => {
        describe("getA1Notation", () => {
            type InferredA1Params = Parameters<typeof getA1Notation>[0];

            const createMockSheet = (title = "Sheet1", rowCount = 1000, columnCount = 26) =>
                ({
                    properties: {
                        title,
                        gridProperties: { rowCount, columnCount },
                    },
                }) as GoogleAppsScript.Sheets.Schema.Sheet;

            const createParams = (range: GoogleAppsScript.Sheets.Schema.GridRange, overrides: Partial<InferredA1Params> = {}): InferredA1Params => ({
                mappedRange: {
                    range,
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

    describe("Range", () => {
        describe("createSingleCellRange", () => {
            it("should create a 1x1 grid range correctly", () => {
                const range = createSingleCellRange(10, 5, 2);
                expect(range).toEqual({
                    sheetId: 10,
                    startRowIndex: 5,
                    startColumnIndex: 2,
                    endRowIndex: 6,
                    endColumnIndex: 3,
                });
            });
        });

        describe("changeGridRangeSheet", () => {
            it("should return a new grid range with an updated sheetId", () => {
                const original: GoogleAppsScript.Sheets.Schema.GridRange = {
                    sheetId: 1,
                    startRowIndex: 0,
                    endRowIndex: 10,
                };
                const updated = changeGridRangeSheet(original, 99);

                expect(updated.sheetId).toBe(99);
                expect(updated.startRowIndex).toBe(0);
                expect(updated.endRowIndex).toBe(10);

                // Should not mutate original
                expect(original.sheetId).toBe(1);
            });
        });

        describe("offsetGridRange", () => {
            const origin: GoogleAppsScript.Sheets.Schema.GridRange = {
                sheetId: 1,
                startRowIndex: 2,
                endRowIndex: 4,
                startColumnIndex: 2,
                endColumnIndex: 4,
            };

            it("should offset the range by rows and cols and maintain size if dimensions aren't overridden", () => {
                const result = offsetGridRange({ origin, rowOffset: 2, colOffset: 1 });
                expect(result).toEqual({
                    sheetId: 1,
                    startRowIndex: 4,
                    endRowIndex: 6,
                    startColumnIndex: 3,
                    endColumnIndex: 5,
                });
            });

            it("should bound the range to a new size if positive height and width are provided", () => {
                const result = offsetGridRange({ origin, rowOffset: 0, colOffset: 0, height: 10, width: 5 });
                expect(result).toEqual({
                    sheetId: 1,
                    startRowIndex: 2,
                    endRowIndex: 12,
                    startColumnIndex: 2,
                    endColumnIndex: 7,
                });
            });

            it("should safely default missing origin indexes to 0", () => {
                const emptyOrigin: GoogleAppsScript.Sheets.Schema.GridRange = { sheetId: 1 };
                const result = offsetGridRange({ origin: emptyOrigin, rowOffset: 1, colOffset: 1, height: 1, width: 1 });
                expect(result).toEqual({
                    sheetId: 1,
                    startRowIndex: 1,
                    endRowIndex: 2,
                    startColumnIndex: 1,
                    endColumnIndex: 2,
                });
            });

            it("should prevent startRowIndex and startColumnIndex from becoming negative", () => {
                const result = offsetGridRange({ origin, rowOffset: -5, colOffset: -5 });
                expect(result).toEqual({
                    sheetId: 1,
                    endRowIndex: 0,
                    endColumnIndex: 0,
                });
            });

            it("should un-bound the range if height and width are negative", () => {
                const result = offsetGridRange({ origin, rowOffset: 1, colOffset: 1, height: -1, width: -1 });
                expect(result).toEqual({
                    sheetId: 1,
                    startRowIndex: 3,
                    startColumnIndex: 3,
                });
                expect(result.endRowIndex).toBeUndefined();
                expect(result.endColumnIndex).toBeUndefined();
            });

            it("should handle un-bounding only one dimension (e.g. height) while maintaining original width", () => {
                const result = offsetGridRange({ origin, rowOffset: 1, colOffset: 1, height: -1 });
                expect(result).toEqual({
                    sheetId: 1,
                    startRowIndex: 3,
                    startColumnIndex: 3,
                    endColumnIndex: 5,
                });
                expect(result.endRowIndex).toBeUndefined();
            });
        });
    });

    describe("Request", () => {
        describe("buildCopyPasteRequest", () => {
            it("should build a valid copyPaste request", () => {
                const source = { sheetId: 1, startRowIndex: 0 };
                const destination = { sheetId: 2, startRowIndex: 5 };
                const request = buildCopyPasteRequest(source, destination, PasteType.PASTE_FORMAT);

                expect(request?.copyPaste).toBeDefined();
                expect(request?.copyPaste?.source).toBe(source);
                expect(request?.copyPaste?.destination).toBe(destination);
                expect(request?.copyPaste?.pasteType).toBe(PasteType.PASTE_FORMAT);
            });
        });
    });

    describe("Time", () => {
        describe("getSheetsDate", () => {
            it("converts Unix epoch 0 (Jan 1, 1970) to the Sheets epoch offset", () => {
                expect(getSheetsDate(0)).toBe(25569);
            });

            it("converts exactly one day after Unix epoch to the correct Sheets date", () => {
                const oneDayInMs = 86400000;
                expect(getSheetsDate(oneDayInMs)).toBe(25570);
            });

            it("converts a known modern Unix epoch to the correct Sheets date", () => {
                // Jan 1, 2024 = 1704067200000 ms
                expect(getSheetsDate(1704067200000)).toBe(45292);
            });

            it("handles negative Unix epochs (dates before 1970)", () => {
                // Dec 30, 1899 = -2209161600000 ms (Sheets Epoch 0)
                expect(getSheetsDate(-2209161600000)).toBe(0);
            });
        });

        describe("getEpochDate", () => {
            it("converts the Sheets epoch offset back to Unix epoch 0", () => {
                expect(getEpochDate(25569)).toBe(0);
            });

            it("converts exactly one day after the Sheets epoch offset to one day in ms", () => {
                const oneDayInMs = 86400000;
                expect(getEpochDate(25570)).toBe(oneDayInMs);
            });

            it("converts a known modern Sheets date to the correct Unix epoch", () => {
                // Jan 1, 2024 = 45292 in Sheets
                expect(getEpochDate(45292)).toBe(1704067200000);
            });

            it("handles Sheets date 0 (Dec 30, 1899)", () => {
                expect(getEpochDate(0)).toBe(-2209161600000);
            });
        });

        describe("Bidirectional Conversion", () => {
            it("returns the exact original Unix epoch after converting to Sheets and back", () => {
                const originalUnixEpoch = 1718323200000; // Arbitrary timestamp
                const sheetsDate = getSheetsDate(originalUnixEpoch);
                const convertedBack = getEpochDate(sheetsDate);

                expect(convertedBack).toBe(originalUnixEpoch);
            });

            it("returns the exact original Sheets date after converting to Unix and back", () => {
                const originalSheetsDate = 48000; // Arbitrary future Sheets date
                const unixEpoch = getEpochDate(originalSheetsDate);
                const convertedBack = getSheetsDate(unixEpoch);

                expect(convertedBack).toBe(originalSheetsDate);
            });
        });
    });
});
