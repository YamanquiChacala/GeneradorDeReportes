import { PasteType } from "../constants";
import {
    getCellData,
    getCellDataArray,
    getCellEffectiveValue,
    getCellNumber,
    getCellText,
    getCellUnixEpoch,
    getEpochDate,
    getSheetsDate,
    type MappedNamedRange,
} from ".";
import { colorToHex } from "./color";
import { createRequiredGetter, makeUserEntered } from "./helper";
import { getA1Notation, getColumnLetter } from "./notation";
import { parseSpreadsheet } from "./parse";
import { changeGridRangeSheet, createSingleCellRange, offsetGridRange } from "./range";
import { buildCopyPasteRequest } from "./request";

describe("googleAPI Utilities", () => {
    describe("Colors", () => {
        describe("colorToHex", () => {
            it("should convert an RGB color to hex correctly", () => {
                const color: GoogleAppsScript.Sheets.Schema.Color = {
                    red: 1, // 255
                    green: 0.50196, // ~128
                    blue: 0, // 0
                };
                expect(colorToHex(color)).toBe("#FF8000");
            });

            it("should treat missing channels as 0", () => {
                const color: GoogleAppsScript.Sheets.Schema.Color = { red: 1 };
                expect(colorToHex(color)).toBe("#FF0000");
            });

            it("should fallback to the provided fallback if color is undefined", () => {
                expect(colorToHex(undefined, "#000000")).toBe("#000000");
                expect(colorToHex(undefined)).toBe("#FFFFFF"); // Default fallback
            });
        });
    });

    describe("MappedNamedRange", () => {
        // Simulates a 5x5 Named Range (A1:E5) with sparse data chunks, but the sheet is 10x10.
        const mockMappedRange: MappedNamedRange = {
            range: { sheetId: 1, startRowIndex: 0, endRowIndex: 5, startColumnIndex: 0, endColumnIndex: 5 },
            sheet: {
                properties: {
                    gridProperties: { rowCount: 10, columnCount: 10 },
                },
                data: [
                    {
                        // Chunk 1: Implicitly starts at 0, 0
                        rowData: [
                            { values: [{ effectiveValue: { stringValue: "Test String" } }] }, // Row 0, Col 0
                            { values: [{ effectiveValue: { numberValue: 0 } }] }, // Row 1, Col 0 (For epoch test)
                        ],
                    },
                    {
                        // Chunk 2: Offset block starting at Row 3, Col 2
                        startRow: 3,
                        startColumn: 2,
                        rowData: [
                            {
                                values: [
                                    { effectiveValue: { stringValue: "Offset Data 1" } }, // Row 3, Col 2
                                    { effectiveValue: { stringValue: "Offset Data 2" } }, // Row 3, Col 3
                                ],
                            },
                        ],
                    },
                    {
                        // Chunk 3: Outside original row bounds (Row 8, Col 2)
                        startRow: 8,
                        startColumn: 2,
                        rowData: [{ values: [{ effectiveValue: { stringValue: "Unbound Row Data" } }] }],
                    },
                    {
                        // Chunk 4: Outside original column bounds (Row 2, Col 8)
                        startRow: 2,
                        startColumn: 8,
                        rowData: [{ values: [{ effectiveValue: { stringValue: "Unbound Col Data" } }] }],
                    },
                ],
            },
        };

        // Mock an unbounded range (e.g., A:Z) where endRowIndex and endColumnIndex are inherently undefined.
        const mockUnboundedRange: MappedNamedRange = {
            range: { sheetId: 2, startRowIndex: 0, startColumnIndex: 0 }, // No end indexes
            sheet: {
                properties: {
                    gridProperties: { rowCount: 100, columnCount: 26 },
                },
                data: [
                    {
                        startRow: 99,
                        startColumn: 25,
                        rowData: [{ values: [{ effectiveValue: { stringValue: "Bottom Right Cell" } }] }],
                    },
                ],
            },
        };

        describe("getCellData & Utilities", () => {
            it("getCellData should return data if within bounds of the first data chunk", () => {
                const data = getCellData({ mappedRange: mockMappedRange, rowOffset: 0, columnOffset: 0 });
                expect(data?.effectiveValue?.stringValue).toBe("Test String");
            });

            it("getCellData should correctly locate data in an offset data chunk", () => {
                const data = getCellData({ mappedRange: mockMappedRange, rowOffset: 3, columnOffset: 2 });
                expect(data?.effectiveValue?.stringValue).toBe("Offset Data 1");
            });

            it("getCellData should return an empty object for empty cells within bounds", () => {
                const data = getCellData({ mappedRange: mockMappedRange, rowOffset: 0, columnOffset: 4 });
                expect(data).toStrictEqual({});
            });

            it("getCellData should return an empty object if out of bounds", () => {
                const data = getCellData({ mappedRange: mockMappedRange, rowOffset: 10, columnOffset: 0 });
                expect(data).toStrictEqual({});
            });

            it("getCellUnixEpoch should correctly calculate the 1899 Sheets epoch", () => {
                // Sheets epoch (0) is Dec 30, 1899
                const expectedEpoch = new Date(Date.UTC(1899, 11, 30)).getTime();
                const epoch = getCellUnixEpoch({ mappedRange: mockMappedRange, rowOffset: 1, columnOffset: 0 });
                expect(epoch).toBe(expectedEpoch);
            });

            it("getCellData should handle unbounded ranges without throwing out-of-bounds errors", () => {
                // Because endRowIndex/endColumnIndex are undefined, the boundary check should naturally pass
                const data = getCellData({ mappedRange: mockUnboundedRange, rowOffset: 99, columnOffset: 25 });
                expect(data?.effectiveValue?.stringValue).toBe("Bottom Right Cell");
            });

            it("getCellEffectiveValue should return the effectiveValue object of a cell", () => {
                const value = getCellEffectiveValue({ mappedRange: mockMappedRange, rowOffset: 0, columnOffset: 0 });
                expect(value).toStrictEqual({ stringValue: "Test String" });
            });

            it("getCellText should extract the stringValue from a cell", () => {
                const text = getCellText({ mappedRange: mockMappedRange, rowOffset: 0, columnOffset: 0 });
                expect(text).toBe("Test String");
            });

            it("getCellText should return an empty string if the cell has no stringValue", () => {
                const text = getCellText({ mappedRange: mockMappedRange, rowOffset: 1, columnOffset: 0 });
                expect(text).toBe(""); // Row 1, Col 0 has a numberValue, not a stringValue
            });

            it("getCellNumber should extract the numberValue from a cell", () => {
                const num = getCellNumber({ mappedRange: mockMappedRange, rowOffset: 1, columnOffset: 0 });
                expect(num).toBe(0);
            });

            it("getCellNumber should return 0 if the cell has no numberValue", () => {
                const num = getCellNumber({ mappedRange: mockMappedRange, rowOffset: 0, columnOffset: 0 });
                expect(num).toBe(0); // Row 0, Col 0 has a stringValue, not a numberValue
            });
        });

        describe("getCellDataArray", () => {
            it("should return a correctly sized 5x5 array and ignore out-of-bounds data when unbound flags are false", () => {
                const result = getCellDataArray(mockMappedRange);

                // 1. Assert the dimensions strictly match the Named Range (5x5)
                expect(result.length).toBe(5);
                expect(result[0]?.length).toBe(5);

                // 2. Assert Chunk 1 data mapped correctly
                expect(result[0]?.[0]).toStrictEqual({ effectiveValue: { stringValue: "Test String" } });
                expect(result[1]?.[0]).toStrictEqual({ effectiveValue: { numberValue: 0 } });

                // 3. Assert Chunk 2 (Offset data) mapped correctly
                expect(result[3]?.[2]).toStrictEqual({ effectiveValue: { stringValue: "Offset Data 1" } });
                expect(result[3]?.[3]).toStrictEqual({ effectiveValue: { stringValue: "Offset Data 2" } });

                // 4. Assert an untouched cell remains an empty object
                expect(result[4]?.[4]).toStrictEqual({});
            });

            it("should expand rows to the sheet limit and include out-of-bounds row data when unboundRows is true", () => {
                const result = getCellDataArray(mockMappedRange, true, false);

                // Dimensions should be 10x5
                expect(result.length).toBe(10);
                expect(result[0]?.length).toBe(5);

                // Chunk 3 (Row 8, Col 2) should now be mapped
                expect(result[8]?.[2]).toStrictEqual({ effectiveValue: { stringValue: "Unbound Row Data" } });
            });

            it("should expand columns to the sheet limit and include out-of-bounds column data when unboundColumns is true", () => {
                const result = getCellDataArray(mockMappedRange, false, true);

                // Dimensions should be 5x10
                expect(result.length).toBe(5);
                expect(result[0]?.length).toBe(10);

                // Chunk 4 (Row 2, Col 8) should now be mapped
                expect(result[2]?.[8]).toStrictEqual({ effectiveValue: { stringValue: "Unbound Col Data" } });
            });

            it("should expand both rows and columns to the sheet limit when both unbound flags are true", () => {
                const result = getCellDataArray(mockMappedRange, true, true);

                // Dimensions should be 10x10
                expect(result.length).toBe(10);
                expect(result[0]?.length).toBe(10);

                // Both out-of-bounds chunks should be mapped
                expect(result[8]?.[2]).toStrictEqual({ effectiveValue: { stringValue: "Unbound Row Data" } });
                expect(result[2]?.[8]).toStrictEqual({ effectiveValue: { stringValue: "Unbound Col Data" } });
            });

            it("should natively fallback to gridProperties if the range lacks endRowIndex and endColumnIndex", () => {
                const result = getCellDataArray(mockUnboundedRange);

                // 1. Assert dimensions fallback to the gridProperties (100x26)
                expect(result.length).toBe(100);
                expect(result[0]?.length).toBe(26);

                // 2. Assert data at the very edge is mapped correctly
                expect(result[99]?.[25]).toStrictEqual({ effectiveValue: { stringValue: "Bottom Right Cell" } });
            });

            it("should return an empty array if range dimensions are 0", () => {
                const emptyRange: MappedNamedRange = {
                    range: { sheetId: 1, startRowIndex: 0, endRowIndex: 0, startColumnIndex: 0, endColumnIndex: 0 },
                    sheet: { data: [] },
                };
                const result = getCellDataArray(emptyRange);
                expect(result.length).toBe(0);
            });
        });

        describe("parseSpreadsheet", () => {
            const schema = {
                sheets: {
                    config: {
                        sheetName: "Configuration",
                        ranges: {
                            users: "UsersRange",
                        },
                    },
                },
            } as const;

            it("should correctly map sheets, named ranges, and all named ranges belonging to that sheet", () => {
                const mockSpreadsheet: GoogleAppsScript.Sheets.Schema.Spreadsheet = {
                    sheets: [
                        { properties: { sheetId: 1, title: "Configuration" } },
                        { properties: { sheetId: 2, title: "SecretData" } }, // Should be ignored
                    ],
                    namedRanges: [
                        { name: "UsersRange", range: { sheetId: 1 } },
                        { name: "OtherConfigRange", range: { sheetId: 1 } }, // Unmapped in schema, but still belongs to the sheet
                        { name: "HiddenRange", range: { sheetId: 2 } }, // Should be ignored
                    ],
                };

                const result = parseSpreadsheet(mockSpreadsheet, schema);

                // Check strict sheets mapping
                expect(result.sheets[schema.sheets.config.sheetName]).toBeDefined();
                expect(result.sheets[schema.sheets.config.sheetName]?.properties?.title).toBe("Configuration");
                expect("SecretData" in result.sheets).toBe(false);

                // Check strict named ranges mapping
                expect(result.mappedRanges[schema.sheets.config.ranges.users]).toBeDefined();
                expect(result.mappedRanges[schema.sheets.config.ranges.users]?.range.sheetId).toBe(1);
                expect("HiddenRange" in result.mappedRanges).toBe(false);
                expect("OtherConfigRange" in result.mappedRanges).toBe(false); // Valid for the sheet, but not in our explicit schema

                // Check the new sheetNamedRanges array
                const configSheetRanges = result.sheetNamedRanges[schema.sheets.config.sheetName];
                expect(configSheetRanges).toBeDefined();
                expect(configSheetRanges?.length).toBe(2);

                // It should grab BOTH ranges that belong to sheetId 1, regardless of schema
                const rangeNames = configSheetRanges?.map((r) => r.name);
                expect(rangeNames).toContain("UsersRange");
                expect(rangeNames).toContain("OtherConfigRange");
            });

            it("should safely handle undefined data safely", () => {
                const result = parseSpreadsheet(undefined, schema);
                expect(result.sheets).toEqual({});
                expect(result.mappedRanges).toEqual({});
                expect(result.sheetNamedRanges).toEqual({});
            });
        });
        describe("createRequiredGetter", () => {
            describe("when the key exists", () => {
                it("returns the corresponding value", () => {
                    const getter = createRequiredGetter({
                        a: 1,
                        b: 2,
                    });

                    expect(getter("a")).toBe(1);
                    expect(getter("b")).toBe(2);
                });

                it("does not throw for 0", () => {
                    const getter = createRequiredGetter({
                        zero: 0,
                    });

                    expect(getter("zero")).toBe(0);
                });

                it("does not throw for false", () => {
                    const getter = createRequiredGetter({
                        falseValue: false,
                    });

                    expect(getter("falseValue")).toBe(false);
                });

                it("does not throw for empty string", () => {
                    const getter = createRequiredGetter({
                        emptyString: "",
                    });

                    expect(getter("emptyString")).toBe("");
                });

                it("supports number keys", () => {
                    const getter = createRequiredGetter<number, string>({
                        1: "one",
                        2: "two",
                    });

                    expect(getter(1)).toBe("one");
                });

                it("supports symbol keys", () => {
                    const key = Symbol("test");

                    const getter = createRequiredGetter({
                        [key]: 123,
                    });

                    expect(getter(key)).toBe(123);
                });
            });

            describe("when the key does not exist", () => {
                it("throws with the default message when no context is provided", () => {
                    const getter = createRequiredGetter<string, number>({
                        a: 1,
                    });

                    expect(() => getter("b")).toThrow("Falta propidad: b");
                });

                it("throws with the contextualized message", () => {
                    const getter = createRequiredGetter<string, number>(
                        {
                            a: 1,
                        },
                        "rango",
                    );

                    expect(() => getter("b")).toThrow("Falta rango: b");
                });

                it("includes symbol keys in the error message", () => {
                    const key = Symbol("missing");

                    const getter = createRequiredGetter<symbol, number>({}, "symbol key");

                    expect(() => getter(key)).toThrow("Falta symbol key: Symbol(missing)");
                });
            });
        });
    });

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

    describe("makeUserEntered", () => {
        it("should overwrite userEntered values with effective values and delete effectives", () => {
            const data: GoogleAppsScript.Sheets.Schema.CellData[][] = [
                [
                    {
                        userEnteredValue: { stringValue: "old" },
                        effectiveValue: { stringValue: "new" },
                        effectiveFormat: { backgroundColor: { red: 1 } },
                        note: "hello",
                    },
                ],
            ];

            const result = makeUserEntered(data, false);
            const firstRow = result[0];
            const firstCell = firstRow ? firstRow[0] : undefined;

            expect(firstCell).toBeDefined();
            expect(firstCell?.userEnteredValue).toEqual({ stringValue: "new" });
            expect(firstCell?.userEnteredFormat).toEqual({ backgroundColor: { red: 1 } });
            expect(firstCell?.note).toBe("hello");
            expect(firstCell?.effectiveValue).toBeUndefined();
            expect(firstCell?.effectiveFormat).toBeUndefined();
        });

        it("should strip out all other properties if stripOthers is true", () => {
            const data: GoogleAppsScript.Sheets.Schema.CellData[][] = [
                [
                    {
                        effectiveValue: { numberValue: 42 },
                        note: "should be stripped",
                        hyperlink: "http://example.com",
                    },
                ],
            ];

            const result = makeUserEntered(data, true);
            const firstRow = result[0];
            const firstCell = firstRow ? firstRow[0] : undefined;

            expect(firstCell).toBeDefined();
            expect(firstCell?.userEnteredValue).toEqual({ numberValue: 42 });
            expect(firstCell?.note).toBeUndefined();
            expect(firstCell?.hyperlink).toBeUndefined();
        });
    });
});
