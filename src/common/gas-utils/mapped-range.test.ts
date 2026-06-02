import {
    getCellBoolean,
    getCellData,
    getCellDataArray,
    getCellEffectiveValue,
    getCellNumber,
    getCellText,
    getCellUnixEpoch,
    type MappedNamedRange,
    resizeMappedRange,
} from ".";
import { Dimension } from "./types";

describe("MappedNamedRange Utilities", () => {
    // Simulates a 5x5 Named Range (A1:E5) with sparse data chunks, but the sheet is 10x10.
    const mockMappedRange: MappedNamedRange = {
        namedRange: { name: "", range: { sheetId: 1, startRowIndex: 0, endRowIndex: 5, startColumnIndex: 0, endColumnIndex: 5 } },
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
                        { values: [{ effectiveValue: { boolValue: true } }] }, // Row 2, Col 0 (For boolean test)
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
        namedRange: { name: "", range: { sheetId: 2 } }, // No end indexes
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
                {
                    startRow: 3,
                    startColumn: 5,
                },
                {
                    startRow: 7,
                    startColumn: 5,
                    rowData: [{}],
                },
            ],
        },
    };

    describe("getCellDataArray", () => {
        it("should return a correctly sized 5x5 array and ignore out-of-bounds data when unbound flags are false", () => {
            const result = getCellDataArray(mockMappedRange);

            expect(result.length).toBe(5);
            expect(result[0]?.length).toBe(5);
            expect(result[0]?.[0]).toStrictEqual({ effectiveValue: { stringValue: "Test String" } });
            expect(result[1]?.[0]).toStrictEqual({ effectiveValue: { numberValue: 0 } });
            expect(result[3]?.[2]).toStrictEqual({ effectiveValue: { stringValue: "Offset Data 1" } });
            expect(result[4]?.[4]).toStrictEqual({});
        });

        it("should return an empty result matrix if the sheet lacks a data array", () => {
            const emptyDataRange: MappedNamedRange = {
                namedRange: { name: "", range: { sheetId: 1, startRowIndex: 0, endRowIndex: 2, startColumnIndex: 0, endColumnIndex: 2 } },
                sheet: { properties: { gridProperties: { rowCount: 10, columnCount: 10 } } },
            };
            const result = getCellDataArray(emptyDataRange);
            expect(result.length).toBe(2);
            expect(result[0]?.[0]).toStrictEqual({});
        });

        it("should handle missing grid properties safely by falling back to start index", () => {
            const noGridPropsRange: MappedNamedRange = {
                namedRange: { name: "", range: { sheetId: 1, startRowIndex: 2, startColumnIndex: 2 } },
                sheet: { data: [] },
            };
            const result = getCellDataArray(noGridPropsRange, true, true);
            expect(result.length).toBe(0);
        });

        it("should expand rows to the sheet limit and include out-of-bounds row data when unboundRows is true", () => {
            const result = getCellDataArray(mockMappedRange, true, false);
            expect(result.length).toBe(10);
            expect(result[8]?.[2]).toStrictEqual({ effectiveValue: { stringValue: "Unbound Row Data" } });
        });

        it("should expand columns to the sheet limit and include out-of-bounds column data when unboundColumns is true", () => {
            const result = getCellDataArray(mockMappedRange, false, true);
            expect(result[0]?.length).toBe(10);
            expect(result[2]?.[8]).toStrictEqual({ effectiveValue: { stringValue: "Unbound Col Data" } });
        });

        it("should expand both rows and columns to the sheet limit when both unbound flags are true", () => {
            const result = getCellDataArray(mockMappedRange, true, true);
            expect(result.length).toBe(10);
            expect(result[0]?.length).toBe(10);
        });

        it("should natively fallback to gridProperties if the range lacks endRowIndex and endColumnIndex", () => {
            const result = getCellDataArray(mockUnboundedRange);
            expect(result.length).toBe(100);
            expect(result[0]?.length).toBe(26);
        });

        it("should return an empty array if range dimensions are 0", () => {
            const emptyRange: MappedNamedRange = {
                namedRange: { name: "", range: { sheetId: 1, startRowIndex: 0, endRowIndex: 0, startColumnIndex: 0, endColumnIndex: 0 } },
                sheet: { data: [] },
            };
            const result = getCellDataArray(emptyRange);
            expect(result.length).toBe(0);
        });
    });

    describe("getCellData", () => {
        it("should return data if within bounds of the first data chunk", () => {
            const data = getCellData({ mappedRange: mockMappedRange, rowOffset: 0, columnOffset: 0 });
            expect(data?.effectiveValue?.stringValue).toBe("Test String");
        });

        it("should correctly locate data in an offset data chunk", () => {
            const data = getCellData({ mappedRange: mockMappedRange, rowOffset: 3, columnOffset: 2 });
            expect(data?.effectiveValue?.stringValue).toBe("Offset Data 1");
        });

        it("should return an empty object for empty cells within bounds", () => {
            const data = getCellData({ mappedRange: mockMappedRange, rowOffset: 0, columnOffset: 4 });
            expect(data).toStrictEqual({});
        });

        it("should return an empty object if out of bounds", () => {
            const data = getCellData({ mappedRange: mockMappedRange, rowOffset: 10, columnOffset: 0 });
            expect(data).toStrictEqual({});
        });

        it("should safely handle sheets missing the data property entirely", () => {
            const emptySheetRange: MappedNamedRange = { namedRange: { name: "", range: { sheetId: 1 } }, sheet: {} };
            const data = getCellData({ mappedRange: emptySheetRange });
            expect(data).toStrictEqual({});
        });

        it("should handle unbounded ranges without throwing out-of-bounds errors", () => {
            const data = getCellData({ mappedRange: mockUnboundedRange, rowOffset: 99, columnOffset: 25 });
            expect(data?.effectiveValue?.stringValue).toBe("Bottom Right Cell");
        });
    });

    describe("getCellEffectiveValue", () => {
        it("should return the effectiveValue object of a cell", () => {
            const value = getCellEffectiveValue({ mappedRange: mockMappedRange, rowOffset: 0, columnOffset: 0 });
            expect(value).toStrictEqual({ stringValue: "Test String" });
        });

        it("should return an empty object if the cell doesn't have an effective value", () => {
            const value = getCellEffectiveValue({ mappedRange: mockUnboundedRange, rowOffset: 3, columnOffset: 5 });
            expect(value).toStrictEqual({});
        });
    });

    describe("getCellText", () => {
        it("should extract the stringValue from a cell", () => {
            const text = getCellText({ mappedRange: mockMappedRange, rowOffset: 0, columnOffset: 0 });
            expect(text).toBe("Test String");
        });

        it("should return an empty string if the cell has no stringValue", () => {
            const text = getCellText({ mappedRange: mockMappedRange, rowOffset: 1, columnOffset: 0 });
            expect(text).toBe("");
        });
    });

    describe("getCellBoolean", () => {
        it("should extract the boolValue from a cell", () => {
            const bool = getCellBoolean({ mappedRange: mockMappedRange, rowOffset: 2, columnOffset: 0 });
            expect(bool).toBe(true);
        });

        it("should return false if the cell has no boolValue", () => {
            const bool = getCellBoolean({ mappedRange: mockMappedRange, rowOffset: 0, columnOffset: 0 });
            expect(bool).toBe(false);
        });
    });

    describe("getCellNumber", () => {
        it("should extract the numberValue from a cell", () => {
            const num = getCellNumber({ mappedRange: mockMappedRange, rowOffset: 1, columnOffset: 0 });
            expect(num).toBe(0);
        });

        it("should return 0 if the cell has no numberValue", () => {
            const num = getCellNumber({ mappedRange: mockMappedRange, rowOffset: 0, columnOffset: 0 });
            expect(num).toBe(0);
        });
    });

    describe("getCellUnixEpoch", () => {
        it("should correctly calculate the 1899 Sheets epoch", () => {
            const expectedEpoch = new Date(Date.UTC(1899, 11, 30)).getTime();
            const epoch = getCellUnixEpoch({ mappedRange: mockMappedRange, rowOffset: 1, columnOffset: 0 });
            expect(epoch).toBe(expectedEpoch);
        });
    });

    describe("resizeMappedRange", () => {
        let testRange: MappedNamedRange;

        beforeEach(() => {
            testRange = {
                namedRange: { name: "", range: { sheetId: 123, startRowIndex: 0, endRowIndex: 5, startColumnIndex: 0, endColumnIndex: 5 } },
                sheet: {},
            };
        });

        it("should insert dimensions when the target size is larger than the current size", () => {
            const result = resizeMappedRange({
                target: testRange,
                targetRows: 8,
                targetCols: 7,
            });

            expect(result.requests).toHaveLength(2);
            expect(result.requests[0]?.insertDimension?.range).toEqual({
                sheetId: 123,
                dimension: Dimension.ROWS,
                startIndex: 1,
                endIndex: 4,
            });
            expect(testRange.namedRange.range.endRowIndex).toBe(8);
        });

        it("should delete dimensions when the target size is smaller than the current size", () => {
            const result = resizeMappedRange({
                target: testRange,
                targetRows: 2,
                targetCols: 3,
            });

            expect(result.requests).toHaveLength(2);
            expect(result.requests[0]?.deleteDimension?.range).toEqual({
                sheetId: 123,
                dimension: Dimension.ROWS,
                startIndex: 2,
                endIndex: 5,
            });
        });

        it("should destroy the range entirely when collapsing to 0", () => {
            const result = resizeMappedRange({
                target: testRange,
                targetRows: 0,
                targetCols: 0,
            });

            expect(result.requests).toHaveLength(2);
            expect(testRange.namedRange.range.endRowIndex).toBe(0);
        });

        it("should return empty requests and leave dimensions unchanged if target size matches current size", () => {
            const result = resizeMappedRange({ target: testRange });
            expect(result.requests).toHaveLength(0);
        });

        it("should handle offset calculations correctly if the current range is already offset", () => {
            const result = resizeMappedRange({
                target: testRange,
                targetRows: 10,
                rowOffset: 5,
                colOffset: 5,
            });
            expect(result.requests[0]?.insertDimension?.range?.startIndex).toBe(6);
            expect(result.rowOffset).toBe(10);
        });

        // NEW: Covers fallback rule `sheetId = target.range.sheetId ?? 0;`
        it("should default the sheetId to 0 if it is omitted from the target range object", () => {
            const statelessRange: MappedNamedRange = {
                namedRange: { name: "", range: { startRowIndex: 0, endRowIndex: 2, startColumnIndex: 0, endColumnIndex: 2 } }, // Missing sheetId
                sheet: {},
            };
            const result = resizeMappedRange({ target: statelessRange, targetRows: 5 });
            expect(result.requests[0]?.insertDimension?.range?.sheetId).toBe(0);
        });

        // NEW: Covers fallback definitions `actualEndRow = actualRange.endRowIndex ?? 0` and `actualEndCol = actualRange.endColumnIndex ?? 0`
        it("should fall back to 0 end-indexes safely when resizing open, horizontally/vertically unbounded ranges", () => {
            const openRange: MappedNamedRange = {
                namedRange: { name: "", range: { sheetId: 1, startRowIndex: 0, startColumnIndex: 0 } }, // Missing endRowIndex and endColumnIndex
                sheet: {},
            };
            const result = resizeMappedRange({
                target: openRange,
                targetRows: 5, // Triggers insert dimension from baseline 0
                targetCols: 5,
            });
            expect(result.requests).toHaveLength(2);
            expect(result.requests[0]?.insertDimension?.range?.endIndex).toBe(6);
            expect(openRange.namedRange.range.endRowIndex).toBe(5);
        });
    });
});
