import { MappedNamedRange, PasteType } from "../../../common/utils/mappedNameRange";

describe("MappedNamedRange", () => {
    // Simulates a 5x5 Named Range (A1:E5) with sparse data chunks
    const mockMappedRange: MappedNamedRange = {
        range: { sheetId: 1, startRowIndex: 0, endRowIndex: 5, startColumnIndex: 0, endColumnIndex: 5 },
        sheet: {
            data: [
                {
                    // Chunk 1: Implicitly starts at 0, 0
                    rowData: [
                        { values: [{ formattedValue: "Test String", effectiveValue: { numberValue: 42 } }] }, // Row 0, Col 0
                        { values: [{ effectiveValue: { numberValue: 0 } }] }, // Row 1, Col 0 (For epoch test, no formattedValue)
                    ],
                },
                {
                    // Chunk 2: Offset block starting at Row 3, Col 2
                    startRow: 3,
                    startColumn: 2,
                    rowData: [
                        {
                            values: [
                                { formattedValue: "Offset Data 1" }, // Row 3, Col 2
                                { formattedValue: "Offset Data 2" }, // Row 3, Col 3
                            ],
                        },
                    ],
                },
            ],
        },
    };
    // Mock an unbounded range (e.g., A:Z) where endRowIndex and endColumnIndex are undefined.
    // It relies on the sheet's gridProperties to determine its maximum size (100 rows, 26 columns).
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
                    rowData: [{ values: [{ formattedValue: "Bottom Right Cell" }] }],
                },
            ],
        },
    };

    describe("getCellData & Utilities", () => {
        it("getCellData should return data if within bounds of the first data chunk", () => {
            const data = MappedNamedRange.getCellData({ mappedRange: mockMappedRange, rowOffset: 0, columnOffset: 0 });
            expect(data?.formattedValue).toBe("Test String");
        });

        it("getCellData should correctly locate data in an offset data chunk", () => {
            const data = MappedNamedRange.getCellData({ mappedRange: mockMappedRange, rowOffset: 3, columnOffset: 2 });
            expect(data?.formattedValue).toBe("Offset Data 1");
        });

        it("getCellData should return undefined for empty cells within bounds", () => {
            const data = MappedNamedRange.getCellData({ mappedRange: mockMappedRange, rowOffset: 0, columnOffset: 4 });
            expect(data).toBeUndefined();
        });

        it("getCellData should return undefined if out of bounds", () => {
            const data = MappedNamedRange.getCellData({ mappedRange: mockMappedRange, rowOffset: 10, columnOffset: 0 });
            expect(data).toBeUndefined();
        });

        it("getCellUnixEpoch should correctly calculate the 1899 Sheets epoch", () => {
            // Sheets epoch (0) is Dec 30, 1899
            const expectedEpoch = new Date(Date.UTC(1899, 11, 30)).getTime();
            const epoch = MappedNamedRange.getCellUnixEpoch({ mappedRange: mockMappedRange, rowOffset: 1, columnOffset: 0 });
            expect(epoch).toBe(expectedEpoch);
        });
        it("getCellData should handle unbounded ranges without throwing out-of-bounds errors", () => {
            // Because endRowIndex/endColumnIndex are undefined, the boundary check should naturally pass,
            // and it should successfully find the data chunk way down the sheet.
            const data = MappedNamedRange.getCellData({ mappedRange: mockUnboundedRange, rowOffset: 99, columnOffset: 25 });
            expect(data?.formattedValue).toBe("Bottom Right Cell");
        });
    });

    describe("getFormattedValues", () => {
        it("should return a correctly sized 2D array stitched from sparse data chunks", () => {
            const result = MappedNamedRange.getFormattedValues(mockMappedRange);

            // 1. Assert the dimensions match the Named Range (5x5)
            expect(result.length).toBe(5);
            expect(result[0]?.length).toBe(5);

            // 2. Assert Chunk 1 data mapped correctly
            expect(result[0]?.[0]).toBe("Test String");

            // 3. Assert a cell with an effectiveValue but no formattedValue returns an empty string
            expect(result[1]?.[0]).toBe("");

            // 4. Assert Chunk 2 (Offset data) mapped correctly
            expect(result[3]?.[2]).toBe("Offset Data 1");
            expect(result[3]?.[3]).toBe("Offset Data 2");

            // 5. Assert an untouched cell remains an empty string
            expect(result[4]?.[4]).toBe("");
        });

        it("should return an empty array if range dimensions are 0", () => {
            const emptyRange: MappedNamedRange = {
                range: { sheetId: 1, startRowIndex: 0, endRowIndex: 0, startColumnIndex: 0, endColumnIndex: 0 },
                sheet: { data: [] },
            };
            const result = MappedNamedRange.getFormattedValues(emptyRange);
            expect(result.length).toBe(0);
        });
    });

    describe("buildCopyPasteRequest", () => {
        it("should return a valid request object on success", () => {
            const req = MappedNamedRange.buildCopyPasteRequest({
                mappedRange: mockMappedRange,
                destinationSheetId: 99,
                destinationStartRow: 10,
                destinationStartColumn: 20,
                pasteType: PasteType.PASTE_NORMAL,
                width: 2,
            });

            expect(req).toBeDefined();
            expect(req?.copyPaste).toBeDefined();

            // Source assertions
            expect(req?.copyPaste?.source?.sheetId).toBe(1);
            expect(req?.copyPaste?.source?.startRowIndex).toBe(0);
            expect(req?.copyPaste?.source?.endRowIndex).toBe(5);
            expect(req?.copyPaste?.source?.startColumnIndex).toBe(0);
            expect(req?.copyPaste?.source?.endColumnIndex).toBe(2);

            // Destination assertions
            expect(req?.copyPaste?.destination?.sheetId).toBe(99);
            expect(req?.copyPaste?.destination?.startRowIndex).toBe(10);
            expect(req?.copyPaste?.destination?.endRowIndex).toBe(15);
            expect(req?.copyPaste?.destination?.startColumnIndex).toBe(20);
            expect(req?.copyPaste?.destination?.endColumnIndex).toBe(22);
        });

        it("should return undefined if mappedRange is missing", () => {
            const req = MappedNamedRange.buildCopyPasteRequest({
                mappedRange: undefined,
                destinationSheetId: 2,
                destinationStartRow: 0,
                destinationStartColumn: 0,
                pasteType: PasteType.PASTE_NORMAL,
            });
            expect(req).toBeUndefined();
        });

        it("should return undefined if requested range is out of bounds", () => {
            const req = MappedNamedRange.buildCopyPasteRequest({
                mappedRange: mockMappedRange,
                destinationSheetId: 2,
                destinationStartRow: 0,
                destinationStartColumn: 0,
                pasteType: PasteType.PASTE_NORMAL,
                height: 10, // Exceeds the endRowIndex of 5
                width: 1,
            });
            expect(req).toBeUndefined();
        });
    });
});
