import { MappedNamedRange, parseSpreadsheet } from "./mapped-name-range";

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
            const data = MappedNamedRange.getCellData({ mappedRange: mockMappedRange, rowOffset: 0, columnOffset: 0 });
            expect(data?.effectiveValue?.stringValue).toBe("Test String");
        });

        it("getCellData should correctly locate data in an offset data chunk", () => {
            const data = MappedNamedRange.getCellData({ mappedRange: mockMappedRange, rowOffset: 3, columnOffset: 2 });
            expect(data?.effectiveValue?.stringValue).toBe("Offset Data 1");
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
            // Because endRowIndex/endColumnIndex are undefined, the boundary check should naturally pass
            const data = MappedNamedRange.getCellData({ mappedRange: mockUnboundedRange, rowOffset: 99, columnOffset: 25 });
            expect(data?.effectiveValue?.stringValue).toBe("Bottom Right Cell");
        });

        it("getCellEffectiveValue should return the effectiveValue object of a cell", () => {
            const value = MappedNamedRange.getCellEffectiveValue({ mappedRange: mockMappedRange, rowOffset: 0, columnOffset: 0 });
            expect(value).toStrictEqual({ stringValue: "Test String" });
        });

        it("getCellText should extract the stringValue from a cell", () => {
            const text = MappedNamedRange.getCellText({ mappedRange: mockMappedRange, rowOffset: 0, columnOffset: 0 });
            expect(text).toBe("Test String");
        });

        it("getCellText should safely return undefined if the cell has no stringValue", () => {
            const text = MappedNamedRange.getCellText({ mappedRange: mockMappedRange, rowOffset: 1, columnOffset: 0 });
            expect(text).toBeUndefined(); // Row 1, Col 0 has a numberValue, not a stringValue
        });

        it("getCellNumber should extract the numberValue from a cell", () => {
            const num = MappedNamedRange.getCellNumber({ mappedRange: mockMappedRange, rowOffset: 1, columnOffset: 0 });
            expect(num).toBe(0);
        });

        it("getCellNumber should safely return undefined if the cell has no numberValue", () => {
            const num = MappedNamedRange.getCellNumber({ mappedRange: mockMappedRange, rowOffset: 0, columnOffset: 0 });
            expect(num).toBeUndefined(); // Row 0, Col 0 has a stringValue, not a numberValue
        });
    });

    describe("getCellDataArray", () => {
        it("should return a correctly sized 5x5 array and ignore out-of-bounds data when unbound flags are false", () => {
            const result = MappedNamedRange.getCellDataArray(mockMappedRange);

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
            const result = MappedNamedRange.getCellDataArray(mockMappedRange, true, false);

            // Dimensions should be 10x5
            expect(result.length).toBe(10);
            expect(result[0]?.length).toBe(5);

            // Chunk 3 (Row 8, Col 2) should now be mapped
            expect(result[8]?.[2]).toStrictEqual({ effectiveValue: { stringValue: "Unbound Row Data" } });
        });

        it("should expand columns to the sheet limit and include out-of-bounds column data when unboundColumns is true", () => {
            const result = MappedNamedRange.getCellDataArray(mockMappedRange, false, true);

            // Dimensions should be 5x10
            expect(result.length).toBe(5);
            expect(result[0]?.length).toBe(10);

            // Chunk 4 (Row 2, Col 8) should now be mapped
            expect(result[2]?.[8]).toStrictEqual({ effectiveValue: { stringValue: "Unbound Col Data" } });
        });

        it("should expand both rows and columns to the sheet limit when both unbound flags are true", () => {
            const result = MappedNamedRange.getCellDataArray(mockMappedRange, true, true);

            // Dimensions should be 10x10
            expect(result.length).toBe(10);
            expect(result[0]?.length).toBe(10);

            // Both out-of-bounds chunks should be mapped
            expect(result[8]?.[2]).toStrictEqual({ effectiveValue: { stringValue: "Unbound Row Data" } });
            expect(result[2]?.[8]).toStrictEqual({ effectiveValue: { stringValue: "Unbound Col Data" } });
        });

        it("should natively fallback to gridProperties if the range lacks endRowIndex and endColumnIndex", () => {
            const result = MappedNamedRange.getCellDataArray(mockUnboundedRange);

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
            const result = MappedNamedRange.getCellDataArray(emptyRange);
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
            expect(result.namedRanges[schema.sheets.config.ranges.users]).toBeDefined();
            expect(result.namedRanges[schema.sheets.config.ranges.users]?.range.sheetId).toBe(1);
            expect("HiddenRange" in result.namedRanges).toBe(false);
            expect("OtherConfigRange" in result.namedRanges).toBe(false); // Valid for the sheet, but not in our explicit schema

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
            expect(result.namedRanges).toEqual({});
            expect(result.sheetNamedRanges).toEqual({});
        });
    });
});
