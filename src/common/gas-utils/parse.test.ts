import { parseSpreadsheet } from "./parse";

describe("parseSpreadsheet", () => {
    const schema = {
        sheets: {
            config: {
                sheetName: "Configuration",
                ranges: {
                    users: "UsersRange",
                    overlap: "student_static", // Test static vs dynamic priority
                },
                dynamicRanges: {
                    students: "student",
                    temps: "temp_",
                },
            },
            // Adding a sheet without ranges to cover the `if (sheetConfig?.ranges)` branch
            logs: {
                sheetName: "Logs",
            },
        },
    } as const;

    it("should correctly map sheets, named ranges, dynamic ranges, and all named ranges belonging to that sheet", () => {
        const mockSpreadsheet: GoogleAppsScript.Sheets.Schema.Spreadsheet = {
            sheets: [
                { properties: { sheetId: 1, title: "Configuration" } },
                { properties: { sheetId: 2, title: "SecretData" } }, // Should be ignored
                { properties: { sheetId: 3, title: "Logs" } },
            ],
            namedRanges: [
                { name: "UsersRange", range: { sheetId: 1 } },
                { name: "OtherConfigRange", range: { sheetId: 1 } }, // Unmapped in schema, but belongs to the sheet
                { name: "HiddenRange", range: { sheetId: 2 } }, // Should be ignored

                // Dynamic ranges
                { name: "student00", range: { sheetId: 1 } },
                { name: "student01", range: { sheetId: 1 } },
                { name: "temp_A", range: { sheetId: 1 } },
                { name: "temp_B", range: { sheetId: 1 } },

                // Static range that happens to share a dynamic prefix
                { name: "student_static", range: { sheetId: 1 } },
            ],
        };

        const result = parseSpreadsheet(mockSpreadsheet, schema);

        // Check strict sheets mapping
        expect(result.mappedSheets[schema.sheets.config.sheetName]).toBeDefined();
        expect(result.mappedSheets[schema.sheets.config.sheetName]?.properties?.title).toBe("Configuration");
        expect(result.mappedSheets[schema.sheets.logs.sheetName]?.properties?.title).toBe("Logs");
        expect("SecretData" in result.mappedSheets).toBe(false);

        // Check strict named ranges mapping
        expect(result.mappedRanges[schema.sheets.config.ranges.users]).toBeDefined();
        expect(result.mappedRanges[schema.sheets.config.ranges.users]?.namedRange.range.sheetId).toBe(1);
        expect("HiddenRange" in result.mappedRanges).toBe(false);
        expect("OtherConfigRange" in result.mappedRanges).toBe(false);

        // Check dynamic named ranges mapping
        expect(result.dynamicMappedRanges.students).toBeDefined();
        expect(result.dynamicMappedRanges.students?.length).toBe(2);
        expect(result.dynamicMappedRanges.students?.map((r) => r.namedRange.range.sheetId)).toEqual([1, 1]);

        expect(result.dynamicMappedRanges.temps).toBeDefined();
        expect(result.dynamicMappedRanges.temps?.length).toBe(2);

        // Verify Static Priority: "student_static" should be in mappedRanges, NOT in dynamicMappedRanges.students
        expect(result.mappedRanges[schema.sheets.config.ranges.overlap]).toBeDefined();
        expect(result.dynamicMappedRanges.students?.length).toBe(2);

        // Check the sheetNamedRanges array mapping
        const configSheetRanges = result.mappedSheetNamedRanges[schema.sheets.config.sheetName];
        expect(configSheetRanges).toBeDefined();
        expect(configSheetRanges?.length).toBe(7); // 2 static + 1 unmapped + 4 dynamic

        const rangeNames = configSheetRanges?.map((r) => r.name);
        expect(rangeNames).toContain("UsersRange");
        expect(rangeNames).toContain("OtherConfigRange");
        expect(rangeNames).toContain("student00");
    });

    it("should safely handle an undefined spreadsheet or missing sheets array", () => {
        const result1 = parseSpreadsheet(undefined, schema);
        expect(result1.mappedSheets).toEqual({});
        expect(result1.mappedRanges).toEqual({});
        expect(result1.mappedSheetNamedRanges).toEqual({});
        expect(result1.dynamicMappedRanges).toEqual({});

        const result2 = parseSpreadsheet({} as GoogleAppsScript.Sheets.Schema.Spreadsheet, schema);
        expect(result2.mappedSheets).toEqual({});
    });

    it("should gracefully handle a spreadsheet that lacks namedRanges", () => {
        const mockSpreadsheet: GoogleAppsScript.Sheets.Schema.Spreadsheet = {
            sheets: [{ properties: { sheetId: 1, title: "Configuration" } }],
        };

        const result = parseSpreadsheet(mockSpreadsheet, schema);

        expect(result.mappedSheets["Configuration"]).toBeDefined();
        expect(result.mappedRanges).toEqual({});
        expect(result.dynamicMappedRanges).toEqual({});
        expect(result.mappedSheetNamedRanges["Configuration"]).toBeUndefined();
    });

    it("should handle sheets missing properties, titles, or IDs", () => {
        const mockSpreadsheet: GoogleAppsScript.Sheets.Schema.Spreadsheet = {
            sheets: [
                {},
                { properties: {} },
                { properties: { title: "Configuration" } }, // missing ID (falls back to 0)
            ],
            namedRanges: [
                { name: "UsersRange", range: { sheetId: 0 } },
                { name: "student00", range: { sheetId: 0 } },
            ],
        };

        const result = parseSpreadsheet(mockSpreadsheet, schema);

        expect(result.mappedSheets["Configuration"]).toBeDefined();
        expect(result.mappedRanges["UsersRange"]).toBeDefined();
        expect(result.dynamicMappedRanges["students"]).toBeDefined();
    });

    it("should safely ignore malformed named ranges or ranges pointing to unknown sheets", () => {
        const mockSpreadsheet: GoogleAppsScript.Sheets.Schema.Spreadsheet = {
            sheets: [{ properties: { sheetId: 1, title: "Configuration" } }],
            namedRanges: [
                {},
                { name: "UsersRange" },
                { range: { sheetId: 1 } },
                { name: "UsersRange", range: { sheetId: 999 } },
                { name: "student01", range: { sheetId: 999 } },
            ],
        };

        const result = parseSpreadsheet(mockSpreadsheet, schema);

        expect(result.mappedSheets["Configuration"]).toBeDefined();
        expect(result.mappedRanges).toEqual({});
        expect(result.dynamicMappedRanges).toEqual({});
    });

    it("should handle malformed schema definitions safely at runtime", () => {
        const malformedSchema = {
            sheets: {
                validSheet: {
                    sheetName: "ValidSheet",
                    ranges: {
                        validRange: "ValidRange",
                        invalidRange: undefined as unknown as string,
                    },
                    dynamicRanges: {
                        validDynamic: "valid_",
                        invalidDynamic: undefined as unknown as string,
                    },
                },
                invalidSheet: {
                    sheetName: undefined as unknown as string,
                },
            },
        } as const;

        const mockSpreadsheet: GoogleAppsScript.Sheets.Schema.Spreadsheet = {
            sheets: [{ properties: { sheetId: 1, title: "ValidSheet" } }],
            namedRanges: [
                { name: "ValidRange", range: { sheetId: 1 } },
                { name: "valid_01", range: { sheetId: 1 } },
            ],
        };

        const result = parseSpreadsheet(mockSpreadsheet, malformedSchema);

        expect(result.mappedSheets["ValidSheet"]).toBeDefined();
        expect(result.mappedRanges["ValidRange"]).toBeDefined();
        expect(result.dynamicMappedRanges["validDynamic"]).toBeDefined();
    });

    it("should fallback to sheetId 0 when a named range omits the sheetId", () => {
        const fallbackSchema = {
            sheets: {
                default: {
                    sheetName: "DefaultSheet",
                    ranges: {
                        zero: "ZeroRange",
                    },
                    dynamicRanges: {
                        dyn: "dyn_",
                    },
                },
            },
        } as const;

        const mockSpreadsheet: GoogleAppsScript.Sheets.Schema.Spreadsheet = {
            sheets: [{ properties: { sheetId: 0, title: "DefaultSheet" } }],
            namedRanges: [
                {
                    name: "ZeroRange",
                    range: { startRowIndex: 1 }, // Omitted sheetId
                },
                {
                    name: "dyn_A",
                    range: { startRowIndex: 2 }, // Omitted sheetId
                },
            ],
        };

        const result = parseSpreadsheet(mockSpreadsheet, fallbackSchema);

        expect(result.mappedSheets["DefaultSheet"]).toBeDefined();
        expect(result.mappedRanges["ZeroRange"]).toBeDefined();
        expect(result.mappedRanges["ZeroRange"]?.sheet.properties?.title).toBe("DefaultSheet");
        expect(result.dynamicMappedRanges["dyn"]).toBeDefined();
    });
});
