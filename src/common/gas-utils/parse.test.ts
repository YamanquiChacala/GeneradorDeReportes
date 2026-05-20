import { parseSpreadsheet } from ".";

describe("parseSpreadsheet", () => {
    const schema = {
        sheets: {
            config: {
                sheetName: "Configuration",
                ranges: {
                    users: "UsersRange",
                },
            },
            // Adding a sheet without ranges to cover the `if (sheetConfig?.ranges)` branch
            logs: {
                sheetName: "Logs",
            },
        },
    } as const;

    it("should correctly map sheets, named ranges, and all named ranges belonging to that sheet", () => {
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
            ],
        };

        const result = parseSpreadsheet(mockSpreadsheet, schema);

        // Check strict sheets mapping
        expect(result.sheets[schema.sheets.config.sheetName]).toBeDefined();
        expect(result.sheets[schema.sheets.config.sheetName]?.properties?.title).toBe("Configuration");
        expect(result.sheets[schema.sheets.logs.sheetName]?.properties?.title).toBe("Logs");
        expect("SecretData" in result.sheets).toBe(false);

        // Check strict named ranges mapping
        expect(result.mappedRanges[schema.sheets.config.ranges.users]).toBeDefined();
        expect(result.mappedRanges[schema.sheets.config.ranges.users]?.range.sheetId).toBe(1);
        expect("HiddenRange" in result.mappedRanges).toBe(false);
        expect("OtherConfigRange" in result.mappedRanges).toBe(false);

        // Check the sheetNamedRanges array mapping
        const configSheetRanges = result.sheetNamedRanges[schema.sheets.config.sheetName];
        expect(configSheetRanges).toBeDefined();
        expect(configSheetRanges?.length).toBe(2);

        const rangeNames = configSheetRanges?.map((r) => r.name);
        expect(rangeNames).toContain("UsersRange");
        expect(rangeNames).toContain("OtherConfigRange");
    });

    it("should safely handle an undefined spreadsheet or missing sheets array", () => {
        // Entirely undefined
        const result1 = parseSpreadsheet(undefined, schema);
        expect(result1.sheets).toEqual({});
        expect(result1.mappedRanges).toEqual({});
        expect(result1.sheetNamedRanges).toEqual({});

        // Missing sheets array
        const result2 = parseSpreadsheet({} as GoogleAppsScript.Sheets.Schema.Spreadsheet, schema);
        expect(result2.sheets).toEqual({});
    });

    it("should gracefully handle a spreadsheet that lacks namedRanges", () => {
        const mockSpreadsheet: GoogleAppsScript.Sheets.Schema.Spreadsheet = {
            sheets: [{ properties: { sheetId: 1, title: "Configuration" } }],
            // explicitly omitting namedRanges
        };

        const result = parseSpreadsheet(mockSpreadsheet, schema);

        expect(result.sheets["Configuration"]).toBeDefined();
        expect(result.mappedRanges).toEqual({});
        // Returns undefined for the array of named ranges because namedRanges didn't exist to filter
        expect(result.sheetNamedRanges["Configuration"]).toBeUndefined();
    });

    it("should handle sheets missing properties, titles, or IDs", () => {
        const mockSpreadsheet: GoogleAppsScript.Sheets.Schema.Spreadsheet = {
            sheets: [
                {}, // completely empty sheet object
                { properties: {} }, // missing title and ID
                { properties: { title: "Configuration" } }, // missing ID (will fallback to 0)
            ],
            namedRanges: [
                { name: "UsersRange", range: { sheetId: 0 } }, // Maps to the sheet that fell back to ID 0
            ],
        };

        const result = parseSpreadsheet(mockSpreadsheet, schema);

        // It mapped the "Configuration" sheet even without an ID
        expect(result.sheets["Configuration"]).toBeDefined();
        // It mapped the range because the sheet ID fell back to 0, matching the range's sheetId 0
        expect(result.mappedRanges["UsersRange"]).toBeDefined();
    });

    it("should safely ignore malformed named ranges or ranges pointing to unknown sheets", () => {
        const mockSpreadsheet: GoogleAppsScript.Sheets.Schema.Spreadsheet = {
            sheets: [{ properties: { sheetId: 1, title: "Configuration" } }],
            namedRanges: [
                {}, // Missing both name and range
                { name: "UsersRange" }, // Missing range
                { range: { sheetId: 1 } }, // Missing name
                { name: "UsersRange", range: { sheetId: 999 } }, // Points to an unmapped sheet
            ],
        };

        const result = parseSpreadsheet(mockSpreadsheet, schema);

        expect(result.sheets["Configuration"]).toBeDefined();
        // None of the ranges were valid or pointed to a valid sheet, so mappedRanges remains empty
        expect(result.mappedRanges).toEqual({});
    });

    it("should handle malformed schema definitions safely at runtime", () => {
        // We cast to `any` here to simulate runtime edge cases where a schema
        // might be improperly constructed or corrupted, bypassing TS strictness.
        const malformedSchema = {
            sheets: {
                validSheet: {
                    sheetName: "ValidSheet",
                    ranges: {
                        validRange: "ValidRange",
                        invalidRange: undefined as unknown as string,
                    },
                },
                invalidSheet: {
                    sheetName: undefined as unknown as string,
                },
            },
        } as const;

        const mockSpreadsheet: GoogleAppsScript.Sheets.Schema.Spreadsheet = {
            sheets: [{ properties: { sheetId: 1, title: "ValidSheet" } }],
            namedRanges: [{ name: "ValidRange", range: { sheetId: 1 } }],
        };

        const result = parseSpreadsheet(mockSpreadsheet, malformedSchema);

        // Verifies the function didn't crash on the undefined schema parts and processed the valid parts
        expect(result.sheets["ValidSheet"]).toBeDefined();
        expect(result.mappedRanges["ValidRange"]).toBeDefined();
    });

    it("should fallback to sheetId 0 when a named range omits the sheetId", () => {
        const fallbackSchema = {
            sheets: {
                default: {
                    sheetName: "DefaultSheet",
                    ranges: {
                        zero: "ZeroRange",
                    },
                },
            },
        } as const;

        const mockSpreadsheet: GoogleAppsScript.Sheets.Schema.Spreadsheet = {
            sheets: [
                // We create a sheet and explicitly give it the fallback ID of 0
                { properties: { sheetId: 0, title: "DefaultSheet" } },
            ],
            namedRanges: [
                {
                    name: "ZeroRange",
                    // 3. Covers Branch: namedRange.range.sheetId ?? 0
                    // The range object exists, but explicitly omits `sheetId`
                    range: { startRowIndex: 1 },
                },
            ],
        };

        const result = parseSpreadsheet(mockSpreadsheet, fallbackSchema);

        expect(result.sheets["DefaultSheet"]).toBeDefined();
        // Because the range omitted the sheetId, the ?? 0 fallback caught it
        // and successfully linked it to the sheet with ID 0.
        expect(result.mappedRanges["ZeroRange"]).toBeDefined();
        expect(result.mappedRanges["ZeroRange"]?.sheet.properties?.title).toBe("DefaultSheet");
    });
});
