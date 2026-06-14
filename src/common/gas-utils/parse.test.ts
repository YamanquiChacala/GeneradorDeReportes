import { isStrictNameRange, parseSpreadsheet } from "./parse";

describe("GAS Utils, Parse", () => {
    describe("isStrictNameRange", () => {
        it("should return true if namedRangeId, name, and range are all present", () => {
            const validRange: GoogleAppsScript.Sheets.Schema.NamedRange = {
                namedRangeId: "id-123",
                name: "TestRange",
                range: { sheetId: 1 },
            };
            expect(isStrictNameRange(validRange)).toBe(true);
        });

        it("should return false if namedRangeId is missing or null", () => {
            expect(isStrictNameRange({ name: "TestRange", range: { sheetId: 1 } })).toBe(false);
            expect(isStrictNameRange({ namedRangeId: null, name: "TestRange", range: { sheetId: 1 } } as unknown as GoogleAppsScript.Sheets.Schema.NamedRange)).toBe(
                false,
            );
        });

        it("should return false if name is missing or null", () => {
            expect(isStrictNameRange({ namedRangeId: "id-123", range: { sheetId: 1 } })).toBe(false);
            expect(isStrictNameRange({ namedRangeId: "id-123", name: null, range: { sheetId: 1 } } as unknown as GoogleAppsScript.Sheets.Schema.NamedRange)).toBe(false);
        });

        it("should return false if range is missing or null", () => {
            expect(isStrictNameRange({ namedRangeId: "id-123", name: "TestRange" })).toBe(false);
            expect(isStrictNameRange({ namedRangeId: "id-123", name: "TestRange", range: null } as unknown as GoogleAppsScript.Sheets.Schema.NamedRange)).toBe(false);
        });

        it("should return false for completely empty objects", () => {
            expect(isStrictNameRange({})).toBe(false);
        });
    });

    describe("parseSpreadsheet", () => {
        const schema = {
            sheets: {
                config: {
                    sheetName: "Configuration",
                    ranges: {
                        users: "UsersRange",
                        overlap: "student_static",
                    },
                    dynamicRanges: {
                        students: "student",
                        temps: "temp_",
                    },
                },
                logs: {
                    sheetName: "Logs",
                },
            },
        } as const;

        it("should correctly map sheets, named ranges, dynamic ranges, and catch unmapped sheets in extraSheets", () => {
            const mockSpreadsheet: GoogleAppsScript.Sheets.Schema.Spreadsheet = {
                sheets: [
                    { properties: { sheetId: 1, title: "Configuration" } },
                    { properties: { sheetId: 2, title: "SecretData" } },
                    { properties: { sheetId: 3, title: "Logs" } },
                    { properties: { sheetId: 4, title: "Archive" } },
                ],
                namedRanges: [
                    { namedRangeId: "id-1", name: "UsersRange", range: { sheetId: 1 } },
                    { namedRangeId: "id-2", name: "OtherConfigRange", range: { sheetId: 1 } },
                    { namedRangeId: "id-3", name: "HiddenRange", range: { sheetId: 2 } },

                    // Dynamic ranges
                    { namedRangeId: "id-4", name: "student00", range: { sheetId: 1 } },
                    { namedRangeId: "id-5", name: "student01", range: { sheetId: 1 } },
                    { namedRangeId: "id-6", name: "temp_A", range: { sheetId: 1 } },
                    { namedRangeId: "id-7", name: "temp_B", range: { sheetId: 1 } },

                    // Static range that happens to share a dynamic prefix
                    { namedRangeId: "id-8", name: "student_static", range: { sheetId: 1 } },
                ],
            };

            const result = parseSpreadsheet(mockSpreadsheet, schema);

            // Check strict sheets mapping
            expect(result.mappedSheets[schema.sheets.config.sheetName]).toBeDefined();
            expect(result.mappedSheets[schema.sheets.config.sheetName]?.properties?.title).toBe("Configuration");
            expect(result.mappedSheets[schema.sheets.logs.sheetName]?.properties?.title).toBe("Logs");
            expect("SecretData" in result.mappedSheets).toBe(false);

            // Check extraSheets mapping
            expect(result.extraSheets).toHaveLength(2);
            expect(result.extraSheets.map((s) => s.properties?.title)).toEqual(["SecretData", "Archive"]);

            // Check strict named ranges mapping
            expect(result.mappedRanges[schema.sheets.config.ranges.users]).toBeDefined();
            expect(result.mappedRanges[schema.sheets.config.ranges.users]?.namedRange.range.sheetId).toBe(1);
            expect("HiddenRange" in result.mappedRanges).toBe(false);
            expect("OtherConfigRange" in result.mappedRanges).toBe(false);

            // Check dynamic named ranges mapping
            expect(result.dynamicMappedRanges.student).toBeDefined();
            expect(result.dynamicMappedRanges.student?.length).toBe(2);
            expect(result.dynamicMappedRanges.student?.map((r) => r.namedRange.range.sheetId)).toEqual([1, 1]);

            expect(result.dynamicMappedRanges.temp_).toBeDefined();
            expect(result.dynamicMappedRanges.temp_?.length).toBe(2);

            // Verify Static Priority: "student_static" should be in mappedRanges, NOT in dynamicMappedRanges.students
            expect(result.mappedRanges[schema.sheets.config.ranges.overlap]).toBeDefined();
            expect(result.dynamicMappedRanges.student?.length).toBe(2);

            // Check the sheetNamedRanges array mapping
            const configSheetRanges = result.mappedSheetNamedRanges[schema.sheets.config.sheetName];
            expect(configSheetRanges).toBeDefined();
            expect(configSheetRanges?.length).toBe(7); // 2 static + 1 unmapped + 4 dynamic

            const rangeNames = configSheetRanges?.map((r) => r.name);
            expect(rangeNames).toContain("UsersRange");
            expect(rangeNames).toContain("OtherConfigRange");
            expect(rangeNames).toContain("student00");

            // Basic check for usedIds on standard sheet
            expect(result.usedIds.has(1)).toBe(true);
            expect(result.usedIds.has(2)).toBe(true);
        });

        it("should safely handle an undefined spreadsheet or missing sheets array", () => {
            const result1 = parseSpreadsheet(undefined, schema);
            expect(result1.mappedSheets).toEqual({});
            expect(result1.mappedRanges).toEqual({});
            expect(result1.mappedSheetNamedRanges).toEqual({});
            expect(result1.dynamicMappedRanges).toEqual({});
            expect(result1.extraSheets).toEqual([]);
            expect(result1.usedIds).toBeInstanceOf(Set);
            expect(result1.usedIds.size).toBe(1);

            const result2 = parseSpreadsheet({} as GoogleAppsScript.Sheets.Schema.Spreadsheet, schema);
            expect(result2.mappedSheets).toEqual({});
            expect(result2.extraSheets).toEqual([]);
            expect(result2.usedIds).toBeInstanceOf(Set);
            expect(result2.usedIds.size).toBe(1);
        });

        it("should gracefully handle a spreadsheet that lacks namedRanges", () => {
            const mockSpreadsheet: GoogleAppsScript.Sheets.Schema.Spreadsheet = {
                sheets: [{ properties: { sheetId: 1, title: "Configuration" } }],
            };

            const result = parseSpreadsheet(mockSpreadsheet, schema);

            expect(result.mappedSheets["Configuration"]).toBeDefined();
            expect(result.mappedRanges).toEqual({});
            expect(result.dynamicMappedRanges).toEqual({});
            expect(result.mappedSheetNamedRanges["Configuration"]).toEqual([]);
            expect(result.extraSheets).toEqual([]);
            expect(result.usedIds.has(1)).toBe(true);
        });

        it("should handle sheets missing properties, titles, or IDs", () => {
            const mockSpreadsheet: GoogleAppsScript.Sheets.Schema.Spreadsheet = {
                sheets: [
                    {},
                    { properties: {} },
                    { properties: { title: "Configuration" } }, // missing ID (falls back to 0)
                ],
                namedRanges: [
                    { namedRangeId: "id-1", name: "UsersRange", range: { sheetId: 0 } },
                    { namedRangeId: "id-2", name: "student00", range: { sheetId: 0 } },
                ],
            };

            const result = parseSpreadsheet(mockSpreadsheet, schema);

            expect(result.mappedSheets["Configuration"]).toBeDefined();
            expect(result.mappedRanges["UsersRange"]).toBeDefined();
            expect(result.dynamicMappedRanges["student"]).toBeDefined();
            expect(result.extraSheets).toHaveLength(0);

            // Should not have added undefined IDs to the set
            expect(result.usedIds.size).toBe(1);
        });

        it("should collect usedIds from spreadsheet metadata, sheets, protected ranges, charts, banded ranges, and filter views", () => {
            const mockSpreadsheet: GoogleAppsScript.Sheets.Schema.Spreadsheet = {
                developerMetadata: [
                    { metadataId: 111 },
                    { visibility: "DOCUMENT" }, // Missing ID
                ],
                sheets: [
                    {
                        properties: { sheetId: 10, title: "Configuration" },
                        protectedRanges: [{ protectedRangeId: 20 }],
                        charts: [{}, { chartId: 30 }],
                        bandedRanges: [{}, { bandedRangeId: 40 }],
                        filterViews: [{}, { filterViewId: 50 }],
                    },
                    {
                        properties: { sheetId: 11, title: "Logs" },
                        protectedRanges: [{ protectedRangeId: 21 }, {}], // Empty obj missing ID
                        charts: [{ chartId: 31 }],
                    },
                    {
                        // Sheet with no properties or IDs
                        properties: { title: "NoIdSheet" },
                    },
                ],
            };

            const result = parseSpreadsheet(mockSpreadsheet, schema);

            expect(result.usedIds).toBeInstanceOf(Set);

            // Expected IDs: 111 (metadata), 10, 11 (sheets), 20, 21 (protected ranges), 30, 31 (charts), 40 (banded), 50 (filter)
            // Total: 9 unique IDs
            expect(result.usedIds.size).toBe(10);

            const expectedIds = [111, 10, 11, 20, 21, 30, 31, 40, 50];
            for (const id of expectedIds) {
                expect(result.usedIds.has(id)).toBe(true);
            }
        });

        it("should safely ignore malformed named ranges or ranges pointing to unknown sheets", () => {
            const mockSpreadsheet: GoogleAppsScript.Sheets.Schema.Spreadsheet = {
                sheets: [{ properties: { sheetId: 1, title: "Configuration" } }],
                namedRanges: [
                    {}, // Completely malformed
                    { namedRangeId: "id-1", name: "UsersRange" }, // Missing range
                    { namedRangeId: "id-2", range: { sheetId: 1 } }, // Missing name
                    { name: "MissingIDRange", range: { sheetId: 1 } }, // Missing ID
                    { namedRangeId: "id-3", name: "UsersRange", range: { sheetId: 999 } }, // Unknown sheet
                    { namedRangeId: "id-4", name: "student01", range: { sheetId: 999 } }, // Unknown sheet
                ],
            };

            const result = parseSpreadsheet(mockSpreadsheet, schema);

            expect(result.mappedSheets["Configuration"]).toBeDefined();
            expect(result.mappedRanges).toEqual({});
            expect(result.dynamicMappedRanges).toEqual({});
            expect(result.extraSheets).toEqual([]);
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
                    { namedRangeId: "id-1", name: "ValidRange", range: { sheetId: 1 } },
                    { namedRangeId: "id-2", name: "valid_01", range: { sheetId: 1 } },
                ],
            };

            const result = parseSpreadsheet(mockSpreadsheet, malformedSchema);

            expect(result.mappedSheets["ValidSheet"]).toBeDefined();
            expect(result.mappedRanges["ValidRange"]).toBeDefined();
            expect(result.dynamicMappedRanges["valid_"]).toBeDefined();
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
                        namedRangeId: "id-1",
                        name: "ZeroRange",
                        range: { startRowIndex: 1 }, // Omitted sheetId
                    },
                    {
                        namedRangeId: "id-2",
                        name: "dyn_A",
                        range: { startRowIndex: 2 }, // Omitted sheetId
                    },
                ],
            };

            const result = parseSpreadsheet(mockSpreadsheet, fallbackSchema);

            expect(result.mappedSheets["DefaultSheet"]).toBeDefined();
            expect(result.mappedRanges["ZeroRange"]).toBeDefined();
            expect(result.mappedRanges["ZeroRange"]?.sheet.properties?.title).toBe("DefaultSheet");
            expect(result.dynamicMappedRanges["dyn_"]).toBeDefined();
        });
    });
});
