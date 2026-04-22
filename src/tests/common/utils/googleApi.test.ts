import { defineActionParameters, getInputs, MappedNamedRange, PasteType, parseSpreadsheet } from "../../../common/utils/googleAPI";

describe("googleAPI Utilities", () => {
    describe("MappedNamedRange", () => {
        // Mock a basic named range with some dummy data
        const mockMappedRange: MappedNamedRange = {
            range: { sheetId: 1, startRowIndex: 0, endRowIndex: 5, startColumnIndex: 0, endColumnIndex: 5 },
            sheet: {
                data: [
                    {
                        rowData: [
                            { values: [{ formattedValue: "Test String", effectiveValue: { numberValue: 42 } }] }, // Row 0, Col 0
                            { values: [{ effectiveValue: { numberValue: 0 } }] }, // Row 1, Col 0 (For epoch test)
                        ],
                    },
                ],
            },
        };

        it("getCellData should return data if within bounds", () => {
            const data = MappedNamedRange.getCellData({ mappedRange: mockMappedRange, row: 0, column: 0 });
            expect(data?.formattedValue).toBe("Test String");
        });

        it("getCellData should return undefined if out of bounds", () => {
            const data = MappedNamedRange.getCellData({ mappedRange: mockMappedRange, row: 10, column: 0 });
            expect(data).toBeUndefined();
        });

        it("getCellUnixEpoch should correctly calculate the 1899 Sheets epoch", () => {
            // Sheets epoch (0) is Dec 30, 1899
            const expectedEpoch = new Date(Date.UTC(1899, 11, 30)).getTime();
            const epoch = MappedNamedRange.getCellUnixEpoch({ mappedRange: mockMappedRange, row: 1, column: 0 });
            expect(epoch).toBe(expectedEpoch);
        });

        it("buildCopyPasteRequest should return a valid request object on success", () => {
            // We use the mockMappedRange defined earlier in the file
            // It has a startRowIndex of 0, endRowIndex of 5, startColumnIndex of 0, endColumnIndex of 5
            const req = MappedNamedRange.buildCopyPasteRequest({
                mappedRange: mockMappedRange,
                destinationSheetId: 99,
                destinationStartRow: 10,
                destinationStartColumn: 20,
                pasteType: PasteType.PASTE_VALUES,
                height: 3,
                width: 2,
            });

            // 1. Assert it didn't fail and return undefined
            expect(req).toBeDefined();

            // 2. Assert the top-level key exists
            expect(req?.copyPaste).toBeDefined();

            // 3. Assert the Source range calculated correctly
            expect(req?.copyPaste?.source?.sheetId).toBe(1);
            expect(req?.copyPaste?.source?.startRowIndex).toBe(0); // offset defaults to 0
            expect(req?.copyPaste?.source?.endRowIndex).toBe(3); // start + height
            expect(req?.copyPaste?.source?.startColumnIndex).toBe(0);
            expect(req?.copyPaste?.source?.endColumnIndex).toBe(2);

            // 4. Assert the Destination range calculated correctly
            expect(req?.copyPaste?.destination?.sheetId).toBe(99);
            expect(req?.copyPaste?.destination?.startRowIndex).toBe(10);
            expect(req?.copyPaste?.destination?.endRowIndex).toBe(13); // 10 + 3
            expect(req?.copyPaste?.destination?.startColumnIndex).toBe(20);
            expect(req?.copyPaste?.destination?.endColumnIndex).toBe(22); // 20 + 2

            // 5. Assert the formatting enums passed through
            expect(req?.copyPaste?.pasteType).toBe(PasteType.PASTE_VALUES);
            expect(req?.copyPaste?.pasteOrientation).toBe("NORMAL");
        });

        it("buildCopyPasteRequest should return undefined if mappedRange is missing", () => {
            const req = MappedNamedRange.buildCopyPasteRequest({
                mappedRange: undefined,
                destinationSheetId: 2,
                destinationStartRow: 0,
                destinationStartColumn: 0,
                pasteType: PasteType.PASTE_NORMAL,
            });
            expect(req).toBeUndefined();
        });

        it("buildCopyPasteRequest should return undefined if requested range is out of bounds", () => {
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

    describe("defineActionParameters", () => {
        const schema = { id: "number", active: "boolean", name: "string" } as const;
        const actions = defineActionParameters(schema);

        it("should correctly build string records from typed params", () => {
            const built = actions.build({ id: 123, active: true, name: "test" });
            expect(built).toEqual({ id: "123", active: "true", name: "test" });
        });

        it("should correctly parse raw strings back to typed params", () => {
            const parsed = actions.parse({ id: "123", active: "true", name: "test", ignoredField: "hello" });
            // Should cast types and ignore fields not in schema
            expect(parsed).toEqual({ id: 123, active: true, name: "test" });
        });

        it("should handle undefined inputs gracefully", () => {
            const parsed = actions.parse(undefined);
            expect(parsed).toEqual({});
        });
    });

    describe("getInputs", () => {
        const schema = {
            username: "string",
            age: "number",
            isAdmin: "boolean",
            startDate: "date",
            startTime: "time",
            selections: "array",
        } as const;

        it("should correctly extract and cast standard form inputs", () => {
            const mockFormInputs: GoogleAppsScript.Addons.CommonEventObject["formInputs"] = {
                username: { "": {}, stringInputs: { value: ["chacala"] } },
                age: { "": {}, stringInputs: { value: ["30"] } },
                isAdmin: { "": {}, stringInputs: { value: ["true"] } },
                startDate: { "": {}, dateInput: { msSinceEpoch: "1700000000000" } },
                startTime: { "": {}, timeInput: { hours: 4, minutes: 45 } },
                selections: { "": {}, stringInputs: { value: ["uno", "dos", "tres"] } },
            };

            const result = getInputs(mockFormInputs, schema);
            expect(result).toEqual({
                username: "chacala",
                age: 30,
                isAdmin: true,
                startDate: 1700000000000,
                startTime: { hours: 4, minutes: 45 },
                selections: ["uno", "dos", "tres"],
            });
        });

        it("should ignore malformed numbers and fallback dates safely", () => {
            const mockFormInputs: GoogleAppsScript.Addons.CommonEventObject["formInputs"] = {
                age: { "": {}, stringInputs: { value: ["not-a-number"] } },
                startDate: { "": {}, stringInputs: { value: ["not-a-date"] } },
            };

            const result = getInputs(mockFormInputs, schema);
            // Neither should exist in the result because they failed parsing safely
            expect(result.age).toBeUndefined();
            expect(result.startDate).toBeUndefined();
        });

        it("should treat empty string as boolean false", () => {
            // Empty string (default value) gets removed.
            const mockFormInputs: GoogleAppsScript.Addons.CommonEventObject["formInputs"] = {};
            const result1 = getInputs(mockFormInputs, schema);
            expect(result1.isAdmin).toBe(false);
        });

        it("should treat 'false' as boolean false", () => {
            const mockFormInputs: GoogleAppsScript.Addons.CommonEventObject["formInputs"] = {
                isAdmin: { "": {}, stringInputs: { value: ["false"] } }, // explicit false
            };
            const result1 = getInputs(mockFormInputs, schema);
            expect(result1.isAdmin).toBe(false);
        });
    });

    describe("parseSpreadsheet", () => {
        const schema = {
            sheetNames: { config: "Configuration" },
            namedRanges: { users: "UsersRange" },
        } as const;

        it("should correctly map allowed sheets and ranges and ignore others", () => {
            const mockSpreadsheet: GoogleAppsScript.Sheets.Schema.Spreadsheet = {
                sheets: [
                    { properties: { sheetId: 1, title: "Configuration" } },
                    { properties: { sheetId: 2, title: "SecretData" } }, // Should be ignored
                ],
                namedRanges: [
                    { name: "UsersRange", range: { sheetId: 1 } },
                    { name: "HiddenRange", range: { sheetId: 2 } }, // Should be ignored
                ],
            };

            const result = parseSpreadsheet(mockSpreadsheet, schema);

            // Check sheets
            expect(result.sheets[schema.sheetNames.config]).toBeDefined();
            expect(result.sheets[schema.sheetNames.config]?.properties?.title).toBe("Configuration");
            expect("SecretData" in result.sheets).toBe(false);

            // Check ranges
            expect(result.namedRanges[schema.namedRanges.users]).toBeDefined();
            expect(result.namedRanges[schema.namedRanges.users]?.range.sheetId).toBe(1);
            expect("HiddenRange" in result.namedRanges).toBe(false);
        });
    });
});
