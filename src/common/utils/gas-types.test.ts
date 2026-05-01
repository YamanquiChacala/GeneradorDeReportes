import { buildFieldsMask, defineActionParameters, defineInputsSchema, defineRangesDataConfig, getInputs } from "./gas-types";
import { parseSpreadsheet } from "./mapped-name-range";

describe("googleAPI Type Utilities", () => {
    describe("defineRangesDataConfig", () => {
        it("should return the exact same configuration object passed to it", () => {
            const config = {
                myRange: { range: "A1:B2", type: "string" },
                myNumber: { range: "C1:C10", type: "number" },
            } as const;
            const result = defineRangesDataConfig(config);
            expect(result).toBe(config);
        });
    });

    describe("defineInputsSchema", () => {
        it("should return the schema and a fieldName identity function", () => {
            const schema = { name: "string", age: "number" } as const;
            const result = defineInputsSchema(schema);

            expect(result.schema).toBe(schema);
            // Verify the identity function works as expected
            expect(result.fieldName("name")).toBe("name");
            expect(result.fieldName("age")).toBe("age");
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

        it("should correctly extract a date from a dateTimeInput", () => {
            const mockFormInputs: GoogleAppsScript.Addons.CommonEventObject["formInputs"] = {
                startDate: { "": {}, dateTimeInput: { msSinceEpoch: "1700000000001", hasDate: true, hasTime: false } },
            };
            const result = getInputs(mockFormInputs, schema);
            expect(result.startDate).toBe(1700000000001);
        });

        it("should fallback to parsing an epoch string if explicit date inputs are missing", () => {
            const mockFormInputs: GoogleAppsScript.Addons.CommonEventObject["formInputs"] = {
                startDate: { "": {}, stringInputs: { value: ["1700000000002"] } },
            };
            const result = getInputs(mockFormInputs, schema);
            expect(result.startDate).toBe(1700000000002);
        });

        it("should fallback to parsing a standard date string if explicit date inputs are missing", () => {
            const mockFormInputs: GoogleAppsScript.Addons.CommonEventObject["formInputs"] = {
                startDate: { "": {}, stringInputs: { value: ["2024-01-01"] } },
            };
            const result = getInputs(mockFormInputs, schema);
            const expectedTime = new Date("2024-01-01").getTime();
            expect(result.startDate).toBe(expectedTime);
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

        it("should treat missing fields as boolean false", () => {
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

    describe("buildFieldsMask", () => {
        // Mock type to test field paths without needing the full GAS Sheet schema
        type MockGASSchema = {
            spreadsheetId: string;
            properties: { title: string; locale?: string };
            sheets: Array<{ properties: { sheetId: number; title: string } }>;
        };

        it("should join valid property paths with commas", () => {
            const mask = buildFieldsMask<MockGASSchema>("spreadsheetId", "properties.title", "sheets.properties.sheetId");
            expect(mask).toBe("spreadsheetId,properties.title,sheets.properties.sheetId");
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
            expect(result.sheets[schema.sheets.config.sheetName]).toBeDefined();
            expect(result.sheets[schema.sheets.config.sheetName]?.properties?.title).toBe("Configuration");
            expect("SecretData" in result.sheets).toBe(false);

            // Check ranges
            expect(result.namedRanges[schema.sheets.config.ranges.users]).toBeDefined();
            expect(result.namedRanges[schema.sheets.config.ranges.users]?.range.sheetId).toBe(1);
            expect("HiddenRange" in result.namedRanges).toBe(false);
        });
    });
});
