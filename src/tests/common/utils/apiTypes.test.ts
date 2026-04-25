import { defineActionParameters, getInputs } from "../../../common/utils/apiTypes";
import { getColumnLetter, parseSpreadsheet } from "../../../common/utils/mappedNameRange";

describe("googleAPI Utilities", () => {
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
