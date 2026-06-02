import { defineActionParameters, defineInputsSchema, getInputs } from ".";

describe("googleAPI Type Utilities", () => {
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

        it("should skip fields that are missing or undefined in rawParams", () => {
            const parsed = actions.parse({ name: "test" }); // id and active are missing
            expect(parsed.id).toBeUndefined();
            expect(parsed.active).toBeUndefined();
            expect(parsed.name).toBe("test");
        });

        it("should ignore properties that fail to parse as valid numbers", () => {
            const parsed = actions.parse({ id: "not-a-number", name: "test" });
            expect(parsed.id).toBeUndefined();
            expect(parsed.name).toBe("test");
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
            const mockFormInputs: GoogleAppsScript.Addons.CommonEventObject["formInputs"] = {};
            const result1 = getInputs(mockFormInputs, schema);
            expect(result1.isAdmin).toBe(false);
        });

        it("should treat 'false' as boolean false", () => {
            const mockFormInputs: GoogleAppsScript.Addons.CommonEventObject["formInputs"] = {
                isAdmin: { "": {}, stringInputs: { value: ["false"] } },
            };
            const result1 = getInputs(mockFormInputs, schema);
            expect(result1.isAdmin).toBe(false);
        });

        it("should safely ignore date fields completely missing explicit inputs and fallbacks", () => {
            const mockFormInputs: GoogleAppsScript.Addons.CommonEventObject["formInputs"] = {
                startDate: { "": {} }, // No dateInput, dateTimeInput, or stringInputs
            };
            const result = getInputs(mockFormInputs, schema);
            expect(result.startDate).toBeUndefined();
        });

        it("should safely ignore time fields missing the timeInput property", () => {
            const mockFormInputs: GoogleAppsScript.Addons.CommonEventObject["formInputs"] = {
                startTime: { "": {} }, // Explicitly missing timeInput
            };
            const result = getInputs(mockFormInputs, schema);
            expect(result.startTime).toBeUndefined();
        });

        it("should skip text, number, and array fields when stringInputs.value is empty", () => {
            const mockFormInputs: GoogleAppsScript.Addons.CommonEventObject["formInputs"] = {
                username: { "": {}, stringInputs: { value: [] } },
                age: { "": {}, stringInputs: { value: [] } },
                selections: { "": {} },
            };
            const result = getInputs(mockFormInputs, schema);
            expect(result.username).toBeUndefined();
            expect(result.age).toBeUndefined();
            expect(result.selections).toBeUndefined();
        });

        it("should fallback to empty string and handle undefined string inputs", () => {
            const mockFormInputs: GoogleAppsScript.Addons.CommonEventObject["formInputs"] = {
                // By forcing undefined into the array, we force evaluation of the `?? ""` and `?? []` fallbacks.
                // biome-ignore lint/suspicious/noExplicitAny: Simulating malformed Json
                username: { "": {}, stringInputs: { value: [undefined as any] } },
            };
            const result = getInputs(mockFormInputs, schema);
            expect(result.username).toBe("");
        });
    });
});
