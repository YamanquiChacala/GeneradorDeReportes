import { buildFieldsMask, createRequiredGetter, makeUserEntered } from "./helpers";

describe("GAS Utils, Helpers", () => {
    describe("makeUserEntered", () => {
        it("should overwrite userEntered values with effective values and delete effectives", () => {
            const data: GoogleAppsScript.Sheets.Schema.CellData[][] = [
                [
                    {
                        userEnteredValue: { stringValue: "old" },
                        effectiveValue: { stringValue: "new" },
                        effectiveFormat: { backgroundColor: { red: 1 } },
                        note: "hello",
                    },
                ],
            ];

            const result = makeUserEntered(data, false);
            // biome-ignore lint/style/noNonNullAssertion: Test setup guarantees index 0 exists
            const firstCell = result[0]![0]!;

            expect(firstCell.userEnteredValue).toEqual({ stringValue: "new" });
            expect(firstCell.userEnteredFormat).toEqual({ backgroundColor: { red: 1 } });
            expect(firstCell.note).toBe("hello");
            expect(firstCell.effectiveValue).toBeUndefined();
            expect(firstCell.effectiveFormat).toBeUndefined();
        });

        it("should not mutate the original input array or its objects", () => {
            const originalData: GoogleAppsScript.Sheets.Schema.CellData[][] = [
                [
                    {
                        userEnteredValue: { stringValue: "old" },
                        effectiveValue: { stringValue: "new" },
                    },
                ],
            ];

            makeUserEntered(originalData, false);

            // The original object should remain completely untouched
            // biome-ignore lint/style/noNonNullAssertion: Test setup guarantees index 0 exists
            const originalCell = originalData[0]![0]!;
            expect(originalCell.effectiveValue).toEqual({ stringValue: "new" });
            expect(originalCell.userEnteredValue).toEqual({ stringValue: "old" });
        });

        it("should strip out all other properties if stripOthers is true", () => {
            const data: GoogleAppsScript.Sheets.Schema.CellData[][] = [
                [
                    {
                        effectiveValue: { numberValue: 42 },
                        note: "should be stripped",
                        hyperlink: "http://example.com",
                    },
                ],
            ];

            const result = makeUserEntered(data, true);
            // biome-ignore lint/style/noNonNullAssertion: Test setup guarantees index 0 exists
            const firstCell = result[0]![0]!;

            expect(firstCell.userEnteredValue).toEqual({ numberValue: 42 });
            expect(firstCell.note).toBeUndefined();
            expect(firstCell.hyperlink).toBeUndefined();

            // Explicitly verify the object only has exactly the keys we want
            expect(Object.keys(firstCell)).toEqual(["userEnteredValue"]);
        });

        it("should default stripOthers to false when a second argument is not provided", () => {
            const data: GoogleAppsScript.Sheets.Schema.CellData[][] = [
                [
                    {
                        effectiveValue: { numberValue: 1 },
                        note: "keep me",
                    },
                ],
            ];

            const result = makeUserEntered(data); // No second argument

            expect(result[0]?.[0]?.userEnteredValue).toEqual({ numberValue: 1 });
            expect(result[0]?.[0]?.note).toBe("keep me");
        });

        it("should retain userEntered properties if effective properties are undefined", () => {
            const data: GoogleAppsScript.Sheets.Schema.CellData[][] = [
                [
                    {
                        userEnteredValue: { stringValue: "fallback string" },
                        userEnteredFormat: { textFormat: { bold: true } },
                    },
                ],
            ];

            const result = makeUserEntered(data, false);
            const cell = result[0]?.[0];

            expect(cell?.userEnteredValue).toEqual({ stringValue: "fallback string" });
            expect(cell?.userEnteredFormat).toEqual({ textFormat: { bold: true } });
        });

        it("should skip assigning userEntered properties if both effective and userEntered are undefined", () => {
            const data: GoogleAppsScript.Sheets.Schema.CellData[][] = [
                [
                    {
                        note: "just a note, no values or formats",
                    },
                ],
            ];

            const result = makeUserEntered(data, false);
            const cell = result[0]?.[0];

            expect(cell?.userEnteredValue).toBeUndefined();
            expect(cell?.userEnteredFormat).toBeUndefined();
            expect(cell?.note).toBe("just a note, no values or formats");
        });

        it("should handle empty data arrays gracefully", () => {
            expect(makeUserEntered([])).toEqual([]);
            expect(makeUserEntered([[]])).toEqual([[]]);
        });
    });

    describe("createRequiredGetter", () => {
        describe("when the key exists", () => {
            it("returns the corresponding value", () => {
                const getter = createRequiredGetter({ a: 1, b: 2 });
                expect(getter("a")).toBe(1);
                expect(getter("b")).toBe(2);
            });

            it("does not throw for falsy values like 0, false, or empty strings", () => {
                const numberGetter = createRequiredGetter({ zero: 0 });
                const booleanGetter = createRequiredGetter({ falseValue: false });
                const stringGetter = createRequiredGetter({ emptyString: "" });

                expect(numberGetter("zero")).toBe(0);
                expect(booleanGetter("falseValue")).toBe(false);
                expect(stringGetter("emptyString")).toBe("");
            });

            it("does not throw for explicit null values", () => {
                const getter = createRequiredGetter({ nullValue: null });
                expect(getter("nullValue")).toBeNull();
            });

            it("supports number keys", () => {
                const getter = createRequiredGetter<number, string>({ 1: "one", 2: "two" });
                expect(getter(1)).toBe("one");
            });

            it("supports symbol keys", () => {
                const key = Symbol("test");
                const getter = createRequiredGetter({ [key]: 123 });
                expect(getter(key)).toBe(123);
            });
        });

        describe("when the key does not exist or is undefined", () => {
            it("throws with the default message when no context is provided", () => {
                const getter = createRequiredGetter<string, number>({ a: 1 });
                expect(() => getter("b")).toThrow("Falta propidad: b");
            });

            it("throws with the contextualized message", () => {
                const getter = createRequiredGetter<string, number>({ a: 1 }, "rango");
                expect(() => getter("b")).toThrow("Falta rango: b");
            });

            it("throws if the key exists but is explicitly set to undefined", () => {
                const getter = createRequiredGetter({ explicitUndefined: undefined });
                expect(() => getter("explicitUndefined")).toThrow("Falta propidad: explicitUndefined");
            });

            it("includes symbol keys in the error message", () => {
                const key = Symbol("missing");
                const getter = createRequiredGetter<symbol, number>({}, "symbol key");
                expect(() => getter(key)).toThrow("Falta symbol key: Symbol(missing)");
            });
        });
    });

    describe("buildFieldsMask", () => {
        type MockGASSchema = {
            spreadsheetId: string;
            properties: { title: string; locale?: string };
            sheets: Array<{ properties: { sheetId: number; title: string } }>;
        };

        it("should join valid property paths with commas", () => {
            const mask = buildFieldsMask<MockGASSchema>("spreadsheetId", "properties.title", "sheets.properties.sheetId");
            expect(mask).toBe("spreadsheetId,properties.title,sheets.properties.sheetId");
        });

        it("should handle a single path without adding trailing commas", () => {
            const mask = buildFieldsMask<MockGASSchema>("spreadsheetId");
            expect(mask).toBe("spreadsheetId");
        });

        it("should return an empty string when no paths are provided", () => {
            const mask = buildFieldsMask<MockGASSchema>();
            expect(mask).toBe("");
        });
    });
});
