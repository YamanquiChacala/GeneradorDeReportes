import * as Utils from "../../../common/utils/googleAPI";

describe("Google API Utils Module", () => {
    describe("flattenFormInputs()", () => {
        it("should safely return an empty object if inputs are undefined", () => {
            const result = Utils.flattenFormInputs(undefined);
            expect(result).toEqual({});
        });

        it("should extract standard string inputs", () => {
            const mockInputs = {
                groupName: { stringInputs: { value: ["Advanced Rust"] } },
                folderId: { stringInputs: { value: ["folder-12345"] } },
            };

            const result = Utils.flattenFormInputs(mockInputs);

            expect(result).toEqual({
                groupName: "Advanced Rust",
                folderId: "folder-12345",
            });
        });

        it('should convert a single "true" string into a boolean true', () => {
            const mockInputs = {
                attendancePerClass: { stringInputs: { value: ["true"] } },
                averagePerField: { stringInputs: { value: ["true"] } },
                // A normal string that happens to be "false" should NOT become a boolean
                // because GAS switches omit the field entirely when false.
                randomText: { stringInputs: { value: ["false"] } },
            };

            const result = Utils.flattenFormInputs(mockInputs);

            expect(result).toEqual({
                attendancePerClass: true,
                averagePerField: true,
                randomText: "false",
            });
        });

        it("should return an array of strings for multiselect checkboxes", () => {
            const mockInputs = {
                selectedDays: { stringInputs: { value: ["Monday", "Wednesday", "Friday"] } },
            };

            const result = Utils.flattenFormInputs(mockInputs);

            expect(result).toEqual({
                selectedDays: ["Monday", "Wednesday", "Friday"],
            });
        });

        it("should extract and parse Date and DateTime epoch timestamps into strict numbers", () => {
            const mockInputs = {
                // Google sometimes sends timestamps as strings
                dateStart: { dateInput: { msSinceEpoch: "1680307200000" } },
                // And sometimes as raw numbers depending on the event payload
                dateEnd: { dateTimeInput: { msSinceEpoch: 1680393600000 } },
            };

            const result = Utils.flattenFormInputs(mockInputs);

            expect(result).toEqual({
                dateStart: 1680307200000,
                dateEnd: 1680393600000,
            });

            // Explicitly verify the type coercion worked
            expect(typeof result["dateStart"]).toBe("number");
            expect(typeof result["dateEnd"]).toBe("number");
        });

        it("should fallback to the raw object if the data structure is unrecognized", () => {
            const mockInputs = {
                normalString: { stringInputs: { value: ["hello"] } },
                weirdFutureGoogleInput: { weirdInput: { data: "who knows" } },
                // biome-ignore lint/suspicious/noExplicitAny: For testing purposes
                nullInput: null as any, // Testing edge case where a key exists but value is null
            };

            const result = Utils.flattenFormInputs(mockInputs);

            expect(result).toEqual({
                normalString: "hello",
                weirdFutureGoogleInput: { weirdInput: { data: "who knows" } },
            });

            // Ensure the nullInput was safely skipped
            expect(result).not.toHaveProperty("nullInput");
        });
    });

    describe("getDateMs", () => {
        /**
         * Helper to safely generate a mocked Google Sheet object with a specific value at a specific row/col
         * conforming to the GoogleAppsScript.Sheets.Schema types.
         */
        const createMockSheet = (rowIndex: number, colIndex: number, numberValue?: number): GoogleAppsScript.Sheets.Schema.Sheet => {
            const rowData: GoogleAppsScript.Sheets.Schema.RowData[] = [];

            // Build the grid structurally up to the requested row and column
            for (let i = 0; i <= rowIndex; i++) {
                const values: GoogleAppsScript.Sheets.Schema.CellData[] = [];
                if (i === rowIndex) {
                    for (let j = 0; j <= colIndex; j++) {
                        if (j === colIndex && numberValue !== undefined) {
                            values.push({ effectiveValue: { numberValue } });
                        } else {
                            values.push({});
                        }
                    }
                }
                rowData.push({ values });
            }

            return {
                data: [{ rowData }],
            } as GoogleAppsScript.Sheets.Schema.Sheet;
        };

        describe("Null / Undefined checks", () => {
            test("should return null if range is entirely missing", () => {
                expect(Utils.getDateMs(undefined, createMockSheet(0, 0, 45000))).toBeNull();
            });

            test("should return null if range is missing startRowIndex", () => {
                const range: GoogleAppsScript.Sheets.Schema.GridRange = { startColumnIndex: 1 };
                expect(Utils.getDateMs(range, createMockSheet(0, 1, 45000))).toBeNull();
            });

            test("should return null if sheet is missing", () => {
                const range: GoogleAppsScript.Sheets.Schema.GridRange = { startRowIndex: 1, startColumnIndex: 1 };
                expect(Utils.getDateMs(range, undefined)).toBeNull();
            });

            test("should return null if the specific cell does not exist in the sheet data", () => {
                const range: GoogleAppsScript.Sheets.Schema.GridRange = { startRowIndex: 5, startColumnIndex: 5 };
                const emptySheet = { data: [] } as GoogleAppsScript.Sheets.Schema.Sheet;
                expect(Utils.getDateMs(range, emptySheet)).toBeNull();
            });

            test("should return null if the cell exists but has no numberValue (e.g., text)", () => {
                const range: GoogleAppsScript.Sheets.Schema.GridRange = { startRowIndex: 1, startColumnIndex: 1 };

                // Build a strict mock that injects a string instead of a number
                const sheet = {
                    data: [
                        {
                            rowData: [
                                {}, // Row 0
                                { values: [{}, { effectiveValue: { stringValue: "Not a date" } }] }, // Row 1
                            ],
                        },
                    ],
                } as GoogleAppsScript.Sheets.Schema.Sheet;

                expect(Utils.getDateMs(range, sheet)).toBeNull();
            });
        });

        describe("Date Calculations", () => {
            // The Sheets Epoch is Dec 30, 1899 UTC
            const SHEETS_EPOCH_MS = new Date(Date.UTC(1899, 11, 30)).getTime();

            test("should correctly handle a 0-index row and column (Cell A1)", () => {
                const range: GoogleAppsScript.Sheets.Schema.GridRange = { startRowIndex: 0, startColumnIndex: 0 };
                const sheet = createMockSheet(0, 0, 1); // 1 day after epoch

                const expectedMs = SHEETS_EPOCH_MS + 1 * 24 * 60 * 60 * 1000;
                expect(Utils.getDateMs(range, sheet)).toBe(expectedMs);
            });

            test("should correctly convert a known Sheets serial date to Unix Epoch ms", () => {
                const range: GoogleAppsScript.Sheets.Schema.GridRange = { startRowIndex: 2, startColumnIndex: 3 };

                // 44927 is January 1, 2023
                const sheet = createMockSheet(2, 3, 44927);
                const expectedMs = new Date(Date.UTC(2023, 0, 1)).getTime();

                expect(Utils.getDateMs(range, sheet)).toBe(expectedMs);
            });

            test("should handle a raw number of 0 (December 30, 1899)", () => {
                const range: GoogleAppsScript.Sheets.Schema.GridRange = { startRowIndex: 1, startColumnIndex: 1 };
                const sheet = createMockSheet(1, 1, 0);

                expect(Utils.getDateMs(range, sheet)).toBe(SHEETS_EPOCH_MS);
            });

            test("should handle fractional days (date with time)", () => {
                const range: GoogleAppsScript.Sheets.Schema.GridRange = { startRowIndex: 0, startColumnIndex: 0 };

                // 44927.5 is January 1, 2023 at 12:00:00 PM
                const sheet = createMockSheet(0, 0, 44927.5);
                const expectedMs = new Date(Date.UTC(2023, 0, 1, 12, 0, 0)).getTime();

                expect(Utils.getDateMs(range, sheet)).toBe(expectedMs);
            });
        });
    });
});
