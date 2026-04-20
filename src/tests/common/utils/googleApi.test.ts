import * as Utils from "../../../common/utils/googleAPI";

describe("Google API Utils Module", () => {
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
