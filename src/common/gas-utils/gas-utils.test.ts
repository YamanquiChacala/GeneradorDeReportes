import { Dimension, PasteType } from "../gas-enums";
import { getEpochDate, getSheetsDate, makeUserEntered } from ".";
import { colorToHex } from "./color";
import { getColumnLetter } from "./notation";
import { changeGridRangeSheet, createSingleCellRange, offsetGridRange } from "./range";
import { buildCopyPasteRequest, buildTransferRequestsBackup } from "./request";

describe("googleAPI Utilities", () => {
    describe("Colors", () => {
        describe("colorToHex", () => {
            it("should convert an RGB color to hex correctly", () => {
                const color: GoogleAppsScript.Sheets.Schema.Color = {
                    red: 1, // 255
                    green: 0.50196, // ~128
                    blue: 0, // 0
                };
                expect(colorToHex(color)).toBe("#FF8000");
            });

            it("should treat missing channels as 0", () => {
                const color: GoogleAppsScript.Sheets.Schema.Color = { red: 1 };
                expect(colorToHex(color)).toBe("#FF0000");
            });

            it("should fallback to the provided fallback if color is undefined", () => {
                expect(colorToHex(undefined, "#000000")).toBe("#000000");
                expect(colorToHex(undefined)).toBe("#FFFFFF"); // Default fallback
            });
        });
    });

    describe("Notation", () => {
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

    describe("Range", () => {
        describe("createSingleCellRange", () => {
            it("should create a 1x1 grid range correctly", () => {
                const range = createSingleCellRange(10, 5, 2);
                expect(range).toEqual({
                    sheetId: 10,
                    startRowIndex: 5,
                    startColumnIndex: 2,
                    endRowIndex: 6,
                    endColumnIndex: 3,
                });
            });
        });

        describe("changeGridRangeSheet", () => {
            it("should return a new grid range with an updated sheetId", () => {
                const original: GoogleAppsScript.Sheets.Schema.GridRange = {
                    sheetId: 1,
                    startRowIndex: 0,
                    endRowIndex: 10,
                };
                const updated = changeGridRangeSheet(original, 99);

                expect(updated.sheetId).toBe(99);
                expect(updated.startRowIndex).toBe(0);
                expect(updated.endRowIndex).toBe(10);

                // Should not mutate original
                expect(original.sheetId).toBe(1);
            });
        });

        describe("offsetGridRange", () => {
            const origin: GoogleAppsScript.Sheets.Schema.GridRange = {
                sheetId: 1,
                startRowIndex: 2,
                endRowIndex: 4,
                startColumnIndex: 2,
                endColumnIndex: 4,
            };

            it("should offset the range by rows and cols and maintain size if dimensions aren't overridden", () => {
                const result = offsetGridRange({ origin, rowOffset: 2, colOffset: 1 });
                expect(result).toEqual({
                    sheetId: 1,
                    startRowIndex: 4,
                    endRowIndex: 6,
                    startColumnIndex: 3,
                    endColumnIndex: 5,
                });
            });

            it("should bound the range to a new size if positive height and width are provided", () => {
                const result = offsetGridRange({ origin, rowOffset: 0, colOffset: 0, height: 10, width: 5 });
                expect(result).toEqual({
                    sheetId: 1,
                    startRowIndex: 2,
                    endRowIndex: 12,
                    startColumnIndex: 2,
                    endColumnIndex: 7,
                });
            });

            it("should safely default missing origin indexes to 0", () => {
                const emptyOrigin: GoogleAppsScript.Sheets.Schema.GridRange = { sheetId: 1 };
                const result = offsetGridRange({ origin: emptyOrigin, rowOffset: 1, colOffset: 1, height: 1, width: 1 });
                expect(result).toEqual({
                    sheetId: 1,
                    startRowIndex: 1,
                    endRowIndex: 2,
                    startColumnIndex: 1,
                    endColumnIndex: 2,
                });
            });

            it("should prevent startRowIndex and startColumnIndex from becoming negative", () => {
                const result = offsetGridRange({ origin, rowOffset: -5, colOffset: -5 });
                expect(result).toEqual({
                    sheetId: 1,
                    startRowIndex: 0,
                    endRowIndex: 0,
                    startColumnIndex: 0,
                    endColumnIndex: 0,
                });
            });

            it("should un-bound the range if height and width are negative", () => {
                const result = offsetGridRange({ origin, rowOffset: 1, colOffset: 1, height: -1, width: -1 });
                expect(result).toEqual({
                    sheetId: 1,
                    startRowIndex: 3,
                    startColumnIndex: 3,
                });
                expect(result.endRowIndex).toBeUndefined();
                expect(result.endColumnIndex).toBeUndefined();
            });

            it("should handle un-bounding only one dimension (e.g. height) while maintaining original width", () => {
                const result = offsetGridRange({ origin, rowOffset: 1, colOffset: 1, height: -1 });
                expect(result).toEqual({
                    sheetId: 1,
                    startRowIndex: 3,
                    startColumnIndex: 3,
                    endColumnIndex: 5,
                });
                expect(result.endRowIndex).toBeUndefined();
            });
        });
    });

    describe("Request", () => {
        describe("buildCopyPasteRequest", () => {
            it("should return undefined if source or destination are missing", () => {
                const range = { sheetId: 1, startRowIndex: 0 };
                expect(buildCopyPasteRequest(undefined, range, PasteType.PASTE_NORMAL)).toBeUndefined();
                expect(buildCopyPasteRequest(range, undefined, PasteType.PASTE_NORMAL)).toBeUndefined();
            });

            it("should build a valid copyPaste request", () => {
                const source = { sheetId: 1, startRowIndex: 0 };
                const destination = { sheetId: 2, startRowIndex: 5 };
                const request = buildCopyPasteRequest(source, destination, PasteType.PASTE_FORMAT);

                expect(request?.copyPaste).toBeDefined();
                expect(request?.copyPaste?.source).toBe(source);
                expect(request?.copyPaste?.destination).toBe(destination);
                expect(request?.copyPaste?.pasteType).toBe(PasteType.PASTE_FORMAT);
            });
        });

        describe("buildTransferRequest", () => {
            // Setup a 2x3 (asymmetrical) destination range
            const destination: GoogleAppsScript.Sheets.Schema.GridRange = {
                sheetId: 1,
                startRowIndex: 0,
                endRowIndex: 2,
                startColumnIndex: 0,
                endColumnIndex: 3,
            };

            // 2x3 Asymmetrical Data
            const data: GoogleAppsScript.Sheets.Schema.CellData[][] = [
                [{ userEnteredValue: { stringValue: "R1-C1" } }, { userEnteredValue: { stringValue: "R1-C2" } }, { userEnteredValue: { stringValue: "R1-C3" } }],
                [{ userEnteredValue: { stringValue: "R2-C1" } }, { userEnteredValue: { stringValue: "R2-C2" } }, { userEnteredValue: { stringValue: "R2-C3" } }],
            ];

            it("should return an updateCells request with correctly mapped values when adaptRange is false", () => {
                const requests = buildTransferRequestsBackup({ destination, data, fields: "userEnteredValue", adaptRange: false });
                expect(requests.length).toBe(1);

                const updateReq = requests[0]?.updateCells;
                expect(updateReq?.range).toEqual(destination);
                expect(updateReq?.fields).toBe("userEnteredValue");

                // DEEP ASSERTION: Check that both the shape AND the values were mapped correctly
                expect(updateReq?.rows).toEqual([
                    {
                        values: [
                            { userEnteredValue: { stringValue: "R1-C1" } },
                            { userEnteredValue: { stringValue: "R1-C2" } },
                            { userEnteredValue: { stringValue: "R1-C3" } },
                        ],
                    },
                    {
                        values: [
                            { userEnteredValue: { stringValue: "R2-C1" } },
                            { userEnteredValue: { stringValue: "R2-C2" } },
                            { userEnteredValue: { stringValue: "R2-C3" } },
                        ],
                    },
                ]);
            });

            it("should return no requests if data is empty and adaptRange is true", () => {
                const requests = buildTransferRequestsBackup({ destination, data: [], fields: "userEnteredValue", adaptRange: true });
                expect(requests.length).toBe(0);
            });

            it("should return no requests if destination is undefined", () => {
                const requests = buildTransferRequestsBackup({ destination: undefined, data, fields: "userEnteredValue", adaptRange: true });
                expect(requests.length).toBe(0);
            });

            it("should insert dimensions and map correct values if data is larger than destination range and adaptRange is true", () => {
                // Dest is 1x1, but our Data is 2x3
                const smallDest: GoogleAppsScript.Sheets.Schema.GridRange = {
                    sheetId: 1,
                    startRowIndex: 0,
                    endRowIndex: 1,
                    startColumnIndex: 0,
                    endColumnIndex: 1,
                };

                const requests = buildTransferRequestsBackup({ destination: smallDest, data, fields: "userEnteredValue", adaptRange: true });

                // Expected: 1 insertRows + 1 insertCols + 1 updateCells = 3 requests
                expect(requests.length).toBe(3);

                // Assert Row Insertion
                expect(requests[0]?.insertDimension?.range?.dimension).toBe(Dimension.ROWS);
                expect(requests[0]?.insertDimension?.range?.startIndex).toBe(1);
                expect(requests[0]?.insertDimension?.range?.endIndex).toBe(2);

                // Assert Col Insertion
                expect(requests[1]?.insertDimension?.range?.dimension).toBe(Dimension.COLUMNS);
                expect(requests[1]?.insertDimension?.range?.startIndex).toBe(1);
                expect(requests[1]?.insertDimension?.range?.endIndex).toBe(3);

                // DEEP ASSERTION: Validate the content payload survived the mapping logic
                const updateReq = requests[2]?.updateCells;
                expect(updateReq?.rows).toEqual([
                    {
                        values: [
                            { userEnteredValue: { stringValue: "R1-C1" } },
                            { userEnteredValue: { stringValue: "R1-C2" } },
                            { userEnteredValue: { stringValue: "R1-C3" } },
                        ],
                    },
                    {
                        values: [
                            { userEnteredValue: { stringValue: "R2-C1" } },
                            { userEnteredValue: { stringValue: "R2-C2" } },
                            { userEnteredValue: { stringValue: "R2-C3" } },
                        ],
                    },
                ]);
            });
        });
    });

    describe("Time", () => {
        describe("getSheetsDate", () => {
            it("converts Unix epoch 0 (Jan 1, 1970) to the Sheets epoch offset", () => {
                expect(getSheetsDate(0)).toBe(25569);
            });

            it("converts exactly one day after Unix epoch to the correct Sheets date", () => {
                const oneDayInMs = 86400000;
                expect(getSheetsDate(oneDayInMs)).toBe(25570);
            });

            it("converts a known modern Unix epoch to the correct Sheets date", () => {
                // Jan 1, 2024 = 1704067200000 ms
                expect(getSheetsDate(1704067200000)).toBe(45292);
            });

            it("handles negative Unix epochs (dates before 1970)", () => {
                // Dec 30, 1899 = -2209161600000 ms (Sheets Epoch 0)
                expect(getSheetsDate(-2209161600000)).toBe(0);
            });
        });

        describe("getEpochDate", () => {
            it("converts the Sheets epoch offset back to Unix epoch 0", () => {
                expect(getEpochDate(25569)).toBe(0);
            });

            it("converts exactly one day after the Sheets epoch offset to one day in ms", () => {
                const oneDayInMs = 86400000;
                expect(getEpochDate(25570)).toBe(oneDayInMs);
            });

            it("converts a known modern Sheets date to the correct Unix epoch", () => {
                // Jan 1, 2024 = 45292 in Sheets
                expect(getEpochDate(45292)).toBe(1704067200000);
            });

            it("handles Sheets date 0 (Dec 30, 1899)", () => {
                expect(getEpochDate(0)).toBe(-2209161600000);
            });
        });

        describe("Bidirectional Conversion", () => {
            it("returns the exact original Unix epoch after converting to Sheets and back", () => {
                const originalUnixEpoch = 1718323200000; // Arbitrary timestamp
                const sheetsDate = getSheetsDate(originalUnixEpoch);
                const convertedBack = getEpochDate(sheetsDate);

                expect(convertedBack).toBe(originalUnixEpoch);
            });

            it("returns the exact original Sheets date after converting to Unix and back", () => {
                const originalSheetsDate = 48000; // Arbitrary future Sheets date
                const unixEpoch = getEpochDate(originalSheetsDate);
                const convertedBack = getSheetsDate(unixEpoch);

                expect(convertedBack).toBe(originalSheetsDate);
            });
        });
    });

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
            const firstRow = result[0];
            const firstCell = firstRow ? firstRow[0] : undefined;

            expect(firstCell).toBeDefined();
            expect(firstCell?.userEnteredValue).toEqual({ stringValue: "new" });
            expect(firstCell?.userEnteredFormat).toEqual({ backgroundColor: { red: 1 } });
            expect(firstCell?.note).toBe("hello");
            expect(firstCell?.effectiveValue).toBeUndefined();
            expect(firstCell?.effectiveFormat).toBeUndefined();
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
            const firstRow = result[0];
            const firstCell = firstRow ? firstRow[0] : undefined;

            expect(firstCell).toBeDefined();
            expect(firstCell?.userEnteredValue).toEqual({ numberValue: 42 });
            expect(firstCell?.note).toBeUndefined();
            expect(firstCell?.hyperlink).toBeUndefined();
        });
    });
});
