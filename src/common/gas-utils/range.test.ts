import { changeGridRangeSheet, createRange, createSingleCellRange, getRangeHeight, getRangeWidth, offsetGridRange, shrinkRangeWidth } from ".";

describe("Range", () => {
    describe("createRange", () => {
        it("should create an unbound range with no width or height", () => {
            const range = createRange(10, 5, 2);
            expect(range).toEqual({
                sheetId: 10,
                startRowIndex: 5,
                startColumnIndex: 2,
            });
        });

        it("should create a bounded range when height and width are provided", () => {
            const range = createRange(10, 5, 2, 3, 4);
            expect(range).toEqual({
                sheetId: 10,
                startRowIndex: 5,
                startColumnIndex: 2,
                endRowIndex: 8,
                endColumnIndex: 6,
            });
        });
    });

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
            const result = offsetGridRange({ origin, height: 10, width: 5 });
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
                endRowIndex: 0,
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

    describe("getRangeWidth", () => {
        it("should correctly calculate the width of a bounded range", () => {
            const range: GoogleAppsScript.Sheets.Schema.GridRange = { sheetId: 1, startColumnIndex: 2, endColumnIndex: 5 };
            expect(getRangeWidth(range)).toBe(3);
        });

        it("should default to 0 if end bounds are missing", () => {
            const range: GoogleAppsScript.Sheets.Schema.GridRange = { sheetId: 1, startColumnIndex: 2 };
            expect(getRangeWidth(range)).toBe(0);
        });

        it("should default startColumnIndex to 0 if missing", () => {
            const range: GoogleAppsScript.Sheets.Schema.GridRange = { sheetId: 1, endColumnIndex: 4 };
            expect(getRangeWidth(range)).toBe(4);
        });

        it("should return 0 if the start index is greater than the end index", () => {
            const range: GoogleAppsScript.Sheets.Schema.GridRange = { sheetId: 1, startColumnIndex: 5, endColumnIndex: 2 };
            expect(getRangeWidth(range)).toBe(0);
        });
    });

    describe("getRangeHeight", () => {
        it("should correctly calculate the height of a bounded range", () => {
            const range: GoogleAppsScript.Sheets.Schema.GridRange = { sheetId: 1, startRowIndex: 2, endRowIndex: 6 };
            expect(getRangeHeight(range)).toBe(4);
        });

        it("should default to 0 if end bounds are missing", () => {
            const range: GoogleAppsScript.Sheets.Schema.GridRange = { sheetId: 1, startRowIndex: 2 };
            expect(getRangeHeight(range)).toBe(0);
        });

        it("should default startRowIndex to 0 if missing", () => {
            const range: GoogleAppsScript.Sheets.Schema.GridRange = { sheetId: 1, endRowIndex: 4 };
            expect(getRangeHeight(range)).toBe(4);
        });

        it("should return 0 if the start index is greater than the end index", () => {
            const range: GoogleAppsScript.Sheets.Schema.GridRange = { sheetId: 1, startRowIndex: 5, endRowIndex: 2 };
            expect(getRangeHeight(range)).toBe(0);
        });
    });

    describe("shrinkRangeWidth", () => {
        it("should shrink the width by the specified number of columns", () => {
            const range: GoogleAppsScript.Sheets.Schema.GridRange = { sheetId: 1, startColumnIndex: 2, endColumnIndex: 7 };
            const result = shrinkRangeWidth(range, 2);
            expect(result).toEqual({
                sheetId: 1,
                startColumnIndex: 2,
                endColumnIndex: 5, // Width was 5, 5 - 2 = 3. New end is 2 + 3 = 5.
            });
        });

        it("should clamp the width to 0 if shrunk by more than the current width", () => {
            const range: GoogleAppsScript.Sheets.Schema.GridRange = { sheetId: 1, startColumnIndex: 2, endColumnIndex: 4 };
            const result = shrinkRangeWidth(range, 5);
            expect(result).toEqual({
                sheetId: 1,
                startColumnIndex: 2,
                endColumnIndex: 2, // Width becomes 0
            });
        });

        it("should handle ranges with missing bounds safely", () => {
            const range: GoogleAppsScript.Sheets.Schema.GridRange = { sheetId: 1, startColumnIndex: 2 };
            const result = shrinkRangeWidth(range, 1);
            expect(result).toEqual({
                sheetId: 1,
                startColumnIndex: 2,
                endColumnIndex: 2, // Original width calculated as 0, remains 0.
            });
        });
    });
});
