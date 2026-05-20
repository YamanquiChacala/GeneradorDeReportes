import { MergeType, PasteOrientation, PasteType } from "../constants";
import type { MappedNamedRange } from ".";
import { buildAddNamedRangeRequest, buildCopyPasteRequest, buildMergeCellsRequest, buildTransferRequests, buildUpdateCellsRequest } from ".";

describe("Request Utilities", () => {
    describe("buildCopyPasteRequest", () => {
        it("should build a valid copyPaste request with normal orientation", () => {
            const source = { sheetId: 1, startRowIndex: 0, endRowIndex: 5, startColumnIndex: 0, endColumnIndex: 5 };
            const destination = { sheetId: 2, startRowIndex: 5, endRowIndex: 10, startColumnIndex: 0, endColumnIndex: 5 };
            const request = buildCopyPasteRequest(source, destination, PasteType.PASTE_FORMAT);

            expect(request.copyPaste).toBeDefined();
            expect(request.copyPaste?.source).toBe(source);
            expect(request.copyPaste?.destination).toBe(destination);
            expect(request.copyPaste?.pasteType).toBe(PasteType.PASTE_FORMAT);
            expect(request.copyPaste?.pasteOrientation).toBe(PasteOrientation.NORMAL);
        });
    });

    describe("buildMergeCellsRequest", () => {
        it("should build a valid mergeCells request", () => {
            const range = { sheetId: 1, startRowIndex: 0, endRowIndex: 2, startColumnIndex: 0, endColumnIndex: 2 };
            const request = buildMergeCellsRequest(range, MergeType.MERGE_ALL);

            expect(request.mergeCells).toBeDefined();
            expect(request.mergeCells?.range).toBe(range);
            expect(request.mergeCells?.mergeType).toBe(MergeType.MERGE_ALL);
        });
    });

    describe("buildAddNamedRangeRequest", () => {
        it("should build a valid addNamedRange request", () => {
            const range = { sheetId: 1, startRowIndex: 0, endRowIndex: 10, startColumnIndex: 0, endColumnIndex: 2 };
            // Using 'as any' to bypass strict TS generic extraction requirements for the mock string
            const request = buildAddNamedRangeRequest("TestNamedRange", range);

            expect(request.addNamedRange).toBeDefined();
            expect(request.addNamedRange?.namedRange?.name).toBe("TestNamedRange");
            expect(request.addNamedRange?.namedRange?.range).toBe(range);
        });
    });

    describe("buildUpdateCellsRequest", () => {
        it("should return undefined if the range has no rows or columns", () => {
            const zeroRange = { sheetId: 1, startRowIndex: 0, endRowIndex: 0, startColumnIndex: 0, endColumnIndex: 0 };
            const request = buildUpdateCellsRequest({ destination: zeroRange, data: [], fields: "*" });

            expect(request).toBeUndefined();
        });

        it("should natively fallback to 0 for missing dimensions and return undefined", () => {
            // Missing all index properties
            const openRange = { sheetId: 1 };
            const request = buildUpdateCellsRequest({ destination: openRange, data: [], fields: "userEnteredValue" });

            expect(request).toBeUndefined();
        });

        it("should build a valid updateCells request and pad sparse data arrays with empty objects", () => {
            const destination = { sheetId: 1, startRowIndex: 0, endRowIndex: 2, startColumnIndex: 0, endColumnIndex: 3 };

            // Sparse data: only 1 row provided (target needs 2), and only 1 col provided (target needs 3)
            const data: GoogleAppsScript.Sheets.Schema.CellData[][] = [[{ userEnteredValue: { stringValue: "A1" } }]];

            const request = buildUpdateCellsRequest({ destination, data, fields: "userEnteredValue" });

            expect(request?.updateCells).toBeDefined();
            expect(request?.updateCells?.range).toBe(destination);
            expect(request?.updateCells?.fields).toBe("userEnteredValue");

            const rows = request?.updateCells?.rows ?? [];
            expect(rows.length).toBe(2); // Expanded to 2 rows to match destination range

            // Row 1
            expect(rows[0]?.values?.length).toBe(3); // Expanded to 3 columns
            expect(rows[0]?.values?.[0]).toEqual({ userEnteredValue: { stringValue: "A1" } });
            expect(rows[0]?.values?.[1]).toEqual({}); // Padded missing column
            expect(rows[0]?.values?.[2]).toEqual({}); // Padded missing column

            // Row 2 (Entirely missing from source data)
            expect(rows[1]?.values?.length).toBe(3);
            expect(rows[1]?.values?.[0]).toEqual({});
            expect(rows[1]?.values?.[1]).toEqual({});
            expect(rows[1]?.values?.[2]).toEqual({});
        });
    });

    describe("buildTransferRequests", () => {
        let mockDestination: MappedNamedRange;

        beforeEach(() => {
            mockDestination = {
                range: { sheetId: 100, startRowIndex: 0, endRowIndex: 5, startColumnIndex: 0, endColumnIndex: 5 },
                sheet: {},
            };
        });

        it("should offset the range and build update requests when adaptRange is false", () => {
            const data: GoogleAppsScript.Sheets.Schema.CellData[][] = [[{ userEnteredValue: { stringValue: "Data" } }]];

            const result = buildTransferRequests({
                destination: mockDestination,
                data,
                fields: "*",
                adaptRange: false,
                rowOffset: 2,
                colOffset: 2,
            });

            // Expect 1 request total (just the updateCells, no resize requests)
            expect(result.requests).toHaveLength(1);
            expect(result.rowOffset).toBe(2);
            expect(result.colOffset).toBe(2);

            const updateReq = result.requests[0]?.updateCells;
            expect(updateReq).toBeDefined();

            // Check that the offsets were successfully applied to the target range
            expect(updateReq?.range?.startRowIndex).toBe(2); // 0 + 2
            expect(updateReq?.range?.endRowIndex).toBe(7); // 5 + 2
            expect(updateReq?.range?.startColumnIndex).toBe(2); // 0 + 2
            expect(updateReq?.range?.endColumnIndex).toBe(7); // 5 + 2

            // It should map data to the size of the offset range (5x5)
            expect(updateReq?.rows?.length).toBe(5);
        });

        it("should adapt (resize) the range, append resize requests, and format update data when adaptRange is true", () => {
            // Data is 2 rows by 3 columns. The current range is 5x5.
            const data: GoogleAppsScript.Sheets.Schema.CellData[][] = [
                [{ userEnteredValue: { stringValue: "R1C1" } }, {}, {}],
                [{}, {}, { userEnteredValue: { stringValue: "R2C3" } }],
            ];

            const result = buildTransferRequests({
                destination: mockDestination,
                data,
                fields: "userEnteredValue",
                adaptRange: true,
                rowOffset: 0,
                colOffset: 0,
            });

            // 1 row delete request, 1 col delete request, 1 updateCells request
            expect(result.requests).toHaveLength(3);

            // Verify the last request is the updateCells request mapped to the newly adapted 2x3 size
            const updateReq = result.requests[2]?.updateCells;
            expect(updateReq).toBeDefined();
            expect(updateReq?.range?.endRowIndex).toBe(2); // shrunk from 5 to 2
            expect(updateReq?.range?.endColumnIndex).toBe(3); // shrunk from 5 to 3

            expect(updateReq?.rows?.length).toBe(2);
            expect(updateReq?.rows?.[0]?.values?.length).toBe(3);
        });

        it("should correctly calculate dataCols based on the longest row in a jagged array", () => {
            const data: GoogleAppsScript.Sheets.Schema.CellData[][] = [
                [{}], // length 1
                [{}, {}, {}, {}], // length 4 (Max)
                [{}, {}], // length 2
            ];

            const result = buildTransferRequests({
                destination: mockDestination,
                data,
                fields: "*",
                adaptRange: true,
            });

            const updateReq = result.requests[result.requests.length - 1]?.updateCells;
            // The maximum columns width was 4, so the range should be adapted to endColumnIndex 4
            expect(updateReq?.range?.endColumnIndex).toBe(4);
            expect(updateReq?.rows?.[0]?.values?.length).toBe(4);
        });

        it("should gracefully handle destroying the range (0 data) and skip pushing updateCells", () => {
            const data: GoogleAppsScript.Sheets.Schema.CellData[][] = []; // Empty data

            const result = buildTransferRequests({
                destination: mockDestination,
                data,
                fields: "*",
                adaptRange: true,
            });

            // Because it shrinks to 0x0, it generates 2 deleteDimension requests
            expect(result.requests).toHaveLength(2);
            expect(result.requests[0]?.deleteDimension).toBeDefined();
            expect(result.requests[1]?.deleteDimension).toBeDefined();

            // The final range dimensions are 0x0, so it should NOT push an updateCells request
            const hasUpdateReq = result.requests.some((req) => req.updateCells !== undefined);
            expect(hasUpdateReq).toBe(false);

            // The memory destination should be completely collapsed
            expect(mockDestination.range.endRowIndex).toBe(0);
            expect(mockDestination.range.endColumnIndex).toBe(0);
        });

        it("should handle omitted sheet properties seamlessly", () => {
            const statelessDestination: MappedNamedRange = {
                range: { startRowIndex: 0, endRowIndex: 2, startColumnIndex: 0, endColumnIndex: 2 }, // No sheetId
                sheet: {},
            };

            const data: GoogleAppsScript.Sheets.Schema.CellData[][] = [[{}]];

            const result = buildTransferRequests({
                destination: statelessDestination,
                data,
                fields: "*",
                adaptRange: false,
            });

            const updateReq = result.requests[0]?.updateCells;
            // Should fallback to sheetId 0
            expect(updateReq?.range?.sheetId).toBe(0);
        });

        it("should default adaptRange to false if omitted, preventing range resizing", () => {
            const data: GoogleAppsScript.Sheets.Schema.CellData[][] = [[{ userEnteredValue: { stringValue: "Data" } }]];

            // Notice: adaptRange is completely omitted here
            const result = buildTransferRequests({
                destination: mockDestination,
                data,
                fields: "*",
            });

            // It should only have 1 request (the updateCells request), meaning resize was skipped
            expect(result.requests).toHaveLength(1);

            const updateReq = result.requests[0]?.updateCells;
            expect(updateReq).toBeDefined();

            // Dimensions should remain the original 5x5 instead of shrinking to the 1x1 data size
            expect(updateReq?.range?.endRowIndex).toBe(5);
            expect(updateReq?.range?.endColumnIndex).toBe(5);
        });

        it("should safely fall back to 0 if the destination range is completely missing endRowIndex and endColumnIndex", () => {
            const openDestination: MappedNamedRange = {
                range: { sheetId: 100, startRowIndex: 0, startColumnIndex: 0 }, // Explicitly missing end indexes
                sheet: {},
            };

            const data: GoogleAppsScript.Sheets.Schema.CellData[][] = [[{}]];

            const result = buildTransferRequests({
                destination: openDestination,
                data,
                fields: "*",
                adaptRange: false,
            });

            // Because finalEndRow (0) - finalStartRow (0) = 0 finalRows,
            // it fails the `if (finalRows > 0 && finalCols > 0)` check and skips building the update request.
            expect(result.requests).toHaveLength(0);
        });
    });
});
