import { MergeType, PasteOrientation, PasteType } from "./api-types";
import {
    addNewNamedRange,
    addNewSheet,
    buildAddBandingRequest,
    buildCopyPasteRequest,
    buildMergeCellsRequest,
    buildTransferRequests,
    buildUpdateCellsRequest,
    buildUpdateSheetPropertiesRequest,
} from "./request";
import { type MappedNamedRange, type ParsedSpreadsheet, RangeBehavior } from "./types";

// Mock the random ID generator for consistent test assertions
jest.mock("../setup-utils", () => ({
    getRandomId: jest.fn(() => 9999),
}));

describe("GAS Util, Requests", () => {
    beforeAll(() => {
        global.Utilities = {
            getUuid: jest.fn(() => "mock-uuid-1234"),
        } as unknown as typeof Utilities;
    });

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

    describe("buildAddBandingRequest", () => {
        it("should correctly build an addBanding request with the provided range and properties", () => {
            const mockRange = { sheetId: 12345, startRowIndex: 0, endRowIndex: 10, startColumnIndex: 0, endColumnIndex: 5 };
            const mockBandingProperties = {
                headerColor: { red: 0.8, green: 0.8, blue: 0.8 },
                firstBandColor: { red: 1, green: 1, blue: 1 },
                secondBandColor: { red: 0.9, green: 0.9, blue: 0.9 },
            };

            const result = buildAddBandingRequest(mockRange, mockBandingProperties);

            expect(result).toEqual({
                addBanding: {
                    bandedRange: {
                        range: mockRange,
                        rowProperties: mockBandingProperties,
                    },
                },
            });
        });
    });

    describe("buildUpdateSheetPropertiesRequest", () => {
        it("should construct a request with only default hideGridlines if no other params provided", () => {
            const req = buildUpdateSheetPropertiesRequest({ sheetId: 1 });
            expect(req.updateSheetProperties?.properties?.sheetId).toBe(1);
            expect(req.updateSheetProperties?.properties?.gridProperties?.hideGridlines).toBe(true);
            expect(req.updateSheetProperties?.fields).toBe("gridProperties.hideGridlines");
        });

        it("should accurately map sheet and grid properties and compile the fields mask", () => {
            const req = buildUpdateSheetPropertiesRequest({
                sheetId: 5,
                index: 2,
                hidden: true,
                rowCount: 100,
                columnCount: 20,
                frozenRowCount: 1,
                frozenColumnCount: 2,
                hideGridlines: false,
            });

            const props = req.updateSheetProperties?.properties;
            expect(props?.sheetId).toBe(5);
            expect(props?.index).toBe(2);
            expect(props?.hidden).toBe(true);

            const gridProps = props?.gridProperties;
            expect(gridProps?.rowCount).toBe(100);
            expect(gridProps?.columnCount).toBe(20);
            expect(gridProps?.frozenRowCount).toBe(1);
            expect(gridProps?.frozenColumnCount).toBe(2);
            expect(gridProps?.hideGridlines).toBeUndefined();

            const fields = req.updateSheetProperties?.fields?.split(",") ?? [];
            expect(fields).toContain("index");
            expect(fields).toContain("hidden");
            expect(fields).toContain("gridProperties.rowCount");
            expect(fields).toContain("gridProperties.frozenColumnCount");
            expect(fields).not.toContain("gridProperties.hideGridlines");
        });
    });

    describe("buildUpdateCellsRequest", () => {
        it("should return undefined if the range has no rows or columns", () => {
            const zeroRange = { sheetId: 1, startRowIndex: 0, endRowIndex: 0, startColumnIndex: 0, endColumnIndex: 0 };
            const request = buildUpdateCellsRequest({ destination: zeroRange, data: [], fields: "*" });
            expect(request).toBeUndefined();
        });

        it("should natively fallback to 0 for missing dimensions and return undefined", () => {
            const openRange = { sheetId: 1 };
            const request = buildUpdateCellsRequest({ destination: openRange, data: [], fields: "userEnteredValue" });
            expect(request).toBeUndefined();
        });

        it("should build a valid updateCells request and pad sparse data arrays with empty objects", () => {
            const destination = { sheetId: 1, startRowIndex: 0, endRowIndex: 2, startColumnIndex: 0, endColumnIndex: 3 };
            const data: GoogleAppsScript.Sheets.Schema.CellData[][] = [[{ userEnteredValue: { stringValue: "A1" } }]];

            const request = buildUpdateCellsRequest({ destination, data, fields: "userEnteredValue" });

            expect(request?.updateCells?.range).toBe(destination);
            const rows = request?.updateCells?.rows ?? [];
            expect(rows.length).toBe(2);
            expect(rows[0]?.values?.length).toBe(3);
            expect(rows[0]?.values?.[0]).toEqual({ userEnteredValue: { stringValue: "A1" } });
            expect(rows[0]?.values?.[1]).toEqual({});
            expect(rows[1]?.values?.length).toBe(3);
        });
    });

    describe("buildTransferRequests", () => {
        let mockDestination: MappedNamedRange;

        jest.mock("./mapped-range", () => ({
            resizeMappedRange: jest.fn().mockImplementation(({ target, targetRows, targetCols }) => {
                if (targetRows === 0 && targetCols === 0) {
                    target.namedRange.range.endRowIndex = target.namedRange.range.startRowIndex;
                    target.namedRange.range.endColumnIndex = target.namedRange.range.startColumnIndex;
                    return { requests: [{ deleteDimension: {} }, { deleteDimension: {} }], rowOffset: 0, colOffset: 0 };
                }

                target.namedRange.range.endRowIndex = (target.namedRange.range.startRowIndex || 0) + targetRows;
                target.namedRange.range.endColumnIndex = (target.namedRange.range.startColumnIndex || 0) + targetCols;
                return { requests: [{ insertDimension: {} }], rowOffset: 0, colOffset: 0 };
            }),
        }));

        beforeEach(() => {
            mockDestination = {
                namedRange: { namedRangeId: "", name: "", range: { sheetId: 100, startRowIndex: 0, endRowIndex: 5, startColumnIndex: 0, endColumnIndex: 5 } },
                sheet: {},
            };
        });

        it("should offset the range and build update requests when behaviors are IGNORE (default)", () => {
            const data: GoogleAppsScript.Sheets.Schema.CellData[][] = [[{ userEnteredValue: { stringValue: "Data" } }]];

            const result = buildTransferRequests({
                destination: mockDestination,
                data,
                fields: "*",
                rowOffset: 2,
                colOffset: 2,
            });

            expect(result.requests).toHaveLength(1);
            expect(result.rowOffset).toBe(2);

            const updateReq = result.requests[0]?.updateCells;
            expect(updateReq?.range?.startRowIndex).toBe(2);
            expect(updateReq?.range?.endRowIndex).toBe(7);
        });

        it("should adapt (resize) the range, append resize requests, and format update data when behaviors are INSERT_DELETE", () => {
            const data: GoogleAppsScript.Sheets.Schema.CellData[][] = [
                [{ userEnteredValue: { stringValue: "R1C1" } }, {}, {}],
                [{}, {}, { userEnteredValue: { stringValue: "R2C3" } }],
            ];

            const result = buildTransferRequests({
                destination: mockDestination,
                data,
                fields: "userEnteredValue",
                rowBehavior: RangeBehavior.INSERT_DELETE,
                colBehavior: RangeBehavior.INSERT_DELETE,
            });

            expect(result.requests).toHaveLength(3);

            const updateReq = result.requests[2]?.updateCells;
            expect(updateReq?.range?.endRowIndex).toBe(2);
            expect(updateReq?.range?.endColumnIndex).toBe(3);
        });

        it("should gracefully handle destroying the range (0 data) and skip pushing updateCells", () => {
            const data: GoogleAppsScript.Sheets.Schema.CellData[][] = [];

            const result = buildTransferRequests({
                destination: mockDestination,
                data,
                fields: "*",
                rowBehavior: RangeBehavior.INSERT_DELETE,
                colBehavior: RangeBehavior.INSERT_DELETE,
            });

            expect(result.requests).toHaveLength(2);
            const hasUpdateReq = result.requests.some((req) => req.updateCells !== undefined);
            expect(hasUpdateReq).toBe(false);

            expect(mockDestination.namedRange.range.endRowIndex).toBe(0);
        });
    });

    // --- State Mutation Tests Setup ---
    const mockSchema = {
        sheets: {
            template: {
                sheetName: "TemplateSheet",
                ranges: { tempRange: "TempRange" },
            },
            data: {
                sheetName: "DataSheet",
                ranges: { staticKey: "StaticRange" },
                dynamicRanges: { dynPrefix: "DynPrefix_" },
            },
            newTemplate: {
                sheetName: "BoundSheet",
            },
        },
    } as const;

    type MockSchema = typeof mockSchema;

    describe("State Mutation: addNewSheet", () => {
        let mockParsedData: ParsedSpreadsheet<MockSchema>;

        beforeEach(() => {
            mockParsedData = {
                mappedSheets: {},
                mappedSheetNamedRanges: {
                    TemplateSheet: [{ namedRangeId: "nr-1", name: "TempRange", range: { sheetId: 1 } }],
                },
                mappedRanges: {},
                dynamicMappedRanges: {},
                extraSheets: [],
            };
        });

        it("should handle schema-bound duplication, managing state and named ranges", () => {
            const { requests, newSheetIds } = addNewSheet({
                parsedData: mockParsedData,
                sourceSheetTitle: "TemplateSheet",
                sourceSheetId: 1,
                insertSheetIndex: 2,
                schema: mockSchema,
                schemaSheetKey: "newTemplate",
            });

            expect(requests.length).toBeGreaterThan(0);
            expect(requests[0]?.deleteNamedRange?.namedRangeId).toBe("nr-1");
            expect(requests[1]?.duplicateSheet?.newSheetName).toBe("BoundSheet");
            expect(requests[2]?.addNamedRange?.namedRange?.namedRangeId).toBe("nr-1");

            expect(newSheetIds).toEqual([9999]);
            expect(mockParsedData.mappedSheets.BoundSheet).toBeDefined();
            expect(mockParsedData.mappedSheetNamedRanges.BoundSheet).toEqual([]);
            expect(mockParsedData.extraSheets).toHaveLength(0);
        });

        it("should handle non-schema multiple sheet duplications", () => {
            const { requests, newSheetIds } = addNewSheet({
                parsedData: mockParsedData,
                sourceSheetTitle: "TemplateSheet",
                sourceSheetId: 1,
                insertSheetIndex: 2,
                multipleSheetNames: ["ExtraOne", "ExtraTwo"],
            });

            expect(newSheetIds).toEqual([9999, 9999]);
            expect(requests.filter((r) => r.duplicateSheet)).toHaveLength(2);

            expect(mockParsedData.extraSheets).toHaveLength(2);
            expect(mockParsedData.extraSheets[0]?.properties?.title).toBe("ExtraOne");
        });
    });

    describe("State Mutation: addNewNamedRange", () => {
        let mockParsedData: ParsedSpreadsheet<MockSchema>;

        beforeEach(() => {
            mockParsedData = {
                mappedSheets: {
                    DataSheet: { properties: { sheetId: 42, title: "DataSheet" } },
                },
                mappedSheetNamedRanges: {},
                mappedRanges: {},
                dynamicMappedRanges: {},
                extraSheets: [],
            };
        });

        it("should throw an error if the specified sheet does not exist in mappedSheets", () => {
            expect(() => {
                addNewNamedRange({
                    parsedData: mockParsedData,
                    // @ts-expect-error - intentionally passing an invalid sheet title to test runtime validation
                    sheetTitle: "UnknownSheet",
                    gridRange: { startRowIndex: 0 },
                    staticRangeKey: "TempRange",
                });
            }).toThrow("Adding range name to sheet");
        });

        it("should successfully add a static named range and update state", () => {
            const req = addNewNamedRange({
                parsedData: mockParsedData,
                sheetTitle: "DataSheet",
                gridRange: { startRowIndex: 1, endRowIndex: 2 },
                staticRangeKey: "TempRange",
            });

            expect(req.addNamedRange?.namedRange?.namedRangeId).toBe("mock-uuid-1234");
            expect(req.addNamedRange?.namedRange?.name).toBe("TempRange");
            expect(req.addNamedRange?.namedRange?.range?.sheetId).toBe(42);

            expect(mockParsedData.mappedSheetNamedRanges.DataSheet).toHaveLength(1);
            expect(mockParsedData.mappedRanges.TempRange).toBeDefined();
            expect(mockParsedData.mappedRanges.TempRange?.sheet.properties?.sheetId).toBe(42);
        });

        it("should successfully add a dynamic named range and update state", () => {
            const req = addNewNamedRange({
                parsedData: mockParsedData,
                sheetTitle: "DataSheet",
                gridRange: { startRowIndex: 5, endRowIndex: 10 },
                rangeName: "dyn_A_1",
                dynamicRangeKey: "dynPrefix",
            });

            expect(req.addNamedRange?.namedRange?.name).toBe("dyn_A_1");

            expect(mockParsedData.mappedSheetNamedRanges.DataSheet).toHaveLength(1);
            expect(mockParsedData.dynamicMappedRanges.dynPrefix).toHaveLength(1);
            expect(mockParsedData.dynamicMappedRanges.dynPrefix?.[0]?.namedRange.name).toBe("dyn_A_1");
        });
    });
});
