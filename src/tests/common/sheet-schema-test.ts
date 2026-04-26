import { ReportSheetSchema, SetupSheetSchema } from "../../common/sheet-schema";
import { type NestedSheetSchema, parseSpreadsheet } from "../../common/utils/mapped-name-range";
import { GasTestRunner } from "../gas-test-runner";

interface SchemaTestConfig {
    name: string;
    schema: NestedSheetSchema & { templateId: string };
    skipSheets?: string[];
    skipRanges?: string[];
}

const schemasToTest: SchemaTestConfig[] = [
    {
        name: "Setup Sheet",
        schema: SetupSheetSchema,
        skipSheets: [SetupSheetSchema.sheets.calendar.sheetName],
    },
    {
        name: "Report Sheet",
        schema: ReportSheetSchema,
    },
];

export function testSchemaValidation() {
    const runner = new GasTestRunner();
    const { describe, test, expect, beforeAll } = runner;

    for (const config of schemasToTest) {
        describe(`Schema Validation: ${config.name}`, () => {
            let rawSpreadsheet: GoogleAppsScript.Sheets.Schema.Spreadsheet;
            let parsedData: ReturnType<typeof parseSpreadsheet>;

            beforeAll(() => {
                const response = Sheets?.Spreadsheets.get(config.schema.templateId, {
                    fields: "sheets(properties(sheetId,title)),namedRanges",
                });

                if (!response) throw new Error(`Failed to fetch the spreadsheet for ${config.name}.`);

                rawSpreadsheet = response;
                parsedData = parseSpreadsheet(rawSpreadsheet, config.schema);
            });

            test("Should successfully fetch and parse the spreadsheet data", () => {
                expect(rawSpreadsheet).toBeTruthy();
                expect(parsedData).toBeTruthy();
                expect(parsedData.sheets).toBeTruthy();
                expect(parsedData.namedRanges).toBeTruthy();
            });

            for (const [sheetKey, sheetConfig] of Object.entries(config.schema.sheets)) {
                if (config.skipSheets?.includes(sheetConfig.sheetName)) continue;

                test(`[Sheet] "${sheetKey}" must exist`, () => {
                    expect(parsedData.sheets[sheetConfig.sheetName]).toBeTruthy();
                });

                if (sheetConfig.ranges) {
                    for (const [rangeKey, rangeName] of Object.entries(sheetConfig.ranges)) {
                        if (config.skipRanges?.includes(rangeName)) continue;

                        test(`|-- [Range] "${rangeKey}" must exist`, () => {
                            expect(parsedData.namedRanges[rangeName]).toBeTruthy();
                        });
                    }
                }
            }
        });
    }

    runner.execute();
}
