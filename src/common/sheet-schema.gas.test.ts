import { GasTestRunner } from "../testing/gas-test-runner";
import { ReportSheetSchema, SetupSheetSchema } from "./sheet-schema";
import { buildFieldsMask } from "./utils/gas-types";
import { parseSpreadsheet } from "./utils/mapped-name-range";

interface ExpectedSchemaShape {
    readonly templateId: string;
    readonly sheets: Record<
        string,
        {
            readonly sheetName: string;
            readonly ranges?: Record<string, string>;
        }
    >;
}

interface SchemaTestConfig {
    name: string;
    // schema: NestedSheetSchema & { templateId: string };
    schema: ExpectedSchemaShape;
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
        skipSheets: [ReportSheetSchema.sheets.attendance.sheetName],
        skipRanges: [
            ReportSheetSchema.sheets.studentTemplate.ranges.unprotectedAbilities,
            ReportSheetSchema.sheets.studentTemplate.ranges.unprotectedComments,
            ReportSheetSchema.sheets.studentTemplate.ranges.unprotectedTrim1,
            ReportSheetSchema.sheets.studentTemplate.ranges.unprotectedTrim2,
            ReportSheetSchema.sheets.studentTemplate.ranges.unprotectedTrim3,
        ],
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
                    fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.Spreadsheet>("sheets.properties.sheetId", "sheets.properties.title", "namedRanges"),
                });

                if (!response) throw new Error(`Failed to fetch the spreadsheet for ${config.name}.`);

                rawSpreadsheet = response;
                parsedData = parseSpreadsheet(rawSpreadsheet, config.schema);
            });

            test("Should successfully fetch and parse the spreadsheet data", () => {
                expect(rawSpreadsheet).toBeTruthy();
                expect(parsedData).toBeTruthy();
                expect(parsedData.sheets).toBeTruthy();
                expect(parsedData.mappedRanges).toBeTruthy();
            });

            for (const [sheetKey, sheetConfig] of Object.entries(config.schema.sheets)) {
                if (config.skipSheets?.includes(sheetConfig.sheetName)) continue;

                test(`[Sheet] "${sheetKey}" must exist`, () => {
                    expect(parsedData.sheets[sheetConfig.sheetName]).toBeTruthy();
                });

                if (sheetConfig.ranges) {
                    for (const [rangeKey, rangeName] of Object.entries(sheetConfig.ranges)) {
                        if (config.skipRanges?.includes(rangeName)) continue;

                        test(`    [Range] "${rangeKey}" must exist`, () => {
                            expect(parsedData.mappedRanges[rangeName]).toBeTruthy();
                        });
                    }
                }
            }
        });
    }

    runner.execute();
}
