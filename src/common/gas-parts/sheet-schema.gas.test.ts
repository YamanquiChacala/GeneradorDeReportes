import { GasTestRunner } from "../../testing/gas-test-runner";
import { buildFieldsMask, parseSpreadsheet } from "../gas-utils";
import { ReportSheetSchema, SetupSheetSchema } from "./sheet-schema";

interface ExpectedSchemaShape {
    readonly templateId: string;
    readonly sheets: Readonly<
        Record<
            string,
            {
                readonly sheetName: string;
                readonly ranges?: Readonly<Record<string, string>>;
            }
        >
    >;
}

interface SchemaTestConfig {
    readonly name: string;
    readonly schema: ExpectedSchemaShape;
    readonly skipSheets?: readonly string[];
    readonly skipRanges?: readonly string[];
}

const schemasToTest: SchemaTestConfig[] = [
    {
        name: "Setup Spreadsheet",
        schema: SetupSheetSchema,
        skipSheets: [SetupSheetSchema.sheets.calendar.sheetName],
    },
    {
        name: "Report Spreadsheet",
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

    for (const schema of schemasToTest) {
        describe(`Schema Validation: ${schema.name}`, () => {
            let rawSpreadsheet: GoogleAppsScript.Sheets.Schema.Spreadsheet;
            let parsedData: ReturnType<typeof parseSpreadsheet>;

            beforeAll(() => {
                const response = Sheets?.Spreadsheets.get(schema.schema.templateId, {
                    fields: buildFieldsMask<GoogleAppsScript.Sheets.Schema.Spreadsheet>("sheets.properties.sheetId", "sheets.properties.title", "namedRanges"),
                });

                if (!response) throw new Error(`Failed to fetch the spreadsheet for ${schema.name}.`);

                rawSpreadsheet = response;
                parsedData = parseSpreadsheet(rawSpreadsheet, schema.schema);
            });

            test("Should successfully fetch and parse the spreadsheet data", () => {
                expect(rawSpreadsheet).toBeTruthy();
                expect(parsedData).toBeTruthy();
                expect(parsedData.mappedSheets).toBeTruthy();
                expect(parsedData.mappedRanges).toBeTruthy();
            });

            for (const [sheetKey, sheetConfig] of Object.entries(schema.schema.sheets)) {
                if (schema.skipSheets?.includes(sheetConfig.sheetName)) continue;

                test(`[Sheet] "${sheetKey}" must exist`, () => {
                    expect(parsedData.mappedSheets[sheetConfig.sheetName]).toBeTruthy();
                });

                if (sheetConfig.ranges) {
                    for (const [rangeKey, rangeName] of Object.entries(sheetConfig.ranges)) {
                        if (schema.skipRanges?.includes(rangeName)) continue;

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
