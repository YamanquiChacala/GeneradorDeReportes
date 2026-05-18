import type { ExtractRangeNames, ExtractSheetNames, MappedNamedRange, NestedSheetSchema, ParsedSpreadsheet } from ".";

/**
 * Parses a Spreadsheet coming from Sheets API, so that the Sheets and named ranges are easy to find and work with.
 */
export function parseSpreadsheet<T extends NestedSheetSchema>(spreadsheet: GoogleAppsScript.Sheets.Schema.Spreadsheet | undefined, schema: T): ParsedSpreadsheet<T> {
    const mappedSheets: Partial<Record<ExtractSheetNames<T>, GoogleAppsScript.Sheets.Schema.Sheet>> = {};
    const mappedSheetNamedRanges: Partial<Record<ExtractSheetNames<T>, GoogleAppsScript.Sheets.Schema.NamedRange[]>> = {};
    const mappedRanges: Partial<Record<ExtractRangeNames<T>, MappedNamedRange>> = {};

    if (!spreadsheet?.sheets) return { sheets: mappedSheets, sheetNamedRanges: mappedSheetNamedRanges, mappedRanges: mappedRanges };

    const allowedSheetNames = new Set<string>();
    const allowedRangeNames = new Set<string>();

    for (const key of Object.keys(schema.sheets)) {
        const sheetConfig = schema.sheets[key];
        if (sheetConfig) allowedSheetNames.add(sheetConfig.sheetName);

        if (sheetConfig?.ranges) {
            for (const rangeKey of Object.keys(sheetConfig.ranges)) {
                if (sheetConfig.ranges[rangeKey]) allowedRangeNames.add(sheetConfig.ranges[rangeKey]);
            }
        }
    }

    const sheetIdLookup: Record<number, GoogleAppsScript.Sheets.Schema.Sheet> = {};

    for (const sheet of spreadsheet.sheets) {
        sheetIdLookup[sheet.properties?.sheetId ?? 0] = sheet;
        const sheetTitle = sheet.properties?.title;
        if (sheetTitle != null && allowedSheetNames.has(sheetTitle)) {
            mappedSheets[sheetTitle as ExtractSheetNames<T>] = sheet;
            mappedSheetNamedRanges[sheetTitle as ExtractSheetNames<T>] = spreadsheet.namedRanges?.filter(
                (namedRange) => namedRange.range?.sheetId === sheet.properties?.sheetId,
            );
        }
    }

    if (spreadsheet.namedRanges) {
        for (const namedRange of spreadsheet.namedRanges) {
            if (namedRange.name == null || namedRange.range == null) continue;
            const linkedSheet = sheetIdLookup[namedRange.range.sheetId ?? 0];

            if (linkedSheet && allowedRangeNames.has(namedRange.name)) {
                mappedRanges[namedRange.name as ExtractRangeNames<T>] = {
                    range: namedRange.range,
                    sheet: linkedSheet,
                };
            }
        }
    }

    return { sheets: mappedSheets, sheetNamedRanges: mappedSheetNamedRanges, mappedRanges: mappedRanges };
}
