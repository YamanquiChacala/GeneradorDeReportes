import type { ExtractDynamicRangeKeys, ExtractRangeNames, ExtractSheetNames, MappedNamedRange, NestedSheetSchema, ParsedSpreadsheet, StrictNameRange } from "./types";

/**
 * Parses a Spreadsheet coming from Sheets API, so that the Sheets and named ranges are easy to find and work with.
 */
export function parseSpreadsheet<T extends NestedSheetSchema>(spreadsheet: GoogleAppsScript.Sheets.Schema.Spreadsheet | undefined, schema: T): ParsedSpreadsheet<T> {
    const mappedSheets: Partial<Record<ExtractSheetNames<T>, GoogleAppsScript.Sheets.Schema.Sheet>> = {};
    const mappedSheetNamedRanges: Partial<Record<ExtractSheetNames<T>, StrictNameRange[]>> = {};
    const mappedRanges: Partial<Record<ExtractRangeNames<T>, MappedNamedRange>> = {};
    const dynamicMappedRanges: Partial<Record<ExtractDynamicRangeKeys<T>, MappedNamedRange[]>> = {};
    const extraSheets: GoogleAppsScript.Sheets.Schema.Sheet[] = [];

    if (!spreadsheet?.sheets) return { mappedSheets, mappedSheetNamedRanges, mappedRanges, dynamicMappedRanges, extraSheets };

    const allowedSheetNames = new Set<string>();
    const allowedRangeNames = new Set<string>();
    const dynamicRangePrefixes = new Map<string, ExtractDynamicRangeKeys<T>>();

    for (const sheetConfig of Object.values(schema.sheets)) {
        allowedSheetNames.add(sheetConfig.sheetName);

        if (sheetConfig.ranges) {
            for (const rangeName of Object.values(sheetConfig.ranges)) {
                allowedRangeNames.add(rangeName);
            }
        }

        if (sheetConfig.dynamicRanges) {
            for (const [key, prefix] of Object.entries(sheetConfig.dynamicRanges)) {
                dynamicRangePrefixes.set(prefix, key as ExtractDynamicRangeKeys<T>);
            }
        }
    }

    const sheetIdLookup: Record<number, GoogleAppsScript.Sheets.Schema.Sheet> = {};

    for (const sheet of spreadsheet.sheets) {
        sheetIdLookup[sheet.properties?.sheetId ?? 0] = sheet;
        const sheetTitle = sheet.properties?.title;

        if (sheetTitle != null) {
            if (allowedSheetNames.has(sheetTitle)) {
                mappedSheets[sheetTitle as ExtractSheetNames<T>] = sheet;
                mappedSheetNamedRanges[sheetTitle as ExtractSheetNames<T>] =
                    spreadsheet.namedRanges?.filter(
                        (namedRange): namedRange is StrictNameRange => isStrictNameRange(namedRange) && namedRange.range.sheetId === sheet.properties?.sheetId,
                    ) ?? [];
            } else {
                extraSheets.push(sheet);
            }
        }
    }

    if (spreadsheet.namedRanges) {
        for (const namedRange of spreadsheet.namedRanges) {
            if (!isStrictNameRange(namedRange)) continue;
            const linkedSheet = sheetIdLookup[namedRange.range.sheetId ?? 0];

            if (!linkedSheet) continue;

            if (allowedRangeNames.has(namedRange.name)) {
                mappedRanges[namedRange.name as ExtractRangeNames<T>] = {
                    namedRange: namedRange,
                    sheet: linkedSheet,
                };
                continue;
            }

            for (const [prefix, dynamicKey] of dynamicRangePrefixes.entries()) {
                if (namedRange.name.startsWith(prefix)) {
                    if (!dynamicMappedRanges[dynamicKey]) dynamicMappedRanges[dynamicKey] = [];
                    dynamicMappedRanges[dynamicKey].push({
                        namedRange: namedRange,
                        sheet: linkedSheet,
                    });
                    break;
                }
            }
        }
    }

    return { mappedSheets, mappedSheetNamedRanges, mappedRanges, dynamicMappedRanges, extraSheets };
}

/**
 * Tests that a named range is a StrictNameRange
 */
export function isStrictNameRange(namedRange: GoogleAppsScript.Sheets.Schema.NamedRange): namedRange is StrictNameRange {
    return namedRange.namedRangeId != null && namedRange.name != null && namedRange.range != null;
}
