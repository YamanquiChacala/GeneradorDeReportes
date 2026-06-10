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
    const usedIds = new Set<number>();

    if (!spreadsheet?.sheets) return { mappedSheets, mappedSheetNamedRanges, mappedRanges, dynamicMappedRanges, extraSheets, usedIds };

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

    // Collect spreadsheet-level Developer Metadata IDs if present
    if (spreadsheet.developerMetadata) {
        for (const meta of spreadsheet.developerMetadata) {
            if (meta.metadataId != null) usedIds.add(meta.metadataId);
        }
    }

    for (const sheet of spreadsheet.sheets) {
        sheetIdLookup[sheet.properties?.sheetId ?? 0] = sheet;
        const sheetTitle = sheet.properties?.title;

        // Track the sheet ID
        if (sheet.properties?.sheetId != null) {
            usedIds.add(sheet.properties.sheetId);
        }

        // Track Protected Ranges
        if (sheet.protectedRanges) {
            for (const pr of sheet.protectedRanges) {
                if (pr.protectedRangeId != null) usedIds.add(pr.protectedRangeId);
            }
        }

        // Track other elements that consume numeric IDs just to be safe
        if (sheet.charts) {
            for (const chart of sheet.charts) {
                if (chart.chartId != null) usedIds.add(chart.chartId);
            }
        }
        if (sheet.bandedRanges) {
            for (const br of sheet.bandedRanges) {
                if (br.bandedRangeId != null) usedIds.add(br.bandedRangeId);
            }
        }
        if (sheet.filterViews) {
            for (const fv of sheet.filterViews) {
                if (fv.filterViewId != null) usedIds.add(fv.filterViewId);
            }
        }

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

    return { mappedSheets, mappedSheetNamedRanges, mappedRanges, dynamicMappedRanges, extraSheets, usedIds };
}

/**
 * Tests that a named range is a StrictNameRange
 */
export function isStrictNameRange(namedRange: GoogleAppsScript.Sheets.Schema.NamedRange): namedRange is StrictNameRange {
    return namedRange.namedRangeId != null && namedRange.name != null && namedRange.range != null;
}
