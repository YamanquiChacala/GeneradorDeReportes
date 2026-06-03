import type { ExtractDynamicRangeKeys, ExtractRangeNames, ExtractSheetNames, MappedNamedRange, NestedSheetSchema, ParsedSpreadsheet, StrictNameRange } from "./types";

/**
 * Parses a Spreadsheet coming from Sheets API, so that the Sheets and named ranges are easy to find and work with.
 */
export function parseSpreadsheet<T extends NestedSheetSchema>(spreadsheet: GoogleAppsScript.Sheets.Schema.Spreadsheet | undefined, schema: T): ParsedSpreadsheet<T> {
    const mappedSheets: Partial<Record<ExtractSheetNames<T>, GoogleAppsScript.Sheets.Schema.Sheet>> = {};
    const mappedSheetNamedRanges: Partial<Record<ExtractSheetNames<T>, StrictNameRange[]>> = {};
    const mappedRanges: Partial<Record<ExtractRangeNames<T>, MappedNamedRange>> = {};
    const dynamicMappedRanges: Partial<Record<ExtractDynamicRangeKeys<T>, MappedNamedRange[]>> = {};

    if (!spreadsheet?.sheets) return { mappedSheets, mappedSheetNamedRanges, mappedRanges, dynamicMappedRanges };

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
        if (sheetTitle != null && allowedSheetNames.has(sheetTitle)) {
            mappedSheets[sheetTitle as ExtractSheetNames<T>] = sheet;
            mappedSheetNamedRanges[sheetTitle as ExtractSheetNames<T>] = spreadsheet.namedRanges?.filter(
                (namedRange): namedRange is StrictNameRange => isStrictNameRange(namedRange) && namedRange.range.sheetId === sheet.properties?.sheetId,
            );
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

    return { mappedSheets, mappedSheetNamedRanges, mappedRanges, dynamicMappedRanges };
}

interface InsertNewNamedRangeToMemoryParams<T extends NestedSheetSchema> {
    parsedData: ParsedSpreadsheet<T>;
    sheetTitle: ExtractSheetNames<T>;
    rangeNameId: string;
    rangeName: string;
    gridRange: GoogleAppsScript.Sheets.Schema.GridRange;
    staticRangeKey?: ExtractRangeNames<T>;
    dynamicRangeKey?: ExtractDynamicRangeKeys<T>;
}

/**
 * Update the memory representation with a new named range
 */
export function insertNewNamedRangeToMemory<T extends NestedSheetSchema>({
    parsedData,
    sheetTitle,
    rangeNameId,
    rangeName,
    gridRange,
    staticRangeKey,
    dynamicRangeKey,
}: InsertNewNamedRangeToMemoryParams<T>) {
    const sheet = parsedData.mappedSheets[sheetTitle];
    if (!sheet) {
        throw new Error(`Cannot add range. Sheet ${sheetTitle as string} is not in memory.`);
    }

    const newStrictNamedRange: StrictNameRange = {
        namedRangeId: rangeNameId,
        name: rangeName,
        range: gridRange,
    };

    const newMappedNamedRange: MappedNamedRange = {
        namedRange: newStrictNamedRange,
        sheet: sheet,
    };

    if (!parsedData.mappedSheetNamedRanges[sheetTitle]) {
        parsedData.mappedSheetNamedRanges[sheetTitle] = [];
    }
    parsedData.mappedSheetNamedRanges[sheetTitle].push(newStrictNamedRange);

    if (staticRangeKey) {
        const targetMap = parsedData.mappedRanges as Record<string, MappedNamedRange | undefined>;
        targetMap[staticRangeKey] = newMappedNamedRange;
    } else if (dynamicRangeKey) {
        const targetMap = parsedData.dynamicMappedRanges as Record<string, MappedNamedRange[] | undefined>;
        if (!targetMap[dynamicRangeKey]) {
            targetMap[dynamicRangeKey] = [];
        }
        targetMap[dynamicRangeKey].push(newMappedNamedRange);
    }
}

/**
 * Tests that a named range is a StrictNameRange
 */
function isStrictNameRange(namedRange: GoogleAppsScript.Sheets.Schema.NamedRange): namedRange is StrictNameRange {
    return namedRange.name != null && namedRange.range != null;
}
