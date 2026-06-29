import { offsetGridRange } from "./range";
import type { MappedNamedRange } from "./types";

/**
 * Transforms a column number into it's corresponding column letter, using 0-based index.
 */
export function getColumnLetter(column: number): string {
    let temp: number;
    let letter = "";
    while (column >= 0) {
        temp = column % 26;
        letter = String.fromCharCode(temp + 65) + letter;
        column = Math.floor(column / 26) - 1;
    }
    return letter;
}

export interface A1NotationParams {
    mappedRange: MappedNamedRange;
    includeSheetName?: boolean;
    customSheetName?: string;
    lockRows?: boolean;
    lockColumns?: boolean;
    rowOffset?: number;
    colOffset?: number;
    height?: number;
    width?: number;
}

/**
 * Returns the A1 notation of a range.
 */
export function getA1Notation({
    mappedRange,
    includeSheetName = false,
    customSheetName,
    lockRows = false,
    lockColumns = false,
    rowOffset = 0,
    colOffset = 0,
    height,
    width,
}: A1NotationParams): string {
    const adjustedRange = offsetGridRange({ origin: mappedRange.namedRange.range, rowOffset, colOffset, height, width });

    // Missing start indices in GridRange imply 0
    const startRowIndex = adjustedRange.startRowIndex ?? 0;
    const startColIndex = adjustedRange.startColumnIndex ?? 0;

    // Check if endings are explicitly defined
    const hasEndRow = adjustedRange.endRowIndex !== undefined;
    const hasEndCol = adjustedRange.endColumnIndex !== undefined;

    // Resolve indices (0-based) to 1-based coordinates
    const startRow = startRowIndex + 1;
    const endRow = hasEndRow ? adjustedRange.endRowIndex : undefined; // endRowIndex is exclusive, so value matches 1-based inclusive end

    const startColLetter = getColumnLetter(startColIndex);
    // biome-ignore lint/style/noNonNullAssertion: If `hasEndCol` we know it's not undefined
    const endColLetter = hasEndCol ? getColumnLetter(adjustedRange.endColumnIndex! - 1) : undefined;

    // Build lock prefixes
    const rowPrefix = lockRows ? "$" : "";
    const colPrefix = lockColumns ? "$" : "";

    // Pre-construct the individual notation parts
    const sColStr = `${colPrefix}${startColLetter}`;
    const sRowStr = `${rowPrefix}${startRow}`;
    const eColStr = hasEndCol ? `${colPrefix}${endColLetter}` : "";
    const eRowStr = hasEndRow ? `${rowPrefix}${endRow}` : "";

    let a1Notation = "";

    // Formulate the A1 notation based on unbounded state combinations
    if (!hasEndRow && !hasEndCol) {
        // Fallback: Fully unbounded at both ends (e.g., starting at C3 and going to bottom-right).
        // Requires sheet boundaries to explicitly cap.
        const gridProps = mappedRange.sheet.properties?.gridProperties;
        const maxRows = gridProps?.rowCount ?? 1000;
        const maxCols = gridProps?.columnCount ?? 1000;
        const maxColLetter = getColumnLetter(maxCols - 1);

        a1Notation = `${sColStr}${sRowStr}:${colPrefix}${maxColLetter}${rowPrefix}${maxRows}`;
    } else if (!hasEndRow) {
        // Unbounded rows (goes to bottom), bounded columns
        if (startRowIndex === 0) {
            // Entire column(s) -> e.g., A:A or A:B
            a1Notation = `${sColStr}:${eColStr}`;
        } else {
            // Starts at specific row, goes to bottom -> e.g., C3:C
            a1Notation = `${sColStr}${sRowStr}:${eColStr}`;
        }
    } else if (!hasEndCol) {
        // Unbounded columns (goes to right), bounded rows
        if (startColIndex === 0) {
            // Entire row(s) -> e.g., 1:1 or 1:2
            a1Notation = `${sRowStr}:${eRowStr}`;
        } else {
            // Starts at specific column, goes to right -> e.g., C3:3
            a1Notation = `${sColStr}${sRowStr}:${eRowStr}`;
        }
    } else {
        // Fully bounded (standard ranges)
        const startA1 = `${sColStr}${sRowStr}`;
        const endA1 = `${eColStr}${eRowStr}`;

        if (startA1 === endA1) {
            a1Notation = startA1; // Single cell (e.g., A1)
        } else {
            a1Notation = `${startA1}:${endA1}`; // Standard bounded range (e.g., A1:B2)
        }
    }

    // Append sheet name if requested
    if (includeSheetName) {
        if (customSheetName) {
            return `'${customSheetName}'!${a1Notation}`;
        }
        if (mappedRange.sheet.properties?.title) {
            const escapedTitle = mappedRange.sheet.properties.title.replace(/'/g, "''");
            return `'${escapedTitle}'!${a1Notation}`;
        }
    }

    return a1Notation;
}
