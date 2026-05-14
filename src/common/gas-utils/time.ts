import { MS_PER_DAY, SHEETS_EPOCH_OFFSET } from "../constants";

//TODO: Make tests for these 2:
/**
 * Converts a Unix Epoch into a Sheets Epoch.
 */
export function getSheetsDate(epoch: number): number {
    return epoch / MS_PER_DAY + SHEETS_EPOCH_OFFSET;
}

/**
 * Converts a Sheets Epock into a Unix Epoch.
 */

export function getEpochDate(sheetsDate: number): number {
    return (sheetsDate - SHEETS_EPOCH_OFFSET) * MS_PER_DAY;
}
