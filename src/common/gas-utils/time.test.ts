import { getEpochDate, getSheetsDate } from ".";

describe("Time", () => {
    describe("getSheetsDate", () => {
        it("converts Unix epoch 0 (Jan 1, 1970) to the Sheets epoch offset", () => {
            expect(getSheetsDate(0)).toBe(25569);
        });

        it("converts exactly one day after Unix epoch to the correct Sheets date", () => {
            const oneDayInMs = 86400000;
            expect(getSheetsDate(oneDayInMs)).toBe(25570);
        });

        it("converts a known modern Unix epoch to the correct Sheets date", () => {
            // Jan 1, 2024 = 1704067200000 ms
            expect(getSheetsDate(1704067200000)).toBe(45292);
        });

        it("handles negative Unix epochs (dates before 1970)", () => {
            // Dec 30, 1899 = -2209161600000 ms (Sheets Epoch 0)
            expect(getSheetsDate(-2209161600000)).toBe(0);
        });
    });

    describe("getEpochDate", () => {
        it("converts the Sheets epoch offset back to Unix epoch 0", () => {
            expect(getEpochDate(25569)).toBe(0);
        });

        it("converts exactly one day after the Sheets epoch offset to one day in ms", () => {
            const oneDayInMs = 86400000;
            expect(getEpochDate(25570)).toBe(oneDayInMs);
        });

        it("converts a known modern Sheets date to the correct Unix epoch", () => {
            // Jan 1, 2024 = 45292 in Sheets
            expect(getEpochDate(45292)).toBe(1704067200000);
        });

        it("handles Sheets date 0 (Dec 30, 1899)", () => {
            expect(getEpochDate(0)).toBe(-2209161600000);
        });
    });

    describe("Bidirectional Conversion", () => {
        it("returns the exact original Unix epoch after converting to Sheets and back", () => {
            const originalUnixEpoch = 1718323200000; // Arbitrary timestamp
            const sheetsDate = getSheetsDate(originalUnixEpoch);
            const convertedBack = getEpochDate(sheetsDate);

            expect(convertedBack).toBe(originalUnixEpoch);
        });

        it("returns the exact original Sheets date after converting to Unix and back", () => {
            const originalSheetsDate = 48000; // Arbitrary future Sheets date
            const unixEpoch = getEpochDate(originalSheetsDate);
            const convertedBack = getSheetsDate(unixEpoch);

            expect(convertedBack).toBe(originalSheetsDate);
        });
    });
});
