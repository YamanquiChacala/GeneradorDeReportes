import { formatDateRange } from "../utils";
import type { ReportPersistentData } from ".";
import { generatePeriodString, getUpperBoundIndex } from "."; // Update path as needed

// Mock the external date formatting utility
jest.mock("../utils", () => ({
    formatDateRange: jest.fn(),
}));

describe("Period Utilities", () => {
    beforeEach(() => {
        jest.clearAllMocks();
    });

    describe("getUpperBoundIndex", () => {
        const calendar = [10, 20, 30, 40, 50];

        it("should return the index of the first element strictly greater than the threshold", () => {
            expect(getUpperBoundIndex(calendar, 25)).toBe(2); // 30 is at index 2
            expect(getUpperBoundIndex(calendar, 9)).toBe(0); // 10 is at index 0
        });

        it("should skip exact matches and find the strictly greater value", () => {
            expect(getUpperBoundIndex(calendar, 30)).toBe(3); // 40 is at index 3
        });

        it("should return calendar.length if the threshold is greater than or equal to the max element", () => {
            expect(getUpperBoundIndex(calendar, 50)).toBe(5); // threshold equals max
            expect(getUpperBoundIndex(calendar, 100)).toBe(5); // threshold exceeds max
        });

        it("should return 0 if the calendar is empty", () => {
            expect(getUpperBoundIndex([], 10)).toBe(0);
        });
    });

    describe("generatePeriodString", () => {
        // Sample configData dates for boundaries
        // Period 0: 10 to 30
        // Period 1: 30 to 50
        // Period 2: 50 to 70
        const mockConfigData = {
            dates: [10, 30, 50, 70],
        };

        const mockCalendar = [5, 10, 15, 25, 30, 45, 50, 60, 70, 80];

        const mockData = {
            configData: mockConfigData,
            calendar: mockCalendar,
        } as ReportPersistentData;

        it("should correctly resolve dates and call formatDateRange for Period 0", () => {
            (formatDateRange as jest.Mock).mockReturnValue("Formatted Date");

            const result = generatePeriodString(mockData, 0);

            expect(formatDateRange).toHaveBeenCalledWith(10, 30);
            expect(result).toBe("Formatted Date");
        });

        it("should correctly resolve dates and call formatDateRange for Period 1", () => {
            (formatDateRange as jest.Mock).mockReturnValue("Formatted Date");

            // Period 1: startBoundary = 29, endBoundary = 50
            // startIndex = 4 (value 30), endIndex = 6 (value 50) -> actualEnd is index 5 (value 45)
            const result = generatePeriodString(mockData, 1);

            expect(formatDateRange).toHaveBeenCalledWith(45, 50);
            expect(result).toBe("Formatted Date");
        });

        it("should correctly resolve dates and call formatDateRange for Period 2", () => {
            (formatDateRange as jest.Mock).mockReturnValue("Formatted Date");

            // Period 2: startBoundary = 49, endBoundary = 70
            // startIndex = 6 (value 50), endIndex = 8 (value 70) -> actualEnd is index 7 (value 60)
            const result = generatePeriodString(mockData, 2);

            expect(formatDateRange).toHaveBeenCalledWith(60, 70);
            expect(result).toBe("Formatted Date");
        });

        it("should return undefined if the start date cannot be resolved", () => {
            const outOfBoundsData = {
                configData: { dates: [100, 200, 300, 400] },
                calendar: [10, 20, 30], // Calendar is completely before config dates
            } as ReportPersistentData;

            const result = generatePeriodString(outOfBoundsData, 0);

            expect(result).toBeUndefined;
            expect(formatDateRange).not.toHaveBeenCalled();
        });

        it("should return undefined if start is greater than end (crossed dates)", () => {
            const crossedData = {
                configData: { dates: [30, 10, 50, 70] }, // Malformed config where period 0 end is before start
                calendar: [5, 10, 15, 25, 30, 45],
            } as ReportPersistentData;

            const result = generatePeriodString(crossedData, 0);

            expect(result).toBeUndefined;
            expect(formatDateRange).not.toHaveBeenCalled();
        });

        it("should return undefined if the calendar is completely empty", () => {
            const emptyData = {
                configData: mockConfigData,
                calendar: [],
            } as unknown as ReportPersistentData;

            const result = generatePeriodString(emptyData, 0);

            expect(result).toBeUndefined;
            expect(formatDateRange).not.toHaveBeenCalled();
        });
    });
});
