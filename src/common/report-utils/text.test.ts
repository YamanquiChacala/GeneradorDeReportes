import { formatDateRange } from "../utils";
import { generatePeriodString } from "./text";
import { Period, type ReportPersistentData } from "./types";

// Mock the external date formatting utility
jest.mock("../utils", () => {
    // 1. Grab the real module
    const originalModule = jest.requireActual("../utils");

    // 2. Return the real module, but overwrite the one function you want to mock
    return {
        __esModule: true,
        ...originalModule,
        formatDateRange: jest.fn(),
    };
});

describe("Report Utils. Period Utilities", () => {
    beforeEach(() => {
        jest.clearAllMocks();
    });

    describe("generatePeriodString", () => {
        // Build a complete mock object to satisfy strict TS without `as ReportPersistentData`
        const baseMockData: ReportPersistentData = {
            configData: {
                attendancePerClass: true,
                averagePerField: true,
                dates: [10, 30, 50, 70],
                subjectGradingWeights: [1, 1, 1],
            },
            protectedSections: {
                habilities: false,
                comments: false,
                trimesters: [false, false, false],
            },
            academicFields: [],
            subjects: [],
            students: [],
            calendar: [5, 10, 15, 25, 30, 45, 50, 60, 70, 80],
        };

        it("should correctly resolve dates and call formatDateRange for Period FIRST", () => {
            (formatDateRange as jest.Mock).mockReturnValue("Formatted Date");

            const result = generatePeriodString(baseMockData, Period.FIRST);

            expect(formatDateRange).toHaveBeenCalledWith(10, 30);
            expect(result).toBe("Formatted Date");
        });

        it("should correctly resolve dates and call formatDateRange for Period SECOND", () => {
            (formatDateRange as jest.Mock).mockReturnValue("Formatted Date");

            // Period SECOND: startBoundary = 29, endBoundary = 50
            // startIndex = 4 (value 30), endIndex = 6 (value 50) -> actualEnd is index 5 (value 45)
            const result = generatePeriodString(baseMockData, Period.SECOND);

            expect(formatDateRange).toHaveBeenCalledWith(45, 50);
            expect(result).toBe("Formatted Date");
        });

        it("should correctly resolve dates and call formatDateRange for Period THIRD", () => {
            (formatDateRange as jest.Mock).mockReturnValue("Formatted Date");

            // Period THIRD: startBoundary = 49, endBoundary = 70
            // startIndex = 6 (value 50), endIndex = 8 (value 70) -> actualEnd is index 7 (value 60)
            const result = generatePeriodString(baseMockData, Period.THIRD);

            expect(formatDateRange).toHaveBeenCalledWith(60, 70);
            expect(result).toBe("Formatted Date");
        });

        it("should return undefined if the start date cannot be resolved", () => {
            const outOfBoundsData: ReportPersistentData = {
                ...baseMockData,
                configData: {
                    ...baseMockData.configData,
                    dates: [100, 200, 300, 400],
                },
                calendar: [10, 20, 30], // Calendar is completely before config dates
            };

            const result = generatePeriodString(outOfBoundsData, Period.FIRST);

            expect(result).toBeUndefined();
            expect(formatDateRange).not.toHaveBeenCalled();
        });

        it("should return undefined if start is greater than end (crossed dates)", () => {
            const crossedData: ReportPersistentData = {
                ...baseMockData,
                configData: {
                    ...baseMockData.configData,
                    dates: [30, 10, 50, 70], // Malformed config where period 0 end is before start
                },
                calendar: [5, 10, 15, 25, 30, 45],
            };

            const result = generatePeriodString(crossedData, Period.FIRST);

            expect(result).toBeUndefined();
            expect(formatDateRange).not.toHaveBeenCalled();
        });

        it("should return undefined if the calendar is completely empty", () => {
            const emptyData: ReportPersistentData = {
                ...baseMockData,
                calendar: [],
            };

            const result = generatePeriodString(emptyData, Period.FIRST);

            expect(result).toBeUndefined();
            expect(formatDateRange).not.toHaveBeenCalled();
        });
    });
});
