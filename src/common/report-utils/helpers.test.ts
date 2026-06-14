import type { DeepPartial } from "../../testing/helpers";
import { calculateAssistanceTrimRanges } from "./helpers";
import type { FrozenArea, ReportPersistentData } from "./types";

describe("Report Utils. Helpers", () => {
    describe("calculateAssistanceTrimRanges", () => {
        it("correctly splits the calendar", () => {
            const mockArea: FrozenArea = {
                rows: 0,
                cols: 0,
            };
            const mockData: DeepPartial<ReportPersistentData> = {
                configData: { dates: [10, 100, 200, 500] },
                calendar: [1, 9, 10, 11, 20, 30, 50, 99, 100, 101, 150, 199, 200, 201, 300, 400, 499, 500, 501, 600, 700],
            };
            const result = calculateAssistanceTrimRanges(mockData as ReportPersistentData, mockArea);

            expect(result).toEqual({
                trim1: { start: 2, end: 8 },
                trim2: { start: 9, end: 12 },
                trim3: { start: 13, end: 17 },
            });
        });

        it("offsets the ranges by the number of frozen columns", () => {
            const mockArea: FrozenArea = {
                rows: 0,
                cols: 3, // Offset by 3
            };
            const mockData: DeepPartial<ReportPersistentData> = {
                configData: { dates: [10, 100, 200, 500] },
                calendar: [1, 9, 10, 11, 20, 30, 50, 99, 100, 101, 150, 199, 200, 201, 300, 400, 499, 500, 501, 600, 700],
            };
            const result = calculateAssistanceTrimRanges(mockData as ReportPersistentData, mockArea);

            expect(result).toEqual({
                trim1: { start: 5, end: 11 }, // Originally 2 to 8
                trim2: { start: 12, end: 15 }, // Originally 9 to 12
                trim3: { start: 16, end: 20 }, // Originally 13 to 17
            });
        });

        it("returns {-1, -1} for trimesters with no matching calendar entries", () => {
            const mockArea: FrozenArea = { rows: 0, cols: 0 };
            const mockData: DeepPartial<ReportPersistentData> = {
                configData: { dates: [10, 20, 30, 40] },
                // Calendar has entries matching limits, but completely skips the 20-30 gap (Trim 2)
                calendar: [5, 10, 40, 50],
            };
            const result = calculateAssistanceTrimRanges(mockData as ReportPersistentData, mockArea);

            expect(result).toEqual({
                // Trim 1 (10 to 20): start is UB(9)=1, end is UB(20)-1 = 2-1 = 1
                trim1: { start: 1, end: 1 },
                // Trim 2 (20 to 30): start is UB(20)=2, end is UB(30)-1 = 2-1 = 1. (2 <= 1 is false)
                trim2: { start: -1, end: -1 },
                // Trim 3 (30 to 40): start is UB(30)=2, end is UB(40)-1 = 3-1 = 2
                trim3: { start: 2, end: 2 },
            });
        });

        it("returns {-1, -1} for all trimesters when the calendar is completely empty", () => {
            const mockArea: FrozenArea = { rows: 0, cols: 5 };
            const mockData: DeepPartial<ReportPersistentData> = {
                configData: { dates: [10, 100, 200, 500] },
                calendar: [],
            };
            const result = calculateAssistanceTrimRanges(mockData as ReportPersistentData, mockArea);

            // getUpperBoundIndex returns -1.
            // format(-1, -1 - 1) -> -1 <= -2 evaluates to false
            expect(result).toEqual({
                trim1: { start: -1, end: -1 },
                trim2: { start: -1, end: -1 },
                trim3: { start: -1, end: -1 },
            });
        });

        it("returns {-1, -1} when all dates are strictly earlier than the first calendar entry", () => {
            const mockArea: FrozenArea = { rows: 0, cols: 0 };
            const mockData: DeepPartial<ReportPersistentData> = {
                configData: { dates: [10, 20, 30, 40] },
                calendar: [100, 200, 300],
            };
            const result = calculateAssistanceTrimRanges(mockData as ReportPersistentData, mockArea);

            // getUpperBoundIndex returns 0 for all.
            // format(0, 0 - 1) -> 0 <= -1 evaluates to false
            expect(result).toEqual({
                trim1: { start: -1, end: -1 },
                trim2: { start: -1, end: -1 },
                trim3: { start: -1, end: -1 },
            });
        });

        it("returns {-1, -1} when all dates are strictly later than the last calendar entry", () => {
            const mockArea: FrozenArea = { rows: 0, cols: 0 };
            const mockData: DeepPartial<ReportPersistentData> = {
                configData: { dates: [1000, 2000, 3000, 4000] },
                calendar: [10, 20, 30],
            };
            const result = calculateAssistanceTrimRanges(mockData as ReportPersistentData, mockArea);

            // getUpperBoundIndex returns -1 for all.
            // format(-1, -1 - 1) -> -1 <= -2 evaluates to false
            expect(result).toEqual({
                trim1: { start: -1, end: -1 },
                trim2: { start: -1, end: -1 },
                trim3: { start: -1, end: -1 },
            });
        });
    });
});
