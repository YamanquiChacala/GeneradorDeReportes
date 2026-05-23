import { MORE_THAN_A_YEAR } from "../constants";
import { type CalendarDates, calculateCalendarDates, calculateCalendarGrid, validateDates } from ".";

describe("Setup Date Utilities", () => {
    describe("validateDates", () => {
        it("returns null on dates in order", () => {
            expect(validateDates([1000, 2000, 3000, 4000])).toBeNull();
        });

        it("fails with dates not in order", () => {
            expect(validateDates([2000, 1000, 3000, 4000])).toBe("Fechas en desorden."); // First out of order
            expect(validateDates([1000, 3000, 2000, 4000])).toBe("Fechas en desorden."); // Middle flipped
            expect(validateDates([1000, 2000, 4000, 3000])).toBe("Fechas en desorden."); // End flipped
            expect(validateDates([4000, 3000, 2000, 1000])).toBe("Fechas en desorden."); // Completely reversed
            expect(validateDates([1000, 1000, 2000, 3000])).toBe("Fechas en desorden."); // Equal dates (not strictly greater)
            expect(validateDates([1000, 3000, 2000, 1001 + MORE_THAN_A_YEAR])).toBe("Fechas en desorden.");
        });

        it("fails with falsy dates (missing elements)", () => {
            // Using 0 to simulate a falsy/missing date epoch
            expect(validateDates([0, 2000, 3000, 4000])).toBe("Faltan fechas.");
        });

        it("fails with too distant dates", () => {
            const start = 1000;
            const end = start + MORE_THAN_A_YEAR + 1;
            expect(validateDates([start, 2000, 3000, end])).toBe("Calendario demasiado largo.");
        });
    });

    describe("calculateCalendarDates", () => {
        const sunJan4 = Date.UTC(2026, 0, 4); // Sunday
        const wedJan7 = Date.UTC(2026, 0, 7); // Wednesday
        const wedJan14 = Date.UTC(2026, 0, 14); // Wednesday
        const wedJan21 = Date.UTC(2026, 0, 21); // Wednesday
        const wedJan28 = Date.UTC(2026, 0, 28); // Wednesday
        const satJan10 = Date.UTC(2026, 0, 10); // Saturday

        it("throws error on invalid dates", () => {
            expect(() => calculateCalendarDates([wedJan14, wedJan7, wedJan21, wedJan28])).toThrow("Fechas en desorden.");
        });

        it("snaps back to first Sunday before start", () => {
            const result = calculateCalendarDates([wedJan7, wedJan14, wedJan21, wedJan28]);
            expect(result.calStart).toBe(sunJan4);
        });

        it("snaps forward to first Saturday after end", () => {
            const result = calculateCalendarDates([sunJan4, wedJan7, wedJan14, wedJan21]);
            expect(result.calEnd).toBe(Date.UTC(2026, 0, 24)); // Jan 21 is Wed, following Sat is Jan 24
        });

        it("calculates the right number of days", () => {
            const result = calculateCalendarDates([wedJan7, wedJan14, wedJan21, wedJan28]);
            // Jan 4 (Sun) to Jan 31 (Sat) = exactly 28 days
            expect(result.totalDays).toBe(28);
        });

        it("calculates the right number of weeks (totalRows)", () => {
            const result = calculateCalendarDates([wedJan7, wedJan14, wedJan21, wedJan28]);
            // Math.ceil(28/7) + 1 = 5 rows
            expect(result.totalRows).toBe(5);
        });

        it("works with four consecutive days", () => {
            const mon = Date.UTC(2026, 0, 5);
            const tue = Date.UTC(2026, 0, 6);
            const wed = Date.UTC(2026, 0, 7);
            const thu = Date.UTC(2026, 0, 8);

            const result = calculateCalendarDates([mon, tue, wed, thu]);
            expect(result.calStart).toBe(sunJan4);
            expect(result.calEnd).toBe(satJan10);
            expect(result.totalDays).toBe(7);
            expect(result.totalRows).toBe(2);
        });

        it("works with a single week perfectly aligned to Sun/Sat bounds", () => {
            const mon = Date.UTC(2026, 0, 5);
            const tue = Date.UTC(2026, 0, 6);

            const result = calculateCalendarDates([sunJan4, mon, tue, satJan10]);
            expect(result.calStart).toBe(sunJan4);
            expect(result.calEnd).toBe(satJan10);
            expect(result.totalDays).toBe(7);
            expect(result.totalRows).toBe(2);
        });
    });

    describe("calculateCalendarGrid", () => {
        // Mock a strict period for predictable grid testing
        const baseDates: CalendarDates = {
            dateStart: Date.UTC(2026, 3, 20), // Monday, Apr 20, 2026
            dateTrimester1: Date.UTC(2026, 3, 24), // Friday, Apr 24, 2026
            dateTrimester2: Date.UTC(2026, 3, 29), // Wednesday, Apr 29, 2026
            dateEnd: Date.UTC(2026, 4, 5), // Tuesday, May 5, 2026
            calStart: Date.UTC(2026, 3, 19), // Sunday, Apr 19, 2026
            calEnd: Date.UTC(2026, 4, 9), // Saturday, May 9, 2026
            totalDays: 21,
            totalRows: 4,
        };

        it("generates the correct number of weeks and days", () => {
            const grid = calculateCalendarGrid(baseDates);
            expect(grid.weeks).toHaveLength(3);

            for (const week of grid.weeks) {
                expect(week.days).toBeDefined();
                expect(week.days).toHaveLength(7);
            }
        });

        it("correctly identifies month transitions based on Wednesdays", () => {
            const grid = calculateCalendarGrid(baseDates);
            // Week 1 Wed = Apr 22 -> Month 3 (April)
            // Week 1 Wed = Apr 29 -> Month 3 (April)
            // Week 2 Wed = May 6 -> Month 4 (May)
            expect(grid.monthBlocks).toHaveLength(2);

            const firstBlock = grid.monthBlocks[0];
            expect(firstBlock?.monthIndex).toBe(3);
            expect(firstBlock?.startRow).toBe(1);
            expect(firstBlock?.endRow).toBe(3);

            const secondBlock = grid.monthBlocks[1];
            expect(secondBlock?.monthIndex).toBe(4);
            expect(secondBlock?.startRow).toBe(3);
            expect(secondBlock?.endRow).toBe(4);
        });

        it("assigns the correct day types based on date bounds", () => {
            const grid = calculateCalendarGrid(baseDates);
            const week1 = grid.weeks[0];
            const week2 = grid.weeks[1];
            expect(week1).toBeDefined();
            expect(week2).toBeDefined();

            if (week1) {
                const sunday = week1.days[0]; // Apr 19 (Out of user bounds)
                expect(sunday?.dayType).toBe("rest");
                expect(sunday?.inBounds).toBe(false);

                const monday = week1.days[1]; // Apr 20 (dateStart -> tr1)
                expect(monday?.dayType).toBe("trimester1");
                expect(monday?.inBounds).toBe(true);

                const friday = week1.days[5]; // Apr 24 (dateTrimester1 boundary -> tr1)
                expect(friday?.dayType).toBe("trimester1");
                expect(friday?.inBounds).toBe(true);
            }

            if (week2) {
                const wednesday = week2.days[3]; // Apr 29 (> end trimester2)
                expect(wednesday?.dayType).toBe("trimester2");

                const thursday = week2.days[4]; // Apr 30 (> dateTrimester2 -> tr3)
                expect(thursday?.dayType).toBe("trimester3");
            }
        });

        it("forces weekends to be rest days even when in bounds", () => {
            const grid = calculateCalendarGrid(baseDates);
            const week1 = grid.weeks[0];

            if (week1) {
                const saturday = week1.days[6]; // Aprl 25 (in bounds, but is a weekend)

                expect(saturday?.isWeekday).toBe(false);
                expect(saturday?.inBounds).toBe(true);
                expect(saturday?.dayType).toBe("rest");
            }
        });
    });
});
