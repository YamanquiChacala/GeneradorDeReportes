import { gcd, getStep, getUpperBoundIndex } from "./math";

describe("Math Utilities", () => {
    describe("gcd (Greatest Common Divisor)", () => {
        it("should calculate the correct GCD for standard positive integers", () => {
            expect(gcd(10, 5)).toBe(5);
            expect(gcd(14, 21)).toBe(7);
            expect(gcd(48, 18)).toBe(6);
            expect(gcd(100, 10)).toBe(10);
        });

        it("should return 1 when the numbers are coprime", () => {
            expect(gcd(17, 3)).toBe(1);
            expect(gcd(25, 9)).toBe(1);
            expect(gcd(8, 15)).toBe(1);
        });

        it("should return the number itself when both arguments are the same", () => {
            expect(gcd(7, 7)).toBe(7);
            expect(gcd(42, 42)).toBe(42);
        });

        it("should handle cases where one of the parameters is 0", () => {
            expect(gcd(10, 0)).toBe(10);
            expect(gcd(0, 5)).toBe(5);
            // Even 0,0 technically returns 0 in this algorithm
            expect(gcd(0, 0)).toBe(0);
        });

        it("should fall back to returning 1 for zero or negative inputs", () => {
            // For 0: target is 0. k1=0 (fails > 0 check), k2=0 (fails < n check). Loop ends.
            expect(getStep(0)).toBe(1);

            // For -5: target is -3. The loop condition `delta <= target` (0 <= -3) immediately fails.
            expect(getStep(-5)).toBe(1);
        });
    });

    describe("getStep", () => {
        it("should return 0 or 1 for very small values of n", () => {
            expect(getStep(1)).toBe(0);
            expect(getStep(2)).toBe(1);
            expect(getStep(3)).toBe(1);
        });

        it("should return the largest coprime near n/2 for even numbers", () => {
            // For 4: target is 2. gcd(4, 2) is 2. Falls back to 1.
            expect(getStep(4)).toBe(1);

            // For 6: target is 3. gcd(6, 3) is 3, gcd(6, 2) is 2, gcd(6, 4) is 2. Falls back to 1.
            expect(getStep(6)).toBe(1);

            // For 8: target is 4. gcd(8, 4) is 4. Next closest is 3. gcd(8, 3) is 1.
            expect(getStep(8)).toBe(3);

            // For 10: target is 5. gcd(10, 5) is 5, gcd(10, 4) is 2, gcd(10, 3) is 1.
            expect(getStep(10)).toBe(3);
        });

        it("should return the largest coprime near n/2 for odd numbers", () => {
            // For 5: target is 2. gcd(5, 2) is 1.
            expect(getStep(5)).toBe(2);

            // For 7: target is 3. gcd(7, 3) is 1.
            expect(getStep(7)).toBe(3);

            // For 9: target is 4. gcd(9, 4) is 1.
            expect(getStep(9)).toBe(4);
        });

        it("should reliably generate a step that visits all elements in a cycle (Coprime validation)", () => {
            // A quick programmatic check to ensure the step value generated
            // is ALWAYS entirely coprime with the input n, ensuring a full cyclic iteration.
            const testSizes = [12, 15, 20, 32, 50, 100];

            for (const size of testSizes) {
                const step = getStep(size);
                expect(gcd(size, step)).toBe(1);
            }
        });
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
});
