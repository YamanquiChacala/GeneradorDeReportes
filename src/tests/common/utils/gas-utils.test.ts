import { getColumnLetter } from "../../../common/utils/gas-utils";

describe("googleAPI Type Utilities", () => {
    describe("getColumnLetter", () => {
        it("should correctly map a single letter", () => {
            expect(getColumnLetter(1 - 1)).toBe("A");
            expect(getColumnLetter(13 - 1)).toBe("M");
            expect(getColumnLetter(26 - 1)).toBe("Z");
        });

        it("should correctly map multiple letters", () => {
            expect(getColumnLetter(1 * 26 + 1 - 1)).toBe("AA");
            expect(getColumnLetter(2 * 26 + 6 - 1)).toBe("BF");
            expect(getColumnLetter(26 * 26 + 26 - 1)).toBe("ZZ");
            expect(getColumnLetter(4 * 26 * 26 + 26 + 26 - 1)).toBe("DAZ");
        });

        it("should ignore negative numbers", () => {
            expect(getColumnLetter(-1)).toBe("");
            expect(getColumnLetter(-10000)).toBe("");
        });
    });
});