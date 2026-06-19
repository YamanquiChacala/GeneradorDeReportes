import { typedEntries, zip } from "./helpers";

describe("Helper Utilities", () => {
    describe("typedEntries", () => {
        it("should return an array of key-value pairs for a standard object", () => {
            const obj = { a: 1, b: "two", c: true };
            const result = typedEntries(obj);

            expect(result).toEqual([
                ["a", 1],
                ["b", "two"],
                ["c", true],
            ]);
        });

        it("should return an empty array for an empty object", () => {
            expect(typedEntries({})).toEqual([]);
        });

        it("should convert number keys to string keys", () => {
            // In JavaScript, object keys are always strings (or Symbols).
            // Object.entries inherently casts numbers to strings.
            const obj = { 1: "one", 2: "two" };
            const result = typedEntries(obj);

            expect(result).toEqual([
                ["1", "one"],
                ["2", "two"],
            ]);
        });

        it("should only include own enumerable properties (ignoring inherited)", () => {
            const proto = { inherited: "value" };
            const obj = Object.create(proto);
            obj.own = "property";

            const result = typedEntries(obj);

            // It should catch 'own' but completely ignore 'inherited'
            expect(result).toEqual([["own", "property"]]);
        });

        it("should process arrays by treating indexes as string keys", () => {
            const arr = ["first", "second"];
            const result = typedEntries(arr);

            expect(result).toEqual([
                ["0", "first"],
                ["1", "second"],
            ]);
        });
    });
    describe("zip generator", () => {
        it("zips two arrays of the exact same length", () => {
            const arr1 = [1, 2, 3];
            const arr2 = ["a", "b", "c"];

            // Spread the generator into a concrete array for easy testing
            const result = [...zip(arr1, arr2)];

            expect(result).toEqual([
                [1, "a"],
                [2, "b"],
                [3, "c"],
            ]);
        });

        it("stops at the length of the shorter array (first array is shorter)", () => {
            const arr1 = [1, 2];
            const arr2 = ["a", "b", "c", "d"];

            const result = [...zip(arr1, arr2)];

            expect(result).toEqual([
                [1, "a"],
                [2, "b"],
            ]);
        });

        it("stops at the length of the shorter array (second array is shorter)", () => {
            const arr1 = [1, 2, 3, 4];
            const arr2 = ["a", "b"];

            const result = [...zip(arr1, arr2)];

            expect(result).toEqual([
                [1, "a"],
                [2, "b"],
            ]);
        });

        it("returns an empty iterable if both arrays are empty", () => {
            const arr1: number[] = [];
            const arr2: string[] = [];

            const result = [...zip(arr1, arr2)];

            expect(result).toEqual([]);
        });

        it("returns an empty iterable if only one array is empty", () => {
            const arr1 = [1, 2, 3];
            const arr2: string[] = [];

            const result = [...zip(arr1, arr2)];

            expect(result).toEqual([]);
        });

        it("skips indices where values are explicitly undefined", () => {
            // Because of our strict noUncheckedIndexedAccess safeguard,
            // the generator intentionally drops pairs where an element is undefined.
            const arr1: (number | undefined)[] = [1, undefined, 3];
            const arr2 = ["a", "b", "c"];

            const result = [...zip(arr1, arr2)];

            // Notice the pair for index 1 is entirely skipped, not yielded as [undefined, "b"]
            expect(result).toEqual([
                [1, "a"],
                [3, "c"],
            ]);
        });
    });
});
