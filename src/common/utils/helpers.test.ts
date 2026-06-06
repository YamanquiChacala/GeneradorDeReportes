import { typedEntries } from "./helpers";

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
