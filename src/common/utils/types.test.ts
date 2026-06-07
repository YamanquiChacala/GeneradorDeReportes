import { Icon, isColorKey } from "./types";

describe("Types for General Utils", () => {
    describe("Icon", () => {
        it.each(Object.entries(Icon))("'%s' should be a properly formatted Iconify string", (_key, icon) => {
            // Must be a non-empty string
            expect(icon.trim().length).toBeGreaterThan(0);

            // Cannot start or end with a slash
            expect(icon.endsWith("/")).toBe(false);
            expect(icon.startsWith("/")).toBe(false);

            // Must contain exactly one slash to separate the prefix (e.g., 'mdi') from the name (e.g., 'home')
            const parts = icon.split("/");
            expect(parts.length).toBe(2);

            // Neither the prefix nor the name can be empty
            expect(parts[0]?.length).toBeGreaterThan(0);
            expect(parts[1]?.length).toBeGreaterThan(0);
        });
    });

    describe("isColorKey", () => {
        it("returns true for valid CSS color keys", () => {
            expect(isColorKey("aliceblue")).toBe(true);
            expect(isColorKey("red")).toBe(true);
            expect(isColorKey("chartreuse")).toBe(true);
            expect(isColorKey("rebeccapurple")).toBe(true);
        });

        it("returns false for invalid color names or random strings", () => {
            expect(isColorKey("notacolor")).toBe(false);
            expect(isColorKey("")).toBe(false);
            expect(isColorKey("123456")).toBe(false);
        });

        it("returns false for hex codes (since they are values, not keys)", () => {
            expect(isColorKey("#FFFFFF")).toBe(false);
            expect(isColorKey("#000000")).toBe(false);
            expect(isColorKey("#FF0000")).toBe(false);
        });

        it("returns false for uppercase/mixed case valid strings (keys are strictly lowercase)", () => {
            expect(isColorKey("AliceBlue")).toBe(false);
            expect(isColorKey("RED")).toBe(false);
            expect(isColorKey("RebeccaPurple")).toBe(false);
        });
    });
});
