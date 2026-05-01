import { Colors, CssColorMap, FileType, Icon, Templates, Urls } from "./enums";

describe("Enums Validation", () => {
    describe("FileType", () => {
        it.each(Object.entries(FileType))("'%s' should be a non-empty string", (_key, value) => {
            expect(typeof value).toBe("string");
            expect(value.trim().length).toBeGreaterThan(0);
        });
    });

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

    describe("Urls", () => {
        it.each(Object.entries(Urls))("'%s' should be a valid HTTP/HTTPS URL", (_key, url) => {
            // Using Node's built-in URL constructor to validate the string
            expect(() => new URL(url)).not.toThrow();

            // Ensure it uses a secure or valid web protocol
            const parsedUrl = new URL(url);
            expect(["http:", "https:"].includes(parsedUrl.protocol)).toBe(true);
        });
    });

    describe("Templates", () => {
        it.each(Object.entries(Templates))("'%s' should be a non-empty string", (_key, value) => {
            expect(typeof value).toBe("string");
            expect(value.trim().length).toBeGreaterThan(0);

            // Templates shouldn't have spaces if they are meant to match file names or script IDs
            expect(value.includes(" ")).toBe(false);
        });
    });

    describe("Colors", () => {
        it.each(Object.entries(Colors))("'%s' should be a valid 6-digit Hex color code", (_key, color) => {
            // Must exactly match the pattern: # followed by exactly 6 letters/numbers
            expect(color).toMatch(/^#[A-Fa-f0-9]{6}$/);
        });
    });

    describe("Web Colors", () => {
        it.each(Object.entries(CssColorMap))("'%s' should be a valid 6-digit Hex color code", (_key, color) => {
            // Must exactly match the pattern: # followed by exactly 6 letters/numbers
            expect(color).toMatch(/^#[A-Fa-f0-9]{6}$/);
        });
    });
});
