import { Colors, Urls } from "./types";

describe("GAS Utils, Types & Enums", () => {
    describe("Urls Enums", () => {
        it.each(Object.entries(Urls))("'%s' should be a valid HTTP/HTTPS URL", (_key, url) => {
            // Using Node's built-in URL constructor to validate the string
            expect(() => new URL(url)).not.toThrow();

            // Ensure it uses a secure or valid web protocol
            const parsedUrl = new URL(url);
            expect(["http:", "https:"].includes(parsedUrl.protocol)).toBe(true);
        });
    });

    describe("Colors Enums", () => {
        it.each(Object.entries(Colors))("'%s' should be a valid 6-character hex color code", (_key, hex) => {
            // Matches standard 6-character hex codes starting with '#'
            const hexRegex = /^#[0-9a-fA-F]{6}$/;
            expect(hex).toMatch(hexRegex);
        });
    });
});
