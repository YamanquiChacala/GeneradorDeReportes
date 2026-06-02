import { Urls } from "./types";

describe("Types for GAS Utils", () => {
    describe("Urls", () => {
        it.each(Object.entries(Urls))("'%s' should be a valid HTTP/HTTPS URL", (_key, url) => {
            // Using Node's built-in URL constructor to validate the string
            expect(() => new URL(url)).not.toThrow();

            // Ensure it uses a secure or valid web protocol
            const parsedUrl = new URL(url);
            expect(["http:", "https:"].includes(parsedUrl.protocol)).toBe(true);
        });
    });
});
