import type { Icon } from "../constants";
import { iconifyUrl } from "./iconify";

describe("Image Utils Module", () => {
    describe("iconifyUrl()", () => {
        it("should generate a basic URL and extract the name", () => {
            const result = iconifyUrl({ iconName: "mdi/home" as unknown as Icon });
            expect(result.url).toBe("https://api.iconify.design/mdi/home.svg");
            expect(result.name).toBe("home");
        });

        it("should properly sanitize a hex color starting with #", () => {
            const result = iconifyUrl({
                iconName: "mdi/home" as unknown as Icon,
                color: "#FF0000",
            });
            expect(result.url).toBe("https://api.iconify.design/mdi/home.svg?color=%23FF0000");
        });

        it("should construct URL with all optional parameters", () => {
            const result = iconifyUrl({
                iconName: "logos/google" as unknown as Icon,
                color: "#0000FF",
                width: 24,
                height: 24,
                box: true,
            });
            expect(result.url).toBe("https://api.iconify.design/logos/google.svg?color=%230000FF&width=24&height=24&box=1");
            expect(result.name).toBe("google");
        });

        it("should handle iconName without slashes", () => {
            const result = iconifyUrl({ iconName: "simpleicon" as unknown as Icon });
            expect(result.url).toBe("https://api.iconify.design/simpleicon.svg");
            expect(result.name).toBe("simpleicon");
        });

        it("should omit box parameter if box is explicitly false", () => {
            const result = iconifyUrl({
                iconName: "mdi/home" as unknown as Icon,
                box: false,
            });
            expect(result.url).toBe("https://api.iconify.design/mdi/home.svg");
        });

        it("should correctly join parameters if only width is provided", () => {
            const result = iconifyUrl({
                iconName: "mdi/home" as unknown as Icon,
                width: 48,
            });
            expect(result.url).toBe("https://api.iconify.design/mdi/home.svg?width=48");
        });

        it("should gracefully handle missing or undefined color", () => {
            const result = iconifyUrl({
                iconName: "mdi/home" as unknown as Icon,
                color: undefined,
                height: 32,
            });
            expect(result.url).toBe("https://api.iconify.design/mdi/home.svg?height=32");
        });

        it("should gracefully handle malformed color", () => {
            const result = iconifyUrl({
                iconName: "mdi/home" as unknown as Icon,
                color: "#hello",
                width: 32,
            });
            expect(result.url).toBe("https://api.iconify.design/mdi/home.svg?width=32");
        });
    });
});
