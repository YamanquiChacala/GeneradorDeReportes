import * as Utils from "../../../common/utils/image";

describe("Image Utils Module", () => {
    describe("iconifyUrl()", () => {
        test("should generate a basic URL and extract the name", () => {
            const result = Utils.iconifyUrl({ iconName: "mdi/home" });
            expect(result.url).toBe("https://api.iconify.design/mdi/home.svg");
            expect(result.name).toBe("home");
        });

        test("should properly sanitize a hex color starting with #", () => {
            const result = Utils.iconifyUrl({ iconName: "mdi/home", color: "#FF0000" });
            expect(result.url).toBe("https://api.iconify.design/mdi/home.svg?color=%23FF0000");
        });

        test("should construct URL with all optional parameters", () => {
            const result = Utils.iconifyUrl({
                iconName: "logos/google",
                color: "blue", // Assuming webColor transforms this to #0000FF
                width: 24,
                height: 24,
                box: true,
            });
            expect(result.url).toBe("https://api.iconify.design/logos/google.svg?color=%230000FF&width=24&height=24&box=1");
            expect(result.name).toBe("google");
        });

        // --- New Edge Cases Below ---

        test("should handle iconName without slashes", () => {
            const result = Utils.iconifyUrl({ iconName: "simpleicon" });
            expect(result.url).toBe("https://api.iconify.design/simpleicon.svg");
            expect(result.name).toBe("simpleicon");
        });

        test("should omit box parameter if box is explicitly false", () => {
            const result = Utils.iconifyUrl({ iconName: "mdi/home", box: false });
            expect(result.url).toBe("https://api.iconify.design/mdi/home.svg");
            // If it falsely triggered, it would have "?box=1" at the end
        });

        test("should correctly join parameters if only width is provided", () => {
            const result = Utils.iconifyUrl({ iconName: "mdi/home", width: 48 });
            // Checks that the '?' is added correctly for a single param
            expect(result.url).toBe("https://api.iconify.design/mdi/home.svg?width=48");
        });

        test("should gracefully handle missing or undefined color", () => {
            const result = Utils.iconifyUrl({ iconName: "mdi/home", color: undefined, height: 32 });
            expect(result.url).toBe("https://api.iconify.design/mdi/home.svg?height=32");
        });
    });
});
