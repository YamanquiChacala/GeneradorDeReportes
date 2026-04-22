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
                color: "blue",
                width: 24,
                height: 24,
                box: true,
            });
            expect(result.url).toBe("https://api.iconify.design/logos/google.svg?color=%230000FF&width=24&height=24&box=1");
            expect(result.name).toBe("google");
        });
    });
});
