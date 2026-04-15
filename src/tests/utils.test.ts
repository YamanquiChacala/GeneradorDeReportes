import * as Utils from "../common/utils";

describe("Utils Module", () => {
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
            expect(result.url).toBe("https://api.iconify.design/logos/google.svg?color=blue&width=24&height=24&box=1");
            expect(result.name).toBe("google");
        });
    });

    describe("sanitizeFileName()", () => {
        test("should remove OS-prohibited characters", () => {
            const input = 'My <Cool> : "File" / \\ | ? * .txt';
            expect(Utils.sanitizeFileName(input)).toBe("My Cool File .txt");
        });

        test("should use default fallback if input becomes empty after sanitization", () => {
            expect(Utils.sanitizeFileName(' < > : " / \\ | ? * ')).toBe("Grupo");
        });

        test("should use custom fallback if provided", () => {
            expect(Utils.sanitizeFileName("", "DefaultFile")).toBe("DefaultFile");
        });
    });

    describe("sanitizeSheetName()", () => {
        test("should replace quotes and apostrophes with spaces", () => {
            expect(Utils.sanitizeSheetName('O\'Conner "Math"')).toBe("O Conner Math");
        });

        test("should remove Spreadsheet-prohibited characters", () => {
            expect(Utils.sanitizeSheetName("Data [2024] : * ? / \\ |")).toBe("Data 2024");
        });

        test("should collapse double spaces created by quote replacements", () => {
            expect(Utils.sanitizeSheetName("A'B\"'C'\"")).toBe("A B C"); // A space B space C
        });

        test("should enforce 31 character limit for Excel compatibility", () => {
            const longName = "This Is A Very Long Sheet Name That Exceeds Thirty One Characters";
            const result = Utils.sanitizeSheetName(longName);
            expect(result.length).toBeLessThanOrEqual(31);
            expect(result).toBe("This Is A Very Long Sheet Name"); // Truncated and trimmed
        });
    });

    describe("flattenFormInputs()", () => {
        it("should safely return an empty object if inputs are undefined", () => {
            const result = Utils.flattenFormInputs(undefined);
            expect(result).toEqual({});
        });

        it("should extract standard string inputs", () => {
            const mockInputs = {
                groupName: { stringInputs: { value: ["Advanced Rust"] } },
                folderId: { stringInputs: { value: ["folder-12345"] } },
            };

            const result = Utils.flattenFormInputs(mockInputs);

            expect(result).toEqual({
                groupName: "Advanced Rust",
                folderId: "folder-12345",
            });
        });

        it('should convert a single "true" string into a boolean true', () => {
            const mockInputs = {
                attendancePerClass: { stringInputs: { value: ["true"] } },
                averagePerField: { stringInputs: { value: ["true"] } },
                // A normal string that happens to be "false" should NOT become a boolean
                // because GAS switches omit the field entirely when false.
                randomText: { stringInputs: { value: ["false"] } },
            };

            const result = Utils.flattenFormInputs(mockInputs);

            expect(result).toEqual({
                attendancePerClass: true,
                averagePerField: true,
                randomText: "false",
            });
        });

        it("should return an array of strings for multiselect checkboxes", () => {
            const mockInputs = {
                selectedDays: { stringInputs: { value: ["Monday", "Wednesday", "Friday"] } },
            };

            const result = Utils.flattenFormInputs(mockInputs);

            expect(result).toEqual({
                selectedDays: ["Monday", "Wednesday", "Friday"],
            });
        });

        it("should extract and parse Date and DateTime epoch timestamps into strict numbers", () => {
            const mockInputs = {
                // Google sometimes sends timestamps as strings
                dateStart: { dateInput: { msSinceEpoch: "1680307200000" } },
                // And sometimes as raw numbers depending on the event payload
                dateEnd: { dateTimeInput: { msSinceEpoch: 1680393600000 } },
            };

            const result = Utils.flattenFormInputs(mockInputs);

            expect(result).toEqual({
                dateStart: 1680307200000,
                dateEnd: 1680393600000,
            });

            // Explicitly verify the type coercion worked
            expect(typeof result["dateStart"]).toBe("number");
            expect(typeof result["dateEnd"]).toBe("number");
        });

        it("should fallback to the raw object if the data structure is unrecognized", () => {
            const mockInputs = {
                normalString: { stringInputs: { value: ["hello"] } },
                weirdFutureGoogleInput: { weirdInput: { data: "who knows" } },
                // biome-ignore lint/suspicious/noExplicitAny: For testing purposes
                nullInput: null as any, // Testing edge case where a key exists but value is null
            };

            const result = Utils.flattenFormInputs(mockInputs);

            expect(result).toEqual({
                normalString: "hello",
                weirdFutureGoogleInput: { weirdInput: { data: "who knows" } },
            });

            // Ensure the nullInput was safely skipped
            expect(result).not.toHaveProperty("nullInput");
        });
    });
});
