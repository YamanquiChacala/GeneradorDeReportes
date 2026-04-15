import * as Utils from "../../../common/utils/text";

describe("Text Utils Module", () => {
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

    describe("webColor", () => {
        describe("Hex color validation", () => {
            it("should return uppercase 6-digit hex for valid 6-digit hex inputs", () => {
                expect(Utils.webColor("#FFAA00")).toBe("#FFAA00");
                expect(Utils.webColor("#ffaa00")).toBe("#FFAA00");
                expect(Utils.webColor("#123456")).toBe("#123456");
            });

            it("should expand valid 3-digit hex to 6-digit uppercase hex", () => {
                expect(Utils.webColor("#FA0")).toBe("#FFAA00");
                expect(Utils.webColor("#fa0")).toBe("#FFAA00");
                expect(Utils.webColor("#123")).toBe("#112233");
            });

            it("should handle leading and trailing whitespace", () => {
                expect(Utils.webColor("  #AABBCC  ")).toBe("#AABBCC");
                expect(Utils.webColor("\t#aB3\n")).toBe("#AABB33");
            });

            it("should return null for invalid hex formats", () => {
                expect(Utils.webColor("FFAA00")).toBeNull(); // Missing #
                expect(Utils.webColor("#FFAA0")).toBeNull(); // 5 digits
                expect(Utils.webColor("#FFAA001")).toBeNull(); // 7 digits
                expect(Utils.webColor("#FGAA00")).toBeNull(); // Invalid hex characters
                expect(Utils.webColor("#12")).toBeNull(); // 2 digits
            });
        });

        describe("Named color to Hex translation", () => {
            it("should return correct 6-digit hex for valid named colors", () => {
                expect(Utils.webColor("red")).toBe("#FF0000");
                expect(Utils.webColor("black")).toBe("#000000");
                expect(Utils.webColor("  ALICEBLUE  ")).toBe("#F0F8FF");
                expect(Utils.webColor("YellowGreen")).toBe("#9ACD32");
            });

            it("should return null for invalid named colors", () => {
                expect(Utils.webColor("notacolor")).toBeNull();
                expect(Utils.webColor("lightblack")).toBeNull();
            });

            it("should return null for 'transparent' as it cannot map to a standard 6-digit hex without alpha", () => {
                expect(Utils.webColor("transparent")).toBeNull();
            });
        });

        describe("Edge cases", () => {
            it("should return null for empty string calls or undefined", () => {
                expect(Utils.webColor()).toBeNull();
                expect(Utils.webColor(undefined)).toBeNull();
                expect(Utils.webColor("")).toBeNull();
                expect(Utils.webColor("   ")).toBeNull();
            });
        });
    });
});
