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
});
