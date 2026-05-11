import * as Utils from "./text-utils";

// Helper to create UTC timestamps safely
function utcMs(year: number, monthIndex: number, day: number): number {
    return Date.UTC(year, monthIndex, day);
}

describe("Text Utils Module", () => {
    describe("sanitizeFileName()", () => {
        it("should remove OS-prohibited characters", () => {
            const input = 'My <Cool> : "File" / \\ | ? * .txt';
            expect(Utils.sanitizeFileName(input)).toBe("My Cool File .txt");
        });

        it("should use default fallback if input becomes empty after sanitization", () => {
            expect(Utils.sanitizeFileName(' < > : " / \\ | ? * ')).toBe("Grupo");
        });

        it("should use custom fallback if provided", () => {
            expect(Utils.sanitizeFileName("", "DefaultFile")).toBe("DefaultFile");
        });

        it("should handle undefined or whitespace-only inputs", () => {
            expect(Utils.sanitizeFileName(undefined)).toBe("Grupo");
            expect(Utils.sanitizeFileName("   ")).toBe("Grupo");
            expect(Utils.sanitizeFileName(undefined, "Fallback")).toBe("Fallback");
        });

        it("should normalize unicode characters (NFKC) and keep Spanish accents intact", () => {
            // \uFB01 is the ligature "ﬁ"
            // \u0065\u0301 is a decomposed "e" + "´"
            // NFKC should convert the ligature to "fi" and compose the accent to "é"
            const input = "file\uFB01name \u0065\u0301";
            expect(Utils.sanitizeFileName(input)).toBe("filefiname é");
        });
    });

    describe("sanitizeSheetName()", () => {
        it("should replace quotes, backticks, and apostrophes with spaces", () => {
            expect(Utils.sanitizeSheetName('O\'Conner `Math` "Data"')).toBe("O Conner Math Data");
        });

        it("should remove Spreadsheet-prohibited characters", () => {
            expect(Utils.sanitizeSheetName("Data [2024] : * ? / \\ |")).toBe("Data 2024");
        });

        it("should collapse double spaces created by quote replacements", () => {
            expect(Utils.sanitizeSheetName("A'B\"'C'\"")).toBe("A B C"); // A space B space C
        });

        it("should enforce 31 character limit for Excel compatibility", () => {
            const longName = "This Is A Very Long Spreadsheet Name That Exceeds Thirty One Characters";
            const result = Utils.sanitizeSheetName(longName);
            expect(result.length).toBeLessThanOrEqual(31);
            expect(result).toBe("This Is A Very Long Spreadsheet");
        });

        it("should trim properly if truncation leaves a trailing space at character 31", () => {
            // A 30-character string, followed by a space, followed by an X.
            // "123456789012345678901234567890 X"
            // Substring to 31 grabs the trailing space, so the final .trim() must remove it.
            const input = "123456789012345678901234567890 X";
            const result = Utils.sanitizeSheetName(input);
            expect(result.length).toBe(30);
            expect(result).toBe("123456789012345678901234567890");
        });

        it("should return an empty string if all characters are stripped", () => {
            // Note: Google Sheets requires tab names to be >= 1 character.
            // It's good to know this returns an empty string so upstream logic can handle it if needed.
            expect(Utils.sanitizeSheetName("[:*?/\\|]")).toBe("");
        });

        it("should handle an initially empty string gracefully", () => {
            expect(Utils.sanitizeSheetName("")).toBe("");
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
                // Assuming CssColorMap maps these properly
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

    describe("formatDateRange", () => {
        it("formats a range within the same year", () => {
            const start = utcMs(2025, 0, 1); // 1 Jan 2025
            const end = utcMs(2025, 1, 15); // 15 Feb 2025

            expect(Utils.formatDateRange(start, end)).toBe("Del 1 de enero al 15 de febrero de 2025");
        });

        it("formats a range across different years", () => {
            const start = utcMs(2024, 11, 31); // 31 Dec 2024
            const end = utcMs(2025, 0, 1); // 1 Jan 2025

            expect(Utils.formatDateRange(start, end)).toBe("Del 31 de diciembre de 2024 al 1 de enero de 2025");
        });

        it("formats a range within the same month", () => {
            const start = utcMs(2025, 4, 10); // 10 May 2025
            const end = utcMs(2025, 4, 20); // 20 May 2025

            expect(Utils.formatDateRange(start, end)).toBe("Del 10 de mayo al 20 de mayo de 2025");
        });

        it("formats a single-day range", () => {
            const date = utcMs(2025, 6, 7); // 7 Jul 2025

            expect(Utils.formatDateRange(date, date)).toBe("Del 7 de julio al 7 de julio de 2025");
        });

        it("uses UTC dates instead of local timezone", () => {
            const start = Date.UTC(2025, 0, 1, 23, 0, 0);
            const end = Date.UTC(2025, 0, 2, 1, 0, 0);

            expect(Utils.formatDateRange(start, end)).toBe("Del 1 de enero al 2 de enero de 2025");
        });
    });
});
