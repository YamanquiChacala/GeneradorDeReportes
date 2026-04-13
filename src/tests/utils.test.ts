import * as Utils from '../common/utils'

describe('Utils Module', () => {

    describe('iconifyUrl()', () => {
        test('should generate a basic URL and extract the name', () => {
            const result = Utils.iconifyUrl({ iconName: 'mdi/home' });
            expect(result.url).toBe('https://api.iconify.design/mdi/home.svg');
            expect(result.name).toBe('home');
        });

        test('should properly sanitize a hex color starting with #', () => {
            const result = Utils.iconifyUrl({ iconName: 'mdi/home', color: '#FF0000' });
            expect(result.url).toBe('https://api.iconify.design/mdi/home.svg?color=%23FF0000');
        });

        test('should construct URL with all optional parameters', () => {
            const result = Utils.iconifyUrl({
                iconName: 'logos/google',
                color: 'blue',
                width: 24,
                height: 24,
                box: true
            });
            expect(result.url).toBe('https://api.iconify.design/logos/google.svg?color=blue&width=24&height=24&box=1');
            expect(result.name).toBe('google');
        });
    });

    describe('sanitizeFileName()', () => {
        test('should remove OS-prohibited characters', () => {
            const input = 'My <Cool> : "File" / \\ | ? * .txt';
            expect(Utils.sanitizeFileName(input)).toBe('My Cool File .txt');
        });

        test('should use default fallback if input becomes empty after sanitization', () => {
            expect(Utils.sanitizeFileName(' < > : " / \\ | ? * ')).toBe('Grupo');
        });

        test('should use custom fallback if provided', () => {
            expect(Utils.sanitizeFileName('', 'DefaultFile')).toBe('DefaultFile');
        });
    });

    describe('sanitizeSheetName()', () => {
        test('should replace quotes and apostrophes with spaces', () => {
            expect(Utils.sanitizeSheetName("O'Conner \"Math\"")).toBe('O Conner Math');
        });

        test('should remove Spreadsheet-prohibited characters', () => {
            expect(Utils.sanitizeSheetName('Data [2024] : * ? / \\ |')).toBe('Data 2024');
        });

        test('should collapse double spaces created by quote replacements', () => {
            expect(Utils.sanitizeSheetName("A'B\"\'C\'\"")).toBe('A B C'); // A space B space C
        });

        test('should enforce 31 character limit for Excel compatibility', () => {
            const longName = 'This Is A Very Long Sheet Name That Exceeds Thirty One Characters';
            const result = Utils.sanitizeSheetName(longName);
            expect(result.length).toBeLessThanOrEqual(31);
            expect(result).toBe('This Is A Very Long Sheet Name'); // Truncated and trimmed
        });
    });
});