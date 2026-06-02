import { GasTestRunner } from "../../testing/gas-test-runner";
import * as fileValidation from "./file-validation";
import { FileType } from "./types";

export function testFileValidation() {
    const runner = new GasTestRunner();
    const { describe, test, expect, beforeAll, afterAll } = runner;

    describe("File Validation Integration Test", () => {
        let tempFileId: string;

        beforeAll(() => {
            const tempFile = DriveApp.createFile("Temp test file.txt", "Testing File Validation");
            tempFileId = tempFile.getId();
        });

        afterAll(() => {
            if (tempFileId) DriveApp.getFileById(tempFileId).setTrashed(true);
        });

        test("Initial State should have undefined FileType", () => {
            expect(fileValidation.getFileType(tempFileId)).toBeUndefined();
            expect(fileValidation.isFileType(tempFileId, FileType.SETUP)).toBe(false);
        });

        test("Setting the File Type to SETUP", () => {
            fileValidation.setFileType(tempFileId, FileType.SETUP);
            Utilities.sleep(500); // Wait for Drive API propagation

            expect(fileValidation.getFileType(tempFileId)).toBe(FileType.SETUP);
            expect(fileValidation.isFileType(tempFileId, FileType.SETUP)).toBeTruthy();
            expect(fileValidation.isFileType(tempFileId, FileType.REPORT)).toBeFalsy();
        });

        test("Setting the File Type to REPORT", () => {
            fileValidation.setFileType(tempFileId, FileType.REPORT);
            Utilities.sleep(500);

            expect(fileValidation.getFileType(tempFileId)).toBe(FileType.REPORT);
            expect(fileValidation.isFileType(tempFileId, FileType.SETUP)).toBeFalsy();
            expect(fileValidation.isFileType(tempFileId, FileType.REPORT)).toBeTruthy();
        });

        test("Removing the File Type should clear it", () => {
            fileValidation.removeFileType(tempFileId);
            Utilities.sleep(500);

            expect(fileValidation.getFileType(tempFileId)).toBeUndefined();
        });

        test("Invalid ID should run silently and return undefined", () => {
            expect(fileValidation.getFileType("")).toBeUndefined();
        });
    });

    return runner.execute();
}
