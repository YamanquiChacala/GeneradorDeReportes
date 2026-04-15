import { FileType } from "../../../common/enums";
import * as fileValidation from "../../../common/utils/fileValidation";

export function testFileValidation() {
    let passed = 0;
    let failed = 0;
    const logs: string[] = [];

    // Simple logging and assertion helpers

    const log = (msg: string) => logs.push(msg);

    const assert = (condition: boolean, msg: string) => {
        if (condition) {
            passed++;
            log(`✅ PASSED: ${msg}`);
        } else {
            failed++;
            log(`❌ FAILED: ${msg}`);
        }
    };

    let tempFileId = null;

    try {
        log("--- Starting FileValidation Integration Test ---");

        // 1. Setup: Create a temporary file in Drive using DriveApp
        // (DriveApp is easier for quick creation/teardown than the Advanced Drive Service)
        const tempFile = DriveApp.createFile("Test_Montessori_Report.txt", "Temporary test file");
        tempFileId = tempFile.getId();
        log(`Created temp file: ${tempFileId}`);

        // 2. Test Initial State
        assert(fileValidation.getFileType(tempFileId) === undefined, "File should initially have an undefined FileType");
        assert(fileValidation.isFileType(tempFileId, FileType.SETUP) === false, "isFileType should return false for an unmarked file");

        // 3. Test Setting the File Type
        fileValidation.setFileType(tempFileId, FileType.SETUP);

        // Wait briefly to ensure Drive API propagates the property (sometimes it's not strictly synchronous)
        Utilities.sleep(500);

        assert(fileValidation.getFileType(tempFileId) === FileType.SETUP, "getFileType should return INIT after setting it");
        assert(fileValidation.isFileType(tempFileId, FileType.SETUP) === true, "isFileType should return true for INIT after setting it");
        assert(fileValidation.isFileType(tempFileId, FileType.REPORT) === false, "isFileType should still return false for REPORT");

        // 4. Test Setting the File Type
        fileValidation.setFileType(tempFileId, FileType.REPORT);

        // Wait briefly to ensure Drive API propagates the property (sometimes it's not strictly synchronous)
        Utilities.sleep(500);

        assert(fileValidation.getFileType(tempFileId) === FileType.REPORT, "getFileType should return REPORT after setting it");
        assert(fileValidation.isFileType(tempFileId, FileType.SETUP) === false, "isFileType should return false for INIT after setting it");
        assert(fileValidation.isFileType(tempFileId, FileType.REPORT) === true, "isFileType should return true for REPORT");

        // 5. Test Removing the File Type
        fileValidation.removeFileType(tempFileId);
        Utilities.sleep(500);

        // Note: Setting an appProperty to null in the Drive API deletes the key.
        // Therefore, retrieving it should yield undefined again.
        assert(fileValidation.getFileType(tempFileId) === undefined, "getFileType should return undefined after removal");

        assert(fileValidation.getFileType("") === undefined, "getFileType should run silently on invalid ID");
    } catch (e) {
        const errorMessage = e instanceof Error ? e.message : String(e);
        log(`🚨 CRITICAL ERROR: ${errorMessage}`);
        failed++;
    } finally {
        // 5. Teardown: Always clean up the temporary file, even if tests fail
        if (tempFileId) {
            try {
                DriveApp.getFileById(tempFileId).setTrashed(true);
                log(`Successfully deleted temp file: ${tempFileId}`);
            } catch (cleanupErr) {
                const errorMessage = cleanupErr instanceof Error ? cleanupErr.message : String(cleanupErr);
                log(`⚠️ FAILED TO CLEAN UP temp file: ${errorMessage}`);
            }
        }
    }

    // Print summary and return for clasp
    log(`--- Test Complete: ${passed} Passed, ${failed} Failed ---`);
    const finalOutput = logs.join("\n");
    console.log(finalOutput);

    return {
        success: failed === 0,
        passed,
        failed,
        logs,
    };
}
