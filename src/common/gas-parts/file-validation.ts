import { buildFieldsMask, FILE_VALIDATION_KEY } from "../gas-utils";
import type { FileType } from "../gas-utils/types";

export function setFileType(fileId: string, fileType: FileType) {
    Drive?.Files.update(
        {
            appProperties: {
                [FILE_VALIDATION_KEY]: fileType,
            },
        },
        fileId,
        // biome-ignore lint/suspicious/noExplicitAny: Lack in @types/google-apps-script
        null as any,
        {
            supportsAllDrives: true,
        },
    );
}

/**
 * Gets the FileType of the given Drive file.
 */
export function getFileType(fileId: string): FileType | undefined {
    if (!fileId) return undefined;
    const fileData = Drive?.Files.get(fileId, { fields: buildFieldsMask<GoogleAppsScript.Drive_v3.Drive.V3.Schema.File>("appProperties"), supportsAllDrives: true });

    return fileData?.appProperties?.[FILE_VALIDATION_KEY];
}

/**
 * Returns wheater the file has the given FileType.
 */
export function isFileType(fileId: string, fileType: FileType): boolean {
    const resultFileType = getFileType(fileId);

    return fileType === resultFileType;
}

/**
 * Removes FileType mark
 */
export function removeFileType(fileId: string) {
    Drive?.Files.update(
        {
            appProperties: {
                [FILE_VALIDATION_KEY]: null,
            },
        },
        fileId,
        // biome-ignore lint/suspicious/noExplicitAny: Lack in @types/google-apps-script
        null as any,
        {
            supportsAllDrives: true,
        },
    );
}
