import type { FileType } from "../constants";
import { buildFieldsMask } from "../gas-utils";
import { FILE_VALIDATION_KEY } from ".";

export function setFileType(fileId: string, fileType: FileType) {
    Drive?.Files.update(
        {
            appProperties: {
                [FILE_VALIDATION_KEY]: fileType,
            },
        },
        fileId,
    );
}

/**
 * Gets the FileType of the given Drive file.
 */
export function getFileType(fileId: string): FileType | undefined {
    if (!fileId) return undefined;
    const fileData = Drive?.Files.get(fileId, { fields: buildFieldsMask<GoogleAppsScript.Drive_v3.Drive.V3.Schema.File>("appProperties") });

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
    );
}
