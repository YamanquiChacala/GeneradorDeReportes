import { FileType } from "./enums";

const key = "fileType";

export function setFileType(fileId: string, fileType: FileType) {
    Drive?.Files.update({
        appProperties: {
            [key]: fileType
        }
    }, fileId);
}

/**
 * Gets the FileType of the given Drive file.
 */
export function getFileType(fileId: string): FileType | undefined {
    if (!fileId) return undefined;
    const fileData = Drive?.Files.get(fileId, { fields: 'appProperties' });

    return fileData?.appProperties?.[key]
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
    Drive?.Files.update({
        appProperties: {
            [key]: null
        }
    }, fileId);
}
