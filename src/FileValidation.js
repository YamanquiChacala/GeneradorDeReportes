/** @enum {string} */
const FileType = {
    INIT: "MontessoriChacalaSchoolReportInitialization",
    REPORT: "MontessoriChacalaSchoolReport",
}

const FileValidation = {
    KEY: "fileType",

    /**
     * Marks a Drive file as the given FileType
     * 
     * @param {string} fileId 
     * @param {FileType} fileType 
     */
    setFileType(fileId, fileType) {
        Drive?.Files.update({
            appProperties: {
                [FileValidation.KEY]: fileType
            }
        }, fileId);
    },

    /**
     * Gets the FileType of the given Drive file.
     * 
     * @param {string} [fileId] 
     * @returns {FileType | undefined}
     */
    getFileType(fileId) {
        if (!fileId) return undefined;
        const fileData = Drive?.Files.get(fileId, { fields: 'appProperties' });

        return fileData?.appProperties?.[FileValidation.KEY]
    },

    /**
     * Returns wheater the file has the given FileType.
     * 
     * @param {string} fileId 
     * @param {FileType} fileType 
     * @returns {boolean}
     */
    isFileType(fileId, fileType) {
        const resultFileType = FileValidation.getFileType(fileId);

        return fileType === resultFileType;
    },

    /**
     * Removes 
     * 
     * @param {string} fileId 
     */
    removeFileType(fileId) {
        Drive?.Files.update({
            appProperties: {
                [FileValidation.KEY]: null
            }
        }, fileId);
    },
}