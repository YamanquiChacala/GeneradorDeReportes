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
     * @param {string} fileId 
     * @returns {FileType | undefined}
     */
    getFileType(fileId) {
        const fileData = Drive?.Files.get(fileId, { fields: 'appProperties' });

        return fileData?.appProperties?.[FileValidation.KEY]
    }
}