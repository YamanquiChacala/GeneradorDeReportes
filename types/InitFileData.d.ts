declare global {
    interface InitFileData {
        folderId: string,
        groupName: string,
        attendancePerClass: boolean,
        averagePerField: boolean,
        dateStart: number,
        dateEndTrimester1: number,
        dateEndTrimester2: number,
        dateEnd: number,
    }
}

export { }