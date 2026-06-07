import { buildSetupFileUpdatePayload } from "./requests";
import type { SetupFileData } from "./types";

jest.mock("../gas-parts", () => ({
    SetupSheetSchema: {
        sheets: {
            groupData: {
                ranges: {
                    groupName: "MockedGroupNameRange",
                    attendancePerClass: "MockedAttendanceRange",
                    averagePerField: "MockedAverageRange",
                    dates: "MockedDatesRange",
                },
            },
        },
    },
}));

describe("Setup Utils. Requests", () => {
    describe("buildSetupFileUpdatePayload", () => {
        it("should correctly map simple properties and format dates into a GAS ValueRange payload", () => {
            // Arrange: Prepare dummy data using explicit UTC times to ensure consistency
            const mockInitData: SetupFileData = {
                folderId: "false",
                groupName: "Test Group Alpha",
                attendancePerClass: false,
                averagePerField: true,
                dateStart: new Date("2026-01-15T00:00:00Z").getTime(),
                dateEndTrimester1: new Date("2026-04-15T00:00:00Z").getTime(),
                dateEndTrimester2: new Date("2026-07-15T00:00:00Z").getTime(),
                dateEnd: new Date("2026-12-20T00:00:00Z").getTime(),
            };

            // Act: Run the target function
            const result = buildSetupFileUpdatePayload(mockInitData);

            // Assert: Verify the structural output length
            expect(result).toHaveLength(4);

            // Assert: Validate individual range payloads
            expect(result[0]).toEqual({
                range: "MockedGroupNameRange",
                values: [["Test Group Alpha"]],
            });

            expect(result[1]).toEqual({
                range: "MockedAttendanceRange",
                values: [[false]],
            });

            expect(result[2]).toEqual({
                range: "MockedAverageRange",
                values: [[true]],
            });

            // Assert: Validate that dates were transformed into 'YYYY-MM-DD' 2D arrays
            expect(result[3]).toEqual({
                range: "MockedDatesRange",
                values: [["2026-01-15"], ["2026-04-15"], ["2026-07-15"], ["2026-12-20"]],
            });
        });
    });
});
