import * as GasEnums from "./api-types";

describe("GAS Enums Validation", () => {
    // GasEnums represents the entire file.
    // Object.entries(GasEnums) gives us arrays like: ["PasteType", { PASTE_NORMAL: "PASTE_NORMAL", ... }]
    Object.entries(GasEnums).forEach(([enumName, enumObject]) => {
        describe(enumName, () => {
            // We use the dynamically grabbed enumObject to run our key-value check
            it.each(Object.entries(enumObject))("'%s' key should exactly match its string value", (key, value) => {
                expect(key).toBe(value);
            });
        });
    });
});
