import { Icon } from "./types";

describe("Types for General Utils", () => {
    describe("Icon", () => {
        it.each(Object.entries(Icon))("'%s' should be a properly formatted Iconify string", (_key, icon) => {
            // Must be a non-empty string
            expect(icon.trim().length).toBeGreaterThan(0);

            // Cannot start or end with a slash
            expect(icon.endsWith("/")).toBe(false);
            expect(icon.startsWith("/")).toBe(false);

            // Must contain exactly one slash to separate the prefix (e.g., 'mdi') from the name (e.g., 'home')
            const parts = icon.split("/");
            expect(parts.length).toBe(2);

            // Neither the prefix nor the name can be empty
            expect(parts[0]?.length).toBeGreaterThan(0);
            expect(parts[1]?.length).toBeGreaterThan(0);
        });
    });
});
