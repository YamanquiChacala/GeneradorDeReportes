import fs from "node:fs";

const CSS_FILE = "templates/report.css";
const TS_FILE = "src/templates/report.css.ts";

try {
    // 1. Read the CSS file
    const cssContent = fs.readFileSync(CSS_FILE, "utf-8");

    // 2. Compress the CSS by removing newlines and extra spaces
    const minifiedCss = cssContent
        .replace(/\r?\n|\r/g, "") // Remove all line breaks
        .replace(/\s{2,}/g, " ") // Collapse multiple spaces into a single space
        .trim();

    // 3. Construct the TypeScript file content
    const tsContent = `import { Base64Fonts } from "../common/utils/base64-constants";

export const REPORT_CSS_STYLE = \`${minifiedCss}\`;`;

    // 4. Write the output file
    fs.writeFileSync(TS_FILE, tsContent, "utf-8");
    console.log(`✅ Successfully generated ${TS_FILE} from ${CSS_FILE}`);
} catch (error) {
    console.error(`❌ Error generating CSS TS file:`, error.message);
    process.exit(1);
}
