import gasPlugin from "@gas-plugin/unplugin/esbuild";
import { build } from "esbuild";

build({
    entryPoints: ["src/main.ts"],
    bundle: true,
    outfile: "dist/Code.js",
    format: "esm",
    target: "es2019", // TODO: Try 2020, but it could break things.
    jsx: "automatic",
    jsxImportSource: "preact",
    loader: {
        ".png": "dataurl",
        ".jpg": "dataurl",
        ".jpeg": "dataurl",
        ".svg": "dataurl",
        ".woff": "dataurl",
        ".woff2": "dataurl",
        ".ttf": "dataurl",
    },
    plugins: [
        gasPlugin({
            manifest: "src/appsscript.json",
            autoGlobals: false,
        }),
    ],
}).catch(() => process.exit(1));
