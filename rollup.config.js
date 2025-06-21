import resolve from "@rollup/plugin-node-resolve";
import commonjs from "@rollup/plugin-commonjs";
import typescript from "@rollup/plugin-typescript";
import dts from "rollup-plugin-dts";
import terser from "@rollup/plugin-terser";

import { createRequire } from "module";
const require = createRequire(import.meta.url);
const { dependencies } = require("./package.json");

export default [
    // JS build
    {
        input: "src/index.ts",
        output: [
            { file: "dist/index.esm.js", format: "esm", sourcemap: true },
            { file: "dist/index.cjs.js", format: "cjs", sourcemap: true },
        ],
        plugins: [resolve(), commonjs(), typescript({ tsconfig: "./tsconfig.json" }), terser()],
        external: [...Object.keys(dependencies || {})],
    },

    {
        input: "src/index.ts",
        output: {
            file: "dist/index.umd.js",
            format: "umd",
            name: "toXlsx",
            globals: {
                exceljs: "ExcelJS",
                "file-saver": "saveAs",
            },
            sourcemap: true,
        },
        plugins: [resolve({ browser: true }), commonjs(), typescript(), terser()],
    },

    // Type declarations
    {
        input: "dist/types/index.d.ts",
        output: [{ file: "dist/index.d.ts", format: "es" }],
        plugins: [dts()],
    },
];
