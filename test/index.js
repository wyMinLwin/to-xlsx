import { exportToXlsx } from "../dist/index.esm.js";
async function main() {
    exportToXlsx({
        fileName: "test-export",
        data: [
            { name: "wai", age: 21, job: "Software Developer" },
            { name: "L", age: 21, job: "Detective" },
            { name: "Wainus", age: 34, job: "Professional Chess Player" },
        ],
        columnHeaders: { name: "username" },
        columnSizes: { job: 40 },
        sheetsBy: {
            namePattern: "Age - $key",
            key: "age",
        },
        title: {
            text: "Text Title",
            bg: "37bf5c",
            color: "000000",
            fontSize: 17,
        },
        columnsOrder: ["age"],
        columnsStyle: {
            bg: "11a1fa",
            color: "000000",
            fontSize: 14,
        },
    });
}
main();
