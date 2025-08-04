import { exportToXlsx } from "../dist/index.esm.js";
async function main() {
    const sampleData = [
        { name: "John", age: 18, department: "IT", salary: 45000 },
        { name: "Jane", age: 25, department: "HR", salary: 55000 },
        { name: "Bob", age: 17, department: "IT", salary: 35000 },
        { name: "Alice", age: 30, department: "Marketing", salary: 60000 },
        { name: "Charlie", age: 22, department: "IT", salary: 50000 },
        { name: "Diana", age: 28, department: "HR", salary: 58000 },
        { name: "Eve", age: 19, department: "Marketing", salary: 42000 },
        { name: "Frank", age: 35, department: "IT", salary: 70000 },
    ];
    try {
        exportToXlsx({
            data: sampleData,
            fileName: "employees-grouped-by-age",
            title: {
                text: "Employee Report - Grouped by Age",
                bg: "4472C4",
                color: "FFFFFF",
                fontSize: 18,
                border: {
                    all: {
                        style: "thick",
                        color: "000000",
                    },
                },
            },

            columnsStyle: {
                bg: "70AD47",
                color: "FFFFFF",
                fontSize: 12,
            },
            columnHeaders: {
                name: "Employee Name",
                age: "Age",
                department: "Department",
                salary: "Annual Salary",
            },
            columnSizes: {
                name: 30,
                age: 10,
                department: 20,
                salary: 15,
            },
            columnsMerge: [
                {
                    keys: {
                        startColumn: "name",
                        endColumn: "age",
                    },
                    columnName: "Personal",
                },
                {
                    keys: {
                        startColumn: "department",
                        endColumn: "salary",
                    },
                    columnName: "Company",
                },
            ],
            groupBy: {
                conditions: [
                    {
                        label: "Under 20",
                        filter: (item) => item.age < 20,
                    },
                    {
                        label: "20-30",
                        filter: (item) => item.age >= 20 && item.age < 30,
                    },
                    {
                        label: "30+",
                        filter: (item) => item.age >= 30,
                    },
                ],
                subtitleStyle: {
                    bg: "BDD7EE",
                    color: "000000",
                    fontSize: 14,
                    border: {
                        bottom: { style: "medium", color: "0070C0" },
                        left: { style: "thin", color: "0070C0" },
                    },
                },
                showSubtotals: true,
                subtotalStyle: {
                    bg: "E6F3FF",
                    color: "000000",
                    fontSize: 11,
                    border: {
                        all: { style: "thin", color: "0070C0" },
                    },
                },
            },
            totals: {
                columns: ["salary"],
                showGrandTotal: true,
                subtotalLabel: "Age Group Subtotal",
                grandTotalLabel: "Company Total",
                operations: {
                    salary: "sum",
                },
                grandTotalStyle: {
                    bg: "4472C4",
                    color: "FFFFFF",
                    fontSize: 13,
                    border: {
                        all: { style: "thick", color: "000000" },
                    },
                },
            },
        });
        console.log("✅ Successfully exported!");
    } catch {
        console.log("❎ Failed to export excel!");
    }
}
main();
