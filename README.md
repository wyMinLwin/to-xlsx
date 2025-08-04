# to-xlsx

A powerful JavaScript/TypeScript library for exporting data to Excel XLSX files with advanced
formatting, grouping, and calculation features.

[![npm version](https://img.shields.io/npm/v/to-xlsx.svg)](https://www.npmjs.com/package/to-xlsx)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

## ✨ Features

- 📊 **Export Arrays to Excel** - Convert JavaScript/TypeScript arrays to XLSX files
- 🎨 **Advanced Styling** - Customize colors, fonts, backgrounds, and borders
- 📝 **Custom Headers & Titles** - Add custom column headers and report titles
- 📏 **Column Management** - Set custom widths, reorder, merge, or exclude columns
- 🗂️ **Data Grouping** - Group data with custom conditions and styled subtitles
- 📄 **Multi-Sheet Support** - Split data across multiple worksheets
- 🧮 **Calculations** - Automatic subtotals and grand totals with multiple operations
- 🎯 **TypeScript Support** - Full type safety and IntelliSense support

## 📦 Installation

```bash
# npm
npm install to-xlsx

# pnpm
pnpm add to-xlsx

# yarn
yarn add to-xlsx
```

## 🚀 Quick Start

```javascript
import { exportToXlsx } from "to-xlsx";

const data = [
    { name: "John", age: 28, department: "IT", salary: 45000 },
    { name: "Jane", age: 32, department: "HR", salary: 55000 },
    { name: "Bob", age: 25, department: "IT", salary: 35000 },
];

// Basic export
exportToXlsx({
    data,
    fileName: "employees",
});
```

## 📚 Usage Examples

### 🎨 Styled Export with Custom Headers

```javascript
exportToXlsx({
    data: employees,
    fileName: "employee-report",
    title: {
        text: "Employee Directory",
        bg: "4472C4",
        color: "FFFFFF",
        fontSize: 16,
        border: {
            all: { style: "thick", color: "000000" },
        },
    },
    columnHeaders: {
        name: "Full Name",
        age: "Age",
        department: "Department",
        salary: "Annual Salary",
    },
    columnSizes: {
        name: 25,
        age: 10,
        department: 20,
        salary: 15,
    },
    columnsStyle: {
        bg: "70AD47",
        color: "FFFFFF",
        fontSize: 12,
    },
});
```

### 🗂️ Data Grouping with Subtotals

```javascript
exportToXlsx({
    data: employees,
    fileName: "employees-by-age-group",
    groupBy: {
        conditions: [
            {
                label: "Young (Under 30)",
                filter: (item) => item.age < 30,
            },
            {
                label: "Mid-Career (30-40)",
                filter: (item) => item.age >= 30 && item.age < 40,
            },
            {
                label: "Senior (40+)",
                filter: (item) => item.age >= 40,
            },
        ],
        subtitleStyle: {
            bg: "BDD7EE",
            color: "000000",
            fontSize: 14,
            border: {
                bottom: { style: "medium", color: "0070C0" },
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
        subtotalLabel: "Group Subtotal",
        grandTotalLabel: "Total Company Payroll",
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
```

### 🧮 Advanced Calculations

```javascript
exportToXlsx({
    data: salesData,
    fileName: "sales-analysis",
    totals: {
        columns: ["quantity", "revenue", "profit"],
        showGrandTotal: true,
        operations: {
            quantity: "sum", // Total units sold
            revenue: "sum", // Total revenue
            profit: "avg", // Average profit margin
        },
        grandTotalLabel: "SUMMARY TOTALS",
    },
});
```

### 🔗 Column Merging

```javascript
exportToXlsx({
    data: employees,
    columnsMerge: [
        {
            keys: {
                startColumn: "firstName",
                endColumn: "lastName",
            },
            columnName: "Personal Info",
        },
        {
            keys: {
                startColumn: "department",
                endColumn: "salary",
            },
            columnName: "Work Details",
        },
    ],
});
```

### 📄 Multi-Sheet Export

```javascript
exportToXlsx({
    data: employees,
    sheetsBy: {
        key: "department",
        namePattern: "$key Department", // Creates sheets like "IT Department", "HR Department"
    },
});
```

## 📖 API Reference

### `exportToXlsx(props: Props<T>)`

Main function to export data to Excel.

#### Props

| Property         | Type                     | Description                          | Default         |
| ---------------- | ------------------------ | ------------------------------------ | --------------- |
| `data`           | `T[]`                    | Array of data objects to export      | **Required**    |
| `fileName`       | `string`                 | Output file name (without extension) | `"ExportSheet"` |
| `columnHeaders`  | `Record<string, string>` | Custom column headers                | `null`          |
| `columnSizes`    | `Record<string, number>` | Column widths                        | `null`          |
| `columnsStyle`   | `ColumnsStyleType`       | Header row styling                   | `null`          |
| `columnsOrder`   | `string[]`               | Custom column order                  | `null`          |
| `columnsMerge`   | `ColumnsMergeType`       | Merge column headers                 | `null`          |
| `excludeColumns` | `string[]`               | Columns to exclude                   | `null`          |
| `sheetsBy`       | `SheetsByType`           | Split into multiple sheets           | `null`          |
| `title`          | `TitleType`              | Report title configuration           | `null`          |
| `groupBy`        | `GroupByType<T>`         | Data grouping configuration          | `null`          |
| `totals`         | `TotalsType`             | Calculations configuration           | `null`          |

### Type Definitions

#### `TitleType`

```typescript
{
    text: string;
    bg?: string;           // Background color (hex)
    color?: string;        // Text color (hex)
    fontSize?: number;     // Font size
    border?: BorderType;   // Border styling
}
```

#### `GroupByType<T>`

```typescript
{
    conditions: GroupCondition<T>[];
    subtitleStyle?: {
        bg?: string;
        color?: string;
        fontSize?: number;
        border?: BorderType;
    };
    showSubtotals?: boolean;      // Enable subtotal rows
    subtotalStyle?: {             // Subtotal row styling
        bg?: string;
        color?: string;
        fontSize?: number;
        border?: BorderType;
    };
}
```

#### `TotalsType`

```typescript
{
    columns: string[];                    // Columns to calculate
    showGrandTotal?: boolean;            // Show grand total row
    subtotalLabel?: string;              // Subtotal row label
    grandTotalLabel?: string;            // Grand total row label
    operations?: {                       // Calculation operations
        [columnName: string]: 'sum' | 'avg' | 'count' | 'min' | 'max';
    };
    grandTotalStyle?: {                  // Grand total styling
        bg?: string;
        color?: string;
        fontSize?: number;
        border?: BorderType;
    };
}
```

#### `BorderType`

```typescript
{
    top?: BorderStyleType;
    left?: BorderStyleType;
    bottom?: BorderStyleType;
    right?: BorderStyleType;
    all?: BorderStyleType;    // Shorthand for all borders
}
```

#### `BorderStyleType`

```typescript
{
    style?: 'thin' | 'dotted' | 'dashDot' | 'hair' | 'dashDotDot' |
            'slantDashDot' | 'mediumDashed' | 'mediumDashDotDot' |
            'mediumDashDot' | 'medium' | 'double' | 'thick';
    color?: string;           // Border color (hex)
}
```

## 🎨 Styling Guide

### Colors

Use hex color codes without the `#` symbol:

- `"FF0000"` for red
- `"00FF00"` for green
- `"0000FF"` for blue
- `"FFFFFF"` for white
- `"000000"` for black

### Border Styles

Available border styles in order of thickness:

- `hair` → `thin` → `medium` → `thick`
- `dotted`, `dashDot`, `dashDotDot` for patterns
- `double` for double lines

## 🧮 Calculation Operations

| Operation | Description               | Example Use Case        |
| --------- | ------------------------- | ----------------------- |
| `sum`     | Addition of all values    | Total sales, quantities |
| `avg`     | Average of all values     | Average price, rating   |
| `count`   | Count of non-empty values | Number of items         |
| `min`     | Minimum value             | Lowest price            |
| `max`     | Maximum value             | Highest score           |

## 🔧 Advanced Features

### Data Filtering for Groups

Use custom filter functions for flexible grouping:

```javascript
groupBy: {
    conditions: [
        {
            label: "High Performers",
            filter: (employee) => employee.rating >= 4.5 && employee.salary > 60000,
        },
        {
            label: "New Hires",
            filter: (employee) => new Date(employee.hireDate) > new Date("2024-01-01"),
        },
    ];
}
```

### Multiple Calculation Types

Different operations on different columns:

```javascript
totals: {
    columns: ["quantity", "price", "rating"],
    operations: {
        quantity: "sum",    // Total units
        price: "avg",       // Average price
        rating: "max"       // Best rating
    }
}
```

## 🛠️ Dependencies

- [**exceljs**](https://github.com/exceljs/exceljs) - Excel workbook management
- [**runtime-save**](https://github.com/wyMinLwin/runtime-save) - Cross-platform file saving

## 🤝 Contributing

Contributions are welcome! Please see our [Contributing Guide](CONTRIBUTING.md) for details.

## 📋 Code of Conduct

This project follows our [Code of Conduct](CODE_OF_CONDUCT.md). Please read it before contributing.

## 📄 License

Licensed under the [MIT License](LICENSE).
