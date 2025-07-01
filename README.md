# to-xlsx

A lightweight JavaScript/TypeScript library to export data to Excel XLSX files with advanced
formatting options.

[![npm version](https://img.shields.io/npm/v/to-xlsx.svg)](https://www.npmjs.com/package/to-xlsx)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

## Features

- Export JavaScript/TypeScript arrays to Excel XLSX files
- Customizable column headers and sizes
- Support for styling (colors, fonts, etc.)
- Group data with subtitles
- Split data across multiple sheets
- Exclude specific columns
- Order columns as needed

## Installation

```bash
# npm
npm install to-xlsx

# pnpm
pnpm add to-xlsx
```

## Usage

```javascript
import { exportToXlsx } from "to-xlsx";

// Your data array
const employees = [
    { name: "John", age: 18, department: "IT", salary: 45000 },
    { name: "Jane", age: 25, department: "HR", salary: 55000 },
    { name: "Bob", age: 17, department: "IT", salary: 35000 },
    // ...more data
];

// Basic usage
exportToXlsx({
    data: employees,
    fileName: "employees-report",
});

// Advanced usage with styling and grouping
exportToXlsx({
    data: employees,
    fileName: "employees-grouped-by-age",
    title: {
        text: "Employee Report - Grouped by Age",
        bg: "4472C4",
        color: "FFFFFF",
        fontSize: 18,
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
    groupBy: {
        field: "age",
        ranges: [
            { min: 0, max: 18, label: "Under 18" },
            { min: 18, max: 25, label: "18-25" },
            { min: 25, max: 35, label: "25-35" },
            { min: 35, max: Infinity, label: "35+" },
        ],
        subtitleStyle: {
            bg: "BDD7EE",
            color: "000000",
            fontSize: 14,
        },
    },
});
```

## API Reference

### exportToXlsx(props)

Main function to export data to Excel.

#### Props

| Property       | Type                           | Description                      | Default       |
| -------------- | ------------------------------ | -------------------------------- | ------------- |
| data           | Array<T>                       | The data array to export         | Required      |
| fileName       | string                         | Name of the output file          | "ExportSheet" |
| columnHeaders  | Record<string, string> \| null | Custom headers for columns       | null          |
| columnSizes    | Record<string, number> \| null | Custom widths for columns        | null          |
| columnsStyle   | ColumnsStyleType               | Style for column headers         | null          |
| columnsOrder   | string[]                       | Order of columns in the output   | null          |
| excludeColumns | string[]                       | Columns to exclude from export   | null          |
| sheetsBy       | SheetsByType                   | Split data into multiple sheets  | null          |
| title          | TitleType                      | Title with optional borders      | null          |
| subtitle       | SubTitleType                   | Subtitle with optional borders   | null          |
| groupBy        | GroupByType<T>                 | Group data with optional borders | null          |

## Border Customization

You can now add custom borders to Title, Subtitle, and GroupBy sections:

```javascript
exportToXlsx({
    // ...other props
    title: {
        text: "Employee Report",
        bg: "4472C4",
        color: "FFFFFF",
        fontSize: 18,
        border: {
            // Apply the same border to all sides
            all: {
                style: "thick",
                color: "000000",
            },
            // Or specify individual sides
            // top: { style: "thin", color: "FF0000" },
            // left: { style: "dotted", color: "00FF00" },
            // bottom: { style: "medium", color: "0000FF" },
            // right: { style: "dashed", color: "FFFF00" }
        },
    },
    // GroupBy with custom borders for subtitles
    groupBy: {
        // ...other groupBy properties
        subtitleStyle: {
            bg: "BDD7EE",
            color: "000000",
            fontSize: 14,
            border: {
                bottom: { style: "medium", color: "0070C0" },
            },
        },
    },
});
```

### Available Border Styles

- `thin` - A thin border (default if style is not specified)
- `dotted` - A dotted border
- `dashDot` - A dash-dot border
- `hair` - A hair (very thin) border
- `dashDotDot` - A dash-dot-dot border
- `slantDashDot` - A slant dash-dot border
- `mediumDashed` - A medium dashed border
- `mediumDashDotDot` - A medium dash-dot-dot border
- `mediumDashDot` - A medium dash-dot border
- `medium` - A medium border
- `double` - A double border
- `thick` - A thick border

## Dependencies

- [exceljs](https://github.com/exceljs/exceljs) - Excel workbook manager
- [runtime-save](https://github.com/wyMinLwin/runtime-save) - File saving utility

## Contributing

Please see [CONTRIBUTING Guide](CONTRIBUTING.md) for details on how to contribute to this project.

## Code of Conduct

This project adheres to a [Code of Conduct](CODE_OF_CONDUCT.md). By participating, you are expected
to uphold this code.

## License

This project is licensed under the [MIT License](LICENSE).
