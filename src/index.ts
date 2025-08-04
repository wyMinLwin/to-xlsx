import ExcelJS, { Worksheet } from "exceljs";

import { saveFile } from "runtime-save";
import { Props } from "./types";
import {
    addTitle,
    getUniqueFields,
    getWorksheetColumns,
    addSubtitle,
    groupDataByConditions,
    addSubtotalRow,
    addGrandTotalRow,
} from "./utils";
import { getProcessPath } from "./process";

export function exportToXlsx<T>(props: Props<T>): void {
    const {
        data,
        excludeColumns = null,
        fileName = "ExportSheet",
        columnHeaders = null,
        columnSizes = null,
        sheetsBy = null,
        columnsOrder = null,
        columnsMerge = null,
        title = null,
        columnsStyle = null,
        groupBy = null,
        totals = null,
    } = props;

    const workbook = new ExcelJS.Workbook();

    // Helper function to add data with optional grouping
    const addDataToWorksheet = (worksheet: Worksheet, dataToAdd: T[]) => {
        const columns = getWorksheetColumns(
            dataToAdd,
            columnHeaders,
            columnSizes,
            excludeColumns,
            columnsOrder
        );

        addTitle(worksheet, columns.length, title);

        if (groupBy) {
            // Group data by conditions and add each group with subtitle
            const groups = groupDataByConditions(dataToAdd, groupBy);

            groups.forEach((group) => {
                // Add subtitle for each group
                addSubtitle(worksheet, columns.length, group.label, groupBy.subtitleStyle);

                // Add header row for each group
                const headerRow = worksheet.addRow(columns.map((col) => col.header));
                //const columnIndex = columns.findIndex(c => c.header === (columnHeaders?.['name'] ?? 'name'))+1;
                //const columnIndex2 = columns.findIndex(c => c.header === (columnHeaders?.['salary'] ?? 'salary'))+1;
                //worksheet.mergeCells(`${headerRow.getCell(columnIndex).address}:${headerRow.getCell(columnIndex2).address}`);
                //worksheet.getCell(headerRow.getCell(columnIndex).address).text =
                if (columnsMerge) {
                    columnsMerge.forEach((cm) => {
                        const startCol =
                            columns.findIndex(
                                (c) =>
                                    c.header ===
                                    (columnHeaders?.[cm.keys.startColumn] ?? cm.keys.startColumn)
                            ) + 1;
                        const endCol =
                            columns.findIndex(
                                (c) =>
                                    c.header ===
                                    (columnHeaders?.[cm.keys.endColumn] ?? cm.keys.endColumn)
                            ) + 1;
                        worksheet.mergeCells(
                            `${headerRow.getCell(startCol).address}:${headerRow.getCell(endCol).address}`
                        );
                        headerRow.getCell(startCol).value = cm.columnName;
                    });
                }
                // Apply styles to header row
                if (columnsStyle) {
                    headerRow.eachCell((cell) => {
                        cell.fill = {
                            type: "pattern",
                            pattern: "solid",
                            fgColor: { argb: columnsStyle.bg },
                        };
                        cell.font = {
                            color: { argb: columnsStyle.color },
                            size: columnsStyle.fontSize,
                        };
                        cell.alignment = {
                            vertical: "middle",
                            horizontal: "center",
                        };
                    });
                }
                if (columnSizes) {
                    worksheet.columns = columns.map((col) => ({
                        key: col.key,
                        width: col.width,
                        style: col.style,
                    }));
                }

                // Add data rows for this group
                group.data.forEach((row) => {
                    const rowValues = columns.map((col) => row[col.key as keyof typeof row]);
                    worksheet.addRow(rowValues);
                });

                // Add subtotal row if enabled
                if (groupBy.showSubtotals && totals && totals.columns.length > 0) {
                    addSubtotalRow(worksheet, group.data, totals, columns, groupBy);
                }
            });

            // Add grand total row for grouped data if enabled
            if (totals && totals.showGrandTotal && totals.columns.length > 0) {
                addGrandTotalRow(worksheet, dataToAdd, totals, columns);
            }
        } else {
            // Regular mode without grouping
            const headerRow = worksheet.addRow(columns.map((col) => col.header));
            if (columnsMerge) {
                columnsMerge.forEach((cm) => {
                    const startCol =
                        columns.findIndex(
                            (c) =>
                                c.header ===
                                (columnHeaders?.[cm.keys.startColumn] ?? cm.keys.startColumn)
                        ) + 1;
                    const endCol =
                        columns.findIndex(
                            (c) =>
                                c.header ===
                                (columnHeaders?.[cm.keys.endColumn] ?? cm.keys.endColumn)
                        ) + 1;
                    worksheet.mergeCells(
                        `${headerRow.getCell(startCol).address}:${headerRow.getCell(endCol).address}`
                    );
                    headerRow.getCell(startCol).value = cm.columnName;
                });
            }

            // Apply styles to header row
            if (columnsStyle) {
                headerRow.eachCell((cell) => {
                    cell.fill = {
                        type: "pattern",
                        pattern: "solid",
                        fgColor: { argb: columnsStyle.bg },
                    };
                    cell.font = {
                        color: { argb: columnsStyle.color },
                        size: columnsStyle.fontSize,
                    };
                    cell.alignment = {
                        vertical: "middle",
                        horizontal: "center",
                    };
                });
            }
            if (columnSizes) {
                worksheet.columns = columns.map((col) => ({
                    key: col.key,
                    width: col.width,
                    style: col.style,
                }));
            }

            // Add all data rows
            dataToAdd.forEach((row) => {
                worksheet.addRow(row);
            });
        }

        // Add grand total row if enabled
        if (totals && totals.showGrandTotal && totals.columns.length > 0) {
            addGrandTotalRow(worksheet, dataToAdd, totals, columns);
        }
    };
    if (sheetsBy) {
        const uniqueFields = getUniqueFields<T, keyof T>(data, sheetsBy.key as keyof T);
        uniqueFields.forEach((uniqueField) => {
            const namePattern = sheetsBy.namePattern;
            const worksheet = workbook.addWorksheet(
                namePattern.includes("$key")
                    ? namePattern.replaceAll("$key", String(uniqueField))
                    : String(uniqueField)
            );

            const filteredData = data.filter((d) => d[sheetsBy.key as keyof T] == uniqueField);
            addDataToWorksheet(worksheet, filteredData);
        });
    } else {
        const worksheet = workbook.addWorksheet(fileName);
        addDataToWorksheet(worksheet, data);
    }

    workbook.xlsx.writeBuffer().then((buffer) => {
        const blob = new Blob([buffer], {
            type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        });
        saveFile(fileName + ".xlsx", blob, getProcessPath());
    });
}
