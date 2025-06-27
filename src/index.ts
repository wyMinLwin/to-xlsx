import ExcelJS, { Worksheet } from "exceljs";

import { saveFile } from "runtime-save";
import { Props } from "./types";
import {
    addTitle,
    getUniqueFields,
    getWorksheetColumns,
    addSubtitle,
    groupDataByConditions,
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
        title = null,
        columnsStyle = null,
        groupBy = null,
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
            });
        } else {
            // Regular mode without grouping
            const headerRow = worksheet.addRow(columns.map((col) => col.header));

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
