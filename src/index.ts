import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import { Props } from "./types";
import { addTitle, getUniqueFields, getWorksheetColumns } from "./utils";

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
    } = props;

    const workbook = new ExcelJS.Workbook();
    if (sheetsBy) {
        const uniqueFields = getUniqueFields<T, keyof T>(data, sheetsBy.key as keyof T);
        uniqueFields.forEach((uniqueField) => {
            const namePattern = sheetsBy.namePattern;
            const worksheet = workbook.addWorksheet(
                namePattern.includes("$key")
                    ? namePattern.replaceAll("$key", String(uniqueField))
                    : String(uniqueField)
            );

            const columns = getWorksheetColumns(
                data,
                columnHeaders,
                columnSizes,
                excludeColumns,
                columnsOrder
            );
            addTitle(worksheet, columns.length, title);
            // Add header row manually after title
            const headerRow = worksheet.addRow(columns.map((col) => col.header));
            // Apply styles only to header row
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
            // Optionally set column widths
            if (columnSizes) {
                worksheet.columns = columns.map((col) => ({
                    key: col.key,
                    width: col.width,
                    style: col.style,
                }));
            }

            data.filter((d) => d[sheetsBy.key as keyof T] == uniqueField).forEach((row) => {
                worksheet.addRow(row);
            });
        });
    } else {
        const worksheet = workbook.addWorksheet(fileName);

        const columns = getWorksheetColumns(
            data,
            columnHeaders,
            columnSizes,
            excludeColumns,
            columnsOrder
        );
        addTitle(worksheet, columns.length, title);
        // Add header row manually after title
        const headerRow = worksheet.addRow(columns.map((col) => col.header));
        // Apply styles only to header row
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
        // Optionally set column widths
        if (columnSizes) {
            worksheet.columns = columns.map((col) => ({
                key: col.key,
                width: col.width,
                style: col.style,
            }));
        }
        data.forEach((row) => {
            worksheet.addRow(row);
        });
    }

    workbook.xlsx.writeBuffer().then((buffer) => {
        const blob = new Blob([buffer], {
            type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        });
        saveAs(blob, fileName);
    });
}
