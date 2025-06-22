import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import { Props, SheetsGroupByType } from "./types";
import { generateColumnSize, generateHeader, getUniqueFields, getWorksheetColumns } from "./utils";

export function exportToXlsx<T>(props: Props<T>): void {
    const {
        data,
        excludeColumns = null,
        fileName = "export",
        headers = null,
        columnSizes = null,
        sheetsGroupBy = null,
    } = props;

    const workbook = new ExcelJS.Workbook();
    if (sheetsGroupBy) {
        const uniqueFields = getUniqueFields<T, keyof T>(data, sheetsGroupBy.key as keyof T);
        uniqueFields.forEach((uniqueField) => {
            const namePattern = sheetsGroupBy.namePattern;
            const worksheet = workbook.addWorksheet(
                namePattern.includes("$key")
                    ? namePattern.replaceAll("$key", String(uniqueField))
                    : String(uniqueField)
            );

            worksheet.columns = getWorksheetColumns(data, headers, columnSizes, excludeColumns);

            data.filter((d) => d[sheetsGroupBy.key as keyof T] == uniqueField).forEach((row) => {
                worksheet.addRow(row);
            });
        });
    } else {
        const worksheet = workbook.addWorksheet("Sheet 1");
        worksheet.columns = getWorksheetColumns(data, headers, columnSizes, excludeColumns);

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
