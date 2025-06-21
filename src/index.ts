import ExcelJS from "exceljs";
import { saveAs } from "file-saver";
import { Props } from "./types";
import { generateHeader } from "./utils";

export function exportToXlsx<T>(props: Props<T>): void {
    const { data, excludeColumns, fileName = "export", headers = null } = props;

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Sheet 1");
    const columns = Object.keys(data[0] as object).filter(
        (column) => !excludeColumns?.includes(column)
    );
    worksheet.columns = columns.map((column) => ({
        header: generateHeader(headers, column),
        key: column,
        width: 50,
    }));

    data.forEach((row) => {
        worksheet.addRow(row);
    });

    workbook.xlsx.writeBuffer().then((buffer) => {
        const blob = new Blob([buffer], {
            type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        });
        saveAs(blob, fileName);
    });
}
