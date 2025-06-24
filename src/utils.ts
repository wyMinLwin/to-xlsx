import { Worksheet } from "exceljs";
import { ColumnSizesType, ColumnsOrderType, ColumnHeadersType, TitleType } from "./types";

export const getWorksheetColumns = <T>(
    data: T[],
    columnHeaders: ColumnHeadersType,
    columnSizes: ColumnSizesType,
    excludeColumns: string[] | null,
    columnsOrder: ColumnsOrderType // e.g. ['age']
) => {
    if (!data.length) return [];

    const allColumns = Object.keys(data[0] as object);

    // Filter out excluded columns
    const filteredColumns = allColumns.filter((column) => !excludeColumns?.includes(column));

    // colums order by user
    const orderedColumns = columnsOrder
        ? [
              ...columnsOrder.filter((col) => filteredColumns.includes(col)),
              ...filteredColumns.filter((col) => !columnsOrder.includes(col)),
          ]
        : filteredColumns;

    return orderedColumns.map((column) => ({
        header: generateHeader(columnHeaders, column),
        key: column,
        width: generateColumnSize(columnSizes, column),
    }));
};

export const generateHeader = (columnHeaders: ColumnHeadersType, header: string): string => {
    return columnHeaders?.[header] ?? header.toUpperCase().slice(0, 1) + header.slice(1);
};

export const generateColumnSize = (columnSizes: ColumnSizesType, header: string): number => {
    return columnSizes?.[header] ?? 20;
};

export const getUniqueFields = <T, K extends keyof T>(arr: T[], key: K) => {
    return [...new Set(arr.map((item) => item[key]))];
};

export const addTitle = (worksheet: Worksheet, length: number, title: TitleType) => {
    const lastRow = worksheet.lastRow;
    const lastRowNumber = lastRow?.number ?? 1;
    const lastCellCount = lastRow?.cellCount ?? 0;
    const firstCell = `${indexToLetter(lastCellCount <= 0 ? lastCellCount : lastCellCount - 1)}${lastRowNumber}`;
    const lastCell = `${indexToLetter(length - 1)}${lastRowNumber}`;
    worksheet.mergeCells(`${firstCell}:${lastCell}`);
    worksheet.getCell(firstCell).value = title?.text;
    worksheet.getCell(firstCell).font = {
        color: { argb: title?.color ?? "000000" },
    };
    worksheet.getCell(firstCell).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: title?.bg ?? "FFFFFF" },
    };
    worksheet.getCell(firstCell).alignment = { horizontal: "center", vertical: "middle" };
};

export const indexToLetter = (index: number) => {
    let letters = "";
    while (index >= 0) {
        letters = String.fromCharCode((index % 26) + 65) + letters;
        index = Math.floor(index / 26) - 1;
    }
    return letters;
};
