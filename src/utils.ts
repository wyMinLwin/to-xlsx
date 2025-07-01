import { Column, Worksheet } from "exceljs";
import {
    ColumnSizesType,
    ColumnsOrderType,
    ColumnHeadersType,
    TitleType,
    GroupByType,
    BorderType,
} from "./types";

export const getWorksheetColumns = <T>(
    data: T[],
    columnHeaders: ColumnHeadersType,
    columnSizes: ColumnSizesType,
    excludeColumns: string[] | null,
    columnsOrder: ColumnsOrderType // e.g. ['age']
): Partial<Column>[] => {
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
    if (!title) return;
    const lastRow = worksheet.lastRow;
    const lastRowNumber = lastRow?.number ?? 1;
    const lastCellCount = lastRow?.cellCount ?? 0;
    const firstCell = `${indexToLetter(lastCellCount <= 0 ? lastCellCount : lastCellCount - 1)}${lastRowNumber}`;
    const lastCell = `${indexToLetter(length - 1)}${lastRowNumber}`;
    worksheet.mergeCells(`${firstCell}:${lastCell}`);
    worksheet.getCell(firstCell).value = title?.text;
    worksheet.getCell(firstCell).font = {
        color: { argb: title?.color ?? "000000" },
        size: title?.fontSize ?? 16,
    };
    worksheet.getCell(firstCell).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: title?.bg ?? "FFFFFF" },
    };
    worksheet.getCell(firstCell).alignment = { horizontal: "center", vertical: "middle" };

    // Apply border if specified
    if (title.border) {
        applyBorders(worksheet, firstCell, title.border);
    }
};

export const addSubtitle = (
    worksheet: Worksheet,
    length: number,
    subtitle: string,
    subtitleStyle?: {
        bg?: string;
        color?: string;
        fontSize?: number;
        border?: BorderType;
    }
) => {
    const lastRow = worksheet.lastRow;
    const lastRowNumber = (lastRow?.number ?? 0) + 1;
    const firstCell = `A${lastRowNumber}`;
    const lastCell = `${indexToLetter(length - 1)}${lastRowNumber}`;

    worksheet.mergeCells(`${firstCell}:${lastCell}`);
    worksheet.getCell(firstCell).value = subtitle;
    worksheet.getCell(firstCell).font = {
        color: { argb: subtitleStyle?.color ?? "333333" },
        size: subtitleStyle?.fontSize ?? 12,
        bold: true,
    };
    worksheet.getCell(firstCell).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: subtitleStyle?.bg ?? "E6E6E6" },
    };
    worksheet.getCell(firstCell).alignment = { horizontal: "left", vertical: "middle" };

    // Apply border if specified
    if (subtitleStyle?.border) {
        applyBorders(worksheet, firstCell, subtitleStyle.border);
    }
};

export const groupDataByConditions = <T>(
    data: T[],
    groupBy: GroupByType<T>
): Array<{ label: string; data: T[] }> => {
    if (!groupBy) return [];

    return groupBy.conditions
        .map((condition) => ({
            label: condition.label,
            data: data.filter(condition.filter),
        }))
        .filter((group) => group.data.length > 0); // Only return groups with data
};

export const indexToLetter = (index: number) => {
    let letters = "";
    while (index >= 0) {
        letters = String.fromCharCode((index % 26) + 65) + letters;
        index = Math.floor(index / 26) - 1;
    }
    return letters;
};

/**
 * Applies border styles to a cell or range of cells
 */
export const applyBorders = (worksheet: Worksheet, cellRef: string, borderProps?: BorderType) => {
    if (!borderProps) return;

    const cell = worksheet.getCell(cellRef);
    const border: Partial<{
        top: { style: string; color: { argb: string } };
        left: { style: string; color: { argb: string } };
        bottom: { style: string; color: { argb: string } };
        right: { style: string; color: { argb: string } };
    }> = {};

    // Apply specific borders if defined
    if (borderProps.top) {
        border.top = {
            style: borderProps.top.style || "thin",
            color: { argb: borderProps.top.color || "000000" },
        };
    }

    if (borderProps.left) {
        border.left = {
            style: borderProps.left.style || "thin",
            color: { argb: borderProps.left.color || "000000" },
        };
    }

    if (borderProps.bottom) {
        border.bottom = {
            style: borderProps.bottom.style || "thin",
            color: { argb: borderProps.bottom.color || "000000" },
        };
    }

    if (borderProps.right) {
        border.right = {
            style: borderProps.right.style || "thin",
            color: { argb: borderProps.right.color || "000000" },
        };
    }

    // Apply the same border to all sides if 'all' is defined
    if (borderProps.all) {
        const allBorder = {
            style: borderProps.all.style || "thin",
            color: { argb: borderProps.all.color || "000000" },
        };

        border.top = allBorder;
        border.left = allBorder;
        border.bottom = allBorder;
        border.right = allBorder;
    }

    // Type assertion needed because ExcelJS has a more specific type for border styles
    cell.border = border as unknown as typeof cell.border;
};
