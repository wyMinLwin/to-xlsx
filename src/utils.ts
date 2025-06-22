import { ColumnSizesType, HeadersType } from "./types";

export const getWorksheetColumns = <T>(
    data: T[],
    headers: HeadersType,
    columnSizes: ColumnSizesType,
    excludeColumns: string[] | null
) => {
    const columns = Object.keys(data[0] as object).filter(
        (column) => !excludeColumns?.includes(column)
    );
    return columns.map((column) => ({
        header: generateHeader(headers, column),
        key: column,
        width: generateColumnSize(columnSizes, column),
    }));
};

export const generateHeader = (headers: HeadersType, header: string): string => {
    return headers?.[header] ?? header.toUpperCase().slice(0, 1) + header.slice(1);
};

export const generateColumnSize = (columnSizes: ColumnSizesType, header: string): number => {
    return columnSizes?.[header] ?? 20;
};

export const getUniqueFields = <T, K extends keyof T>(arr: T[], key: K) => {
    return [...new Set(arr.map((item) => item[key]))];
};
