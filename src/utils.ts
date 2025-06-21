import { ColumnSizesType, HeadersType } from "./types";

export const generateHeader = (headers: HeadersType, header: string): string => {
    return headers?.[header] ?? header.toUpperCase().slice(0, 1) + header.slice(1);
};

export const generateColumnSize = (columnSizes: ColumnSizesType, header: string): number => {
    return columnSizes?.[header] ?? 20;
};
