import { Color } from "exceljs";

export type Props<T> = {
    data: T[];
    fileName?: string;
    // columns customization
    columnHeaders?: ColumnHeadersType;
    columnSizes?: ColumnSizesType;
    columnsStyle?: ColumnsStyleType;
    columnsOrder?: ColumnsOrderType;
    excludeColumns?: string[];
    // split by sheets
    sheetsBy?: SheetsByType;
    title?: TitleType;
    subtitle?: SubTitleType;
};

export type ColumnHeadersType = Record<string, string> | null;

export type ColumnSizesType = Record<string, number> | null;

export type ColumnsStyleType = {
    bg?: string;
    color?: string;
    fontSize?: number;
} | null;

export type SheetsByType = {
    namePattern: string;
    key: string;
} | null;

export type ColumnsOrderType = string[] | null;

export type TitleType = {
    text: string;
    bg?: string;
    color?: string;
    fontSize?: number;
} | null;

export type SubTitleType = {
    text: string;
    bg?: Partial<Color>;
    color?: Partial<Color>;
} | null;
