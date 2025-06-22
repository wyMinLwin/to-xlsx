export type Props<T> = {
    data: T[];
    excludeColumns?: string[];
    fileName?: string;
    headers?: HeadersType;
    columnSizes?: ColumnSizesType;
    sheetsGroupBy?: SheetsGroupByType;
};

export type HeadersType = Record<string, string> | null;
export type ColumnSizesType = Record<string, number> | null;
export type SheetsGroupByType = {
    namePattern: string;
    key: string;
} | null;
