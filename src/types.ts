export type Props<T> = {
    data: T[];
    excludeColumns?: string[];
    fileName?: string;
    headers?: HeadersType;
    columnSizes?: ColumnSizesType;
};

export type HeadersType = Record<string, string> | null;
export type ColumnSizesType = Record<string, number> | null;
