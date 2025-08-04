import { Color } from "exceljs";

export type Props<T> = {
    data: T[];
    fileName?: string;
    // columns customization
    columnHeaders?: ColumnHeadersType;
    columnSizes?: ColumnSizesType;
    columnsStyle?: ColumnsStyleType;
    columnsOrder?: ColumnsOrderType;
    columnsMerge?: ColumnsMergeType;
    excludeColumns?: string[];
    // split by sheets
    sheetsBy?: SheetsByType;
    title?: TitleType;
    subtitle?: SubTitleType;
    // group by functionality
    groupBy?: GroupByType<T>;
    // totals functionality
    totals?: TotalsType;
};

export type ColumnHeadersType = Record<string, string> | null;

export type ColumnSizesType = Record<string, number> | null;

export type ColumnsStyleType = {
    bg?: string;
    color?: string;
    fontSize?: number;
} | null;

export type ColumnsOrderType = string[] | null;

export type ColumnsMergeType =
    | {
          keys: {
              startColumn: string;
              endColumn: string;
          };
          columnName: string;
      }[]
    | null;

export type SheetsByType = {
    namePattern: string;
    key: string;
} | null;

export type TitleType = {
    text: string;
    bg?: string;
    color?: string;
    fontSize?: number;
    border?: BorderType;
} | null;

export type SubTitleType = {
    text: string;
    bg?: Partial<Color>;
    color?: Partial<Color>;
    border?: BorderType;
} | null;

export type GroupByType<T> = {
    conditions: GroupCondition<T>[];
    subtitleStyle?: {
        bg?: string;
        color?: string;
        fontSize?: number;
        border?: BorderType;
    };
    showSubtotals?: boolean;
    subtotalStyle?: {
        bg?: string;
        color?: string;
        fontSize?: number;
        border?: BorderType;
    };
} | null;

export type GroupCondition<T> = {
    label: string;
    filter: (item: T) => boolean;
};

export type BorderType = {
    top?: BorderStyleType;
    left?: BorderStyleType;
    bottom?: BorderStyleType;
    right?: BorderStyleType;
    all?: BorderStyleType; // shorthand to set all borders at once
};

export type BorderStyleType = {
    style?:
        | "thin"
        | "dotted"
        | "dashDot"
        | "hair"
        | "dashDotDot"
        | "slantDashDot"
        | "mediumDashed"
        | "mediumDashDotDot"
        | "mediumDashDot"
        | "medium"
        | "double"
        | "thick";
    color?: string;
};

export type TotalsType = {
    columns: string[]; // Array of column names to calculate totals for
    showGrandTotal?: boolean;
    grandTotalLabel?: string;
    subtotalLabel?: string;
    grandTotalStyle?: {
        bg?: string;
        color?: string;
        fontSize?: number;
        border?: BorderType;
    };
    operations?: {
        [columnName: string]: "sum" | "avg" | "count" | "min" | "max";
    };
} | null;
