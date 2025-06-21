export type Props<T> = {
    data: T[];
    excludeColumns?: string[];
    fileName?: string;
    headers?: HeaderType;
};

export type HeaderType = Record<string, string> | null;
