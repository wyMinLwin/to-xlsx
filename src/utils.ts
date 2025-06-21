import { HeaderType } from "./types";

export const generateHeader = (headers: HeaderType, header: string): string => {
    return headers?.[header] ?? header.toUpperCase().slice(0, 1) + header.slice(1);
};
