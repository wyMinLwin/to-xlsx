// Polyfill for Node.js process in browser environments
export const getProcessPath = (): string => {
    if (typeof process !== "undefined" && process.cwd) {
        // Node.js environment
        return process.cwd() + "/test/output/";
    }
    // Browser environment - return empty string or a browser-specific path
    return "";
};
