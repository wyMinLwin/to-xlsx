console.log("This is interesting");

export function exportToXlsx<T>(data: T[]): void {
    console.log("Exporting to XLSX:", data);
}
