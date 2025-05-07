/**
 * Interface cho kết quả đọc file Excel
 */
export interface IExcelReadResult {
    documents: Record<string, unknown>[];
    rowCount: number;
    fileName: string;
}

/**
 * Interface cho kết quả lưu vào MongoDB
 */
export interface IMongoSaveResult {
    documentsInserted: number;
    errors: string[];
}

/**
 * Interface cho ánh xạ cột
 */
export interface IColumnMapping {
    excelColumn: string;
    mongoField: string;
}

/**
 * Interface cho tham số cấu hình node
 */
export interface INodeParameters {
    database: string;
    collection: string;
    sheetName: string;
    hasHeaders: boolean;
    batchSize: number;
    clearCollection: boolean;
    skipEmptyRows: boolean;
    convertDataTypes: boolean;
    dateFields: string[];
    timezone: string;
    columnMappings: Record<string, string>;
    selectColumns: boolean;
    selectedColumns: string[];
}
