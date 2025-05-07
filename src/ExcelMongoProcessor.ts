import { createReadStream, promises } from 'fs';
import * as ExcelJS from 'exceljs';
import { Collection } from 'mongodb';
import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';
import { IExcelReadResult, IMongoSaveResult } from './interfaces';
import { Readable } from 'stream';
import * as crypto from 'crypto';

/**
 * Lớp xử lý Excel và MongoDB
 */
export class ExcelMongoProcessor {
    /**
     * Kiểm tra xem đường dẫn có phải là URL không
     */
    private isUrl(path: string): boolean {
        return path.startsWith('http://') || path.startsWith('https://');
    }

    /**
     * Tải file từ URL và lưu vào thư mục tạm thời
     */

    private async downloadFile(url: string): Promise<string> {
        return new Promise(async (resolve, reject) => {
            const tempDir = os.tmpdir();
            const fileName = `${crypto.randomUUID()}.xlsx`
            const tempFilePath = path.join(tempDir, fileName);
            const fileStream = fs.createWriteStream(tempFilePath);

            let totalBytes = 0;

            try {
                // Thực hiện request với fetch
                const response = await fetch(url, {
                    headers: {
                        'User-Agent': 'Mozilla/5.0', // Giả lập trình duyệt
                    },
                    redirect: 'follow', // Tự động theo dõi redirect
                });

                // Kiểm tra mã trạng thái
                if (!response.ok) {
                    throw new Error(`Failed to download file: HTTP status code ${response.status} - ${response.statusText}`);
                }

                // Kiểm tra Content-Type
                const contentType = response.headers.get('content-type') || '';
                if (!contentType.includes('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') &&
                    !contentType.includes('application/octet-stream')) {
                    throw new Error(`Invalid file type. Expected .xlsx, but received Content-Type: ${contentType}`);
                }

                // Kiểm tra response.body
                if (!response.body) {
                    throw new Error('No response body received. Unable to download file.');
                }

                // Sử dụng Readable.fromWeb để chuyển đổi ReadableStream từ fetch thành Node.js stream
                const nodeReadable = Readable.fromWeb(response.body as any);

                // Pipe stream vào file
                nodeReadable.pipe(fileStream);

                // Theo dõi tiến trình tải xuống
                nodeReadable.on('data', (chunk) => {
                    totalBytes += chunk.length;
                });

                // Xử lý khi hoàn thành
                await new Promise<void>((resolveStream, rejectStream) => {
                    fileStream.on('finish', () => {
                        promises.stat(tempFilePath)
                            .then((stats) => {
                                const fileSize = stats.size;
                                const contentLength = response.headers.get('content-length');
                                if (contentLength && fileSize !== parseInt(contentLength)) {
                                    fs.unlinkSync(tempFilePath);
                                    rejectStream(new Error('Downloaded file is incomplete'));
                                } else {
                                    resolveStream();
                                }
                            })
                            .catch((err) => rejectStream(err));
                    });

                    fileStream.on('error', (err) => {
                        fs.unlink(tempFilePath, () => { });
                        rejectStream(err);
                    });

                    nodeReadable.on('error', (err) => {
                        fs.unlink(tempFilePath, () => { });
                        rejectStream(err);
                    });
                });

                resolve(tempFilePath);
            } catch (error) {
                fs.unlink(tempFilePath, () => { });
                reject(error instanceof Error ? error : new Error(String(error)));
            }
        });
    }

    /**
     * Trích xuất tên file từ đường dẫn hoặc URL
     */
    private extractFileName(filePath: string): string {
        if (this.isUrl(filePath)) {
            return new URL(filePath).pathname.split('/').pop() || 'downloaded_file.xlsx';
        } else {
            return filePath.split(/[\\/]/).pop() || '';
        }
    }

    /**
     * Chuyển đổi giá trị ngày tháng từ Excel
     */
    private convertExcelDate(value: any, timezone: string = 'UTC'): Date {
        if (typeof value === 'number') {
            // Excel lưu ngày dưới dạng số ngày kể từ 1/1/1900
            // Phần nguyên là số ngày, phần thập phân là thời gian trong ngày
            const excelEpoch = new Date(Date.UTC(1899, 11, 30));
            const days = Math.floor(value);
            const fraction = value - days;

            // Chuyển đổi phần thập phân thành milliseconds
            // 1 ngày = 24 giờ = 24 * 60 * 60 * 1000 milliseconds
            const milliseconds = Math.round(fraction * 24 * 60 * 60 * 1000);

            // Tạo đối tượng Date từ số ngày và milliseconds
            const utcDate = new Date(excelEpoch.getTime() + days * 24 * 60 * 60 * 1000 + milliseconds);

            // Điều chỉnh múi giờ nếu cần
            if (timezone !== 'UTC') {
                // Sử dụng thư viện như moment-timezone hoặc luxon để điều chỉnh múi giờ
                // Ở đây tôi sẽ sử dụng cách đơn giản hơn
                const offset = new Date().getTimezoneOffset() * 60000;
                return new Date(utcDate.getTime() - offset);
            } else {
                return utcDate;
            }
        } else if (value instanceof Date) {
            return value;
        } else if (typeof value === 'string') {
            const parsedDate = new Date(value);
            if (!isNaN(parsedDate.getTime())) {
                return parsedDate;
            }
        }
        return value; // Trả về giá trị gốc nếu không thể chuyển đổi
    }

    /**
     * Chuyển đổi kiểu dữ liệu của giá trị
     */
    private convertDataType(value: any): any {
        if (value === null || value === undefined) {
            return value;
        }

        // Chuyển đổi số
        if (typeof value === 'string' && !isNaN(Number(value))) {
            return Number(value);
        }

        // Chuyển đổi boolean
        if (typeof value === 'string' &&
            (value.toLowerCase() === 'true' || value.toLowerCase() === 'false')) {
            return value.toLowerCase() === 'true';
        }

        return value;
    }

    /**
     * Áp dụng ánh xạ tên cột sang tên trường MongoDB
     */
    private applyColumnMapping(header: string, columnMappings: { [key: string]: string }): string {
        return columnMappings[header] || header;
    }

    /**
     * Xử lý giá trị ô dữ liệu từ Excel
     */
    private processCellValue(
        value: any,
        header: string,
        dateFields: string[],
        convertDataTypes: boolean,
        timezone: string
    ): any {
        // Xử lý trường hợp đặc biệt cho các trường ngày tháng
        if (dateFields.includes(header) && value !== null && value !== undefined) {
            return this.convertExcelDate(value, timezone);
        }

        // Chuyển đổi kiểu dữ liệu khác nếu được yêu cầu
        if (convertDataTypes && value !== null && value !== undefined) {
            return this.convertDataType(value);
        }

        return value;
    }

    /**
     * Kiểm tra xem cột có được chọn để lưu không
     */
    private isColumnSelected(
        header: string,
        selectColumns: boolean,
        selectedColumns: string[]
    ): boolean {
        // Nếu không chọn cột cụ thể, luôn trả về true
        if (!selectColumns || selectedColumns.length === 0) {
            return true;
        }

        // Kiểm tra xem cột có trong danh sách được chọn không
        return selectedColumns.includes(header);
    }

    /**
     * Tạo document từ dữ liệu hàng Excel
     */
    private createDocumentFromRow(
        rowData: any[],
        headers: string[],
        rowNumber: number,
        fileName: string,
        originalFileName: string,
        dateFields: string[],
        convertDataTypes: boolean,
        timezone: string,
        columnMappings: { [key: string]: string },
        selectColumns: boolean = false,
        selectedColumns: string[] = []
    ): any {
        // Tạo document với các trường mặc định
        let document: any = {
            _source_file: fileName,
            _row_number: rowNumber
        };
        
        // Thêm trường tên file gốc nếu có
        if (originalFileName) {
            document.old_file_name = originalFileName;
        }

        // Nếu có header, tạo object với key là header và value là giá trị tương ứng
        if (headers.length > 0) {
            headers.forEach((header, index) => {
                if (header) {
                    // Kiểm tra xem cột có được chọn để lưu không
                    if (this.isColumnSelected(header, selectColumns, selectedColumns)) {
                        // Xử lý giá trị ô
                        let value = this.processCellValue(
                            rowData[index],
                            header,
                            dateFields,
                            convertDataTypes,
                            timezone
                        );

                        // Áp dụng ánh xạ tên cột
                        const customKey = this.applyColumnMapping(header, columnMappings);
                        document[customKey] = value;
                    }
                }
            });
        } else {
            // Nếu không có header, sử dụng mảng giá trị
            document._data = rowData;
        }

        return document;
    }

    /**
     * Đọc file Excel từ đường dẫn cục bộ hoặc URL
     */
    public async readExcelFile(
        excelFilePath: string,
        originalFileName: string,
        sheetName: string,
        hasHeaders: boolean,
        skipEmptyRows: boolean,
        convertDataTypes: boolean,
        dateFields: string[] = [],
        timezone: string = 'UTC',
        columnMappings: Record<string, string> = {},
        selectColumns: boolean = false,
        selectedColumns: string[] = []
    ): Promise<IExcelReadResult> {
        let filePath = excelFilePath;
        let tempFile = false;

        // Nếu là URL, tải file về
        if (this.isUrl(excelFilePath)) {
            try {
                filePath = await this.downloadFile(excelFilePath);
                tempFile = true;
            } catch (error) {
                throw new Error(`Failed to download Excel file from URL: ${(error as Error).message}`);
            }
        }

        // Lấy tên file từ đường dẫn
        const fileName = this.extractFileName(filePath);

        try {
            // Tạo stream đọc file Excel
            const stream = createReadStream(filePath);
            const workbook = new ExcelJS.Workbook();

            try {
                stream.on('error', (error) => {
                    throw new Error(`Stream error: ${error.message}`);
                });

                await workbook.xlsx.read(stream);
            } catch (error) {
                throw new Error(`Invalid Excel file format`);
            }

            // Lấy worksheet từ workbook
            const worksheet = sheetName ?
                workbook.getWorksheet(sheetName) :
                workbook.worksheets[0];

            if (!worksheet) {
                throw new Error(`Worksheet ${sheetName || 'default'} not found`);
            }

            let headers: string[] = [];
            let documents: any[] = [];
            let rowCount = 0;

            // Xử lý từng hàng trong worksheet
            worksheet.eachRow((row, rowNumber) => {
                // Nếu là hàng đầu tiên và có header, lưu lại header
                if (rowNumber === 1 && hasHeaders) {
                    // Bỏ qua phần tử đầu tiên vì ExcelJS đánh số từ 1
                    headers = Array.isArray(row.values) ? row.values.slice(1).map((h: any) =>
                        h !== null && h !== undefined ? String(h).trim() : `column_${rowCount}`
                    ) : [];
                    return;
                }

                // Lấy dữ liệu từ hàng, bỏ qua phần tử đầu tiên
                const rowData = Array.isArray(row.values) ? row.values.slice(1) : [];

                // Kiểm tra hàng trống nếu cần
                if (skipEmptyRows && rowData.every(cell => cell === null || cell === undefined || cell === '')) {
                    return;
                }

                rowCount++;

                // Tạo document từ dữ liệu hàng, đã tích hợp lọc cột
                const document = this.createDocumentFromRow(
                    rowData,
                    headers,
                    rowNumber,
                    fileName,
                    originalFileName,
                    dateFields,
                    convertDataTypes,
                    timezone,
                    columnMappings,
                    selectColumns,
                    selectedColumns
                );

                // Thêm document vào danh sách
                documents.push(document);
            });

            return { documents, rowCount, fileName };
        } finally {
            // Xóa file tạm thời nếu đã tải từ URL
            if (tempFile) {
                try {
                    fs.unlinkSync(filePath);
                } catch (error) {
                    // Bỏ qua lỗi khi xóa file tạm thời
                }
            }
        }
    }

    /**
     * Lưu documents vào MongoDB theo batch
     */
    public async saveToMongoDB(
        collection: Collection,
        documents: Record<string, unknown>[],
        batchSize: number
    ): Promise<IMongoSaveResult> {
        let documentsInserted = 0;
        let errors: string[] = [];

        // Xử lý theo batch
        for (let i = 0; i < documents.length; i += batchSize) {
            const batch = documents.slice(i, i + batchSize);
            try {
                const result = await collection.insertMany(batch);
                documentsInserted += result.insertedCount;
            } catch (error) {
                errors.push(`Error inserting batch: ${error instanceof Error ? error.message : String(error)}`);
            }
        }

        return { documentsInserted, errors };
    }
}
