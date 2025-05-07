import { IExecuteFunctions, INodeExecutionData, NodeOperationError } from 'n8n-workflow';
import { existsSync } from 'fs';
import { MongoClient, MongoClientOptions, Collection, Db } from 'mongodb';
import { ExcelMongoProcessor } from './ExcelMongoProcessor';
import { INodeParameters } from './interfaces';
import * as os from 'os';
import * as path from 'path';
import * as fs from 'fs';

/**
 * Lớp xử lý quá trình thực thi node
 */
export class NodeExecutionHandler {
    private executeFunctions: IExecuteFunctions;
    private processor: ExcelMongoProcessor;

    constructor(executeFunctions: IExecuteFunctions) {
        this.executeFunctions = executeFunctions;
        this.processor = new ExcelMongoProcessor();
    }

    /**
     * Kiểm tra và lấy đường dẫn file Excel
     */
    private async getExcelFilePath(): Promise<string> {
        // Lấy URL của file Excel
        const excelFileUrl = this.executeFunctions.getNodeParameter('excelFileUrl', 0) as string;
        
        if (!excelFileUrl) {
            throw new NodeOperationError(
                this.executeFunctions.getNode(),
                'Excel file URL is required!'
            );
        }
        
        // Kiểm tra URL hợp lệ
        try {
            new URL(excelFileUrl);
        } catch (error) {
            throw new NodeOperationError(
                this.executeFunctions.getNode(),
                'Invalid URL format! Please provide a valid URL.'
            );
        }
        
        return excelFileUrl;
    }

    /**
     * Lấy các tham số cấu hình từ node
     */
    private getNodeParameters(): INodeParameters {
        const database = this.executeFunctions.getNodeParameter('database', 0) as string;
        const collection = this.executeFunctions.getNodeParameter('collection', 0) as string;
        const sheetName = this.executeFunctions.getNodeParameter('sheetName', 0) as string;
        const hasHeaders = this.executeFunctions.getNodeParameter('hasHeaders', 0) as boolean;
        const batchSize = this.executeFunctions.getNodeParameter('batchSize', 0) as number;
        const clearCollection = this.executeFunctions.getNodeParameter('clearCollection', 0) as boolean;
        const skipEmptyRows = this.executeFunctions.getNodeParameter('skipEmptyRows', 0) as boolean;
        const convertDataTypes = this.executeFunctions.getNodeParameter('convertDataTypes', 0) as boolean;
        
        // Các tham số khác
        const dateFields = (this.executeFunctions.getNodeParameter('dateFields', 0) as string || '')
            .split(',')
            .map(field => field.trim())
            .filter(field => field !== '');
        
        const timezone = this.executeFunctions.getNodeParameter('timezone', 0) as string;
        
        // Xử lý column mappings
        const columnMappingsData = this.executeFunctions.getNodeParameter('columnMappings', 0) as {
            mapping?: Array<{ excelColumn: string; mongoField: string }>;
        };
        
        const columnMappings: Record<string, string> = {};
        if (columnMappingsData.mapping && columnMappingsData.mapping.length > 0) {
            for (const mapping of columnMappingsData.mapping) {
                if (mapping.excelColumn && mapping.mongoField) {
                    columnMappings[mapping.excelColumn] = mapping.mongoField;
                }
            }
        }
        
        // Xử lý selected columns
        const selectColumns = this.executeFunctions.getNodeParameter('selectColumns', 0) as boolean;
        const selectedColumns: string[] = [];
        
        if (selectColumns) {
            const selectedColumnsStr = this.executeFunctions.getNodeParameter('selectedColumns', 0) as string;
            if (selectedColumnsStr) {
                selectedColumnsStr.split(',').forEach(column => {
                    const trimmedColumn = column.trim();
                    if (trimmedColumn) {
                        selectedColumns.push(trimmedColumn);
                    }
                });
            }
        }
        
        return {
            database,
            collection,
            sheetName,
            hasHeaders,
            batchSize,
            clearCollection,
            skipEmptyRows,
            convertDataTypes,
            dateFields,
            timezone,
            columnMappings,
            selectColumns,
            selectedColumns,
        };
    }

    /**
     * Tạo kết nối đến MongoDB
     */
    private async connectToMongoDB(): Promise<MongoClient> {
        // Tùy chọn kết nối MongoDB
        const mongoOptions: MongoClientOptions = {
            connectTimeoutMS: 10000, // Timeout kết nối 10 giây
            socketTimeoutMS: 45000,  // Timeout socket 45 giây
        };
        
        // Kiểm tra phương thức kết nối
        const useMongoDbUri = this.executeFunctions.getNodeParameter('useMongoDbUri', 0) as boolean;
        
        let mongoDbUri: string;
        
        if (useMongoDbUri) {
            // Sử dụng URI được cung cấp trực tiếp
            mongoDbUri = this.executeFunctions.getNodeParameter('mongoDbUri', 0) as string;
            
            if (!mongoDbUri) {
                throw new NodeOperationError(
                    this.executeFunctions.getNode(),
                    'MongoDB URI is required when using direct URI connection!'
                );
            }
        } else {
            // Sử dụng thông tin xác thực
            const credentials = await this.executeFunctions.getCredentials('mongoDb');
            
            if (!credentials) {
                throw new NodeOperationError(
                    this.executeFunctions.getNode(),
                    'MongoDB credentials are required!'
                );
            }
            
            // Kiểm tra xem credentials có chứa URI trực tiếp không
            if (credentials.connectionString) {
                mongoDbUri = credentials.connectionString as string;
            } else {
                // Xây dựng URI từ thông tin xác thực
                const server = credentials.server as string || 'localhost';
                const port = credentials.port as number || 27017;
                const database = credentials.database as string || '';
                const user = credentials.user as string || '';
                const password = credentials.password as string || '';
                const authSource = credentials.authSource as string || 'admin';
                const ssl = credentials.ssl as boolean || false;
                const replicaSet = credentials.replicaSet as string || '';
                
                // Xác định loại kết nối (mongodb:// hoặc mongodb+srv://)
                const isSrv = server.includes('.mongodb.net') || (credentials.isSrv as boolean || false);
                const protocol = isSrv ? 'mongodb+srv://' : 'mongodb://';
                
                // Xây dựng URI
                let uri = protocol;
                
                // Thêm thông tin xác thực nếu có
                if (user && password) {
                    uri += `${encodeURIComponent(user)}:${encodeURIComponent(password)}@`;
                }
                
                // Thêm server và port (nếu không phải SRV)
                if (isSrv) {
                    uri += server;
                } else {
                    uri += `${server}:${port}`;
                }
                
                // Thêm database nếu có
                if (database) {
                    uri += `/${database}`;
                }
                
                // Thêm các tham số
                const params = [];
                
                if (authSource) {
                    params.push(`authSource=${authSource}`);
                }
                
                if (ssl) {
                    params.push('ssl=true');
                }
                
                if (replicaSet) {
                    params.push(`replicaSet=${replicaSet}`);
                }
                
                // Thêm các tham số vào URI
                if (params.length > 0) {
                    uri += `?${params.join('&')}`;
                }
                
                mongoDbUri = uri;
            }
        }
        
        try {
            console.log('Connecting to MongoDB with URI:', mongoDbUri);
            const client = new MongoClient(mongoDbUri, mongoOptions);
            await client.connect();
            console.log('Successfully connected to MongoDB');
            return client;
        } catch (error) {
            console.error('MongoDB connection error:', error);
            throw new NodeOperationError(
                this.executeFunctions.getNode(),
                `Failed to connect to MongoDB: ${error instanceof Error ? error.message : String(error)}`
            );
        }
    }

    /**
     * Xóa dữ liệu trong collection nếu được yêu cầu
     */
    private async clearCollectionIfRequested(
        collection: Collection,
        shouldClear: boolean
    ): Promise<void> {
        if (shouldClear) {
            await collection.deleteMany({});
        }
    }

    /**
     * Tạo kết quả trả về cho node
     */
    private createNodeOutput(
        success: boolean,
        rowCount: number,
        documentsInserted: number,
        errors: string[],
        fileName: string,
        columnMappingsCount: number,
        selectedColumnsCount: number = 0
    ): INodeExecutionData[] {
        return [{
            json: {
                success,
                rowsProcessed: rowCount,
                documentsInserted,
                errors,
                fileName,
                columnMappingsApplied: columnMappingsCount,
                selectedColumnsCount,
                message: errors.length === 0
                    ? `Successfully processed ${rowCount} rows and inserted ${documentsInserted} documents into MongoDB from file ${fileName}${selectedColumnsCount > 0 ? ` (with ${selectedColumnsCount} selected columns)` : ''}`
                    : `Processed with errors: ${rowCount} rows processed, ${documentsInserted} documents inserted, ${errors.length} errors occurred from file ${fileName}${selectedColumnsCount > 0 ? ` (with ${selectedColumnsCount} selected columns)` : ''}`,
            },
        }];
    }

    /**
     * Thực thi quá trình xử lý
     */
    public async execute(): Promise<INodeExecutionData[][]> {
        let client: MongoClient | undefined;
        
        try {
            // Lấy đường dẫn file Excel
            const excelFile = await this.getExcelFilePath();
            
            // Lấy các tham số cấu hình
            const params = this.getNodeParameters();
            
            // Kết nối đến MongoDB
            client = await this.connectToMongoDB();
            const db: Db = client.db(params.database);
            const collection: Collection = db.collection(params.collection);
            
            // Xóa dữ liệu trong collection nếu được yêu cầu
            await this.clearCollectionIfRequested(collection, params.clearCollection);
            
            // Đọc file Excel
            const { documents, rowCount, fileName } = await this.processor.readExcelFile(
                excelFile,
                params.sheetName,
                params.hasHeaders,
                params.skipEmptyRows,
                params.convertDataTypes,
                params.dateFields,
                params.timezone,
                params.columnMappings,
                params.selectColumns,
                params.selectedColumns
            );
            
            // Lưu vào MongoDB
            const { documentsInserted, errors } = await this.processor.saveToMongoDB(
                collection,
                documents,
                params.batchSize
            );
            
            // Trả về kết quả
            return [this.createNodeOutput(
                errors.length === 0,
                rowCount,
                documentsInserted,
                errors,
                fileName,
                Object.keys(params.columnMappings).length,
                params.selectColumns ? params.selectedColumns.length : 0
            )];
            
        } catch (error) {
            // Xử lý lỗi
            if (error instanceof Error) {
                throw new NodeOperationError(
                    this.executeFunctions.getNode(),
                    error,
                    { message: error.message }
                );
            } else {
                throw new NodeOperationError(
                    this.executeFunctions.getNode(),
                    String(error)
                );
            }
        } finally {
            // Đóng kết nối MongoDB
            if (client) {
                try {
                    await client.close();
                } catch (error) {
                    // Bỏ qua lỗi khi đóng kết nối
                }
            }
        }
    }
}















