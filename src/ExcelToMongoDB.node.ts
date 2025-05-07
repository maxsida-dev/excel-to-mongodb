import { IExecuteFunctions, INodeType, INodeTypeDescription, INodeExecutionData, NodeConnectionType, INodePropertyOptions, ILoadOptionsFunctions } from 'n8n-workflow';
import { NodeExecutionHandler } from './NodeExecutionHandler';

/**
 * Node n8n tùy chỉnh nâng cao để đọc file Excel theo dạng stream và lưu vào MongoDB
 */
export class ExcelToMongoDB implements INodeType {
    description: INodeTypeDescription = {
        displayName: 'Excel to MongoDB',
        name: 'excelToMongoDB',
        icon: 'file:../icons/excel-to-mongodb.svg',
        group: ['transform'],
        version: 1,
        description: 'Import Excel files to MongoDB with advanced options',
        defaults: {
            name: 'Excel to MongoDB',
        },
        inputs: [NodeConnectionType.Main],
        outputs: [NodeConnectionType.Main],
        credentials: [
            {
                name: 'mongoDb',
                required: true,
                displayOptions: {
                    show: {
                        useMongoDbUri: [false],
                    },
                },
            },
        ],
        properties: [
            // Excel File URL
            {
                displayName: 'Excel File URL',
                name: 'excelFileUrl',
                type: 'string',
                default: '',
                required: true,
                placeholder: 'https://example.com/file.xlsx',
                description: 'URL of the Excel file to download',
            },
            {
                displayName: 'Original File Name',
                name: 'originalFileName',
                type: 'string',
                default: '',
                required: false,
                placeholder: 'sales_data_2023.xlsx',
                description: 'Original file name to store in MongoDB documents (helps identify the source file)',
            },
            
            // MongoDB Connection Options
            {
                displayName: 'Connection Method',
                name: 'useMongoDbUri',
                type: 'boolean',
                default: false,
                description: 'Whether to connect using a MongoDB URI or credentials',
            },
            {
                displayName: 'MongoDB URI',
                name: 'mongoDbUri',
                type: 'string',
                default: '',
                placeholder: 'mongodb://localhost:27017',
                description: 'MongoDB connection URI',
                displayOptions: {
                    show: {
                        useMongoDbUri: [true],
                    },
                },
            },
            {
                displayName: 'Database',
                name: 'database',
                type: 'string',
                default: '',
                description: 'MongoDB database name',
            },
            {
                displayName: 'Collection',
                name: 'collection',
                type: 'string',
                default: '',
                description: 'MongoDB collection name',
            },
            {
                displayName: 'Clear Collection Before Import',
                name: 'clearCollection',
                type: 'boolean',
                default: false,
                description: 'Whether to clear the collection before importing data',
            },
            
            // Excel Options
            {
                displayName: 'Sheet Name',
                name: 'sheetName',
                type: 'string',
                default: '',
                description: 'Name of the sheet to read (leave empty for first sheet)',
            },
            {
                displayName: 'Has Headers',
                name: 'hasHeaders',
                type: 'boolean',
                default: true,
                description: 'Whether the first row contains headers',
            },
            {
                displayName: 'Skip Empty Rows',
                name: 'skipEmptyRows',
                type: 'boolean',
                default: true,
                description: 'Whether to skip empty rows',
            },
            
            // Data Conversion Options
            {
                displayName: 'Convert Data Types',
                name: 'convertDataTypes',
                type: 'boolean',
                default: true,
                description: 'Whether to automatically convert data types (e.g., string to number)',
            },
            {
                displayName: 'Date Fields',
                name: 'dateFields',
                type: 'string',
                default: '',
                description: 'Comma-separated list of fields to be treated as dates',
                placeholder: 'birthDate,createdAt,updatedAt',
            },
            {
                displayName: 'Timezone',
                name: 'timezone',
                type: 'string',
                default: 'UTC',
                description: 'Timezone for date conversion',
            },
            
            // Column Mapping
            {
                displayName: 'Column Mappings',
                name: 'columnMappings',
                placeholder: 'Add Column Mapping',
                type: 'fixedCollection',
                typeOptions: {
                    multipleValues: true,
                },
                default: {},
                options: [
                    {
                        name: 'mapping',
                        displayName: 'Mapping',
                        values: [
                            {
                                displayName: 'Excel Column',
                                name: 'excelColumn',
                                type: 'string',
                                default: '',
                                description: 'Name of the column in Excel',
                            },
                            {
                                displayName: 'MongoDB Field',
                                name: 'mongoField',
                                type: 'string',
                                default: '',
                                description: 'Name of the field in MongoDB',
                            },
                        ],
                    },
                ],
            },
            
            // Select Columns
            {
                displayName: 'Select Specific Columns',
                name: 'selectColumns',
                type: 'boolean',
                default: false,
                description: 'Whether to select specific columns to save to MongoDB',
            },
            {
                displayName: 'Columns to Save',
                name: 'selectedColumns',
                type: 'string',
                default: '',
                description: 'Comma-separated list of column names to save to MongoDB',
                placeholder: 'name,email,phone',
                displayOptions: {
                    show: {
                        selectColumns: [true],
                    },
                },
            },
            
            // Performance Options
            {
                displayName: 'Batch Size',
                name: 'batchSize',
                type: 'number',
                default: 100,
                description: 'Number of documents to insert in one batch',
            },
        ],
    };

    async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
        const handler = new NodeExecutionHandler(this);
        return handler.execute();
    }
}









