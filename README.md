# Excel to MongoDB Node for n8n Workflow

This custom n8n node allows you to import Excel files into MongoDB with advanced options for data transformation, column mapping, and batch processing. It's optimized for handling large Excel files efficiently.

## Features

- **Multiple Sources**: Import Excel files from local filesystem or URLs
- **High Performance**: Optimized for processing large Excel files with minimal memory usage
- **Data Transformation**: Automatically convert data types and handle date fields
- **Column Mapping**: Map Excel columns to MongoDB fields with custom names
- **Column Selection**: Choose specific columns to import
- **Batch Processing**: Control batch size for optimal performance
- **Collection Management**: Option to clear collection before import
- **Timeout Control**: Set execution timeout for handling very large files
- **Error Handling**: Robust error handling with detailed error messages

## Performance Optimizations

- **Efficient Memory Usage**: Processes Excel files using optimized methods to minimize memory consumption
- **Batch Inserts**: Inserts documents in batches to reduce MongoDB overhead
- **Timeout Management**: Configurable timeout settings to handle large files without blocking workflows
- **Stream Processing**: Uses streaming techniques for handling large files when possible

## Installation

### Local Installation

1. Clone this repository to your n8n custom nodes directory:
```bash
cd ~/.n8n/custom
git clone https://github.com/maxsida-dev/excel-to-mongodb.git
```

2. Install dependencies and build:
```bash
cd excel-to-mongodb
npm install
npm run build
```

3. Restart n8n

## Usage

1. Add the "Excel to MongoDB" node to your workflow
2. Configure the node with your Excel file source and MongoDB connection details
3. Set up data transformation options as needed
4. Configure performance options for large files
5. Run the workflow

## Configuration Options

### Excel File Source
- **Excel File URL**: URL of the Excel file to download
- **Original File Name**: Original file name to store in MongoDB documents (helps identify the source file)

### MongoDB Connection
- **MongoDB URI**: Connection string for MongoDB
- **Database**: Name of the database
- **Collection**: Name of the collection
- **Clear Collection Before Import**: Whether to clear the collection before importing data

### Excel Options
- **Sheet Name**: Name of the sheet to read (leave empty for first sheet)
- **Has Headers**: Whether the first row contains headers
- **Skip Empty Rows**: Whether to skip empty rows

### Data Conversion Options
- **Convert Data Types**: Automatically convert data types (e.g., string to number)
- **Date Fields**: Comma-separated list of fields to be treated as dates
- **Timezone**: Timezone for date conversion

### Column Mapping
- **Column Mappings**: Map Excel columns to MongoDB fields with custom names

### Column Selection
- **Select Specific Columns**: Choose specific columns to save to MongoDB
- **Columns to Save**: Comma-separated list of column names to save

### Performance Options
- **Batch Size**: Number of documents to insert in one batch (higher values may improve performance for small documents)

## Handling Large Files

When working with large Excel files (100MB+), consider these best practices:

1. **Increase Batch Size**: For small documents, larger batch sizes (500-1000) can improve performance
2. **Set Appropriate Timeout**: Set a timeout value that allows enough time for processing
3. **Select Specific Columns**: Only import the columns you need to reduce memory usage
4. **Use Data Type Conversion**: Enable data type conversion to optimize MongoDB storage

## Output

The node outputs a JSON object with the following properties:

- **success**: Whether the operation was successful
- **rowsProcessed**: Number of rows processed from the Excel file
- **documentsInserted**: Number of documents inserted into MongoDB
- **errors**: Array of error messages, if any
- **fileName**: Name of the source Excel file
- **columnMappingsApplied**: Number of column mappings applied
- **selectedColumnsCount**: Number of columns selected for import
- **message**: Summary message of the operation

## Example Workflow

1. **Trigger Node** (e.g., Manual Trigger)
2. **Excel to MongoDB Node**
   - Excel File: `/path/to/data.xlsx`
   - MongoDB URI: `mongodb://localhost:27017`
   - Database: `mydb`
   - Collection: `customers`
   - Has Headers: `true`
   - Convert Data Types: `true`
   - Date Fields: `birthDate,registrationDate`
   - Column Mappings:
     - Excel Column: `Name`, MongoDB Field: `fullName`
     - Excel Column: `Email`, MongoDB Field: `emailAddress`
   - Batch Size: `500`

## License

MIT

