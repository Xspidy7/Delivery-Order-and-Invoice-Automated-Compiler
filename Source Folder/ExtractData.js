const XLSX = require('xlsx');
const fs = require('fs');

function extractData(filePath, parsingPoints) {
    const workbook = XLSX.readFile(filePath);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];

    // Extracting header information
    const headerDocumentNumber = getCellValue(sheet, parsingPoints.HEADER_DOCUMENT_NUMBER);
    const headerDocumentDate = getCellValue(sheet, parsingPoints.HEADER_DOCUMENT_DATE);
    const headerSerialOrderNumber = getCellValue(sheet, parsingPoints.HEADER_SERIAL_ORDER_NUMBER);
    const headerSupplierName = extractFirstRowText(sheet);

    // Extracting the table of data
    const startRow = getRowIndex(sheet, parsingPoints.COLUMN_PART_NUMBER) + 2;
    let data = [];
    let rowIndex = startRow;
    while (true) {
        let partNumber = getCellValueByRow(sheet, parsingPoints.COLUMN_PART_NUMBER, rowIndex);
        if (!partNumber) break;
        let customerPo = getCellValueByRow(sheet, parsingPoints.COLUMN_CUSTOMER_PURCHASE_ORDER, rowIndex);
        let quantity = getCellValueByRow(sheet, parsingPoints.COLUMN_QUANTITY, rowIndex);

        data.push({
            Supplier: headerSupplierName,
            DocumentNumber: headerDocumentNumber,
            DocumentDate: headerDocumentDate,
            SerialOrderNumber: headerSerialOrderNumber,
            CustomerPO: customerPo,
            PartNumber: partNumber,
            Quantity: quantity
        });
        rowIndex++;
    }
    return data;
}

function extractFirstRowText(sheet) {
    const firstRow = [];
    let colIndex = 1;
    while (sheet[`A${colIndex}`]) {
        let cellValue = sheet[`A${colIndex}`] ? sheet[`A${colIndex}`].v : '';
        firstRow.push(cellValue);
        colIndex++;
    }
    return firstRow.join(', ');
}

function getCellValue(sheet, searchValue) {
    const range = XLSX.utils.decode_range(sheet['!ref']);
    for (let row = range.s.r; row <= range.e.r; row++) {
        for (let col = range.s.c; col <= range.e.c; col++) {
            let cellAddress = {c: col, r: row};
            let cellRef = XLSX.utils.encode_cell(cellAddress);
            let cell = sheet[cellRef];
            if (cell && cell.v === searchValue) {
                let valueCell = {c: col + 1, r: row};
                return sheet[XLSX.utils.encode_cell(valueCell)].v;
            }
        }
    }
}

function getRowIndex(sheet, searchValue) {
    const range = XLSX.utils.decode_range(sheet['!ref']);
    for (let row = range.s.r; row <= range.e.r; row++) {
        for (let col = range.s.c; col <= range.e.c; col++) {
            let cellAddress = {c: col, r: row};
            let cellRef = XLSX.utils.encode_cell(cellAddress);
            let cell = sheet[cellRef];
            if (cell && cell.v === searchValue) {
                return row;
            }
        }
    }
}

function getCellValueByRow(sheet, columnName, rowIndex) {
    const rowStart = getRowIndex(sheet, columnName) + 2;
    let cellAddress = {c: getColumnIndex(sheet, columnName), r: rowIndex};
    let cellRef = XLSX.utils.encode_cell(cellAddress);
    let cell = sheet[cellRef];
    return cell ? cell.v : '';
}

function getColumnIndex(sheet, columnName) {
    const range = XLSX.utils.decode_range(sheet['!ref']);
    for (let row = range.s.r; row <= range.e.r; row++) {
        for (let col = range.s.c; col <= range.e.c; col++) {
            let cellAddress = {c: col, r: row};
            let cellRef = XLSX.utils.encode_cell(cellAddress);
            let cell = sheet[cellRef];
            if (cell && cell.v === columnName) {
                return col;
            }
        }
    }
}

function processFiles(destinationFile, files, parsingPoints) {
    let allData = [];

    files.forEach(file => {
        console.log(`Processing file '${file}'...`);
        const data = extractData(file, parsingPoints);
        allData.push(...data);
    });

    // Convert data to Excel
    const newSheet = XLSX.utils.json_to_sheet(allData);
    const newWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newWorkbook, newSheet, 'Sheet1');
    XLSX.writeFile(newWorkbook, destinationFile);

    console.log(`All files processed. Data written to ${destinationFile}`);
}

function main() {
    const files = ['file1.xlsx', 'file2.xlsx'];  // List of files to process
    const destinationFile = 'output.xlsx';  // Output Excel file

    // Parsing points (based on your VBA constants)
    const parsingPoints = {
        HEADER_DOCUMENT_NUMBER: 'DO No:',
        HEADER_DOCUMENT_DATE: 'DO Date:',
        HEADER_SERIAL_ORDER_NUMBER: 'S/O No:',
        COLUMN_PART_NUMBER: 'PART NO / DESCRIPTION',
        COLUMN_CUSTOMER_PURCHASE_ORDER: 'CUST-PO',
        COLUMN_QUANTITY: 'QTY'
    };

    processFiles(destinationFile, files, parsingPoints);
}

main();
