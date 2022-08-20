
const path = require('path');
const ExcelJS = require('../../lib/exceljs.nodejs.js');

const TEST_XLSX_FILE_NAME = path.resolve(__dirname, '../data/splice.xlsx');
const filenameOut = path.resolve(__dirname, '../data/splice-out.xlsx');
const options = {
    filename: TEST_XLSX_FILE_NAME,
    useStyles: true,
};

const test = async () => {

    const wb = new ExcelJS.Workbook(options);
    await wb.xlsx.readFile(TEST_XLSX_FILE_NAME);
    const ws = wb.getWorksheet('Sheet1');
    ws.spliceColumns(4, 2);
    
    await wb.xlsx.writeFile(filenameOut);
};
test();