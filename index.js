var XLSX = require('xlsx');
var XLSX_EXTEND = require('./xlsx.utils');
var NWISE = require('./nwise.utils');
var json = require('./sample.json');
var result = NWISE.utils.flatten_object(json);

var workbook = XLSX.utils.create_empty_workbook();
var worksheet = XLSX.utils.object_array_to_sheet(result);

workbook.SheetNames.push('Sheet1');
workbook.Sheets['Sheet1'] = worksheet;
XLSX.writeFile(workbook, 'out.xlsx');

/*
var xlsxFile = 'wb1.xlsx';
var jsonFile = 'sample.json';
var sheetName = 'Sheet1';

var workbook = XLSX.readFile(xlsxFile);
var worksheet = workbook.Sheets[sheetName];
var objArray = XLSX.utils.sheet_to_row_object_array(worksheet, { header: 1 });
console.log(objArray);



workbook = XLSX.utils.create_empty_workbook();
worksheet = XLSX.utils.object_array_to_sheet(objArray);

workbook.SheetNames.push('Sheet1');
workbook.Sheets['Sheet1'] = worksheet;
XLSX.writeFile(workbook, 'out.xlsx');
*/