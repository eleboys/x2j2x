
var x2j2x = {};
(function make_nwise(x2j2x) {

    var XLSX = require('xlsx');
    var XLSX_EXTEND = require('./xlsx.utils');

    function join_object_path(p1, p2) {
        if (!p1) {
            return p2;
        }
        if (!p2) {
            return p1;
        }
        return `${p1}.${p2}`;
    }

    function flatten_object(jsonData, basePath, output) {
        var output = output || [];
        basePath = basePath || '';
        if (!Object.isExtensible(jsonData)) {
            output.push([basePath, jsonData]);
            return output;
        }
        for (var key in jsonData) {
            var value = jsonData[key];
            if (Object.isExtensible(value)) {
                if (!Array.isArray(value)) {
                    flatten_object(value, join_object_path(basePath, key), output);
                } else {
                    value.forEach((item, i) => {
                        flatten_object(item, join_object_path(basePath, key) + `[${i}]`, output);
                    });
                }
            } else {
                output.push([join_object_path(basePath, key), value]);
            }
        }

        return output;
    }

    function object_array_to_sheet(data, opts) {
        var ws = {};
        var range = { s: { c: 10000000, r: 10000000 }, e: { c: 0, r: 0 } };
        for (var R = 0; R != data.length; ++R) {
            for (var C = 0; C != data[R].length; ++C) {
                if (range.s.r > R) range.s.r = R;
                if (range.s.c > C) range.s.c = C;
                if (range.e.r < R) range.e.r = R;
                if (range.e.c < C) range.e.c = C;
                var cell = { v: data[R][C] };
                if (cell.v == null) continue;
                var cell_ref = XLSX.utils.encode_cell({ c: C, r: R });

                if (typeof cell.v === 'number') cell.t = 'n';
                else if (typeof cell.v === 'boolean') cell.t = 'b';
                else if (cell.v instanceof Date) {
                    cell.t = 'n'; cell.z = XLSX.SSF._table[14];
                    cell.v = datenum(cell.v);
                }
                else cell.t = 's';

                ws[cell_ref] = cell;
            }
        }
        if (range.s.c < 10000000) ws['!ref'] = XLSX.utils.encode_range(range);
        ws['!merges'] = [];
        return ws;
    }

    function create_empty_workbook() {
        return {
            SheetNames: [],
            Sheets: {}
        }
    }

    /*
    * inputPath: string - excel file path to read key and values form
    * sheetName: string - save data on this sheetNane (by default 'Sheet1')
    * keyColNum: numeber - column number on which keys are stored (by default 0)
    * valColNum: numeber - column number on which values are stored (by default 1)
    **/
    function json_to_excel(object, outPath, sheetName) {
        sheetName = sheetName || 'Sheet1'
        var flattenObj = flatten_object(object);
        var workbook = create_empty_workbook();
        var worksheet = object_array_to_sheet(flattenObj);

        workbook.SheetNames.push(sheetName);
        workbook.Sheets[sheetName] = worksheet;
        XLSX.writeFile(workbook, outPath);
    }

    /*
    * object: any - a javascript/json object to be converted to excel file
    * outPath: string - excel file path to save on
    * sheetName: string - save data on this sheetNane (by default 'Sheet1')
    **/
    function extend_object(keyParts, value, object) {
        if (!keyParts.length) return;
        var firstPart = keyParts[0];
        var index;
        if (index = /\[\d\]$/.exec(firstPart)){
            firstPart = firstPart.replace(index,"");
            index = parseInt(/\d/.exec(index[0])[0]);
            object[firstPart] = object[firstPart]||[];
            object[firstPart][index] = keyParts.length==1 ? value : (object[firstPart][index]||{});
            extend_object(keyParts.splice(1), value, object[firstPart][index]);
        } else {
            object[firstPart] = keyParts.length==1 ? value : (object[firstPart]||{});
            extend_object(keyParts.splice(1), value, object[firstPart]);
        }
    }

    function excel_to_json(inputPath, sheetName, keyColNum, valColNum) {
        sheetName = sheetName || 'Sheet1';
        keyColNum = keyColNum || 0;
        valColNum = valColNum || 1;
        var workbook = XLSX.readFile(inputPath);
        var worksheet = workbook.Sheets[sheetName];
        var objArray = XLSX.utils.sheet_to_row_object_array(worksheet, { header: 1 })
                                 .filter(o => o && o.length);
        var finalObject = {};
        for (var i = 0; i < objArray.length; i++) {
            var key = objArray[i][keyColNum];
            var value = objArray[i][valColNum];
            var keyParts = key.split('.');
            extend_object(keyParts, value, finalObject)
        }
        return finalObject;
    }

    x2j2x.json_to_excel = json_to_excel;
    x2j2x.excel_to_json = excel_to_json;

})(typeof exports !== 'undefined' ? exports : x2j2x);