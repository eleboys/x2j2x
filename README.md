# x2j2x
Converting json file to excel and vice versa using nodejs.

Supporting json standards like nested objects and arrays. It converts json file to a two column excel file by which the first column will contain a flat dot seperated path to each value and the second column will contain the value. Later on you can convert the excel file agian to the same json structure using the reverse function.

## Install
```
npm install x2j2x
```

## Usages
I use it when I have provided a json file for my multiple language applications, containing keys and values in a base language and I want to hand in them to some translator (mashine/human) and when I have translations (new values) in a diffrent column I can retransform it to json format using the existing key and new value columns.
But you can use it for whatever you like :)

```javascript
var x2j2x = require('x2j2x'); 
// for converting an excel file to json object
var json_object = x2j2x.excel_to_json('input.xlsx');
// for converting a json_object to excel file
x2j2x.json_to_excel(json_object, 'out.xlsx');
```

## Docs
So obvious!

```javascript
/*
 * object: any - a javascript/json object to be converted to excel file
 * outPath: string - excel file path to save on
 * sheetName: string - save data on this sheetNane (by default 'Sheet1')
**/
json_to_excel(object, outPath, sheetName = 'Sheet1')
/*
 * inputPath: string - excel file path to read key and values form
 * sheetName: string - save data on this sheetNane (by default 'Sheet1')
 * keyColNum: numeber - column number on which keys are stored (by default 0)
 * valColNum: numeber - column number on which values are stored (by default 1)
**/
excel_to_json(inputPath, sheetName='Sheet1', keyColNum=0, valColNum=1)
```