var XLSX = require('xlsx')
var fs = require('fs');
var path = require('path');

var mapPath = "Cyber RPG";
var propertiesToInclude = ["utip", "uhpm", "ua1b", "ua1c", "ugol", "ureq", "udef", "urqa", "ulum", "upap"];

var excelPath = "./democracy3/policies.csv";
var outputExcelPath = "./democracy3/policies.xlsx";

module.exports = { "init": init };

function init() {
    var unitBalanceWorkBook = XLSX.readFile(excelPath, { type: 'string', raw: true });

    var unitBalanceSheet = unitBalanceWorkBook.Sheets[unitBalanceWorkBook.SheetNames[0]];

    //var unitBalanceJsonSheet = XLSX.utils.sheet_to_json(unitBalanceSheet);

    var arrayOfArrays = sheet2arr(unitBalanceSheet);

    modifyArray(arrayOfArrays);

    var workSheet = XLSX.utils.aoa_to_sheet(arrayOfArrays);

    createFormulas(workSheet);

    //var workSheet = XLSX.utils.json_to_sheet(unitBalanceJsonSheet);

    XLSX.utils.book_append_sheet(unitBalanceWorkBook, workSheet);

    XLSX.writeFile(unitBalanceWorkBook, outputExcelPath);
};

function modifyArray(arrayOfArrays) {
    var effectsJson = {};
    var header = arrayOfArrays[0];

    for (var i = 1; i < arrayOfArrays.length; i++) {
        var array = arrayOfArrays[i];

        var sliceIndex = 19;
        var part1 = array.slice(0, sliceIndex);
        var effects = array.slice(sliceIndex, array.length);

        part1.push("=1");
        header.push(undefined);
        sliceIndex++;

        for (var j = 0; j < effects.length; j++) {
            var effect = effects[j];

            if (effect && effect.split && !Number.isInteger(effect)) {
                var effectParts = effect.split(",");
                var effectName = effectParts[0];
                var effectFormula = effectParts[1];

                if (effectFormula) {
                    effectFormula = "=" + effectFormula.replace('x', 'T' + (i + 1));

                    var regex = /[\(\)]/g;
                    var matches = effectFormula.match(regex);
                    if (matches && matches.length % 2 == 1) {
                        effectFormula = effectFormula.slice(0, effectFormula.length - 1);
                    }
                }

                var insertIndex = sliceIndex + Object.keys(effectsJson).length;
                if (effectsJson[effectName]) {
                    insertIndex = effectsJson[effectName];
                }
                else {
                    effectsJson[effectName] = insertIndex;
                }

                if (insertIndex >= part1.length) {
                    part1.length = insertIndex + 1;
                }

                if (insertIndex >= header.length) {
                    header.length = insertIndex + 1;
                }

                header[insertIndex] = effectName;
                part1[insertIndex] = effectFormula;
            }
        }

        arrayOfArrays[i] = part1;
    }

    createSumRow(header, arrayOfArrays);

    substituteFormulaVariables(effectsJson, arrayOfArrays);
}

function createSumRow(header, arrayOfArrays) {
    // create row with sum
    arrayOfArrays.push(new Array(header.length));
    var sumRow = arrayOfArrays[arrayOfArrays.length - 1];
    for (var i = 0; i < header.length; i++) {
        var rowChar = numToAlpha(i);
        sumRow[i] = "=SUM(" + rowChar + "2:" + rowChar + (arrayOfArrays.length - 1) + ")";
    }
}

function substituteFormulaVariables(effectsJson, arrayOfArrays) {
    for (var i = 1; i < arrayOfArrays.length; i++) {
        var array = arrayOfArrays[i];

        for (var j = 0; j < array.length; j++) {
            var element = array[j];

            if (element && element[0] == "=") {
                var regex = /[a-zA-Z]{2,}/g;
                var matches = element.match(regex);

                if (matches) {
                    var insertIndex = effectsJson[matches[0]];
                    if (insertIndex) {
                        array[j] = element.replace(regex, numToAlpha(insertIndex) + (arrayOfArrays.length));
                    }
                }
            }
        }
    }
}

function sheet2arr(sheet) {
    var result = [];
    var row;
    var rowNum;
    var colNum;
    var range = XLSX.utils.decode_range(sheet['!ref']);
    for (rowNum = range.s.r; rowNum <= range.e.r; rowNum++) {
        row = [];
        for (colNum = range.s.c; colNum <= range.e.c; colNum++) {
            var nextCell = sheet[
                XLSX.utils.encode_cell({ r: rowNum, c: colNum })
            ];
            if (typeof nextCell === 'undefined') {
                row.push(void 0);
            } else row.push(nextCell.v);
        }
        result.push(row);
    }
    return result;
};

function jsonSheetToArrayOfArrays(jsonSheet) {
    var arrayOfArrays = [];

    for (var i = 0; i < jsonSheet.length; i++) {
        var jsonSheetRow = jsonSheet[i];

        var keys = Object.keys(jsonSheetRow);
        if (!arrayOfArrays[0] || arrayOfArrays[0].length < keys.length) {
            arrayOfArrays[0] = keys;
        }

        var values = Object.values(jsonSheetRow);

        //validateColumnsExistance(arrayOfArrays[0], keys, values)

        arrayOfArrays.push(values);
    }

    return arrayOfArrays;
}

function createFormulas(workSheet) {
    var keys = Object.keys(workSheet);
    for (var i = 0; i < keys.length; i++) {
        var key = keys[i];

        var cell = workSheet[key];

        if (cell.v && cell.v[0] == "=") {
            cell.f = cell.v;
            cell.v = undefined;
        }
    }
}

function numToAlpha(num) {
    var alpha = '';

    for (; num >= 0; num = parseInt(num / 26, 10) - 1) {
        alpha = String.fromCharCode(num % 26 + 0x41) + alpha;
    }

    return alpha;
}