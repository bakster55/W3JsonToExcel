var XLSX = require('xlsx')
var fs = require('fs');
var path = require('path');

var mapPath = "Cyber RPG";
var propertiesToInclude = ["utip", "uhpm", "ua1b", "ua1c", "ugol", "ureq", "udef", "urqa", "ulum", "upap"];

var w3ResourcesPath = "./w3-resources/";
var mapResourcesPath = path.resolve(w3ResourcesPath + mapPath);

module.exports = { "init": init };

function init() {
    var unitBalanceWorkBook = XLSX.readFile(mapResourcesPath + "/UnitBalance.slk");
    var unitWeaponsWorkBook = XLSX.readFile(mapResourcesPath + "/UnitWeapons.slk");

    var unitBalanceSheet = unitBalanceWorkBook.Sheets[unitBalanceWorkBook.SheetNames[0]];
    var unitWeaponsSheet = unitWeaponsWorkBook.Sheets[unitWeaponsWorkBook.SheetNames[0]];

    var unitBalanceJsonSheet = XLSX.utils.sheet_to_json(unitBalanceSheet);
    var unitWeaponsJsonSheet = XLSX.utils.sheet_to_json(unitWeaponsSheet);;

    var merjedJsonSheet = mergeSheets(unitBalanceJsonSheet, "unitBalanceID", unitWeaponsJsonSheet, "unitWeapID");

    var cardSummonsList = require(mapResourcesPath + "/card-summons-list.js")
    var filteredRowsJsonSheet = filterSheetRows(merjedJsonSheet, "unitBalanceID", cardSummonsList);
    var filteredColumnsJsonSheet = filterSheetColumns(filteredRowsJsonSheet, ["unitBalanceID", "HP", "def", "dmgplus1"]);

    var workSheet = XLSX.utils.json_to_sheet(filteredColumnsJsonSheet);

    XLSX.utils.book_append_sheet(unitBalanceWorkBook, workSheet, "CardSummons");

    XLSX.writeFile(unitBalanceWorkBook, mapResourcesPath + "/resources.xlsx");
};

function filterSheetRows(jsonSheet, columnName, rowValues) {
    var result = jsonSheet.filter(sr => {
        return rowValues.some(value => value == sr[columnName]);
    });

    return result;
}

function filterSheetColumns(jsonSheet, columnNames) {
    jsonSheet.forEach(row => {
        var rowColumnNames = Object.keys(row);

        rowColumnNames.forEach(columnName => {
            if (!columnNames.includes(columnName)) {
                delete row[columnName];
            }
        });
    });

    return jsonSheet;
}

function mergeSheets(sheetArray1, key1, sheetArray2, key2) {
    //var sheetHashSet1 = getSheetHashSet(sheetArray1, key1);
    var sheetHashSet2 = getSheetHashSet(sheetArray2, key2);

    sheetArray1.forEach(row => {
        var row2 = sheetHashSet2[row[key1]];

        Object.assign(row, row2);
    })

    return sheetArray1;
}

function getSheetHashSet(sheetArray, key) {
    var sheetHashSet = {};

    sheetArray.forEach(row => {
        sheetHashSet[row[key]] = row;
    })

    return sheetHashSet;
}