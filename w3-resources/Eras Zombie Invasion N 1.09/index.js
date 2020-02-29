var XLSX = require('xlsx')
var fs = require('fs');
var path = require('path');

var mapPath = "Eras Zombie Invasion 0.83.2T";

var w3ResourcesPath = "./w3-resources/";
var mapResourcesPath = path.resolve(w3ResourcesPath + mapPath);

module.exports = { "init": init };

function init() {
    var workBook = XLSX.utils.book_new();

    workBook = getUnitsWorkBook(workBook);

    XLSX.writeFile(workBook, mapResourcesPath + "/resources.xlsx");
};

function getUnitsWorkBook(workBook) {
    var w3uPath = path.resolve(mapResourcesPath + '/war3map.w3u.json');
    var w3uDataRaw = fs.readFileSync(w3uPath);
    var w3uJson = JSON.parse(w3uDataRaw);

    var propertiesToInclude = ["utip", "uhpm", "ua1b", "ua1c", "ugol", "ureq", "udef", "urqa", "ulum", "upap"];
    var workSheet = jsonToWorkSheet(w3uJson, propertiesToInclude)
    XLSX.utils.book_append_sheet(workBook, workSheet, "Units");

    return workBook;
}

function jsonToWorkSheet(w3Json, propertiesToInclude) {
    var arrayOfArrays = [[]];

    if (propertiesToInclude) {
        propertiesToInclude = propertiesToInclude.reduce(function (accumulator, value) {
            accumulator[value] = true;
            return accumulator
        }, {});
    }

    var header = { "id": 0 };
    var headerLength = 1;

    var customKeys = Object.keys(w3Json.custom);
    for (var i = 0; i < customKeys.length; i++) {
        var customKey = customKeys[i];
        var customProperties = w3Json.custom[customKey];

        var row = [customKey];

        for (var j = 0; j < customProperties.length; j++) {
            var customProperty = customProperties[j];

            if (!propertiesToInclude || propertiesToInclude.hasOwnProperty(customProperty.id)) {
                if (!Number.isInteger(header[customProperty.id])) {
                    header[customProperty.id] = headerLength;
                    headerLength++;
                }

                row[header[customProperty.id]] = customProperty.value;
            }
        }

        arrayOfArrays.push(row);
    }

    arrayOfArrays.unshift(Object.keys(header));

    var workSheet = XLSX.utils.aoa_to_sheet(arrayOfArrays);

    return workSheet;
}