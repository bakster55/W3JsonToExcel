var XLSX = require('xlsx')
var fs = require('fs');
var path = require('path');

var mapPath = "Eras Zombie Invasion 0.83.2T";
var propertiesToInclude = ["utip", "uhpm", "ua1b", "ua1c", "ugol", "ureq", "udef", "urqa", "ulum", "upap"];

var w3ResourcesPath = "./w3-resources/";
var mapResourcesPath = path.resolve(w3ResourcesPath + mapPath);

(function () {
    var workBook = XLSX.utils.book_new();

    createUnitsWorksheet(workBook);

    XLSX.writeFile(workBook, mapResourcesPath + "/resources.xlsx");
})();

function createUnitsWorksheet(workBook) {
    var w3uPath = path.resolve(mapResourcesPath + '/war3map.w3u.json');
    var arrayOfArrays = [[]];

    var header = { "id": 0 };
    var headerLength = 1;

    var w3uDataRaw = fs.readFileSync(w3uPath);
    var w3uData = JSON.parse(w3uDataRaw);

    propertiesToInclude = propertiesToInclude.reduce(function (accumulator, value) {
        accumulator[value] = true;
        return accumulator
    }, {});

    var customKeys = Object.keys(w3uData.custom);
    for (var i = 0; i < customKeys.length; i++) {
        var customKey = customKeys[i];
        var unitProperties = w3uData.custom[customKey];

        var row = [customKey];

        for (var j = 0; j < unitProperties.length; j++) {
            var unitProperty = unitProperties[j];

            if (propertiesToInclude[unitProperty.id]) {
                if (!Number.isInteger(header[unitProperty.id])) {
                    header[unitProperty.id] = headerLength;
                    headerLength++;
                }

                row[header[unitProperty.id]] = unitProperty.value;
            }
        }

        arrayOfArrays.push(row);
    }

    arrayOfArrays.unshift(Object.keys(header));

    var workSheet = XLSX.utils.aoa_to_sheet(arrayOfArrays);
    XLSX.utils.book_append_sheet(workBook, workSheet, "Units");
}