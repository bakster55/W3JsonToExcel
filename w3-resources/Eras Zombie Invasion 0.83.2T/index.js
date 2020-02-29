var XLSX = require('xlsx')
var fs = require('fs');
var path = require('path');

var mapPath = "Eras Zombie Invasion N 1.09";

var w3ResourcesPath = "./w3-resources/";
var mapResourcesPath = path.resolve(w3ResourcesPath + mapPath);

module.exports = { "init": init };

function init() {
    var workBook = XLSX.utils.book_new();

    // Strings
    var w3JsonPath = path.resolve(mapResourcesPath + '/war3map.wts.json');
    var w3JsonString = fs.readFileSync(w3JsonPath);
    var w3StringsJson = JSON.parse(w3JsonString);
    w3StringsJson = ProcessStringsJson(w3StringsJson);

    // Units
    var w3JsonPath = path.resolve(mapResourcesPath + '/war3map.w3u.json');
    var propertiesToInclude = ["utip", "uhpm", "ua1b", "ua1c", "ugol", "ureq", "udef", "urqa", "ulum", "upap"];
    var workSheet = getWorkSheet(w3JsonPath, propertiesToInclude, w3StringsJson);
    XLSX.utils.book_append_sheet(workBook, workSheet, "Units");

    // Upgrades
    var w3JsonPath = path.resolve(mapResourcesPath + '/war3map.w3q.json');
    var propertiesToInclude = null;
    var workSheet = getWorkSheet(w3JsonPath, propertiesToInclude, w3StringsJson);
    XLSX.utils.book_append_sheet(workBook, workSheet, "Upgrades");

    XLSX.writeFile(workBook, mapResourcesPath + "/war3map.xlsx");
};

function getWorkSheet(w3JsonPath, propertiesToInclude, w3StringsJson) {
    var w3JsonString = fs.readFileSync(w3JsonPath);
    var w3Json = JSON.parse(w3JsonString);
    var arraysOfArraysWorkSheet = jsonToArraysOfArraysWorkSheet(w3Json, propertiesToInclude, w3StringsJson);
    var workSheet = XLSX.utils.aoa_to_sheet(arraysOfArraysWorkSheet);

    return workSheet;
}

function jsonToArraysOfArraysWorkSheet(w3Json, propertiesToInclude, w3StringsJson) {
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

                if (w3StringsJson.hasOwnProperty(customProperty.value)) {
                    row[header[customProperty.id]] = w3StringsJson[customProperty.value];
                }
                else {
                    row[header[customProperty.id]] = customProperty.value;
                }
            }
        }

        arrayOfArrays.push(row);
    }

    arrayOfArrays.unshift(Object.keys(header));

    return arrayOfArrays;
}

function ProcessStringsJson(w3StringsJson) {
    var keys = Object.keys(w3StringsJson);
    for (let index = 0; index < keys.length; index++) {
        const key = keys[index];
        const value = w3StringsJson[key];

        delete w3StringsJson[key];
        w3StringsJson["TRIGSTR_" + key] = value;
    }

    return w3StringsJson;
}