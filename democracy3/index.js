var XLSX = require('xlsx')
var fs = require('fs');
var path = require('path');

var policiesExcelPath = "./democracy3/policies.csv";
var situationsExcelPath = "./democracy3/situations.csv";
var initialPoliciesExcelPath = "./democracy3/initialPolicies.xlsx";

var outputExcelPath = "./democracy3/policies.xlsx";

var sliceIndexConst = 19;

var minCostColumnName = "B";
var maxCostColumnName = "C";
var costMultiplierColumnName = "D";

var minIncomeColumnName = "E";
var maxIncomeColumnName = "F";
var incomeMultiplierColumnName = "G";

var policiesSliderColumnName = "J";
var isPolicyActiveColumnName = "K";

module.exports = { "init": init };

function init() {
    var policiesWorkBook = XLSX.readFile(policiesExcelPath, { type: 'string', raw: true });
    var situationsWorkBook = XLSX.readFile(situationsExcelPath, { type: 'string', raw: true });
    var initialPoliciesWorkBook = XLSX.readFile(initialPoliciesExcelPath, { type: 'string', raw: true });

    var policiesSheet = policiesWorkBook.Sheets[policiesWorkBook.SheetNames[0]];
    var situationsSheet = situationsWorkBook.Sheets[situationsWorkBook.SheetNames[0]];
    var initialPoliciesSheet = initialPoliciesWorkBook.Sheets[initialPoliciesWorkBook.SheetNames[0]];

    var policiesArrayOfArrays = sheet2arr(policiesSheet);
    var situationsArrayOfArrays = sheet2arr(situationsSheet);
    var initialPoliciesArrayOfArrays = sheet2arr(initialPoliciesSheet);

    removeColumns(policiesArrayOfArrays, ["A", "C", "D", "E", "F", "G", "H", "I", "J", "S", "K", "O"]);

    mergeArrays(policiesArrayOfArrays, situationsArrayOfArrays);

    modifyArray(policiesArrayOfArrays);

    initPolicies(initialPoliciesArrayOfArrays, policiesArrayOfArrays);

    createSumRow(policiesArrayOfArrays[0], policiesArrayOfArrays);

    var workSheet = XLSX.utils.aoa_to_sheet(policiesArrayOfArrays);

    createFormulas(workSheet);


    XLSX.utils.book_append_sheet(policiesWorkBook, workSheet);

    XLSX.writeFile(policiesWorkBook, outputExcelPath);
};

function initPolicies(initialPoliciesArrayOfArrays, policiesArrayOfArrays) {
    for (var i = 1; i < initialPoliciesArrayOfArrays.length; i++) {
        var initialPolicy = initialPoliciesArrayOfArrays[i];

        var effectName = formatEffectName(initialPolicy[0]);

        var policyRow = policiesArrayOfArrays.find(r => formatEffectName(r[0]) == effectName);
        if (policyRow) {
            policyRow[alphaToNum(isPolicyActiveColumnName)] = 1;
            policyRow[alphaToNum(policiesSliderColumnName)] = initialPolicy[1];
        }
    }
}

function mergeArrays(policiesArrayOfArrays, situationsArrayOfArrays) {
    for (var i = 1; i < situationsArrayOfArrays.length; i++) {
        var situation = situationsArrayOfArrays[i];

        var lastIndexOf = situation.lastIndexOf("#");
        var causes = situation.slice(13, lastIndexOf);
        var effects = situation.slice(lastIndexOf + 1, situation.length);

        for (var j = 0; j < causes.length; j++) {
            var effect = causes[j];

            if (effect) {
                var effectParts = effect.split(",");
                var effectName = formatEffectName(effectParts[0]);


                var newRegEx = new RegExp(effectName, "i");
                effect = effect.replace(newRegEx, situation[1]);

                var policyRow = policiesArrayOfArrays.find(r => formatEffectName(r[0]) == effectName);
                if (policyRow) {
                    policyRow.push(effect);
                }
                else {
                    effect = effect.replace(/(?<=\W)x(?=\W*)/, effectName);

                    var newRow = new Array(sliceIndexConst);
                    newRow[0] = effectName;
                    newRow[1] = newRow[2] = newRow[3] = newRow[4] = newRow[5] = newRow[6] = 0;
                    newRow[10] = 1;
                    newRow.push(effect)
                    policiesArrayOfArrays.push(newRow);
                }
            }
        }

        // Add situation effects
        for (var j = 0; j < effects.length; j++) {
            var effect = effects[j];

            if (effect) {
                var policyRow = policiesArrayOfArrays.find(r => formatEffectName(r[0]) == formatEffectName(situation[1]));
                if (policyRow) {
                    policyRow.push(effect);
                }
                else {
                    var newRow = new Array(sliceIndexConst);
                    newRow[0] = situation[1];
                    newRow[1] = newRow[2] = newRow[3] = newRow[4] = newRow[5] = newRow[6] = 0;
                    newRow[10] = 1;
                    newRow.push(effect)
                    policiesArrayOfArrays.push(newRow);
                }
            }
        }
    }
}

function modifyArray(arrayOfArrays) {
    var allEffectsIndexes = {};
    var header = arrayOfArrays[0];

    // TOTAL COST
    header.splice(sliceIndexConst, 0, "TOTAL COST");

    // TOTAL INCOME
    header.splice(sliceIndexConst + 1, 0, "TOTAL INCOME");

    // POLICIES SLIDER
    header.splice(sliceIndexConst + 2, 0, "POLICIES SLIDER");

    // POLICY ACTIVE
    header.splice(sliceIndexConst + 3, 0, "POLICY ACTIVE");

    for (var i = 1; i < arrayOfArrays.length; i++) {
        var array = arrayOfArrays[i];

        var sliceIndex = sliceIndexConst;
        var part1 = array.slice(0, sliceIndex);
        var effects = array.slice(sliceIndex, array.length);

        // TOTAL COST
        part1.push("=(" + getColumnAddress(minCostColumnName, i + 1) + "+("
            + getColumnAddress(maxCostColumnName, i + 1) + "-"
            + getColumnAddress(minCostColumnName, i + 1) + ")*("
            + getColumnAddress(costMultiplierColumnName, i + 1) + "+"
            + getColumnAddress(policiesSliderColumnName, i + 1) + "))*"
            + getColumnAddress(isPolicyActiveColumnName, i + 1));
        sliceIndex++;

        // TOTAL INCOME
        part1.push("=(" + getColumnAddress(minIncomeColumnName, i + 1) + "+("
            + getColumnAddress(maxIncomeColumnName, i + 1) + "-"
            + getColumnAddress(minIncomeColumnName, i + 1) + ")*("
            + getColumnAddress(incomeMultiplierColumnName, i + 1) + "+"
            + getColumnAddress(policiesSliderColumnName, i + 1) + "))*"
            + getColumnAddress(isPolicyActiveColumnName, i + 1));
        sliceIndex++;

        // POLICIES SLIDER
        part1.push("=0");
        sliceIndex++;

        // POLICY ACTIVE
        part1.push("=0");
        sliceIndex++;

        for (var j = 0; j < effects.length; j++) {
            var effect = effects[j];

            if (effect && !effect.split) {
                var a = 5;
            }

            if (effect) {
                var effectParts = effect.split(",");
                var effectName = formatEffectName(effectParts[0]);
                var effectFormula = effectParts[1];

                if (effectFormula) {
                    effectFormula = fixInvalidEffectFormulaString(effectFormula);

                    effectFormula = "(" + effectFormula + ")*" + getColumnAddress(isPolicyActiveColumnName, i + 1);

                    effectFormula = setFormulaBoundaries(effectFormula.replace(/(?<=\W)x(?=\W*)/, getColumnAddress(policiesSliderColumnName, i + 1)));
                }

                var insertIndex = sliceIndex + Object.keys(allEffectsIndexes).length;
                if (allEffectsIndexes[effectName]) {
                    insertIndex = allEffectsIndexes[effectName];
                }
                else {
                    allEffectsIndexes[effectName] = insertIndex;
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

    parseOtherFormulas(arrayOfArrays);

    substituteFormulaVariables(allEffectsIndexes, arrayOfArrays);
}

function getColumnAddress(columnName, index) {
    return columnName + index;
}

function removeColumns(arrayOfArrays, columnNames) {
    var columnIndexes = columnNames.map(v => alphaToNum(v))
        .sort(function (a, b) { return b - a });

    for (var i = 0; i < arrayOfArrays.length; i++) {
        var array = arrayOfArrays[i];

        for (var j = 0; j < columnIndexes.length; j++) {
            var columnIndex = columnIndexes[j];

            arrayOfArrays[i].splice(columnIndex, 1)
        }
    }

    sliceIndexConst = sliceIndexConst - columnIndexes.length;
}

function createSumRow(header, arrayOfArrays) {
    arrayOfArrays.push(new Array(header.length));
    var sumRow = arrayOfArrays[arrayOfArrays.length - 1];

    for (var i = 0; i < header.length; i++) {
        var rowChar = numToAlpha(i);
        sumRow[i] = setFormulaBoundaries("SUM(" + rowChar + "2:" + rowChar + (arrayOfArrays.length - 1) + ")");
    }

    var headerIndexes = [
        alphaToNum("H"),
        alphaToNum("I"),
    ];

    for (var i = 0; i < headerIndexes.length; i++) {
        var headerIndex = headerIndexes[i];
        var rowChar = numToAlpha(headerIndex);
        sumRow[headerIndex] = "=SUM(" + rowChar + "2:" + rowChar + (arrayOfArrays.length - 1) + ")";
    }
}

// Create formulas from values
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

function substituteFormulaVariables(allEffectsIndexes, arrayOfArrays) {
    for (var i = 1; i < arrayOfArrays.length; i++) {
        var array = arrayOfArrays[i];

        if (i == 14) {
            var a = 5;
        }

        for (var j = 0; j < array.length; j++) {
            var element = array[j];

            if (element && element[0] == "=") {
                var regex = /[a-zA-Z_]{2,}/g;
                var matches = element.match(regex);

                if (matches) {
                    matches = matches.filter(m => m != "MIN" && m != "MAX");

                    if (matches.length) {
                        matches.forEach(m => {
                            var effect = formatEffectName(m);
                            var insertIndex = allEffectsIndexes[effect];

                            if (insertIndex) {
                                var newRegEx = new RegExp(effect, "i");
                                element = element.replace(newRegEx, numToAlpha(insertIndex) + (arrayOfArrays.length + 1));
                            }
                            else {
                                element = "=0";
                            }
                        });

                        array[j] = element;
                    }
                }
            }
        }
    }
}

function formatEffectName(effectName) {
    if (effectName) {
        // if (effectName[0] == "_") {
        //     effectName = effectName.slice(1, effectName.length);
        // }

        // if (effectName[effectName.length - 1] == "_") {
        //     effectName = effectName.slice(0, effectName.length - 1);
        // }

        return effectName.toLowerCase();
    }
}

function parseOtherFormulas(arrayOfArrays) {
    var formulasIndexes = [
        alphaToNum("D"),
        alphaToNum("G"),
    ];

    for (var i = 1; i < arrayOfArrays.length; i++) {
        var row = arrayOfArrays[i];

        formulasIndexes.forEach(function (j) {
            var formula = row[j];

            if (formula) {
                row[j] = parseMultiplierFormulas(formula)
            }
        });
    }
}

function parseMultiplierFormulas(formula) {
    if (formula) {
        var effects = formula.split(";");
        effects = effects.map((effect) => {
            var effectParts = effect.split(",");

            if (effectParts) {
                var effectName = effectParts[0];
                var effectFormula = effectParts[1];

                if (effectName && effectFormula) {
                    effectFormula = effectFormula.replace('x', effectName);

                    return effectFormula;
                }
                else {
                    // Invalud value
                    return "0";
                }
            }
        });

        var effectsFormula = setFormulaBoundaries(effects.join("+"), 0, 10);

        return effectsFormula;
    }
}

function setFormulaBoundaries(formula, a, b) {
    a = Number.isInteger(a) ? a : -1;
    b = Number.isInteger(b) ? b : 1;

    if (formula) {
        return "=MIN(" + b + ", MAX(" + a + "," + formula + "))";
    }
}

function fixInvalidEffectFormulaString(effectFormula) {
    var regex = /[\(\)]/g;
    var matches = effectFormula.match(regex);
    if (matches && matches.length % 2 == 1) {
        effectFormula = effectFormula.slice(0, effectFormula.length - 1);
    }

    return effectFormula;
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

function numToAlpha(num) {
    var alpha = '';

    for (; num >= 0; num = parseInt(num / 26, 10) - 1) {
        alpha = String.fromCharCode(num % 26 + 0x41) + alpha;
    }

    return alpha;
}

function alphaToNum(alpha) {

    var i = 0,
        num = 0,
        len = alpha.length;

    for (; i < len; i++) {
        num = num * 26 + alpha.charCodeAt(i) - 0x40;
    }

    return num - 1;
}