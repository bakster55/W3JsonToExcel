var XLSX = require('xlsx')
var fs = require('fs');
var path = require('path');

//
var policiesExcelPath = "./democracy3/data/simulation/policies.csv";
var situationsExcelPath = "./democracy3/data/simulation/situations.csv";

//
var clonesDronesPoliciesExcelPath = "./democracy3/clonesdrones/data/simulation/policies.csv";
var clonesDronesSituationsExcelPath = "./democracy3/clonesdrones/data/simulation/situations.csv";

//
var socialEngineeringPoliciesExcelPath = "./democracy3/socialengineering/data/simulation/policies.csv";

//
var extremismPoliciesExcelPath = "./democracy3/extremism/data/simulation/policies.csv";
var extremismSituationsExcelPath = "./democracy3/extremism/data/simulation/situations.csv";

//
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

var policiesArrayOfArrays = undefined;
var situationsArrayOfArrays = undefined;

var allEffectsIndexes = {};
var allSituationIndexes = {};

module.exports = { "init": init };

function init() {
    var policiesWorkBook = XLSX.readFile(policiesExcelPath, { type: 'string', raw: true });
    var situationsWorkBook = XLSX.readFile(situationsExcelPath, { type: 'string', raw: true });

    var clonesDronesPoliciesWorkBook = XLSX.readFile(clonesDronesPoliciesExcelPath, { type: 'string', raw: true });
    var clonesDronesSituationsWorkBook = XLSX.readFile(clonesDronesSituationsExcelPath, { type: 'string', raw: true });

    var socialEngineeringPoliciesWorkBook = XLSX.readFile(socialEngineeringPoliciesExcelPath, { type: 'string', raw: true });

    var extremismPoliciesWorkBook = XLSX.readFile(extremismPoliciesExcelPath, { type: 'string', raw: true });
    var extremismSituationsWorkBook = XLSX.readFile(extremismSituationsExcelPath, { type: 'string', raw: true });

    var initialPoliciesWorkBook = XLSX.readFile(initialPoliciesExcelPath, { type: 'string', raw: true });


    //
    var policiesSheet = policiesWorkBook.Sheets[policiesWorkBook.SheetNames[0]];
    var situationsSheet = situationsWorkBook.Sheets[situationsWorkBook.SheetNames[0]];

    var clonesDronesPoliciesSheet = clonesDronesPoliciesWorkBook.Sheets[clonesDronesPoliciesWorkBook.SheetNames[0]];
    var clonesDronesSituationsSheet = clonesDronesSituationsWorkBook.Sheets[clonesDronesSituationsWorkBook.SheetNames[0]];

    var socialEngineeringPoliciesSheet = socialEngineeringPoliciesWorkBook.Sheets[socialEngineeringPoliciesWorkBook.SheetNames[0]];

    var extremismPoliciesSheet = extremismPoliciesWorkBook.Sheets[extremismPoliciesWorkBook.SheetNames[0]];
    var extremismSituationsSheet = extremismSituationsWorkBook.Sheets[extremismSituationsWorkBook.SheetNames[0]];

    var initialPoliciesSheet = initialPoliciesWorkBook.Sheets[initialPoliciesWorkBook.SheetNames[0]];


    //
    var resultArrayOfArrays = sheet2arr(policiesSheet);
    policiesArrayOfArrays = sheet2arr(policiesSheet);
    situationsArrayOfArrays = sheet2arr(situationsSheet);

    var clonesDronesPoliciesArrayOfArrays = sheet2arr(clonesDronesPoliciesSheet);
    var clonesDronesSituationsArrayOfArrays = sheet2arr(clonesDronesSituationsSheet);

    var socialEngineeringPoliciesArrayOfArrays = sheet2arr(socialEngineeringPoliciesSheet);

    var extremismPoliciesArrayOfArrays = sheet2arr(extremismPoliciesSheet);
    var extremismSituationsArrayOfArrays = sheet2arr(extremismSituationsSheet);

    var initialresultArrayOfArrays = sheet2arr(initialPoliciesSheet);


    //
    clonesDronesPoliciesArrayOfArrays.splice(0, 1);
    clonesDronesSituationsArrayOfArrays.splice(0, 1);

    socialEngineeringPoliciesArrayOfArrays.splice(0, 1);

    extremismPoliciesArrayOfArrays.splice(0, 1);
    extremismSituationsArrayOfArrays.splice(0, 1);


    //
    removeColumns(policiesArrayOfArrays, ["A", "C", "D", "E", "F", "G", "H", "I", "J", "S", "K", "O"], false);
    removeColumns(clonesDronesPoliciesArrayOfArrays, ["A", "C", "D", "E", "F", "G", "H", "I", "J", "S", "K", "O"], false);
    removeColumns(socialEngineeringPoliciesArrayOfArrays, ["A", "C", "D", "E", "F", "G", "H", "I", "J", "S", "K", "O"], false);
    removeColumns(extremismPoliciesArrayOfArrays, ["A", "C", "D", "E", "F", "G", "H", "I", "J", "S", "K", "O"], false);
    removeColumns(resultArrayOfArrays, ["A", "C", "D", "E", "F", "G", "H", "I", "J", "S", "K", "O"]);


    //
    resultArrayOfArrays = resultArrayOfArrays.concat(socialEngineeringPoliciesArrayOfArrays)
        .concat(extremismPoliciesArrayOfArrays).concat(clonesDronesPoliciesArrayOfArrays);
    policiesArrayOfArrays = policiesArrayOfArrays.concat(socialEngineeringPoliciesArrayOfArrays)
        .concat(extremismPoliciesArrayOfArrays).concat(clonesDronesPoliciesArrayOfArrays);
    situationsArrayOfArrays = situationsArrayOfArrays.concat(extremismSituationsArrayOfArrays)
        .concat(clonesDronesSituationsArrayOfArrays);


    mergeArrays(resultArrayOfArrays, situationsArrayOfArrays);

    modifyArray(resultArrayOfArrays);

    initPolicies(initialresultArrayOfArrays, resultArrayOfArrays);

    createSumRow(resultArrayOfArrays[0], resultArrayOfArrays);

    var workSheet = XLSX.utils.aoa_to_sheet(resultArrayOfArrays);

    createFormulas(workSheet);


    XLSX.utils.book_append_sheet(policiesWorkBook, workSheet);

    XLSX.writeFile(policiesWorkBook, outputExcelPath);
};

function initPolicies(initialresultArrayOfArrays, resultArrayOfArrays) {
    for (var i = 1; i < initialresultArrayOfArrays.length; i++) {
        var initialPolicy = initialresultArrayOfArrays[i];

        var effectName = formatEffectName(initialPolicy[0]);

        var policyRow = resultArrayOfArrays.find(r => formatEffectName(r[0]) == effectName);
        if (policyRow) {
            policyRow[alphaToNum(isPolicyActiveColumnName)] = 1;
            policyRow[alphaToNum(policiesSliderColumnName)] = initialPolicy[1];
        }
    }
}

function mergeArrays(resultArrayOfArrays, situationsArrayOfArrays) {
    for (var i = 1; i < situationsArrayOfArrays.length; i++) {
        var situation = situationsArrayOfArrays[i];

        if (i == 28) {
            var a = 5;
        }

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

                if (!getPolicy(effectName)) {
                    effect = effect.replace(/(?<=\W)x(?=\W*)/, effectName);
                }

                var policyRow = resultArrayOfArrays.find(r => formatEffectName(r[0]) == effectName);
                if (policyRow) {
                    policyRow.push(effect);
                }
                else {
                    var newRow = new Array(sliceIndexConst);
                    newRow[0] = effectName;
                    newRow[1] = newRow[2] = newRow[3] = newRow[4] = newRow[5] = newRow[6] = 0;
                    newRow.push(effect)
                    resultArrayOfArrays.push(newRow);
                }
            }
        }

        // Add situation effects
        // for (var j = 0; j < effects.length; j++) {
        //     var effect = effects[j];

        //     if (effect) {
        //         var policyRow = resultArrayOfArrays.find(r => formatEffectName(r[0]) == formatEffectName(situation[1]));
        //         if (policyRow) {
        //             policyRow.push(effect);
        //         }
        //         else {
        //             var newRow = new Array(sliceIndexConst);
        //             newRow[0] = situation[1];
        //             newRow[1] = newRow[2] = newRow[3] = newRow[4] = newRow[5] = newRow[6] = 0;
        //             newRow.push(effect)
        //             resultArrayOfArrays.push(newRow);
        //         }
        //     }
        // }
    }
}

function modifyArray(arrayOfArrays) {
    var header = arrayOfArrays[0];

    // ADD TOTAL COST COLUMN
    header.splice(sliceIndexConst, 0, "TOTAL COST");

    // ADD TOTAL INCOME COLUMN
    header.splice(sliceIndexConst + 1, 0, "TOTAL INCOME");

    // ADD POLICIES SLIDER COLUMN
    header.splice(sliceIndexConst + 2, 0, "POLICIES SLIDER");

    // ADD POLICY ACTIVE COLUMN
    header.splice(sliceIndexConst + 3, 0, "POLICY ACTIVE");

    for (var i = 1; i < arrayOfArrays.length; i++) {
        var row = arrayOfArrays[i];

        var sliceIndex = sliceIndexConst;
        var policy = row.slice(0, sliceIndex);
        var effects = row.slice(sliceIndex, row.length);

        // SET TOTAL COST FORMULA
        policy.push("=(" + getColumnAddress(minCostColumnName, i + 1) + "+("
            + getColumnAddress(maxCostColumnName, i + 1) + "-"
            + getColumnAddress(minCostColumnName, i + 1) + ")*("
            + getColumnAddress(costMultiplierColumnName, i + 1) + "+"
            + getColumnAddress(policiesSliderColumnName, i + 1) + "))*"
            + getColumnAddress(isPolicyActiveColumnName, i + 1));
        sliceIndex++;

        // SET TOTAL INCOME FORMULA
        policy.push("=(" + getColumnAddress(minIncomeColumnName, i + 1) + "+("
            + getColumnAddress(maxIncomeColumnName, i + 1) + "-"
            + getColumnAddress(minIncomeColumnName, i + 1) + ")*("
            + getColumnAddress(incomeMultiplierColumnName, i + 1) + "+"
            + getColumnAddress(policiesSliderColumnName, i + 1) + "))*"
            + getColumnAddress(isPolicyActiveColumnName, i + 1));
        sliceIndex++;

        // SET POLICIES SLIDER FORMULA
        policy.push("=0");
        sliceIndex++;

        // SET POLICY ACTIVE FORMULA
        policy.push("=0");
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

                var insertIndex = 0;
                if (getSituation(effectName)) {
                    insertIndex = sliceIndex + Object.keys(allSituationIndexes).length;
                    if (allSituationIndexes[effectName]) {
                        insertIndex = allSituationIndexes[effectName];
                    }
                    else {
                        allSituationIndexes[effectName] = insertIndex;
                    }
                }
                else {
                    insertIndex = sliceIndex + Object.keys(allEffectsIndexes).length + 60;
                    if (allEffectsIndexes[effectName]) {
                        insertIndex = allEffectsIndexes[effectName];
                    }
                    else {
                        allEffectsIndexes[effectName] = insertIndex;
                    }
                }



                if (insertIndex >= policy.length) {
                    policy.length = insertIndex + 1;
                }

                if (insertIndex >= header.length) {
                    header.length = insertIndex + 1;
                }

                header[insertIndex] = effectName;
                policy[insertIndex] = effectFormula;
            }
        }

        arrayOfArrays[i] = policy;
    }

    parseOtherFormulas(arrayOfArrays);

    substituteFormulaVariables(allEffectsIndexes, arrayOfArrays);
}

function getColumnAddress(columnName, index) {
    return columnName + index;
}

function removeColumns(arrayOfArrays, columnNames, enableSliceIndex = true) {
    var columnIndexes = columnNames.map(v => alphaToNum(v))
        .sort(function (a, b) { return b - a });

    for (var i = 0; i < arrayOfArrays.length; i++) {
        var array = arrayOfArrays[i];

        if (i == 123) {
            var a = 5;
        }

        for (var j = 0; j < columnIndexes.length; j++) {
            var columnIndex = columnIndexes[j];

            arrayOfArrays[i].splice(columnIndex, 1)
        }
    }

    if (enableSliceIndex) {
        sliceIndexConst = sliceIndexConst - columnIndexes.length;
    }
}

function createSumRow(header, arrayOfArrays) {
    arrayOfArrays.push(new Array(header.length));
    var sumRow = arrayOfArrays[arrayOfArrays.length - 1];

    // Effects sum
    for (var i = 0; i < header.length; i++) {
        var rowChar = numToAlpha(i);
        sumRow[i] = setFormulaBoundaries("SUM(" + rowChar + "2:" + rowChar + (arrayOfArrays.length - 1) + ")");
    }


    // Multipliers sum
    var headerIndexes = [
        alphaToNum("H"),
        alphaToNum("I"),
    ];

    for (var i = 0; i < headerIndexes.length; i++) {
        var headerIndex = headerIndexes[i];
        var rowChar = numToAlpha(headerIndex);
        sumRow[headerIndex] = "=SUM(" + rowChar + "2:" + rowChar + (arrayOfArrays.length - 1) + ")";
    }

    // SET POLICY ACTIVE FORMULA
    for (var i = 0; i < arrayOfArrays.length; i++) {
        var row = arrayOfArrays[i];

        if (!policiesArrayOfArrays.find(p => p[0] == row[0])) {
            if (getSituation(row[0])) {
                //row[10] = "=IF(" + numToAlpha(allEffectsIndexes[formatEffectName(row[0])]) + (arrayOfArrays.length) + "<0, 1, 0)";
            }
            else {
                row[10] = "=1";
            }
        }
    }
}

function getSituation(name) {
    var nameInt = formatEffectName(name);
    return situationsArrayOfArrays.find(p => formatEffectName(p[1]) == nameInt);
}

function getPolicy(name) {
    var nameInt = formatEffectName(name);
    return policiesArrayOfArrays.find(p => formatEffectName(p[0]) == nameInt);
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