function refreshSheetData(e) {
    if (e.changeType == 'FORMAT') {
        var editedRangeNotation = e.source.getActiveRange().getA1Notation()
        var cellNotationRangesToTrigger = []
        var activeSheet = e.source.getActiveSheet()
        var range = activeSheet.getDataRange()
        var formulas = range.getFormulas()
        for (var row in formulas) {
            for (var col in formulas[row]) {
                var formulaToEvaluate = formulas[row][col]
                var indexOformula = formulaToEvaluate.indexOf('sumOnFontColor')
                if (indexOformula > -1) {
                    formulaToEvaluate = formulaToEvaluate.substring(indexOformula)
                    Logger.log('formulaToEvaluate ' + formulaToEvaluate)
                    var rangeString = formulaToEvaluate.substring(formulaToEvaluate.indexOf('(') + 1)
                    var cellNotationToTrigger = (rangeString.split(","))[0].replace('"', '').replace('"', '') // X1:X23
                    Logger.log('cellNotationToTrigger ' + cellNotationToTrigger)
                    var triggerRange = activeSheet.getRange(cellNotationToTrigger)
                    var numRows = triggerRange.getNumRows()
                    var numCols = triggerRange.getNumColumns()
                    for (var i = 1; i <= numRows; i++) {
                        for (var j = 1; j <= numCols; j++) {
                          var currentNotation = triggerRange.getCell(i, j).getA1Notation()
                          if (editedRangeNotation == currentNotation) {
                              cellNotationRangesToTrigger.push(currentNotation)
                          }
                        }
                    }
                }
            }
        }
        
        cellNotationRangesToTrigger = uniqueArray(cellNotationRangesToTrigger)
        Logger.log('cellNotationRangesToTrigger ' + cellNotationRangesToTrigger)
        
        // Get all the ranges
        var oldValues = []
        for (var idx in cellNotationRangesToTrigger) {
            var rangeToTrigger = activeSheet.getRange(cellNotationRangesToTrigger[idx].trim())
            var existingValue = rangeToTrigger.getValue()
            oldValues.push(existingValue)

        }
      
      Logger.log('Old Values: ' + oldValues)
      

        for (var idx in cellNotationRangesToTrigger) {
            var rangeToTrigger = activeSheet.getRange(cellNotationRangesToTrigger[idx].trim())
            rangeToTrigger.setValue(0)
        }

        SpreadsheetApp.flush()

        for (var idx in cellNotationRangesToTrigger) {
            Logger.log('Notation: ' + cellNotationRangesToTrigger[idx] + ' Value: ' + oldValues[idx])
            var rangeToTrigger = activeSheet.getRange(cellNotationRangesToTrigger[idx].trim())
            rangeToTrigger.setValue(oldValues[idx])
        }

        Logger.log(cellNotationRangesToTrigger)
        Logger.log(oldValues)
    }

}

function uniqueArray(array) {
    var temp = array.reduce(function(previous, current) {
        previous[current] = true;
        return previous;
    }, {})

    return Object.keys(temp)
}

function onEdit(e) {
    Logger.log('Edit happened to the sheet')
    refreshSheetData(e)
}

function onChange(e) {
    Logger.log('Change happened to the sheet')
    refreshSheetData(e)
}