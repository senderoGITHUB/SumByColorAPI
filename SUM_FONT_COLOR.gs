function sumByColor(range, colorRef, toAdd) {
    var sheet = SpreadsheetApp.getActiveSheet();
     var color = sheet.getRange(colorRef).getFontColor();
     var arr = [];
     var rangeSUM = sheet.getRange(range);
     var rangeVal = rangeSUM.getValues();
     var allColors = rangeSUM.getFontColors();

     for (var i = 0; i < allColors.length; i++) {
         for (var j = 0; j < allColors[0].length; j++) {
           if (toAdd) {
              if (allColors[i][j] == color && isTypeNumber(rangeVal[i][j])) arr.push(rangeVal[i][j]);
           } else if (!toAdd) {
              if (allColors[i][j] != color && isTypeNumber(rangeVal[i][j])) arr.push(rangeVal[i][j]);
           }
         }
     }

     return arr.reduce(function (a, b) {
         return a + b;
     }, 0)
 }

function sumOnFontColorOmit(rangeOmit, toOmitColorRef) {
   var colorRef = toOmitColorRef;
   var range = rangeOmit;
  
     return sumByColor(range, colorRef, false);
 }

 function sumOnFontColor(rangeBy, bycolorRef) {
   var colorRef = bycolorRef;
   var range = rangeBy;
   
     return sumByColor(range, colorRef, true);
 }


 function isTypeNumber(arg) {
     return typeof arg == 'number';
 }
