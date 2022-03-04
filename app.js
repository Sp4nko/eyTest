const ExcelJS = require('exceljs');
const wb = new ExcelJS.Workbook();

const fileName = 'DevExercise.xlsx';

let sheetSumCol = []
let rentObj = {}
let numbersSum = [] //array to store every row sum


wb.xlsx.readFile(fileName).then(() => {
    // Iterate over all sheets
    wb.eachSheet(function (worksheet, sheetId) {       

        const lastCol = worksheet.getColumn(worksheet.columnCount + 1);

        //First iterate over all rows that have values in a worksheet to write sums and fill "Rent"
        worksheet.eachRow(function (row, rowNumber) {

            //sum all the numbers and output the result to a global arr (numbersSum)
            sumRow(row.values)

            //check if the row on the loop is "Rent" and is true filling yellow background
            if (row.values[2] == "Rent") {
                worksheet.getRow(rowNumber).fill = {
                    type: 'pattern',
                    pattern: 'darkVertical',
                    fgColor: {
                        argb: 'FFFF00'
                    }
                };
            }
        })

        numbersSum.shift();// erasing col title (currently 0)
        numbersSum.unshift('Sum'); // setting col title to: 'Sum'
        lastCol.header = numbersSum;

        getMaxNumAndFillRow(worksheet)
        numbersSum = []; // clear global array for next sheet;
    });

    wb.addWorksheet('RentTotal');
    const ws = wb.getWorksheet('RentTotal');
    ws.columns = [
        { header: 'Sum', key: 'sum', width: 32 },
        { header: 'From sheet', key: 'sheet', width: 10, outlineLevel: 1 }
      ];
      
      const sumCol = ws.getColumn(1);
      sumCol.header = Object.values(rentObj)

      const sheetName = ws.getColumn(2);
      sheetName.header = Object.keys(rentObj)



    wb.xlsx.writeFile("copy" + fileName)
        .then(function () {
            // done
        });

}).catch(err => {
    console.log(err.message);
});

//function to detect if all given values are numbers and sum them.
const sumRow = (arr) => {
    let rowTotal = 0;
    for (let cell in arr) {
        if (!isNaN(arr[cell])) {
        rowTotal += arr[cell];
        } 
      }
      numbersSum.push(rowTotal);
    return rowTotal;
}

//Getting biggest number from sum col and filling row with orange color.
const getMaxNumAndFillRow = (worksheet) => {
    let maxResult;
    sheetSumCol = worksheet.getColumn(worksheet.columnCount).values
    sheetSumCol.shift() //Poping empty index from the array.
    sheetSumCol.shift() //Poping empty index from the array.
    maxResult = Math.max.apply(null, sheetSumCol);
    sheetSumCol.indexOf(maxResult)

    rentObj[worksheet.name] = maxResult;
    worksheet.getRow(sheetSumCol.indexOf(maxResult) + 2).fill = {
        type: 'pattern',
        pattern: 'darkVertical',
        fgColor: {
            argb: 'FFD580'
        }
    };

}