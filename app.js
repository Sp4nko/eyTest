const ExcelJS = require('exceljs');
const wb = new ExcelJS.Workbook();

const fileName = 'DevExercise.xlsx';

let sheetSumCol = [];
let rentObj = {}


wb.xlsx.readFile(fileName).then(() => {
    // Iterate over all sheets
    wb.eachSheet(function (worksheet, sheetId) {
        console.log(worksheet.columnCount)
       

        const testCol = worksheet.getColumn(worksheet.columnCount + 1);
        testCol.header = "testCell"
        

        //First iterate over all rows that have values in a worksheet to write sums and fill "Rent"
        worksheet.eachRow(function (row, rowNumber) {
            //sum all the numbers and output the result at the last cell.
            worksheet.getCell('G' + rowNumber).value = sumRow(row.values[3], row.values[4], row.values[5], row.values[6])
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
        getMaxNumAndFillRow(worksheet)

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

    //   console.log(ws.columnCount)




    // for (const prop in rentObj) {
    //     ws.getCell('A1').value = rentObj[prop]
    //     ws.getCell('A2').value = prop



    //     console.log(`rentObj.${prop} = ${rentObj[prop]}`);
    //   }


    wb.xlsx.writeFile("copy" + fileName)
        .then(function () {
            // done
        });

}).catch(err => {
    console.log(err.message);
});

//function to detect if all given values are numbers and sum them.
const sumRow = (a, b, c, d) => {
    if (!isNaN(a) && !isNaN(b) && !isNaN(c) && !isNaN(d)) {
        return a + b + c + d;
    } else {
        // console.log("NaN");
    }

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