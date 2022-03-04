const ExcelJS = require('exceljs');
const wb = new ExcelJS.Workbook();

const fileName = 'DevExercise.xlsx'; // The Excel test file we're manipulating 

let sheetSumCol = [] //  Array that holds all rows in the sum Col.
let rentObj = {} //   Obj that holds rent sums orderd by sheet.
let numbersSum = [] //    Array to store every row sum.
let sumOfSums = {} //     Obj of all sums by catagory.

//open the excel file
wb.xlsx.readFile(fileName).then(() => {
    // Iterate over all sheets
    wb.eachSheet(function (worksheet, sheetId) {
        const lastCol = worksheet.getColumn(worksheet.columnCount + 1); // Targeting last Col of the sheet.

        //First iterate over all rows that have values
        worksheet.eachRow(function (row, rowNumber) {

            //sum all the numbers and output the result to a global arr (numbersSum).
            let summedRow = sumRow(row.values)

            //check if the row on the loop is "Rent"
            if (row.values[2] == "Rent") {
                rentObj[worksheet.name] = summedRow; // Adds rent Sum to rentObj by sheet name.

                // Fill yellow color to rent row.
                worksheet.getRow(rowNumber).fill = {
                    type: 'pattern',
                    pattern: 'darkVertical',
                    fgColor: {
                        argb: 'FFFF00'
                    }
                };
            }
        })

        numbersSum.shift(); // Erasing col title
        numbersSum.unshift('Sum'); //  Setting col title to: 'Sum'
        lastCol.header = numbersSum; //   Writing data

        // Detects the biggest sum value on sheet and filling the row with orange.
        getMaxNumAndFillRow(worksheet) 
        numbersSum = []; // Clear global array for next sheet use;
    });


    //creating rentTotal sheet
    wb.addWorksheet('RentTotal');
    const ws = wb.getWorksheet('RentTotal');

    //writting data into sheet
    const sumCol = ws.getColumn(1);
    sumCol.header = Object.values(rentObj)
    const sheetName = ws.getColumn(2);
    sheetName.header = Object.keys(rentObj)


    //creating All_sheets_total_by_Catagory sheet
    wb.addWorksheet('All_sheets_total_by_Catagory');
    const lastSheet = wb.getWorksheet('All_sheets_total_by_Catagory');

    //writting data into sheet
    sumOfSums.Department = "All sheets total"
    const totalSumSheetCol1 = lastSheet.getColumn(1);
    totalSumSheetCol1.header = Object.keys(sumOfSums)
    const totalSumSheetCol2 = lastSheet.getColumn(2);
    totalSumSheetCol2.header = Object.values(sumOfSums)

    // Write File.
    wb.xlsx.writeFile("Proccessed_" + fileName)
        .then(function () {
            console.log("Success!");
        });

}).catch(err => {
    console.log(err.message);
});

//function to detect if all given values are numbers and sum them.
const sumRow = (arr) => {
    // console.log(arr[2]);
    let rowTotal = 0;
    for (let cell in arr) {
        if (!isNaN(arr[cell])) {
            rowTotal += arr[cell];
        }
    }
    numbersSum.push(rowTotal);

    // Initialize first value to prevent NaN.
    if (!sumOfSums[arr[2]]) {
        sumOfSums[arr[2]] = 0
    }

    // Adds up the numbers by Catagory.
    sumOfSums[arr[2]] += rowTotal

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

    // Fill biggest number's row to orange.
    worksheet.getRow(sheetSumCol.indexOf(maxResult) + 2).fill = {
        type: 'pattern',
        pattern: 'darkVertical',
        fgColor: {
            argb: 'FFD580'
        }
    };

}