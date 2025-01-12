const Exceljs = require('exceljs');

// const workBook =new Exceljs.Workbook();
// workBook.xlsx.readFile('/Users/shreyanshjha/Downloads/excel-test.xlsx').then(function() {
//     const workSheet = workBook.getWorksheet('Sheet1');
//
//     workSheet.eachRow((row, rowNumber) => {
//         row.eachCell((cell, colNumber) => {
//             console.log(cell.value);
//         });
//     })
// });

// async and await
async  function writeExcelTest(searchText, replaceText, filePath) {
    const workBook =new Exceljs.Workbook();
    await workBook.xlsx.readFile(filePath)
    const workSheet = workBook.getWorksheet('Sheet1');
    const output = await readExcel(workSheet, searchText);

    const cell = workSheet.getCell(output.row, output.col);
    cell.value = replaceText;
    await workBook.xlsx.writeFile(filePath);
}

async function readExcel(workSheet, searchText) {
    let output = {row: -1, col: -1};
    workSheet.eachRow((row, rowNumber) => {
        row.eachCell((cell, colNumber) => {
            if(cell.value === searchText){
                output.row = rowNumber;
                output.col = colNumber;
            }
        });
    });
    return output;
}

writeExcelTest("Orange", "Red", "/Users/shreyanshjha/Downloads/excel-test.xlsx");