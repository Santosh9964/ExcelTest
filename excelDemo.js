const Exceljs = require('exceljs');

async function excelTest() {

const workbook = new Exceljs.Workbook();
await workbook.xlsx.readFile("D://exceldownloadTest.xlsx")

const worksheet = workbook.getWorksheet('Sheet1');
worksheet.eachRow((row,rowNumber) =>
{
   row.eachCell((cell,colNumber) =>
   {
       console.log(cell.value);

   })
})
}

excelTest();




