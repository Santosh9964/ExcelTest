const Exceljs = require('exceljs');

async function excelTest() {

 let output ={row:-1,column:-1};   

const workbook = new Exceljs.Workbook();
await workbook.xlsx.readFile("D://Excelutils//exceldownloadTest.xlsx")

const worksheet = workbook.getWorksheet('Sheet1');
worksheet.eachRow((row,rowNumber) =>
{
   row.eachCell((cell,colNumber) =>
   {
      if(cell.value==="Banana")
      {
        output.row = rowNumber;
        output.column = colNumber;
      }

   })
})

const cell = worksheet.getCell(output.row,output.column);
cell.value = "Republic";
await workbook.xlsx.writeFile("D://Excelutils//exceldownloadTest.xlsx");

}

excelTest();




