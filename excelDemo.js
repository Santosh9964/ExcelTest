const Exceljs = require('exceljs');

async function WriteexcelTest(searchText,replaceText,filePath) {

 

const workbook = new Exceljs.Workbook();
await workbook.xlsx.readFile(filePath);
const worksheet = workbook.getWorksheet('Sheet1');
const output = await readExcel(worksheet,searchText);
const cell = worksheet.getCell(output.row,output.column);
cell.value = replaceText;
await workbook.xlsx.writeFile(filePath);

}

async function readExcel(worksheet,searchText)
 {
    let output ={row:-1,column:-1};  

   worksheet.eachRow((row,rowNumber) =>
{
   row.eachCell((cell,colNumber) =>
   {
      if(cell.value===searchText)
      {
        output.row = rowNumber;
        output.column = colNumber;
      }

   })
   
})
return output;
}

WriteexcelTest("Mango","Rabbit","D://Excelutils//exceldownloadTest.xlsx");




