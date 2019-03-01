//import exceljs and XMLHttpRequest
var Excel = require('exceljs');
var XMLHttpRequest = require("xmlhttprequest").XMLHttpRequest;

//Read Excel file
var workbook = new Excel.Workbook();

//Give Excel file path
var path = <file_path>;
workbook.xlsx.readFile(path).then(function () {
            
//Get sheet by Name
var worksheet=workbook.getWorksheet('<sheet name>');
            

            
//get url value and check response and write response
var i=1;
do{
    //column to be loaded from excel is 'B'
    var C1 ='B';
    //column to be write the result in excel sheet is 'C'
    var C2 ='C';
    var cell_url = C1+i;
    var cell_res = C2+i;
    var url = worksheet.getCell(cell_url).value;
    console.log(url);
      var http = new XMLHttpRequest();
      http.open('HEAD', url, false);
      http.send(null);
      console.log(http.status);
      worksheet.getCell(cell_res).value = http.status;   
i++;
}
while(i<=worksheet.rowCount)



//Save the workbook
return workbook.xlsx.writeFile(path);

});
