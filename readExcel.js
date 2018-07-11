
class Invoice {

  constructor(filename, location) {
    this.filename = filename;
    this.location = location;
  }

  getFinishedFilename() {
    let date = new Date();
    return `${this.filename + this.location}-${date.getFullYear()}-${date.getMonth() + 1}-${date.getDate()}`
  }
}

let invoice = new Invoice('check', 'USA');

//Read a file
var workbook = new Excel.Workbook();
workbook.xlsx.readFile("data/Sample.xlsx").then(function () {
            
//Get sheet by Name
var worksheet=workbook.getWorksheet('Sheet1');
            
//Get Lastrow
var row = worksheet.lastRow;
//Save the workbook
return workbook.xlsx.writeFile("data/checked.xlsx");
 
});