// code by Hetvi: please contact H for any bugs/suggestions/improvements // written July 2021

var blockSize = [15,20];
var maxEmailsRow = 7;

function myFunction() {
  var rows = SpreadsheetApp.getActiveSpreadsheet().getSheets()[1].getDataRange().getValues();
  // allots a 15 cell * 20 cell canvas per person
  // var blockSize = [15,20]; // at 0 num rows, at 1 num cols
  // console.log(rows[0:3][0]);

  var artSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var appendStr = "Canvas of ";

  var allotRanges = [];
  var allEmail = [];
  var allName = [];
  for(var i = 0; i<rows.length; i++){
    // console.log(rows[i]);
    var emailId = rows[i][0];
    // var name = appendStr + rows[i][1];
    var name = rows[i][1];
    allEmail[i] = emailId;
    allName[i] = name;
  }

  // SpreadsheetApp.getActiveSpreadsheet().addEditors(allEmail);

  console.log(allEmail);
  console.log(allName);
  var protection = artSheet.protect().setDescription("art");

  protection.addEditor("indradhanu.iitdelhi@gmail.com");   // replace with admin email

  console.log(protection.getEditors());

    if (protection.canDomainEdit()) {
      protection.setDomainEdit(false);
    }

    var startRow = 1; 
    var startCol = 1;
    // var maxEmailsRow = 7; //max canvass-es to allot per row
    var ranges = [];

  for(var i = 0; i<allEmail.length; i++){

    if(i%maxEmailsRow==0 && i!=0){
        startRow+=blockSize[0];
        startCol = 1;
    }

    var range = artSheet.getRange(startRow,startCol,blockSize[0],blockSize[1]);
    startCol+=blockSize[1];
    ranges[i] = range;
    // if(i%maxEmailsRow==0 && i!=0){
    //     startRow+=blockSize[0];
    //     startCol = 1;
    // }
    var protection1 = range.protect().setDescription(allName[i]);
    protection1.removeEditors(protection1.getEditors());
    if (protection1.canDomainEdit()) {
      protection1.setDomainEdit(false);
    }
    protection1.addEditor(allEmail[i]);
    // colval = getRandomColor();
    // lst = ["E40303","FF8C00","FFED00","008026","004DFF","750787"]
    // colval = lst[i%lst.length];
    // var bgColor = SpreadsheetApp.newColor().setRgbColor(colval).build();
    // console.log(colval);
    // range.setBackgroundObject(bgColor);
  }
    protection.setUnprotectedRanges(ranges);
    console.log(startRow);
    console.log(startCol);
}

  function colourQuiltBg(){
    var startRow = 91; var startCol = 1;
    var artSheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];

    for(var i = 0; i<21; i++){
      if(i%maxEmailsRow==0 && i!=0){
        startRow+=blockSize[0];
        startCol = 1;
      }

      var range = artSheet.getRange(startRow,startCol,blockSize[0],blockSize[1]);
      startCol+=blockSize[1];
      // ranges[i] = range;
      lst = ["E40303","FF8C00","FFED00","008026","004DFF","750787"]
      colval = lst[i%lst.length];
      var bgColor = SpreadsheetApp.newColor().setRgbColor(colval).build();
      console.log(colval);
      range.setBackgroundObject(bgColor);
    }
  }

  // function getRandomColor(var inx) {
  // var letters = '0123456789ABCDEF';
  // var color = '#';
  // for (var i = 0; i < 6; i++) {
  //   color += letters[Math.floor(Math.random() * 16)];
  // }
  // lst = ["E40303","FF8C00","FFED00","008026","004DFF","750787"]
  // return lst[Math.floor(Math.random()*lst.length)];
  // return lst[inx%lst.length];
  // return color;
// }

function removeProtect(){
    var protections = SpreadsheetApp.getActiveSpreadsheet().getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (var i = 0; i < protections.length; i++) {
    protections[i].remove();
  }

}

// refs
// https://developers.google.com/apps-script/reference/spreadsheet/protection
// https://stackoverflow.com/questions/38993561/protect-ranges-with-google-apps-script
// https://cloud.google.com/blog/control-protected-ranges-and-sheets-google-sheets-apps-script
// https://webapps.stackexchange.com/questions/107380/how-to-make-a-copy-and-remove-all-protected-ranges-from-the-copied-sheet
// https://developers.google.com/apps-script/reference/spreadsheet/range#setbackgroundscolor
// https://developers.google.com/apps-script/reference/spreadsheet/color-builder#setRgbColor(String)