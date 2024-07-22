function createPDFManually(){
  let sa = SpreadsheetApp.getActiveSpreadsheet();
  let ui = SpreadsheetApp.getUi();

  var result = ui.alert('Create PDF of Activesheet ?','ensure that you have selected invoice sheet',ui.ButtonSet.OK_CANCEL);
  if (result == ui.Button.CANCEL){
    return;
  }

  let ss = sa.getActiveSheet();
  let invoiceNumber = ss.getRange("G11").getValue();
  let clientName = ss.getRange("B11").getValue();
  //let timeStamp = Utilities.formatDate( new Date() , "GMT", 'ddmmyyhhmmss');

  const pdf = createPDF( sa.getId(), ss, `Invoice#${invoiceNumber}-${clientName}`);
  
  ui.alert("Manual PDF Generated for:" + clientName + "\n" + pdf.getUrl());
}


function createPDF(ssId, sheet, pdfName) {
  const lastRow = sheet.getLastRow();
  const fr = 0, fc = 0, lc = 10, lr = lastRow + 1 ;  //start row start col , end col , end row
  //const fr = 0, fc = 0, lc = 9, lr = 27;
  const url = "https://docs.google.com/spreadsheets/d/" + ssId + "/export" +
    "?format=pdf&" +
    "size=7&" +
    "fzr=true&" +
    "portrait=true&" +
    "fitw=true&" +
    "gridlines=false&" +
    "printtitle=false&" +
    "top_margin=0.5&" +
    "bottom_margin=0.25&" +
    "left_margin=0.5&" +
    "right_margin=0.5&" +
    "sheetnames=false&" +
    "pagenum=UNDEFINED&" +
    "attachment=true&" +
    "gid=" + sheet.getSheetId() + '&' +
    "r1=" + fr + "&c1=" + fc + "&r2=" + lr + "&c2=" + lc;

  //Logger.log("PDFNe>>" + pdfName);
  const params = { method: "GET", headers: { "authorization": "Bearer " + ScriptApp.getOAuthToken() } };
  const blob = UrlFetchApp.fetch(url, params).getBlob().setName(pdfName + '.pdf');

  // Gets the folder in Drive where the PDFs are stored.
  const folder = getFolderByName_(OUTPUT_FOLDER_NAME);

  const pdfFile = folder.createFile(blob);
  
  return pdfFile;
}
