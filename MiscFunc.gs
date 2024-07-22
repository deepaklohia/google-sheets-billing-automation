function uploadDialog() {
    var htmlOutput = HtmlService
        .createHtmlOutputFromFile('importDialog.html')
        .setWidth(300)
        .setHeight(280);
 
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Import Data');
  }
   
function clearAllSheet(){
  let ui = SpreadsheetApp.getUi();
  var result = ui.alert('Are you Sure?','You are about to clear existing data from Clients / Services / Expenses / Invoice sheet!',ui.ButtonSet.OK_CANCEL);
  if (result == ui.Button.CANCEL){ return;}
  clearExpenses();
  clearServices();
  clearClients();

  clearOutputInvoice();
}

function deleteClientSheets(){
  let sa = SpreadsheetApp.getActive();
  let ui = SpreadsheetApp.getUi();

  var result = ui.alert('Are you Sure?','You are about to delete all clients invoice generated!',ui.ButtonSet.OK_CANCEL);
  if (result == ui.Button.CANCEL){ return;}

    sa.getSheets().forEach(
      function (s){
        var shNm = s.getName();
          if (shNm.substring(0, 2) == "XX") {
              var clSheet = sa.getSheetByName(shNm);
              sa.deleteSheet(clSheet);
          }
      }
  )
}

function clearOutputInvoice(){
  let ui = SpreadsheetApp.getUi();
    let sa = SpreadsheetApp.getActive();
    //let shName = shName;
    var cs = sa.getSheetByName(INVOICES_SHEET_NAME);
    var lastRow = cs.getLastRow() + 3;
    cs.getRange('A2:J' + lastRow ).clearContent();
    cs.getRange('J2:J' + lastRow ).setBackground(null);

}

  function clearExpenses(){
    let sa = SpreadsheetApp.getActive();
  
    //let shName = shName;
    var cs = sa.getSheetByName(EXPENSES_SHEET_NAME);
    var lastRow = cs.getLastRow() + 3;

    cs.getRange('A2:G' + lastRow ).clearContent();
    //cs.getRange('B2:F' + lastRow ).clearContent().setBackground(null);
  }

  function clearServices(){
    let sa = SpreadsheetApp.getActive();
  
    //let shName = shName;
    var cs = sa.getSheetByName(SERVICES_SHEET_NAME);
    var lastRow = cs.getLastRow() + 3;

    cs.getRange('A2:K' + lastRow ).clearContent();
    //cs.getRange('B2:F' + lastRow ).clearContent().setBackground(null);
  }

  function clearClients(){
    let sa = SpreadsheetApp.getActive();
  
    //let shName = shName;
    var cs = sa.getSheetByName(CUSTOMERS_SHEET_NAME);
    var lastRow = cs.getLastRow() + 10;

    cs.getRange('A2:AD' + lastRow ).clearContent();
    cs.getRange("C2:C" + lastRow).setBackground(null);//no color

  }

