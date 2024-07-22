const sa = SpreadsheetApp.getActiveSpreadsheet();
var i = 0; 
function processDocuments() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  var result = ui.alert('Create Invoices in Bulk?','You are about to clear client print status',ui.ButtonSet.OK_CANCEL);
  if (result == ui.Button.CANCEL){
    return;
  }

  const clientSheet = ss.getSheetByName(CUSTOMERS_SHEET_NAME);
  const expenseSheet = ss.getSheetByName(EXPENSES_SHEET_NAME);
  const servicesSheet = ss.getSheetByName(SERVICES_SHEET_NAME);
  const invoicesSheet = ss.getSheetByName(INVOICES_SHEET_NAME);
  const invoiceTemplateSheet = ss.getSheetByName(INVOICE_TEMPLATE_SHEET_NAME);

  //show sheet for good
  invoiceTemplateSheet.showSheet()
  
  // Creating Table Chunk
  const clients = dataRangeToObject(clientSheet);
  const expenses = dataRangeToObject(expenseSheet);
  const services = dataRangeToObject(servicesSheet);

  ss.toast('Creating Invoices', APP_TITLE, 1);
  const invoices = [];
  //var cellRow = 1 ;
  var invoiceCount = 0 ;

  var lastRow = clientSheet.getLastRow()
  if (lastRow <= 1) {
    ui.alert("Error","Not enough Data in Client Sheet", ui.ButtonSet.OK);
    return ;
  }
  //clear existing status
  clientSheet.getRange("C2:C" + lastRow).setBackground(null);//no color

  //checking each ID
  clients.forEach(function (client, index) {
    let clientFound = services.filter(function (service) {
      return service.client_id == client.client_id;
    });

    //loop back
    if (clientFound.length <= 0){return; }

    //if inactive customer
    if (client.status != 'Active'){return; }
    
    ss.toast(`Creating Invoice for ${client.client_id}-${client.full_name}`, APP_TITLE, 1);

    //Generating invoice
    let invoice = createInvoiceForCustomer(client, expenses, services, invoiceTemplateSheet, ss.getId());
    invoices.push(invoice);

    clientSheet.getRange(index + 2, 3).setBackground("#b7e1cd"); //green color
    invoiceCount += 1 ;
    
    i +=1;
    if (i => 5) {  
      SpreadsheetApp.flush();
      ss.toast('Waiting to avoid error... ', APP_TITLE, 1);
      i = 0;
      Utilities.sleep(5000);
      SpreadsheetApp.flush();
    }
  });

  if (invoices.length == 0){
    ui.alert("not enough data", "ensure that valid clients exists in client file and other sheets", ui.ButtonSet.OK);
    return ;
  }

  // filling INVOICE STATUS DATA
  invoicesSheet.getRange(2, 1, invoices.length, invoices[0].length).setValues(invoices);  //start row, col , how many rows , lastcol
  clientSheet.activate();

  //hide for good
  var tempSheet = ss.getSheetByName(INVOICE_TEMPLATE_SHEET_NAME);
  tempSheet.hideSheet();

  ui.alert('Done', invoiceCount + " invoice generated (No Duplicates)", ui.ButtonSet.OK);
}

  function createInvoiceForCustomer(client, expenses, services, templateSheet, ssId) {

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    //ss = ss.getSheetByName(templateSheet);

    //GETTING SERVICES INFO
    let serviceItems = [];
    var totalServicesAmount = 0;
    //var servicesFound = false ;

    let clientServices = services.filter(function (service) {
      return service.client_id == client.client_id;
    });

    clientServices.forEach(function (serviceItem) {
      // billable customer items
      if (serviceItem.billable == "Yes" ){
        totalServicesAmount += parseFloat(serviceItem.amount);
        serviceItems.push([serviceItem.biller, serviceItem.date , serviceItem.code, serviceItem.task, serviceItem.description,
        serviceItem.hours, serviceItem.rate, serviceItem.amount]);
      }
    });
    
    //GETTING EXPENSE INFORMATION
    let expenseItems = [];
    var totalExpenseAmount = 0;
    var expensesFound = false ;
    let clientExpenses = expenses.filter(function (expense) {
      return expense.client_id == client.client_id;
    });
    if (clientExpenses.length > 0 ){
      expensesFound = true ;
      clientExpenses.forEach(function (expenseItem) {
        totalExpenseAmount += parseFloat(expenseItem.amount);
        expenseItems.push([expenseItem.date, expenseItem.expense_code , expenseItem.professional_services, expenseItem.invoice_number, " ", " ", expenseItem.amount]);
        //Logger.log( expenseItem.date + ">>" + expenseItem.expense_code + ">>" + expenseItem.invoice_number + ">>" + expenseItem.amount );
      });
    }

    // Generates a random invoice number. You can replace with your own document ID method.
    const invoiceNumber = Math.floor(100000 + Math.random() * 900000);

    // Calulates dates.
    const todaysDate = new Date().toDateString()
    const dueDate = new Date(Date.now() + 1000 * 60 * 60 * 24 * DUE_DATE_NUM_DAYS).toDateString()

    // Sets values in the template.
    //FILLING INVOICE TEMPLATE.
    
    //ADDING A SHEET FOR TEMPLATE
    var clientID = client.client_id;
    var tempSheet = ss.getSheetByName(clientID);
    if (tempSheet != null) {
      ss.deleteSheet(tempSheet);
    }
  
    templateSheet = ss.insertSheet(clientID, ss.getNumSheets(), {template: templateSheet});


    //customer information
    templateSheet.getRange('B11').setValue(client.full_name);
    templateSheet.getRange('B12').setValue(client.home_address);
    templateSheet.getRange('B13').setValue(client.home_city + " " + client.home_zip_code + " " + client.home_state);
    templateSheet.getRange('B14').setValue(client.email);
    templateSheet.getRange('B15').setValue(client.cell_number).setHorizontalAlignment('left');

    //invoice information
    templateSheet.getRange('G11').setValue(invoiceNumber);
    templateSheet.getRange('G13').setValue(todaysDate);
    templateSheet.getRange('G15').setValue(dueDate);

    //var rowsAdded = 0 ;
    var startRow = 0 ;
    var proServCellAdd ;
    var expCellAdd ;
    var totCellAdd ;
    var disCellAdd ;
    var prePymtCellAdd ;
    var pmtsCellAdd;
    var balDueCellAdd;

    //adding rows
    if (serviceItems.length > 1 ){
      templateSheet.insertRowsAfter(21, serviceItems.length -1 );  //row and number of rows
    }
    templateSheet.getRange(20, 2, serviceItems.length, 8).setValues(serviceItems);

    //totals
    endRow = 20 + serviceItems.length + 1   ;
    //templateSheet.getRange('E' + startRow ).setValue("FOR PROFESSIONAL SERVICES RENDERED:");
    templateSheet.getRange('G' + endRow ).setValue('=SUM(G20:G' + (endRow - 2) +  ')');
    proServCellAdd = 'I' + endRow ;
    templateSheet.getRange(proServCellAdd).setValue('=SUM(I20:I' + (endRow - 2) +  ')');

    if (expensesFound == true){
      //feeding expense items
      startRow = endRow + 3 ; //considering the summary
      if (expenseItems.length > 1 ){
        templateSheet.insertRowsAfter(startRow + 1, expenseItems.length -1 );  //row and number of rows
      }
      templateSheet.getRange(startRow, 3, expenseItems.length, 7).setValues(expenseItems);
      var endRow = startRow + expenseItems.length - 1
      expCellAdd = 'I' + (endRow + 2 ) ;
      templateSheet.getRange(expCellAdd).setValue('=SUM(I' + startRow + ':I' + endRow  +  ')');
      endRow += 2 ;
    }
    else{
      //delete rows if no expenses
      //endRow += 2 ;
      templateSheet.deleteRows( endRow + 2, 5);
    }

    //total bill
    endRow += 2;
    totCellAdd = 'I' + endRow ;
    if (expensesFound == true){ templateSheet.getRange(totCellAdd).setValue("=" + proServCellAdd + "+" + expCellAdd ); }
    else{ templateSheet.getRange(totCellAdd).setValue("=" + proServCellAdd );  }

    //previous balance
    endRow += 2;
    prePymtCellAdd = 'I' + endRow ;
    templateSheet.getRange(prePymtCellAdd).setValue(0);

    //discount
    endRow += 2;
    disCellAdd = 'I' + endRow ;
    templateSheet.getRange(disCellAdd).setValue(0);

    //payments
    endRow += 2;
    pmtsCellAdd = 'I' + endRow ;
    templateSheet.getRange(pmtsCellAdd).setValue(0);

    //balance due
    endRow += 3;
    balDueCellAdd = 'I' + endRow;

    //SpreadsheetApp.getUi().alert(disCellAdd);
    //templateSheet.getRange(balDueCellAdd).setValue("=" + totCellAdd + "+" + prePymtCellAdd + "-" + pmtsCellAdd + "-" + disCellAdd );
    templateSheet.getRange(balDueCellAdd).setValue("=" + totCellAdd + "+" + prePymtCellAdd + "-" + pmtsCellAdd + "-" + disCellAdd );

    // Cleans up and creates PDF.
    SpreadsheetApp.flush();
    Utilities.sleep(500); // Using to offset any potential latency in creating .pdf
    const pdf = createPDF(ssId, templateSheet, `Invoice#${invoiceNumber}-${client.client_id}`);
    
    //delete added rows;
    //templateSheet.deleteRows( 20, rowsAdded);
    return [invoiceNumber, todaysDate, client.client_id, client.full_name, client.email, '', totalServicesAmount, dueDate, pdf.getUrl(), 'No'];
}

/**
* Resets the template sheet by clearing out customer data.
* You use this to prepare for the next iteration or to view blank
* the template for design.
* 
* Called by createInvoiceForCustomer() or by the user via custom menu item.
*/
function clearTemplateSheet() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const templateSheet = ss.getSheetByName(INVOICE_TEMPLATE_SHEET_NAME);
  // Clears existing data from the template.
  const rngClear = templateSheet.getRangeList(['B11:B14', 'G11', 'G13', 'G15']).getRanges()
  rngClear.forEach(function (cell) {
    cell.clearContent();
  });
  // This sample only accounts for six rows of data 'B18:G24'. You can extend or make dynamic as necessary.
  //clearing data 
  templateSheet.getRange(20, 2, 9, 4).clearContent();
}

/**
 * Helper function that turns sheet data range into an object. 
 * 
 * @param {SpreadsheetApp.Sheet} sheet - Sheet to process
 * Return {object} of a sheet's datarange as an object 
 */
function dataRangeToObject(sheet) {
  const dataRange = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
  const keys = dataRange.splice(0, 1)[0];
  return getObjects(dataRange, createObjectKeys(keys));
}

/**
 * Utility function for mapping sheet data to objects.
 */
function getObjects(data, keys) {
  let objects = [];
  for (let i = 0; i < data.length; ++i) {
    let object = {};
    let hasData = false;
    for (let j = 0; j < data[i].length; ++j) {
      let cellData = data[i][j];
      if (isCellEmpty(cellData)) {
        continue;
      }
      object[keys[j]] = cellData;
      hasData = true;
    }
    if (hasData) {
      objects.push(object);
    }
  }
  return objects;
}
// Creates object keys for column headers.
function createObjectKeys(keys) {
  return keys.map(function (key) {
    return key.replace(/\W+/g, '_').toLowerCase();
  });
}
// Returns true if the cell where cellData was read from is empty.
function isCellEmpty(cellData) {
  return typeof (cellData) == "string" && cellData == "";
}
