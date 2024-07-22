function onOpen(e) {

const menu = SpreadsheetApp.getUi().createMenu(APP_TITLE)
  menu
     .addItem('STEP1: Prepare Data', 'uploadDialog')

    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('STEP2: Generate Invoice')
      .addItem('Create Bulk Invoice (PDFs)', 'processDocuments')
      .addItem('Create Single Invoice (PDF)', 'createPDFManually'))

    .addSeparator()
    .addItem('STEP3: Send Email', 'sendEmails')

    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('Data Cleaning')
      .addItem('Clear All Input/Output Data', 'clearAllSheet')
      .addItem('DELETE Client Invoice Sheets', 'deleteClientSheets'))
    .addToUi();
}
