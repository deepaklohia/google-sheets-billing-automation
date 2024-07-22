// Email constants
const EMAIL_SUBJECT = 'DLA: Invoice';
const EMAIL_BODY = 'Hello!\rPlease find attached your invoice from Deepak Lohia Automations.';

function sendEmails() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  var emailCount = 0 ;

  if (EMAIL_OVERRIDE){
    var title = "you are about to send PDFs to " + EMAIL_ADDRESS_OVERRIDE + "\n from invoice sheet";
  }
  else{
    var title = "you are about to send email to CUSTOMER\n from invoice sheet";
  }
  
  var result = ui.alert('Send Invoice Emails ?',title,ui.ButtonSet.OK_CANCEL);
  if (result == ui.Button.CANCEL){
    return;
  }

  const invoicesSheet = ss.getSheetByName(INVOICES_SHEET_NAME);
  const invoicesData = invoicesSheet.getRange(1, 1, invoicesSheet.getLastRow(), invoicesSheet.getLastColumn()).getValues();
  const keysI = invoicesData.splice(0, 1)[0];
  const invoices = getObjects(invoicesData, createObjectKeys(keysI));
  ss.toast('Emailing Invoices', APP_TITLE, 1);
  invoices.forEach(function (invoice, index) {

    if (invoice.email_sent != 'Yes') {
      ss.toast(`Emailing Invoice for ${invoice.client_id}-${invoice.customer_name}`, APP_TITLE, 1);

      const fileId = invoice.invoice_link.match(/[-\w]{25,}(?!.*[-\w]{25,})/)
      const attachment = DriveApp.getFileById(fileId);

      let recipient = invoice.email;
      if (EMAIL_OVERRIDE) {
        recipient = EMAIL_ADDRESS_OVERRIDE
      }

      GmailApp.sendEmail(recipient, EMAIL_SUBJECT, EMAIL_BODY, {
        attachments: [attachment.getAs(MimeType.PDF)],
        name: APP_TITLE
      });

      emailCount += 1;
      invoicesSheet.getRange(index + 2, 10).setValue('Yes').setBackground("#b7e1cd"); //green color      
    }
  });

  ui.alert("Done", emailCount + " emails sent.", ui.ButtonSet.OK);
}
