function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Functions')
      .addItem('Foreclosure Mailer', 'createMailer')
      .addToUi();
}

function createMailer() {
  const template = DriveApp.getFileById('1bkKOkeoCyGvn9Z_yO5i19GVkoTI4nmJw9fKlC-og130');
  const destinationFolder = DriveApp.getFolderById('13Ea-OTi0DrvZCEoi3Q89RTg2cMfkC4j_');
  const sheet = SpreadsheetApp.openById('1GjisxWYNefoB1huG7fhUU537BnNuYpEPHxF50Hp21X4').getSheetByName('Sheet1');
  const rows = sheet.getDataRange().getValues();

  rows.forEach(function(row, index) {
    if(index === 0) return;
    if(row[8]) return;
    if(row[0]){} else {return};

    const copy = template.makeCopy(`Foreclosure Mailer - ${row[0]}`, destinationFolder);
    const doc = DocumentApp.openById(copy.getId());
    const body = doc.getBody();
    
    body.replaceText('{{Name}}', row[0]);
    body.replaceText('{{Address1}}', row[1]);
    body.replaceText('{{Address2}}', row[2]);
    body.replaceText('{{Property Value}}', row[6]);

    doc.saveAndClose();

    // Log mailer doc url to sheet
    const url = doc.getUrl();
    sheet.getRange(index + 1, 9).setValue(url);

    Logger.log(`Mailer created: ${row[0]}`);
  })
}