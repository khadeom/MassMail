function onOpen() 
{
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Automation')
      .addItem('send PDF Form', 'sendPDFForm')
      .addItem('send to all', 'sendFormToAll')
      .addToUi();
      console.log(MailApp.getRemainingDailyQuota())
      
}

function sendPDFForm()
{
  var row = SpreadsheetApp.getActiveSheet().getActiveCell().getRow();
  sendEmailWithAttachment(row);
}

function sendEmailWithAttachment(row)
{
 
  
  var client = getClientInfo(row);


 //var filename= 'Atharva Kulkarni Non Technical Event, Collage Making.jpg';
  
  var file = DriveApp.getFilesByName(client.filename);
  
  if (!file.hasNext()) 
  {
    console.error("Could not open file "+client.filename);
    return;
  }

  
  var template = HtmlService
      .createTemplateFromFile('email-template');
  template.client = client;
  var message = template.evaluate().getContent();
  
  
  MailApp.sendEmail({
    to: client.email,
    subject: "Engineer's Day participation Certificate",
    htmlBody: message,
    attachments: [file.next().getAs(MimeType.JPEG)]
  });
  
}

function getClientInfo(row)
{
   var sheet = SpreadsheetApp.getActive().getSheetByName('Sheet1');
   
   var values = sheet.getRange(row,1,1,5).getValues();
   var rec = values[0];
  //console.log(rec)
  var client = 
      {
        filename: rec[4],
        name: rec[1],
        first_name:rec[2],
        email: rec[3]
      };
  //client.name = client.name
  return client;
}

function sendFormToAll()
{
   var sheet = SpreadsheetApp.getActive().getSheetByName('Sheet1');
  
   var last_row = sheet.getDataRange().getLastRow();
    //console.log(last_row)
   for(var row=2; row <= last_row; row++)
   {
     var sent = sheet.getRange(row,7).getValues();
     if (sent!= "email sent"){
        //console.log("okay")
        sendEmailWithAttachment(row);
        Utilities.sleep(10);
        sheet.getRange(row,7).setValue("email sent");
     }



     
   }
}
