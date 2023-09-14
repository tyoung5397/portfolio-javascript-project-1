/** pad */
function pad(n) {
  return n < 10 ? '0' + n : n
}

/**
 *   FUNCTION: Sales_Report_Email()
 *    PURPOSE: Convert to XLSX and email recipientList.
 */
function Sales_Report_Email(recipientList) {

  try {
    var date = new Date();
    var dateString = date.getFullYear() + '-' + pad(date.getMonth()+1) + '-'  + pad(date.getDate());
    var debug = new Boolean(true);
    var emailAddress = recipientList;
    var emailBody = 'Please find the attached ' + dateString + ' sales pipeline report.';
    var emailSubject = 'CompanyPlaceholder Sales Pipeline ' + dateString;
    var fileName = emailSubject;
    var reportName = 'Report - Detail';
    var spreadSheet = SpreadsheetApp.getActive();
   
    // convert to XLSX    
    //var url = 'https://docs.google.com/spreadsheets/d/' + spreadSheet.getId() + '/export?' + 'exportFormat=xlsx&format=xlsx&gid=' + spreadSheet.getSheetByName('Report - Detail').getSheetId();
    var url = 'https://docs.google.com/spreadsheets/d/' + spreadSheet.getId() + '/export?' + 'exportFormat=xlsx&format=xlsx';
  
    var params = {
      method      : "get",
      headers     : {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
      muteHttpExceptions: true
    };
    
    var blob = UrlFetchApp.fetch(url, params).getBlob();
    blob.setName(fileName + ".xlsx");
    
    // email
    var htmlBody = 
        '<table class="MsoNormalTable" border="0" cellspacing="0" cellpadding="0" width="600" style="width:6.25in"><tr><td colspan="2" style="padding:0in 0in 0in 0in"><p><span style="font-size:11.5pt;font-family:&quot;Calibri&quot;,sans-serif;color:black">'
        + '<p>' + emailBody + '<br><br>'
        + 'Regards,<br><br></p></span><span style="font-size:11.5pt;font-family:&quot;Calibri&quot;,sans-serif;color:black">'
        + Session.getActiveUser()
        + '<b></b></span></p><p style="margin-right:0in;margin-bottom:7.5pt;margin-left:0in"><span style="font-size:10.5pt;font-family:&quot;Calibri&quot;,sans-serif;color:#666666"><img border="0" width="250" height="65" style="width:2.6041in;height:.677in" id="_x0000_i1025" src="http://www.CompanyPlaceholder.com/Images/email-signature/CompanyPlaceholder-email-signature-retina.png" alt="CompanyPlaceholder Logo"></span></p></td></tr><tr style="height:15.0pt"><td colspan="2" style="border:none;border-bottom:solid #eeeeee 1.5pt;padding:0in 0in 0in 0in;height:15.0pt"></td></tr></table><p><span style="font-size:8.5pt;font-family:&quot;Calibri&quot;,sans-serif;color:#999999">Confidentiality Note: The information contained in this email and document(s) attached are for the exclusive use of the addressee and may contain confidential, privileged and non-disclosable information. If the recipient of this email is not the addressee, such recipient is strictly prohibited from reading, photocopying, distribution or otherwise using this email or its contents in any way.</span></p></div></div></div></div></div></div></div></div></div></div></div></div></div></div></div></div></div></div></div><p class="MsoNormal">'
    
    MailApp.sendEmail({
      to: emailAddress,
      subject: emailSubject,
      htmlBody: htmlBody,
      attachments: [blob],      
    });
    
  } catch (f) {
    Logger.log(f.toString());
    SpreadsheetApp.getUi().alert(f.toString());
  }  

}
