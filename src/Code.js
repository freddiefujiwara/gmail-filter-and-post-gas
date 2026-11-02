export function myFunction() {
  // the speech URL prefix comes from the SPEECH_URL_PREFIX script property
  const IFTTT = PropertiesService.getScriptProperties().getProperty('SPEECH_URL_PREFIX');
  const URL = '%E9%96%A2%E9%80%A3%E3%81%AE%E3%83%A1%E3%83%BC%E3%83%AB%E3%81%8C%E5%B1%8A%E3%81%84%E3%81%A6%E3%81%84%E3%81%BE%E3%81%99%E3%80%82%E7%A2%BA%E8%AA%8D%E3%81%97%E3%81%A6%E4%B8%8B%E3%81%95%E3%81%84%E3%80%82';
  // get the data from the spreadsheet and check all the labels.
  SpreadsheetApp.getActiveSheet().getDataRange().getValues().forEach((label, v) => {
    // skip blank labels
    if (label.length < 2) {
      return;
    }

    // check all relevant unread messages
    let gotMail = false;
    GmailApp.search('label:' + label[0] + ' label:unread').forEach((thread, t) => {
      thread.getMessages().forEach((message, m) => {
        message.markRead();
        gotMail = true;
        Logger.log(message.getSubject());
      });
    });
    // skip it if there is no target email
    if (!gotMail) {
      return;
    }

    //  POST IFTTT if you've got email from with the label
    Logger.log(IFTTT);
    const res = UrlFetchApp.fetch(IFTTT + encodeURIComponent(label[1]) + URL, {
      'muteHttpExceptions': true,
      'method': 'get'
    });
    Logger.log("Response:" + res);
  });
}
