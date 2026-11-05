// Maximum number of characters of a subject that are spoken aloud.
const SUBJECT_MAX_LENGTH = 20;
// Characters that end the meaningful head of a subject.
const SUBJECT_DELIMITERS = ['：', ':', '｜', '|', '、', '。'];

// Reduce a subject to its meaningful head: drop leading bracketed tags such
// as or [...], cut at the first delimiter, and cap the length.
export function truncateSubject(subject) {
if (!subject) {
return '';
}
const stripped = String(subject)
.trim()
.replace(/^(\s*【[\s\S]*?】\s*|\s*\[[\s\S]*?\]\s*)+/, '');
let cut = stripped.length;
SUBJECT_DELIMITERS.forEach((delimiter) => {
const index = stripped.indexOf(delimiter);
if (index > 0 && index < cut) {
cut = index;
}
});
return stripped.slice(0, cut).trim().slice(0, SUBJECT_MAX_LENGTH);
}

// Extract "H時MM分に登校しました" / "H時MM分に下校しました" from the plain
// body of a HotConpass mail, e.g. "... 2026/11/05 08:23:38 に登校しました。"
export function extractSchoolDetail(body) {
if (!body) {
return '';
}
const match = String(body).match(
/(\d{1,2}):(\d{2}):\d{2}\s*に(登校|下校)しました/
);
if (!match) {
return '';
}
return Number(match[1]) + '時' + match[2] + '分に' + match[3] + 'しました';
}

export function myFunction() {
// the speech URL prefix comes from the SPEECH_URL_PREFIX script property
const IFTTT = PropertiesService.getScriptProperties().getProperty('SPEECH_URL_PREFIX');
// staging sets the DRY_RUN script property to 1: compose and log the
// speech URL but never fetch it, so nothing is actually announced
const DRY_RUN = PropertiesService.getScriptProperties().getProperty('DRY_RUN') === '1';
const URL = '%E9%96%A2%E9%80%A3%E3%81%AE%E3%83%A1%E3%83%BC%E3%83%AB%E3%81%8C%E5%B1%8A%E3%81%84%E3%81%A6%E3%81%84%E3%81%BE%E3%81%99%E3%80%82%E7%A2%BA%E8%AA%8D%E3%81%97%E3%81%A6%E4%B8%8B%E3%81%95%E3%81%84%E3%80%82';
// get the data from the spreadsheet and check all the labels.
SpreadsheetApp.getActiveSheet().getDataRange().getValues().forEach((label, v) => {
// skip blank labels
if (label.length < 2) {
return;
}

// check all relevant unread messages
let gotMail = false;
// the newest message (last message of the first thread) is the one the
// optional detail in the third column is extracted from
let detailMessage = null;
GmailApp.search('label:' + label[0] + ' label:unread').forEach((thread, t) => {
const messages = thread.getMessages();
if (t === 0 && messages.length > 0) {
detailMessage = messages[messages.length - 1];
}
messages.forEach((message, m) => {
message.markRead();
gotMail = true;
Logger.log(message.getSubject());
});
});
// skip it if there is no target email
if (!gotMail) {
return;
}

// the optional third column selects extra detail to speak:
// "件名" speaks the (truncated) subject, "登下校" speaks the time a
// HotConpass mail reports. Without a recognised mode, or when no detail
// can be extracted, the announcement stays exactly as before.
const mode = label.length > 2? String(label[2]).trim(): '';
let detail = '';
if (mode === '登下校' && detailMessage) {
detail = extractSchoolDetail(detailMessage.getPlainBody());
} else if (mode === '件名' && detailMessage) {
detail = truncateSubject(detailMessage.getSubject());
}

// POST IFTTT if you've got email from with the label
Logger.log(IFTTT);
const url = detail
? IFTTT + encodeURIComponent(label[1] + '、' + detail + '。')
: IFTTT + encodeURIComponent(label[1]) + URL;
if (DRY_RUN) {
Logger.log('DRY_RUN: ' + url);
} else {
const res = UrlFetchApp.fetch(url, {
'muteHttpExceptions': true,
'method': 'get'
});
Logger.log("Response:" + res);
}
});
}
