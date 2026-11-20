// Maximum number of characters of a subject title that are spoken aloud.
const SUBJECT_MAX_LENGTH = 20;
// Characters that end the meaningful head of a subject.
const SUBJECT_DELIMITERS = ['：', ':', '｜', '|', '、', '。'];

// Read a value as text without ever throwing: strings pass through,
// other scalars stringify exactly as String() did, and a value whose
// toString is not callable (which String() throws on) reads as empty.
function outsideText_(value) {
if (typeof value === 'string') {
return value;
}
try {
return String(value);
} catch (e) {
return '';
}
}

// Spoken text is always passed through encodeURIComponent, which
// throws on a lone surrogate — one half of an astral character, left
// when the length cut splits a pair or already present in the mail.
// Lone surrogates are spoken as U+FFFD; well-formed pairs are
// untouched, so every previously encodable text is unchanged.
function speechSafe_(text) {
return text.replace(
/[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/g,
'�'
);
}

// Extract the title to speak from a message subject. Japanese
// notification mails often embed the actual title in corner brackets,
// e.g. "someone posted 「the title」" — that segment wins. Otherwise
// the head of the subject is used: leading tags such as 【...】 or [...]
// are dropped and the rest is cut at the first delimiter. A trailing
// "..." left by a sender that truncated the title itself is removed.
export function extractSubjectTitle(subject) {
if (!subject) {
return '';
}
const stripped = outsideText_(subject)
.trim()
.replace(/^(\s*【[\s\S]*?】\s*|\s*\[[\s\S]*?\]\s*)+/, '');
const quoted = stripped.match(/「([^」]+)」/);
let title = '';
if (quoted) {
title = quoted[1].replace(/(\.\.\.|…)+$/, '').trim();
}
if (!title) {
let cut = stripped.length;
SUBJECT_DELIMITERS.forEach((delimiter) => {
const index = stripped.indexOf(delimiter);
// A delimiter only counts when at least 3 characters precede it,
// so "本日、お荷物をお届けいたします" is not reduced to "本日".
if (index >= 3 && index < cut) {
cut = index;
}
});
title = stripped.slice(0, cut).trim();
}
return speechSafe_(title.slice(0, SUBJECT_MAX_LENGTH));
}

// Extract "H時MM分に登校しました" / "H時MM分に下校しました" from the plain
// body of a school gate mail, e.g. "... 2026/11/05 08:23:38 に登校しました。"
export function extractSchoolDetail(body) {
if (!body) {
return '';
}
const match = outsideText_(body).match(
/(\d{1,2}):(\d{2}):\d{2}\s*に(登校|下校)しました/
);
if (!match) {
return '';
}
// only times that exist on the clock are announced: the pattern
// also matches shapes like 25:99 that no gate mail can carry
if (Number(match[1]) > 23 || Number(match[2]) > 59) {
return '';
}
return Number(match[1]) + '時' + match[2] + '分に' + match[3] + 'しました';
}

function formatTimesWhen(month, day, hour, minute) {
return (
Number(month) +
'月' +
Number(day) +
'日' +
Number(hour) +
'時' +
(minute === '00' ? '' : minute + '分')
);
}

// Whether a datetime captured by the ■日時 patterns below exists:
// the patterns accept shapes like month 13, February 30 or 25:99
// that no real booking can carry, and a Date built from them
// silently rolls over, so the components must survive a round trip.
function isRealDateTime_(match) {
const year = Number(match[1]);
const month = Number(match[2]);
const day = Number(match[3]);
const hour = Number(match[4]);
const minute = Number(match[5]);
const date = new Date(year, month - 1, day, hour, minute);
return (
date.getFullYear() === year &&
date.getMonth() === month - 1 &&
date.getDate() === day &&
date.getHours() === hour &&
date.getMinutes() === minute
);
}

// Extract the booking summary from a Times CAR mail body: the
// ■予約内容 block carries the station, the car and the start /
// return datetimes on the lines after each heading.
export function extractTimesDetail(body) {
if (!body) {
return '';
}
const text = outsideText_(body);
const start = text.match(
/■利用開始日時\s*\n\s*(\d{4})\/(\d{1,2})\/(\d{1,2})\s+(\d{1,2}):(\d{2})/
);
const end = text.match(
/■返却予定日時\s*\n\s*(\d{4})\/(\d{1,2})\/(\d{1,2})\s+(\d{1,2}):(\d{2})/
);
const station = text.match(/■ステーション\s*\n\s*(\S.*)/);
const car = text.match(/■車両\s*\n\s*([^\s（(]+)/);
if (!start || !end || !station || !car) {
return '';
}
if (!isRealDateTime_(start) || !isRealDateTime_(end)) {
return '';
}
const sameDay = start[2] === end[2] && start[3] === end[3];
const endText = sameDay
? Number(end[4]) + '時' + (end[5] === '00' ? '' : end[5] + '分') + 'まで'
: formatTimesWhen(end[2], end[3], end[4], end[5]) + 'まで';
return speechSafe_(
formatTimesWhen(start[2], start[3], start[4], start[5]) +
'から' +
endText +
'、' +
station[1].trim() +
'の' +
car[1] +
'です'
);
}

export function myFunction() {
// the speech URL prefix comes from the SPEECH_URL_PREFIX script property
const SPEECH_URL_PREFIX = PropertiesService.getScriptProperties().getProperty('SPEECH_URL_PREFIX');
// staging sets the DRY_RUN script property to 1: compose and log the
// speech URL but never fetch it, so nothing is actually announced
const DRY_RUN = PropertiesService.getScriptProperties().getProperty('DRY_RUN') === '1';
const HOME_API_TOKEN = PropertiesService.getScriptProperties().getProperty('HOME_API_TOKEN');
const URL = '%E9%96%A2%E9%80%A3%E3%81%AE%E3%83%A1%E3%83%BC%E3%83%AB%E3%81%8C%E5%B1%8A%E3%81%84%E3%81%A6%E3%81%84%E3%81%BE%E3%81%99%E3%80%82%E7%A2%BA%E8%AA%8D%E3%81%97%E3%81%A6%E4%B8%8B%E3%81%95%E3%81%84%E3%80%82';
// get the data from the spreadsheet and check all the labels.
SpreadsheetApp.getActiveSheet().getDataRange().getValues().forEach((label, v) => {
// skip blank rows: an empty label would degenerate the search below
// into matching every unread message in the mailbox
if (!Array.isArray(label) || label.length < 2 || !label[0] || !String(label[0]).trim()) {
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

// the optional third column lists the modes to try, in order,
// separated by "、", "," or spaces. "件名" extracts a title from the
// subject of the newest message; "登下校" extracts the gate time
// from its body; "タイムズ" extracts the booking summary from its
// body. The first mode that yields a detail wins, e.g.
// "登下校、件名" speaks the gate time when present and the subject
// title otherwise. With no recognised mode, or when no mode yields
// a detail, the announcement stays exactly as before.
const modes =
label.length > 2
? String(label[2])
.split(/[,、\s]+/)
.map((mode) => mode.trim())
.filter((mode) => mode.length > 0)
: [];
let detail = '';
if (detailMessage) {
for (const mode of modes) {
if (mode === '登下校') {
detail = extractSchoolDetail(detailMessage.getPlainBody());
} else if (mode === 'タイムズ') {
detail = extractTimesDetail(detailMessage.getPlainBody());
} else if (mode === '件名') {
detail = extractSubjectTitle(detailMessage.getSubject());
}
if (detail) {
break;
}
}
}

// POST to the speech URL if you've got email from with the label
Logger.log(SPEECH_URL_PREFIX);
const url = detail
? SPEECH_URL_PREFIX + encodeURIComponent(label[1] + '、' + detail + '。')
: SPEECH_URL_PREFIX + encodeURIComponent(label[1]) + URL;
if (DRY_RUN) {
Logger.log('DRY_RUN: ' + url);
} else {
const res = UrlFetchApp.fetch(url, {
'muteHttpExceptions': true,
'method': 'get',
'headers': { 'X-Api-Token': HOME_API_TOKEN }
});
Logger.log("Response:" + res);
}
});
}
