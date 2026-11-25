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

// ---- presence derivation from school gate mails ----
// A school gate mail carries the exact time the gate reader was
// tapped. Two kinds of presence events come out of it. The gate tap
// itself is recorded as-is at school: arriving at the 登校 time and
// leaving at the 下校 time. The matching events for home are derived
// with a fixed offset: leaving home is 20 minutes before the 登校
// time and arriving home is 20 minutes after the 下校 time. A derived
// home event is only ever written once its time has come: a 下校
// mail is processed about 20 minutes before the arrival it implies,
// so the event is held in the PRESENCE_GATE_STATE script property
// and flushed by a later run. Everything is a no-op unless the
// PRESENCE_GAS_URL, PRESENCE_API_TOKEN and PRESENCE_GATE_PERSON
// script properties are all set, and DRY_RUN only logs the payload
// instead of posting it.
const GATE_HOME_OFFSET_MS = 20 * 60 * 1000;
const PRESENCE_GATE_STATE_KEY = 'PRESENCE_GATE_STATE';

// Extract the presence events implied by a school gate mail body,
// e.g. "... 2026/11/05 08:23:38 に登校しました。" Returns
// {action: 'leave'|'arrive', at: <epoch ms>, date: 'YYYY-MM-DD',
// school: {action: 'arrive'|'leave', at: <epoch ms>}} where action/at
// describe the derived home event, school describes the gate tap
// itself, and date is the gate day in JST. Returns null when the body
// carries no usable gate time.
export function schoolGatePresenceEvent(body) {
if (!body) {
return null;
}
const match = outsideText_(body).match(
/(\d{4})\/(\d{1,2})\/(\d{1,2})\s+(\d{1,2}):(\d{2}):(\d{2})\s*に(登校|下校)しました/
);
if (!match) {
return null;
}
const year = Number(match[1]);
const month = Number(match[2]);
const day = Number(match[3]);
const hour = Number(match[4]);
const minute = Number(match[5]);
const second = Number(match[6]);
// the pattern also matches shapes no real gate mail can carry
// (month 13, February 30, 25:99), so validate the components
if (month < 1 || month > 12 || hour > 23 || minute > 59 || second > 59) {
return null;
}
const daysInMonth = new Date(Date.UTC(year, month, 0)).getUTCDate();
if (day < 1 || day > daysInMonth) {
return null;
}
// guard against off-pattern taps: a 登校 mail is a morning event and
// a 下校 mail an afternoon one (an evening 登校 mail is off-pattern
// and must not turn into an evening "left home")
const isGateIn = match[7] === '登校';
if (isGateIn && (hour < 4 || hour >= 12)) {
return null;
}
if (!isGateIn && (hour < 12 || hour >= 23)) {
return null;
}
// the gate clock is JST; convert the wall time to an instant
const gateMs = Date.UTC(year, month - 1, day, hour - 9, minute, second);
const date =
match[1] +
'-' +
String(month).padStart(2, '0') +
'-' +
String(day).padStart(2, '0');
return isGateIn
? {
action: 'leave',
at: gateMs - GATE_HOME_OFFSET_MS,
date: date,
school: { action: 'arrive', at: gateMs }
}
: {
action: 'arrive',
at: gateMs + GATE_HOME_OFFSET_MS,
date: date,
school: { action: 'leave', at: gateMs }
};
}

function presenceGateConfig_() {
const props = PropertiesService.getScriptProperties();
const url = props.getProperty('PRESENCE_GAS_URL');
const token = props.getProperty('PRESENCE_API_TOKEN');
const person = props.getProperty('PRESENCE_GATE_PERSON');
if (!url || !token || !person) {
return null;
}
return {
url: url,
token: token,
person: person,
dryRun: props.getProperty('DRY_RUN') === '1'
};
}

function loadPresenceGateState_() {
const raw = PropertiesService.getScriptProperties().getProperty(
PRESENCE_GATE_STATE_KEY
);
try {
const parsed = raw ? JSON.parse(raw) : {};
return {
leaveDate: parsed.leaveDate || null,
arriveDate: parsed.arriveDate || null,
schoolArriveDate: parsed.schoolArriveDate || null,
schoolLeaveDate: parsed.schoolLeaveDate || null,
pending: Array.isArray(parsed.pending) ? parsed.pending : []
};
} catch (e) {
return {
leaveDate: null,
arriveDate: null,
schoolArriveDate: null,
schoolLeaveDate: null,
pending: []
};
}
}

function savePresenceGateState_(state) {
PropertiesService.getScriptProperties().setProperty(
PRESENCE_GATE_STATE_KEY,
JSON.stringify(state)
);
}

function postPresenceEvent_(config, place, action, atMs) {
const payload = {
token: config.token,
person: config.person,
place: place,
action: action,
timestamp: new Date(atMs).toISOString()
};
if (config.dryRun) {
Logger.log('DRY_RUN presence: ' + JSON.stringify(payload));
return true;
}
try {
const res = UrlFetchApp.fetch(config.url, {
method: 'post',
contentType: 'application/json',
muteHttpExceptions: true,
payload: JSON.stringify(payload)
});
if (res.getResponseCode() !== 200) {
return false;
}
return JSON.parse(res.getContentText()).ok === true;
} catch (e) {
return false;
}
}

// Inspect one message seen by the label loop; when it is a gate mail,
// record (or hold) the presence events it implies. Runs inside the
// loop so gate mails are seen exactly once, while still unread.
function collectGatePresence_(message) {
const subject = message.getSubject();
if (subject !== '登校メール' && subject !== '下校メール') {
return;
}
const config = presenceGateConfig_();
if (!config) {
return;
}
const event = schoolGatePresenceEvent(message.getPlainBody());
if (!event) {
return;
}
const state = loadPresenceGateState_();
// the derived home event first, so it lands in the sheet ahead of
// the school event it implies
recordGateEvent_(config, state, 'home', event.action, event.at, event.date);
recordGateEvent_(
config,
state,
'school',
event.school.action,
event.school.at,
event.date
);
savePresenceGateState_(state);
}

// The state field that records, per place and action, the gate day
// already written, so a gate mail never produces the same event
// twice.
function gateDateKey_(place, action) {
if (place === 'school') {
return action === 'arrive' ? 'schoolArriveDate' : 'schoolLeaveDate';
}
return action === 'leave' ? 'leaveDate' : 'arriveDate';
}

// Record one gate-derived presence event: post it now when its time
// has come, otherwise hold it in the pending list and let a later
// run write it once its time has come (a failed post is held the
// same way and retried by the flush).
function recordGateEvent_(config, state, place, action, at, date) {
const key = gateDateKey_(place, action);
if (state[key] === date) {
return;
}
if (
state.pending.some(
(p) =>
(p.place || 'home') === place && p.action === action && p.date === date
)
) {
return;
}
if (at <= Date.now() && postPresenceEvent_(config, place, action, at)) {
state[key] = date;
} else {
state.pending.push({ place: place, action: action, at: at, date: date });
}
}

// Write held presence events whose time has come. Runs at the top of
// every myFunction run, whether or not any new mail arrived.
function flushPendingPresence_() {
const config = presenceGateConfig_();
if (!config) {
return;
}
const state = loadPresenceGateState_();
if (state.pending.length === 0) {
return;
}
const now = Date.now();
const remaining = [];
let changed = false;
state.pending.forEach((p) => {
// pending entries written before places were recorded carry no
// place and are home events
const place = p.place || 'home';
if (p.at <= now && postPresenceEvent_(config, place, p.action, p.at)) {
state[gateDateKey_(place, p.action)] = p.date;
changed = true;
} else {
remaining.push(p);
}
});
if (changed || remaining.length !== state.pending.length) {
state.pending = remaining;
savePresenceGateState_(state);
}
}

export function myFunction() {
flushPendingPresence_();
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
collectGatePresence_(message);
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
