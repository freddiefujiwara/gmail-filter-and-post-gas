import { beforeEach, describe, expect, it, vi } from "vitest";
import fc from "fast-check";
import {
  extractSchoolDetail,
  extractSubjectTitle,
  extractTimesDetail,
  myFunction,
} from "../src/Code.js";

// Property-based tests (fast-check). Generators build strings from
// whole code points only: mail subjects and bodies reach this code as
// decoded text, and a generator must not throw or build ill-formed
// strings before the code under test runs. Lone surrogates that are
// already present in an input are covered deterministically in
// test/abnormal.spec.js instead.

const PREFIX = "https://example.invalid/speech/-s/";
const TOKEN = "dummy-home-api-token";

const hasLoneSurrogate = (text) =>
  /[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/.test(
    text
  );

// A junk value from the JSON-ish domain, including the two shapes
// String() throws on.
const junkValue = fc.oneof(
  fc.string(),
  fc.integer(),
  fc.double({ noNaN: true }),
  fc.boolean(),
  fc.constant(null),
  fc.constant(undefined),
  fc.array(fc.string(), { maxLength: 3 }),
  fc.record({ a: fc.string() }),
  fc.constant({ toString: [] }),
  fc.constant(Object.create(null))
);

// Subject building blocks: tags, quotes, delimiters, ellipses, astral
// characters and plain CJK/ASCII — the shapes real notification mail
// subjects are made of, as whole code points.
const subjectPieces = fc.array(
  fc.constantFrom(
    "お知らせ",
    "テスト",
    "🎉",
    "【タグ】",
    "[tag]",
    "「",
    "」",
    "：",
    "、",
    "…",
    "...",
    " ",
    "abc",
    "123"
  ),
  { maxLength: 12 }
);

// Plain subject characters: no tags, quotes, delimiters or edge
// whitespace, so a well-formed subject must survive extraction whole.
const plainSubject = fc
  .array(
    fc.constantFrom("あ", "い", "う", "漢", "字", "a", "b", "1", "🎉", "ー"),
    { minLength: 1, maxLength: 10 }
  )
  .map((chars) => chars.join(""));

const pad2 = (n) => String(n).padStart(2, "0");

const validTime = fc.record({
  hour: fc.integer({ min: 0, max: 23 }),
  minute: fc.integer({ min: 0, max: 59 }),
  second: fc.integer({ min: 0, max: 59 }),
  verb: fc.constantFrom("登校", "下校"),
});

const invalidTime = fc.oneof(
  fc.record({
    hour: fc.integer({ min: 24, max: 99 }),
    minute: fc.integer({ min: 0, max: 59 }),
  }),
  fc.record({
    hour: fc.integer({ min: 0, max: 23 }),
    minute: fc.integer({ min: 60, max: 99 }),
  })
);

describe("extractor totality over junk input", () => {
  it("no extractor throws and every result is a string", () => {
    fc.assert(
      fc.property(junkValue, (value) => {
        for (const extract of [
          extractSubjectTitle,
          extractSchoolDetail,
          extractTimesDetail,
        ]) {
          expect(typeof extract(value)).toBe("string");
        }
      }),
      { numRuns: 100 }
    );
  });
});

describe("extractSubjectTitle properties", () => {
  it("output is always encodable and within the spoken length", () => {
    fc.assert(
      fc.property(subjectPieces, (pieces) => {
        const title = extractSubjectTitle(pieces.join(""));
        expect(hasLoneSurrogate(title)).toBe(false);
        expect(() => encodeURIComponent(title)).not.toThrow();
        // 20 code units, plus at most one replacement character where
        // the cut met an astral pair
        expect(title.length).toBeLessThanOrEqual(21);
      }),
      { numRuns: 100 }
    );
  });

  it("a plain well-formed subject is returned whole", () => {
    fc.assert(
      fc.property(plainSubject, (subject) => {
        expect(extractSubjectTitle(subject)).toBe(subject);
      }),
      { numRuns: 100 }
    );
  });

  it("the corner-bracket segment wins verbatim when it fits", () => {
    fc.assert(
      fc.property(
        plainSubject,
        fc.constantFrom("先生", "保護者", "事務局"),
        (content, who) => {
          const subject = who + "が「" + content + "」を投稿しました";
          expect(extractSubjectTitle(subject)).toBe(content);
        }
      ),
      { numRuns: 100 }
    );
  });
});

describe("extractSchoolDetail properties", () => {
  it("a valid gate time anywhere in the body is extracted exactly", () => {
    // The filler alphabet excludes digits, colons and the verb
    // characters, so the only gate-time pattern in the body is the
    // inserted one and the first match is deterministic.
    const filler = fc
      .array(fc.constantFrom("あ", "の", "で", "す", "x", " ", "。"), {
        maxLength: 30,
      })
      .map((chars) => chars.join(""));
    fc.assert(
      fc.property(validTime, filler, ({ hour, minute, second, verb }, padding) => {
        const body =
          padding +
          " 2026/11/05 " +
          hour +
          ":" +
          pad2(minute) +
          ":" +
          pad2(second) +
          " に" +
          verb +
          "しました。" +
          padding;
        expect(extractSchoolDetail(body)).toBe(
          hour + "時" + pad2(minute) + "分に" + verb + "しました"
        );
      }),
      { numRuns: 100 }
    );
  });

  it("an out-of-range gate time yields no detail", () => {
    fc.assert(
      fc.property(invalidTime, ({ hour, minute }) => {
        const body =
          "2026/11/05 " + hour + ":" + pad2(minute) + ":00 に登校しました。";
        expect(extractSchoolDetail(body)).toBe("");
      }),
      { numRuns: 100 }
    );
  });

  it("output is empty or a well-formed gate announcement", () => {
    fc.assert(
      fc.property(fc.string(), (body) => {
        const detail = extractSchoolDetail(body);
        expect(
          detail === "" || /^\d{1,2}時\d{2}分に(登校|下校)しました$/.test(detail)
        ).toBe(true);
      }),
      { numRuns: 100 }
    );
  });
});

const timesBody = (start, end, station, car) =>
  [
    "■予約内容",
    "■ステーション",
    station,
    "■車両",
    car + " （ダミー 000 あ 00-00）",
    "■利用開始日時",
    start,
    "■返却予定日時",
    end,
  ].join("\n");

const formatWhen = (month, day, hour, minute) =>
  month +
  "月" +
  day +
  "日" +
  hour +
  "時" +
  (minute === 0 ? "" : pad2(minute) + "分");

const validBooking = fc.record({
  start: fc.date({
    min: new Date(2026, 0, 1),
    max: new Date(2028, 11, 31),
    noInvalidDate: true,
  }),
  startHour: fc.integer({ min: 0, max: 23 }),
  startMinute: fc.constantFrom(0, 5, 30, 45),
  // A booking spans days, never a year: the production same-day
  // check compares month and day only (its long-standing shape), so
  // a same month/day pair in a different year — which only a booking
  // of a year or more could produce — is outside the domain of this
  // mail and is not generated. The return date is the start date
  // plus a few days instead.
  endOffsetDays: fc.integer({ min: 0, max: 3 }),
  endHour: fc.integer({ min: 0, max: 23 }),
  endMinute: fc.constantFrom(0, 5, 30, 45),
  station: fc.constantFrom("ダミー駐車場", "テスト駅前", "サンプル第２"),
  car: fc.constantFrom("ダミーカー", "テスト号", "サンプル"),
});

const stampOf = (date, hour, minute) =>
  date.getFullYear() +
  "/" +
  (date.getMonth() + 1) +
  "/" +
  date.getDate() +
  " " +
  hour +
  ":" +
  pad2(minute);

describe("extractTimesDetail properties", () => {
  it("a well-formed booking is summarised exactly", () => {
    fc.assert(
      fc.property(validBooking, (b) => {
        const end = new Date(
          b.start.getFullYear(),
          b.start.getMonth(),
          b.start.getDate() + b.endOffsetDays
        );
        const body = timesBody(
          stampOf(b.start, b.startHour, b.startMinute),
          stampOf(end, b.endHour, b.endMinute),
          b.station,
          b.car
        );
        const sameDay = b.endOffsetDays === 0;
        const endText = sameDay
          ? b.endHour +
            "時" +
            (b.endMinute === 0 ? "" : pad2(b.endMinute) + "分") +
            "まで"
          : formatWhen(
              end.getMonth() + 1,
              end.getDate(),
              b.endHour,
              b.endMinute
            ) + "まで";
        const expected =
          formatWhen(
            b.start.getMonth() + 1,
            b.start.getDate(),
            b.startHour,
            b.startMinute
          ) +
          "から" +
          endText +
          "、" +
          b.station +
          "の" +
          b.car +
          "です";
        expect(extractTimesDetail(body)).toBe(expected);
      }),
      { numRuns: 100 }
    );
  });

  it("a booking with a deformed datetime yields no detail", () => {
    const deformed = fc.oneof(
      fc.constant("2026/02/30 11:00"),
      fc.constant("2026/04/31 11:00"),
      fc.constant("2026/13/10 11:00"),
      fc.constant("2026/10/32 11:00"),
      fc.constant("2026/10/25 24:00"),
      fc.constant("2026/10/25 11:60"),
      fc
        .record({
          month: fc.integer({ min: 13, max: 99 }),
          day: fc.integer({ min: 1, max: 28 }),
        })
        .map(({ month, day }) => "2026/" + month + "/" + day + " 11:00")
    );
    fc.assert(
      fc.property(deformed, (start) => {
        expect(
          extractTimesDetail(timesBody(start, "2026/10/25 23:00", "甲", "乙"))
        ).toBe("");
      }),
      { numRuns: 100 }
    );
  });

  it("output is empty or a spoken summary ending in です", () => {
    fc.assert(
      fc.property(fc.string(), (body) => {
        const detail = extractTimesDetail(body);
        expect(detail === "" || detail.endsWith("です")).toBe(true);
        expect(hasLoneSurrogate(detail)).toBe(false);
      }),
      { numRuns: 100 }
    );
  });
});

// ---- myFunction batch property ----

let sheetValues;
let threadsByQuery;

const mockSheet = {
  getDataRange: vi.fn(() => ({ getValues: vi.fn(() => sheetValues) })),
};

global.SpreadsheetApp = {
  getActiveSheet: vi.fn(() => mockSheet),
};

global.GmailApp = {
  search: vi.fn((query) => threadsByQuery[query] ?? []),
};

global.UrlFetchApp = {
  fetch: vi.fn(() => "OK"),
};

global.PropertiesService = {
  getScriptProperties: vi.fn(() => ({
    getProperty: vi.fn((name) =>
      name === "SPEECH_URL_PREFIX"
        ? PREFIX
        : name === "HOME_API_TOKEN"
          ? TOKEN
          : null
    ),
  })),
};

global.Logger = {
  log: vi.fn(),
};

beforeEach(() => {
  sheetValues = [];
  threadsByQuery = {};
  vi.clearAllMocks();
});

describe("myFunction batch property", () => {
  it("every label with mail is announced once and every URL is encodable", () => {
    fc.assert(
      fc.property(
        fc.array(subjectPieces, { minLength: 1, maxLength: 6 }),
        (subjectSets) => {
          sheetValues = subjectSets.map((_, i) => [
            "label" + i,
            "文面" + i,
            "件名",
          ]);
          const messages = subjectSets.map((pieces) => ({
            markRead: vi.fn(),
            getSubject: vi.fn(() => pieces.join("")),
            getPlainBody: vi.fn(() => ""),
          }));
          threadsByQuery = {};
          subjectSets.forEach((_, i) => {
            threadsByQuery["label:label" + i + " label:unread"] = [
              { getMessages: vi.fn(() => [messages[i]]) },
            ];
          });
          vi.clearAllMocks();
          expect(() => myFunction()).not.toThrow();
          for (const message of messages) {
            expect(message.markRead).toHaveBeenCalledTimes(1);
          }
          expect(UrlFetchApp.fetch).toHaveBeenCalledTimes(subjectSets.length);
          for (const call of UrlFetchApp.fetch.mock.calls) {
            const decoded = decodeURIComponent(call[0]);
            expect(hasLoneSurrogate(decoded)).toBe(false);
          }
        }
      ),
      { numRuns: 100 }
    );
  });
});
