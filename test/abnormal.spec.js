import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  extractSchoolDetail,
  extractSubjectTitle,
  extractTimesDetail,
  myFunction,
} from "../src/Code.js";

// Abnormal-case tests: mail subjects, mail bodies and sheet rows are
// input from outside and can be deformed in ways well-formed fixtures
// never show. All fixtures here are synthetic (dummy names, dummy
// places, example.invalid URLs) — no real mail content.
//
// The announcement text built from a subject or a body is always
// passed through encodeURIComponent before it is fetched, so any
// extractor output must be encodable: a lone surrogate in it makes
// encodeURIComponent throw URIError and kills the whole run.

const PREFIX = "https://example.invalid/speech/-s/";
const TOKEN = "dummy-home-api-token";
const SUFFIX =
  "%E9%96%A2%E9%80%A3%E3%81%AE%E3%83%A1%E3%83%BC%E3%83%AB%E3%81%8C%E5%B1%8A%E3%81%84%E3%81%A6%E3%81%84%E3%81%BE%E3%81%99%E3%80%82%E7%A2%BA%E8%AA%8D%E3%81%97%E3%81%A6%E4%B8%8B%E3%81%95%E3%81%84%E3%80%82";

// A value whose toString is not callable: String() throws on it.
// Such shapes are producible from JSON-domain data and have broken
// bare String() reads across this fleet.
const UNSTRINGIFIABLE = { toString: [] };
const NULL_PROTO = Object.create(null);

const hasLoneSurrogate = (text) =>
  /[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/.test(
    text
  );

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

const makeMessage = (subject, body) => ({
  markRead: vi.fn(),
  getSubject: vi.fn(() => subject),
  getPlainBody: vi.fn(() => body),
});

const makeThread = (messages) => ({
  getMessages: vi.fn(() => messages),
});

beforeEach(() => {
  sheetValues = [];
  threadsByQuery = {};
  vi.clearAllMocks();
});

describe("extractSubjectTitle with deformed subjects", () => {
  it("reads an unstringifiable subject as empty instead of throwing", () => {
    expect(extractSubjectTitle(UNSTRINGIFIABLE)).toBe("");
    expect(extractSubjectTitle(NULL_PROTO)).toBe("");
  });

  it("does not split a surrogate pair at the 20 character cut", () => {
    const title = extractSubjectTitle("あ".repeat(19) + "🎉" + "あ".repeat(5));
    expect(hasLoneSurrogate(title)).toBe(false);
    expect(() => encodeURIComponent(title)).not.toThrow();
    // the pair does not fit whole, so it is not spoken half-way:
    // either it is dropped or replaced, never left as a lone half
    expect(title.startsWith("あ".repeat(19))).toBe(true);
    expect(title.length).toBeLessThanOrEqual(21);
  });

  it("does not split a surrogate pair inside the corner-bracket segment", () => {
    const title = extractSubjectTitle(
      "先生が「" + "あ".repeat(19) + "🎉のお知らせ」を投稿しました"
    );
    expect(hasLoneSurrogate(title)).toBe(false);
    expect(() => encodeURIComponent(title)).not.toThrow();
  });

  it("keeps an astral character that fits whole inside the limit", () => {
    // 18 BMP characters + one astral pair = exactly 20 code units
    expect(extractSubjectTitle("あ".repeat(18) + "🎉")).toBe(
      "あ".repeat(18) + "🎉"
    );
  });

  it("neutralises a lone surrogate already present in the subject", () => {
    const title = extractSubjectTitle("お知らせ\uD83Cです、補足です");
    expect(hasLoneSurrogate(title)).toBe(false);
    expect(() => encodeURIComponent(title)).not.toThrow();
  });

  it("stringifies scalar subjects as before", () => {
    expect(extractSubjectTitle(42)).toBe("42");
    expect(extractSubjectTitle(true)).toBe("true");
  });
});

describe("extractSchoolDetail with deformed bodies", () => {
  it("reads an unstringifiable body as empty instead of throwing", () => {
    expect(extractSchoolDetail(UNSTRINGIFIABLE)).toBe("");
    expect(extractSchoolDetail(NULL_PROTO)).toBe("");
  });

  it("rejects times that do not exist on the clock", () => {
    expect(
      extractSchoolDetail("ダミー学園 2026/11/05 25:99:00 に登校しました。")
    ).toBe("");
    expect(
      extractSchoolDetail("ダミー学園 2026/11/05 24:00:00 に下校しました。")
    ).toBe("");
    expect(
      extractSchoolDetail("ダミー学園 2026/11/05 08:60:00 に登校しました。")
    ).toBe("");
  });

  it("still extracts the last valid time of day", () => {
    expect(
      extractSchoolDetail("ダミー学園 2026/11/05 23:59:59 に下校しました。")
    ).toBe("23時59分に下校しました");
  });

  it("finds the gate time at the end of a very long body", () => {
    const body =
      "ダミーの本文です。".repeat(5000) +
      "\n2026/11/05 07:05:09 に登校しました。";
    expect(extractSchoolDetail(body)).toBe("7時05分に登校しました");
  });

  it("reads a non-string body without throwing", () => {
    expect(extractSchoolDetail(123)).toBe("");
    expect(extractSchoolDetail(undefined)).toBe("");
  });
});

const timesBody = (start, end, station = "ダミー駐車場", car = "ダミーカー") =>
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

describe("extractTimesDetail with deformed bodies", () => {
  it("reads an unstringifiable body as empty instead of throwing", () => {
    expect(extractTimesDetail(UNSTRINGIFIABLE)).toBe("");
    expect(extractTimesDetail(NULL_PROTO)).toBe("");
  });

  it("rejects booking dates that do not exist on the calendar", () => {
    expect(
      extractTimesDetail(timesBody("2026/02/30 11:00", "2026/02/30 23:00"))
    ).toBe("");
    expect(
      extractTimesDetail(timesBody("2026/04/31 11:00", "2026/04/31 23:00"))
    ).toBe("");
    expect(
      extractTimesDetail(timesBody("2026/13/10 11:00", "2026/13/10 23:00"))
    ).toBe("");
    expect(
      extractTimesDetail(timesBody("2026/00/10 11:00", "2026/00/10 23:00"))
    ).toBe("");
    expect(
      extractTimesDetail(timesBody("2026/10/00 11:00", "2026/10/00 23:00"))
    ).toBe("");
    // 2026 is not a leap year
    expect(
      extractTimesDetail(timesBody("2026/02/29 11:00", "2026/02/29 23:00"))
    ).toBe("");
  });

  it("rejects a booking whose return datetime does not exist", () => {
    expect(
      extractTimesDetail(timesBody("2026/10/25 11:00", "2026/02/30 23:00"))
    ).toBe("");
  });

  it("rejects booking times that do not exist on the clock", () => {
    expect(
      extractTimesDetail(timesBody("2026/10/25 25:00", "2026/10/25 23:00"))
    ).toBe("");
    expect(
      extractTimesDetail(timesBody("2026/10/25 11:99", "2026/10/25 23:00"))
    ).toBe("");
  });

  it("still summarises a leap-day booking", () => {
    expect(
      extractTimesDetail(timesBody("2028/02/29 11:00", "2028/02/29 23:00"))
    ).toBe("2月29日11時から23時まで、ダミー駐車場のダミーカーです");
  });

  it("keeps the spoken summary encodable when the station line is deformed", () => {
    const detail = extractTimesDetail(
      timesBody("2026/10/25 11:00", "2026/10/25 23:00", "ダミー\uD83C駐車場")
    );
    expect(hasLoneSurrogate(detail)).toBe(false);
    expect(() => encodeURIComponent(detail)).not.toThrow();
  });
});

describe("myFunction with deformed sheet rows", () => {
  it("skips non-array rows and still processes the valid rows", () => {
    sheetValues = [null, "not-a-row", 42, ["dummy", "ダミー文"]];
    const message = makeMessage("件名ダミー", "");
    threadsByQuery["label:dummy label:unread"] = [makeThread([message])];
    expect(() => myFunction()).not.toThrow();
    expect(message.markRead).toHaveBeenCalledTimes(1);
    expect(UrlFetchApp.fetch).toHaveBeenCalledTimes(1);
  });

  it("passes a numeric label cell into the search query as before", () => {
    sheetValues = [[123, "ダミー文"]];
    myFunction();
    expect(GmailApp.search).toHaveBeenCalledWith("label:123 label:unread");
  });

  it("ignores a numeric modes cell and keeps the legacy announcement", () => {
    sheetValues = [["dummy", "ダミー文", 123]];
    const message = makeMessage("下校メール", "");
    threadsByQuery["label:dummy label:unread"] = [makeThread([message])];
    myFunction();
    expect(UrlFetchApp.fetch).toHaveBeenCalledWith(
      PREFIX + encodeURIComponent("ダミー文") + SUFFIX,
      {
        muteHttpExceptions: true,
        method: "get",
        headers: { "X-Api-Token": TOKEN },
      }
    );
  });
});

describe("myFunction with deformed messages", () => {
  it("one emoji-boundary subject does not kill the run for later labels", () => {
    sheetValues = [
      ["first", "第一", "件名"],
      ["second", "第二"],
    ];
    const emojiMessage = makeMessage(
      "あ".repeat(19) + "🎉" + "あ".repeat(5),
      ""
    );
    threadsByQuery["label:first label:unread"] = [makeThread([emojiMessage])];
    const normalMessage = makeMessage("普通の件名", "");
    threadsByQuery["label:second label:unread"] = [
      makeThread([normalMessage]),
    ];
    expect(() => myFunction()).not.toThrow();
    expect(emojiMessage.markRead).toHaveBeenCalledTimes(1);
    expect(normalMessage.markRead).toHaveBeenCalledTimes(1);
    expect(UrlFetchApp.fetch).toHaveBeenCalledTimes(2);
    const firstUrl = UrlFetchApp.fetch.mock.calls[0][0];
    expect(hasLoneSurrogate(decodeURIComponent(firstUrl))).toBe(false);
  });

  it("falls back to the legacy announcement for a null subject", () => {
    sheetValues = [["dummy", "ダミー文", "件名"]];
    const message = makeMessage(null, null);
    threadsByQuery["label:dummy label:unread"] = [makeThread([message])];
    myFunction();
    expect(UrlFetchApp.fetch).toHaveBeenCalledWith(
      PREFIX + encodeURIComponent("ダミー文") + SUFFIX,
      {
        muteHttpExceptions: true,
        method: "get",
        headers: { "X-Api-Token": TOKEN },
      }
    );
  });

  it("falls back to the legacy announcement for a null plain body", () => {
    sheetValues = [["dummy", "ダミー文", "登下校"]];
    const message = makeMessage("下校メール", null);
    threadsByQuery["label:dummy label:unread"] = [makeThread([message])];
    myFunction();
    expect(UrlFetchApp.fetch).toHaveBeenCalledWith(
      PREFIX + encodeURIComponent("ダミー文") + SUFFIX,
      {
        muteHttpExceptions: true,
        method: "get",
        headers: { "X-Api-Token": TOKEN },
      }
    );
  });

  it("announces once per label even when a thread is returned twice", () => {
    sheetValues = [["dummy", "ダミー文"]];
    const message = makeMessage("件名ダミー", "");
    const thread = makeThread([message]);
    threadsByQuery["label:dummy label:unread"] = [thread, thread];
    myFunction();
    expect(UrlFetchApp.fetch).toHaveBeenCalledTimes(1);
  });

  it("does not announce for a label whose threads are all empty", () => {
    sheetValues = [["dummy", "ダミー文", "件名"]];
    threadsByQuery["label:dummy label:unread"] = [makeThread([])];
    myFunction();
    expect(UrlFetchApp.fetch).not.toHaveBeenCalled();
  });
});
