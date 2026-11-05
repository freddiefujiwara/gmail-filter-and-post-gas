import { beforeEach, describe, expect, it, vi } from "vitest";
import {
  extractSchoolDetail,
  extractSubjectTitle,
  extractTimesDetail,
  myFunction,
} from "../src/Code.js";

// ---- Mocks for Google Apps Script globals ----
const PREFIX = "https://example.invalid/speech/-s/";
const SUFFIX =
  "%E9%96%A2%E9%80%A3%E3%81%AE%E3%83%A1%E3%83%BC%E3%83%AB%E3%81%8C%E5%B1%8A%E3%81%84%E3%81%A6%E3%81%84%E3%81%BE%E3%81%99%E3%80%82%E7%A2%BA%E8%AA%8D%E3%81%97%E3%81%A6%E4%B8%8B%E3%81%95%E3%81%84%E3%80%82";

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
    getProperty: vi.fn((name) => (name === "SPEECH_URL_PREFIX" ? PREFIX : null)),
  })),
};

global.Logger = {
  log: vi.fn(),
};

const makeMessage = (subject) => ({
  markRead: vi.fn(),
  getSubject: vi.fn(() => subject),
});

const makeThread = (messages) => ({
  getMessages: vi.fn(() => messages),
});

beforeEach(() => {
  sheetValues = [];
  threadsByQuery = {};
  vi.clearAllMocks();
});

describe("myFunction", () => {
  it("skips rows with fewer than 2 columns without searching Gmail", () => {
    sheetValues = [["lonely"], []];
    myFunction();
    expect(GmailApp.search).not.toHaveBeenCalled();
    expect(UrlFetchApp.fetch).not.toHaveBeenCalled();
  });

  it("does not fetch when a label has no unread threads", () => {
    sheetValues = [["kuronekoyamato", "クロネコヤマト"]];
    myFunction();
    expect(GmailApp.search).toHaveBeenCalledTimes(1);
    expect(GmailApp.search).toHaveBeenCalledWith(
      "label:kuronekoyamato label:unread"
    );
    expect(UrlFetchApp.fetch).not.toHaveBeenCalled();
  });

  it("marks every message read, logs subjects, and fetches once per label with mail", () => {
    const text = "武南学園";
    // the percent-encoded form the old =ENCODEURL("武南学園") cells produced
    const encoded = "%E6%AD%A6%E5%8D%97%E5%AD%A6%E5%9C%92";
    sheetValues = [["bunan", text]];
    const m1 = makeMessage("subject one");
    const m2 = makeMessage("subject two");
    const m3 = makeMessage("subject three");
    threadsByQuery["label:bunan label:unread"] = [
      makeThread([m1, m2]),
      makeThread([m3]),
    ];
    myFunction();
    expect(m1.markRead).toHaveBeenCalledTimes(1);
    expect(m2.markRead).toHaveBeenCalledTimes(1);
    expect(m3.markRead).toHaveBeenCalledTimes(1);
    expect(Logger.log).toHaveBeenCalledWith("subject one");
    expect(Logger.log).toHaveBeenCalledWith("subject two");
    expect(Logger.log).toHaveBeenCalledWith("subject three");
    expect(UrlFetchApp.fetch).toHaveBeenCalledTimes(1);
    expect(UrlFetchApp.fetch).toHaveBeenCalledWith(PREFIX + encoded + SUFFIX, {
      muteHttpExceptions: true,
      method: "get",
    });
  });

  it("builds the search query exactly, passing a trailing-space label through as is", () => {
    sheetValues = [["times-coming ", "タイムズ"]];
    myFunction();
    expect(GmailApp.search).toHaveBeenCalledTimes(1);
    expect(GmailApp.search).toHaveBeenCalledWith(
      "label:times-coming  label:unread"
    );
  });

  it("processes multiple rows in order, fetching only for labels with mail", () => {
    sheetValues = [
      ["first", "あ"],
      ["second", "い"],
      ["third", "う"],
    ];
    threadsByQuery["label:first label:unread"] = [
      makeThread([makeMessage("a")]),
    ];
    threadsByQuery["label:third label:unread"] = [
      makeThread([makeMessage("c")]),
    ];
    myFunction();
    expect(GmailApp.search.mock.calls.map((c) => c[0])).toEqual([
      "label:first label:unread",
      "label:second label:unread",
      "label:third label:unread",
    ]);
    expect(UrlFetchApp.fetch.mock.calls.map((c) => c[0])).toEqual([
      PREFIX + "%E3%81%82" + SUFFIX,
      PREFIX + "%E3%81%86" + SUFFIX,
    ]);
  });
});

describe("extractSubjectTitle", () => {
  it("drops a leading bracketed tag and cuts at a delimiter", () => {
    expect(
      extractSubjectTitle("【楽天市場】発送完了のお知らせ：注文番号123-456")
    ).toBe("発送完了のお知らせ");
  });

  it("drops a leading half-width bracketed tag", () => {
    expect(extractSubjectTitle("[重要] 振込のお知らせ")).toBe("振込のお知らせ");
  });

  it("cuts at a Japanese comma", () => {
    expect(extractSubjectTitle("タイトル、補足の説明です")).toBe("タイトル");
  });

  it("does not cut when fewer than 3 characters precede the delimiter", () => {
    expect(extractSubjectTitle("本日、お荷物をお届けいたします")).toBe(
      "本日、お荷物をお届けいたします"
    );
    expect(extractSubjectTitle("お荷物お届けのお知らせ")).toBe(
      "お荷物お届けのお知らせ"
    );
  });

  it("caps the result at 20 characters", () => {
    expect(extractSubjectTitle("あ".repeat(30))).toBe("あ".repeat(20));
  });

  it("returns an empty string for empty or bracket-only subjects", () => {
    expect(extractSubjectTitle("")).toBe("");
    expect(extractSubjectTitle(null)).toBe("");
    expect(extractSubjectTitle("【楽天市場】")).toBe("");
  });

  it("keeps a short plain subject as-is", () => {
    expect(extractSubjectTitle("下校メール")).toBe("下校メール");
    expect(extractSubjectTitle("登校メール")).toBe("登校メール");
  });

  it("prefers the segment in corner brackets as the title", () => {
    expect(
      extractSubjectTitle(
        "石川 美和子 先生が「この度、日本スポーツ振興センターの制度改...」をR8第2学年【保護者】に投稿しました。"
      )
    ).toBe("この度、日本スポーツ振興センターの制度改");
  });

  it("removes a trailing ellipsis the sender left in the title", () => {
    expect(
      extractSubjectTitle("山田 先生が「校外学習のお知らせ…」を投稿しました")
    ).toBe("校外学習のお知らせ");
  });
});

describe("myFunction with a detail mode in the third column", () => {
  it("speaks the extracted subject title for the 件名 mode", () => {
    sheetValues = [["rakuten-shipping", "楽天市場", "件名"]];
    const message = makeMessage(
      "【楽天市場】発送完了のお知らせ：注文番号123-456"
    );
    threadsByQuery["label:rakuten-shipping label:unread"] = [
      makeThread([message]),
    ];
    myFunction();
    expect(UrlFetchApp.fetch).toHaveBeenCalledWith(
      PREFIX + encodeURIComponent("楽天市場、発送完了のお知らせ。"),
      { muteHttpExceptions: true, method: "get" }
    );
  });

  it("speaks the corner-bracket title of a school post notification", () => {
    sheetValues = [["bunan", "武南学園", "件名"]];
    const message = makeMessage(
      "石川 美和子 先生が「この度、日本スポーツ振興センターの制度改...」をR8第2学年【保護者】に投稿しました。"
    );
    threadsByQuery["label:bunan label:unread"] = [makeThread([message])];
    myFunction();
    expect(UrlFetchApp.fetch).toHaveBeenCalledWith(
      PREFIX +
        encodeURIComponent(
          "武南学園、この度、日本スポーツ振興センターの制度改。"
        ),
      { muteHttpExceptions: true, method: "get" }
    );
  });

  it("speaks the plain subject of a school gate mail", () => {
    sheetValues = [["bunan", "武南学園", "件名"]];
    const message = makeMessage("下校メール");
    threadsByQuery["label:bunan label:unread"] = [makeThread([message])];
    myFunction();
    expect(message.markRead).toHaveBeenCalledTimes(1);
    expect(UrlFetchApp.fetch).toHaveBeenCalledWith(
      PREFIX + encodeURIComponent("武南学園、下校メール。"),
      { muteHttpExceptions: true, method: "get" }
    );
  });

  it("uses the last message of the first thread for the title", () => {
    sheetValues = [["bunan", "武南学園", "件名"]];
    const older = makeMessage("登校メール");
    const newer = makeMessage("下校メール");
    threadsByQuery["label:bunan label:unread"] = [
      makeThread([older, newer]),
    ];
    myFunction();
    expect(UrlFetchApp.fetch).toHaveBeenCalledWith(
      PREFIX + encodeURIComponent("武南学園、下校メール。"),
      { muteHttpExceptions: true, method: "get" }
    );
  });

  it("keeps the legacy announcement for an unknown mode", () => {
    sheetValues = [["bunan", "武南学園", "その他"]];
    const message = makeMessage("下校メール");
    threadsByQuery["label:bunan label:unread"] = [makeThread([message])];
    myFunction();
    expect(UrlFetchApp.fetch).toHaveBeenCalledWith(
      PREFIX + "%E6%AD%A6%E5%8D%97%E5%AD%A6%E5%9C%92" + SUFFIX,
      { muteHttpExceptions: true, method: "get" }
    );
  });

  it("does not fetch for a mode row without unread mail", () => {
    sheetValues = [["bunan", "武南学園", "件名"]];
    myFunction();
    expect(UrlFetchApp.fetch).not.toHaveBeenCalled();
  });
});

const makeBodyMessage = (subject, body) => ({
  markRead: vi.fn(),
  getSubject: vi.fn(() => subject),
  getPlainBody: vi.fn(() => body),
});

const TOUKOU_BODY =
  "藤原　千咲 さん\n\n武南高等学校からのお知らせです。\n\n中学・高等学校に 2026/11/05 08:23:38 に登校しました。\n\n\n本メールは送信専用のため返信はお受けできません。";
const GEKOU_BODY =
  "藤原　千咲 さん\n\n武南高等学校からのお知らせです。\n\n中学・高等学校より 2026/11/05 17:47:34 に下校しました。\n\n\n本メールは送信専用のため返信はお受けできません。";

describe("extractSchoolDetail", () => {
  it("extracts the arrival time from a 登校 mail body", () => {
    expect(extractSchoolDetail(TOUKOU_BODY)).toBe("8時23分に登校しました");
  });

  it("extracts the leaving time from a 下校 mail body", () => {
    expect(extractSchoolDetail(GEKOU_BODY)).toBe("17時47分に下校しました");
  });

  it("returns an empty string when the body does not match", () => {
    expect(extractSchoolDetail("関係のない本文です")).toBe("");
    expect(extractSchoolDetail("")).toBe("");
  });
});

const TIMES_BODY = [
  "藤原史和　様",
  "",
  "┏━━■予約内容■━━━━━┓",
  "",
  "■予約番号",
  "231551462",
  "■ステーション",
  "坂下第３駐車場",
  "https://share.timescar.jp/view/station/detail.jsp?scd=NM29",
  "■車両",
  "ハスラー （川口 580 ワ 664:ダークブルー）",
  "■利用開始日時",
  "2026/10/25 11:00",
  "■返却予定日時",
  "2026/10/25 23:00 ",
  "■追加運転者",
  "なし",
].join("\n");

describe("extractTimesDetail", () => {
  it("summarises a same-day booking from the real mail layout", () => {
    expect(extractTimesDetail(TIMES_BODY)).toBe(
      "10月25日11時から23時まで、坂下第３駐車場のハスラーです"
    );
  });

  it("spells out the return date for a multi-day booking", () => {
    const body = TIMES_BODY.replace(
      "2026/10/25 23:00 ",
      "2026/10/27 09:30"
    );
    expect(extractTimesDetail(body)).toBe(
      "10月25日11時から10月27日9時30分まで、坂下第３駐車場のハスラーです"
    );
  });

  it("returns an empty string when the block is missing", () => {
    expect(extractTimesDetail("予約内容のない本文")).toBe("");
    expect(extractTimesDetail("")).toBe("");
  });
});

describe("myFunction with the タイムズ mode", () => {
  it("speaks the booking summary", () => {
    sheetValues = [["times-coming ", "タイムズ", "タイムズ"]];
    const message = makeBodyMessage("【Times CAR】予約確認", TIMES_BODY);
    threadsByQuery["label:times-coming  label:unread"] = [
      makeThread([message]),
    ];
    myFunction();
    expect(UrlFetchApp.fetch).toHaveBeenCalledWith(
      PREFIX +
        encodeURIComponent(
          "タイムズ、10月25日11時から23時まで、坂下第３駐車場のハスラーです。"
        ),
      { muteHttpExceptions: true, method: "get" }
    );
  });
});

describe("myFunction with the 登下校 mode", () => {
  it("speaks the gate time from the body", () => {
    sheetValues = [["bunan", "武南学園", "登下校"]];
    const message = makeBodyMessage("下校メール", GEKOU_BODY);
    threadsByQuery["label:bunan label:unread"] = [makeThread([message])];
    myFunction();
    expect(message.markRead).toHaveBeenCalledTimes(1);
    expect(UrlFetchApp.fetch).toHaveBeenCalledWith(
      PREFIX + encodeURIComponent("武南学園、17時47分に下校しました。"),
      { muteHttpExceptions: true, method: "get" }
    );
  });

  it("stays on the legacy announcement when the body has no gate time", () => {
    sheetValues = [["bunan", "武南学園", "登下校"]];
    const message = makeBodyMessage(
      "石川 美和子 先生が「この度、日本スポーツ振興センターの制度改...」をR8第2学年【保護者】に投稿しました。",
      "投稿のお知らせです。"
    );
    threadsByQuery["label:bunan label:unread"] = [makeThread([message])];
    myFunction();
    expect(UrlFetchApp.fetch).toHaveBeenCalledWith(
      PREFIX + "%E6%AD%A6%E5%8D%97%E5%AD%A6%E5%9C%92" + SUFFIX,
      { muteHttpExceptions: true, method: "get" }
    );
  });

  it("falls back to the legacy announcement when neither works", () => {
    sheetValues = [["bunan", "武南学園", "登下校"]];
    const message = makeBodyMessage("", "時刻の入っていない本文");
    threadsByQuery["label:bunan label:unread"] = [makeThread([message])];
    myFunction();
    expect(UrlFetchApp.fetch).toHaveBeenCalledWith(
      PREFIX + "%E6%AD%A6%E5%8D%97%E5%AD%A6%E5%9C%92" + SUFFIX,
      { muteHttpExceptions: true, method: "get" }
    );
  });
});

describe("myFunction with combined modes", () => {
  it("登下校、件名 speaks the subject title when the body has no gate time", () => {
    sheetValues = [["bunan", "武南学園", "登下校、件名"]];
    const message = makeBodyMessage(
      "石川 美和子 先生が「この度、日本スポーツ振興センターの制度改...」をR8第2学年【保護者】に投稿しました。",
      "投稿のお知らせです。"
    );
    threadsByQuery["label:bunan label:unread"] = [makeThread([message])];
    myFunction();
    expect(UrlFetchApp.fetch).toHaveBeenCalledWith(
      PREFIX +
        encodeURIComponent(
          "武南学園、この度、日本スポーツ振興センターの制度改。"
        ),
      { muteHttpExceptions: true, method: "get" }
    );
  });

  it("登下校、件名 still prefers the gate time when present", () => {
    sheetValues = [["bunan", "武南学園", "登下校、件名"]];
    const message = makeBodyMessage("登校メール", TOUKOU_BODY);
    threadsByQuery["label:bunan label:unread"] = [makeThread([message])];
    myFunction();
    expect(UrlFetchApp.fetch).toHaveBeenCalledWith(
      PREFIX + encodeURIComponent("武南学園、8時23分に登校しました。"),
      { muteHttpExceptions: true, method: "get" }
    );
  });

  it("件名、登下校 prefers the subject title when both are available", () => {
    sheetValues = [["bunan", "武南学園", "件名、登下校"]];
    const message = makeBodyMessage("下校メール", GEKOU_BODY);
    threadsByQuery["label:bunan label:unread"] = [makeThread([message])];
    myFunction();
    expect(UrlFetchApp.fetch).toHaveBeenCalledWith(
      PREFIX + encodeURIComponent("武南学園、下校メール。"),
      { muteHttpExceptions: true, method: "get" }
    );
  });

  it("ignores unknown modes mixed into the list", () => {
    sheetValues = [["bunan", "武南学園", "不明、件名"]];
    const message = makeBodyMessage("下校メール", GEKOU_BODY);
    threadsByQuery["label:bunan label:unread"] = [makeThread([message])];
    myFunction();
    expect(UrlFetchApp.fetch).toHaveBeenCalledWith(
      PREFIX + encodeURIComponent("武南学園、下校メール。"),
      { muteHttpExceptions: true, method: "get" }
    );
  });

  it("speaks a delivery subject for a 件名 row", () => {
    sheetValues = [["kuronekoyamato", "クロネコヤマト", "件名"]];
    const message = makeBodyMessage(
      "本日、お荷物をお届けいたします",
      "お届け予定の本文"
    );
    threadsByQuery["label:kuronekoyamato label:unread"] = [
      makeThread([message]),
    ];
    myFunction();
    expect(UrlFetchApp.fetch).toHaveBeenCalledWith(
      PREFIX +
        encodeURIComponent("クロネコヤマト、本日、お荷物をお届けいたします。"),
      { muteHttpExceptions: true, method: "get" }
    );
  });
});

describe("myFunction with DRY_RUN", () => {
  it("logs the composed URL and never fetches when DRY_RUN is 1", () => {
    PropertiesService.getScriptProperties.mockImplementation(() => ({
      getProperty: (name) =>
        name === "SPEECH_URL_PREFIX"
          ? PREFIX
          : name === "DRY_RUN"
            ? "1"
            : null,
    }));
    try {
      sheetValues = [["bunan", "武南学園", "件名"]];
      const message = makeMessage("下校メール");
      threadsByQuery["label:bunan label:unread"] = [makeThread([message])];
      myFunction();
      expect(message.markRead).toHaveBeenCalledTimes(1);
      expect(UrlFetchApp.fetch).not.toHaveBeenCalled();
      expect(Logger.log).toHaveBeenCalledWith(
        "DRY_RUN: " +
          PREFIX +
          encodeURIComponent("武南学園、下校メール。")
      );
    } finally {
      PropertiesService.getScriptProperties.mockImplementation(() => ({
        getProperty: (name) =>
          name === "SPEECH_URL_PREFIX" ? PREFIX : null,
      }));
    }
  });
});

describe("myFunction with blank label rows", () => {
  it("ignores rows whose label cell is empty instead of matching all unread mail", () => {
    sheetValues = [
      ["", "", ""],
      ["bunan", "武南学園", ""],
    ];
    const stray = makeMessage("関係ないメール");
    threadsByQuery["label: label:unread"] = [makeThread([stray])];
    const bunanMessage = makeMessage("下校メール");
    threadsByQuery["label:bunan label:unread"] = [
      makeThread([bunanMessage]),
    ];
    myFunction();
    expect(stray.markRead).not.toHaveBeenCalled();
    expect(bunanMessage.markRead).toHaveBeenCalledTimes(1);
    expect(UrlFetchApp.fetch).toHaveBeenCalledTimes(1);
  });
});
