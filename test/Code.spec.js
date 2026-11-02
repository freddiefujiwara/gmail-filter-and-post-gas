import { beforeEach, describe, expect, it, vi } from "vitest";
import { myFunction } from "../src/Code.js";

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
