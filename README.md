# gmail-filter-and-post-gas

The back side of a small mail notifier. It watches a list of Gmail
labels kept in a Google Sheets sheet and announces new mail on a
home speaker.

## How it works

- The spreadsheet bound to this project is the label list. There is
  no header row.
  - Column A: a Gmail label name.
  - Column B: the announcement text for that label, in plain text.
    The code URL-encodes it when building the speech request, so
    the sheet stays readable and editable as normal Japanese.
- `myFunction` runs every minute from a time-driven trigger.
  - For each row with at least 2 columns, it searches Gmail for
    `label:<label> label:unread`.
  - If unread mail exists, it marks every message in those threads
    as read, logs each subject, and sends one GET request to
    `SPEECH_URL_PREFIX` + the row's text + a fixed suffix (a
    sentence meaning "related mail has arrived, please check it").
  - If there is no unread mail for the label, it does nothing for
    that row. Rows are processed in sheet order.

## Setup

1. Open the Apps Script project bound to the label sheet.
2. In the project settings, set this script property. Its value is
   private, so it is not in this repository.
   - `SPEECH_URL_PREFIX`: URL prefix of the speech command, ending
     with `/-s/`. The announcement text is appended to it.
3. Set `.clasp.json` to that project's script ID.
4. Run `npm install`, `npm run build`, and `npm run deploy`
   (clasp push) to send `dist/` to the project.
5. Add a time-driven trigger: `myFunction`, every minute.

This project is not a web app, so there is no web app deployment.
GitHub Actions (Node CI) can build and deploy on push when the
repository secret `CLASPRC_JSON` is set. No `DEPLOYMENT_ID` secret
is needed.

## Test

```
npm test
```

Unit tests mock the Apps Script services (SpreadsheetApp, GmailApp,
UrlFetchApp, PropertiesService, Logger) and check that short rows
are skipped, that nothing is fetched when a label has no unread
mail, that messages are marked read and announced exactly once with
the expected URL and options, that the search query is built
exactly (including labels with a trailing space), and that rows are
processed in order. The build test checks that `dist/Code.gs` is
written without `export` statements and that the manifest is
copied.
