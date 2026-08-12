# Drive Chip Automator for Google Sheets

Drive Chip Automator is a Google Apps Script that converts Google Drive and
Google Docs URLs into native File Smart Chips in Google Sheets. It also
rebuilds existing File Smart Chips so their displayed names update after the
underlying Drive files are renamed.

The script adds a **🚀 Drive Chip Automator** menu to the spreadsheet and can:

- test and refresh one selected cell;
- scan every sheet for supported columns;
- convert plain Drive links into File Smart Chips;
- refresh the displayed names of existing File Smart Chips;
- save its place before the Apps Script runtime limit and continue later; and
- fall back to slower individual updates if a bulk request fails.

## Important safety note

The script refreshes a chip by wiping and reconstructing the target cell. Test
it on a copy of your spreadsheet first, especially if target cells contain
formulas, notes, or other content in addition to a Drive link or File Smart
Chip.

## Supported columns

The script scans columns whose row 1 header contains any of these words,
case-insensitively:

- `attachment`
- `attachments`
- `exhibit`

It starts processing at row 2 and checks every sheet in the spreadsheet.

## Installation

1. Make a backup copy of your spreadsheet.
2. In Google Sheets, open **Extensions → Apps Script**.
3. Replace the contents of `Code.gs` with the contents of this repository's
   [`Code.gs`](Code.gs), then save the project.
4. In the Apps Script editor, find **Services** in the left sidebar and click
   the **+** button.
5. Select **Google Sheets API** and click **Add**.
6. Return to the spreadsheet and reload the page.
7. When **🚀 Drive Chip Automator** appears in the menu bar, select a Drive URL
   or File Smart Chip and run **🧪 Test Update on Selected Cell**.
8. Review and approve Google's authorization prompt if it appears.
9. After the selected-cell test succeeds, run **Convert & Refresh Smart Chips
   (Entire Sheet)**.

The included [`appsscript.json`](appsscript.json) documents the Google Sheets
API v4 dependency for users who manage the complete Apps Script project
manifest.

## Continuing after a timeout

Apps Script executions have runtime limits. When this script approaches its
configured four-minute limit, it saves the current sheet and row in your user
properties and displays a message. Run **Convert & Refresh Smart Chips (Entire
Sheet)** again to continue.

To discard saved progress and restart at the beginning, use **Reset Script
Progress**.

## How it works

The script reads each target cell through the Google Sheets API. It accepts:

- an existing File Smart Chip URI;
- a cell hyperlink; or
- a plain-text `drive.google.com` or `docs.google.com` URL.

For each recognized URL, it clears the cell and writes a new `chipRuns` value
using the original Drive URL. Reconstructing the chip causes Google Sheets to
resolve the file's current name.

## Requirements

- A Google Sheet with a spreadsheet-bound Apps Script project
- Permission to edit the spreadsheet
- Access to the linked Drive files
- The **Google Sheets API** service enabled in Apps Script
- V8 Apps Script runtime

## Troubleshooting

### `Sheets is not defined`

In the Apps Script editor, click **Services +**, select **Google Sheets API**,
and click **Add**.

### The custom menu does not appear

Save the Apps Script project and reload the spreadsheet. The `onOpen` function
creates the menu when the spreadsheet opens.

### No columns are updated

Confirm that the relevant column headers are in row 1 and contain
`attachment`, `attachments`, or `exhibit`.

### A cell is skipped

The cell must contain an existing File Smart Chip, a hyperlink, or a plain
Google Drive/Docs URL. The executing Google account must also have access to
the linked file.

## Privacy

The repository contains no spreadsheet IDs, Drive file IDs, credentials, or
user data. The script operates inside the spreadsheet where you install it.

## License

This project is licensed under the [MIT License](LICENSE).
