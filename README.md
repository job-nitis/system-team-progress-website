# System Team Project Progress Site

Static website for showing yearly System Team project progress.

## What is included

- `index.html` - the full website in one file
- `google-apps-script-progress-only.js` - Apps Script for the `ProgressRows` sheet tab
- `google-apps-script-highlight-data-only.js` - Apps Script for the `Highlight information` sheet tab and Drive picture upload
- `google-apps-script-license-time-only.js` - Apps Script for the `License detail` sheet tab
- `google-apps-script-users-only.js` - Apps Script for the `Users` sheet tab and email permission lookup
- Year tabs for switching between project years
- Editable progress summary table
- Editable license table with remaining-time color status
- `Add Project`, `Edit`, and `Delete` actions for table rows
- Highlight cards with month filter, image upload, topic, and detail
- Email login with Viewer, Editor, and Admin role controls
- Highlight data saved to Google Sheets and pictures saved to Google Drive when Apps Script is deployed
- Project cards with status, owner, overall project, and progress
- Summary metrics for total, completed, active, and watch projects
- Timeline and decision sections for annual reporting

## How to update project data

Open `index.html` and find this section near the bottom:

```js
const projectYears = [
```

Edit the placeholder projects inside that list. For each project, you can update:

- `name`
- `description`
- `status`
- `progress`
- `owner`
- `quarter`
- `impact`
- `tech`
- table fields such as `tor`, `pr`, `sapPr`, and `po`

Use these status values for the built-in colors:

- `On track`
- `Watch`
- `Done`
- `Blocked`

## How table storage works

The progress table saves edits in the browser immediately using `localStorage`.
Deleting a project removes it from the active year after a confirmation popup.
The website also loads table rows from Google Sheets on page load, when the year changes, and once every minute while open.

Highlight data uses a spreadsheet named `Highlight data` with a sheet tab named `Highlight information` and these columns:

```text
ID | Year | Month | Topic | Detail | Image URL | Created At
```

Highlight pictures are uploaded to the configured Google Drive folder in `google-apps-script-highlight-data-only.js`.

License data uses a spreadsheet named `License Time` with a sheet tab named `License detail` and these columns:

```text
ID | License Name | Expired Date | Remaining Time | Updated At
```

The website and Apps Script calculate Remaining Time automatically from Expired Date.

User permission data uses a spreadsheet named `System Website User` with a sheet tab named `Users` and these columns:

```text
Email | Name | Role | Active
```

Roles:

```text
Viewer = view only
Editor = add and edit
Admin = add, edit, and delete
```

For team/shared storage, use Google Sheets:

1. Create a Google Sheet.
2. Open `Extensions` > `Apps Script`.
3. Paste the code from `google-apps-script-highlight-data-only.js` for Highlight data, or `google-apps-script-progress-only.js` for ProgressRows.
4. Deploy it as a web app.
5. Copy the web app URL.
6. In `index.html`, find:

```js
const GOOGLE_APPS_SCRIPT_URL = "";
```

Paste your web app URL between the quotes.

The Google Sheet will store these columns:

```text
Year | ID | Initiative Project | Owner | Tech Preparation | TOR | PR | SAP PR | PO | Delivery | FAT/SAT | Close MOC | Overall Plan | Updated At
```

## How to publish with GitHub Pages

1. Upload these files to a GitHub repository.
2. In GitHub, open repository `Settings`.
3. Go to `Pages`.
4. Under `Build and deployment`, choose `Deploy from a branch`.
5. Select the branch that contains `index.html`.
6. Save, then wait for GitHub to provide the Pages URL.
