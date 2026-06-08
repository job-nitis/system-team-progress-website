# System Team Project Progress Site

Static website for showing yearly System Team project progress.

## What is included

- `index.html` - the full website in one file
- `google-apps-script.js` - optional Google Sheets backend script
- Year tabs for switching between project years
- Editable progress summary table
- Project cards with status, owner, quarter, area, progress, and impact
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

For team/shared storage, use Google Sheets:

1. Create a Google Sheet.
2. Open `Extensions` > `Apps Script`.
3. Paste the code from `google-apps-script.js`.
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
