# System Team Project Progress Site

Static website for showing yearly System Team project progress.

## What is included

- `index.html` - the full website in one file
- Year tabs for switching between project years
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

Use these status values for the built-in colors:

- `On track`
- `Watch`
- `Done`
- `Blocked`

## How to publish with GitHub Pages

1. Upload these files to a GitHub repository.
2. In GitHub, open repository `Settings`.
3. Go to `Pages`.
4. Under `Build and deployment`, choose `Deploy from a branch`.
5. Select the branch that contains `index.html`.
6. Save, then wait for GitHub to provide the Pages URL.
