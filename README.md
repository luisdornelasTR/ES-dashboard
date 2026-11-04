# ES Dashboard

A read-only dashboard for the Embedded Specialist program. Import an individual customer's Excel tracking spreadsheet to view tickets, meeting topics, and objectives.

The application is built with TypeScript, Vite, and SheetJS. Spreadsheet processing happens in the browser.

## Requirements

- Node.js 22.12+ or a newer supported LTS version.
- npm, included with Node.js.
- Git, if cloning the repository.

Verify your installation:

```bash
node -v
npm -v
git --version
```

On Windows, if PowerShell blocks `npm.ps1`, use `npm.cmd` instead of `npm` in every command below.

## Run locally

### 1. Download the project

On the repository's GitHub page, click **Code** and copy its HTTPS URL.

Run the following command, replacing `<repository-url>` with that URL:

```bash
git clone <repository-url>
```

Open the downloaded project folder in VSCode:

1. Select **File → Open Folder**.
2. Select the folder containing `package.json`.
3. Select **Terminal → New Terminal**.


### 2. Install dependencies

From the folder containing `package.json`, run:

```bash
npm install
```

On Windows PowerShell:

```powershell
npm.cmd install
```

### 3. Start the development server

```bash
npm run dev
```

On Windows PowerShell:

```powershell
npm.cmd run dev
```

The terminal will display a local URL, usually:

```text
http://localhost:5173/
```

Open the exact URL shown in the terminal. If port 5173 is already in use, Vite may choose another port.


## Use the dashboard

1. Click **Import spreadsheet**.
2. Select an individual customer tracking file in `.xlsx` format.
3. Navigate between **Tickets**, **Meetings**, and **Objectives**.

Import the individual customer spreadsheet, not the master tracking spreadsheet.

### Tickets

- Click the total or a status count to filter tickets. The horizontal bar shows the share of each status and also supports filtering.
- Only statuses present in the imported spreadsheet appear in the overview.
- View descriptions, status, priority, links, notes, and updates.
- Filter by multiple ticket types and statuses.
- Search ticket content.
- Expand a ticket to view its details.
- Click a ticket ID badge (for example, **TR-15364**) to open its associated Jira link in a new tab without expanding the ticket.
- In an expanded ticket, use **Copy** next to the Jira or Freshdesk link to copy its full URL. The button confirms when the link has been copied.
- Open tickets appear before completed, rejected, or closed tickets.

Ticket statuses are read from the spreadsheet. *The dashboard does not edit tickets or update Jira.*

### Meetings

- See discussed, resolved, pending, and unrecorded topic counts across all meetings and for each date.
- Use the pending count at the top to open all pending topics, or a meeting's pending shortcut to focus on that date.
- Each meeting shows its resolution percentage, including while collapsed.
- All meeting dates are selected by default.
- Use the date dropdown to select or deselect multiple dates.
- Expand a date to view its topics, then expand a topic for details.
- Filter topics by resolved, pending, or unrecorded status.

### Objectives

- View each objective's description, current state, importance, and status.
- Objectives are ordered by importance: Must Have, Highly Desirable, then New opportunity.
- Target dates and progress percentages are not displayed.

### Filters

Dropdown filters support:

- Selecting multiple options.
- **Select all** to select every option.
- **Clear** to deselect every option.
- Searching within the available options.

Clearing a filter selects nothing, so no matching records will appear until options are selected again.

Overview counts always describe all imported records. The result count shows how many match the current filters and search. Clicking an overview count or pending shortcut clears previous filters and search so the results match the selected count; the dropdowns update to reflect the selection.

## Spreadsheet format

The source of data needs to be an individual ES tracking sheet.


## Project files

| Path | Purpose |
| --- | --- |
| `src/main.ts` | Spreadsheet import, filtering, and interface rendering |
| `src/style.css` | Dashboard styling |
| `src/assets/tr-logo.png` | testRigor logo for styling |
| `index.html` | Application entry page |
| `package.json` | Dependencies and npm commands |
