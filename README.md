# TestCanvas User Guide

## 0) Quick Start (60 Seconds)

1. Enter Run name, tester name, and version/build.
2. Click Add Test.
3. Fill Test Scenario and UAT Results.
4. Add evidence:
   - Screenshot: click Attach Screenshot (inside a test card), or
   - Video: click Record for this test.
5. Press Ctrl+S or click Save.
6. Click Export Run JSON before you close browser.

## 1) What This Tool Saves Automatically

- TestCanvas is local-first. Working data is saved in your browser on your machine.
- Save creates a stronger checkpoint snapshot for recovery.
- Save status is shown near the top (for example: Saved locally, Summary updated, Ready).

Important:
- Browser-local data can be lost if cache/site data is cleared, profile changes, or storage is full.
- Do not rely on local autosave alone for important runs.

## 2) High-Risk Actions That Can Cause Data Loss

1. Clicking New Test Run without exporting first.
2. Clearing browser cache/site data.
3. Switching browser profile/device and expecting local data to follow.
4. Closing browser after major edits without creating backup files.

## 3) Safe Working Routine (Recommended)

1. Start or continue your run.
2. Click Save after major changes.
3. Press Ctrl+S regularly (same full manual save as Save button).
4. Before risky actions (New Test Run, browser cleanup, end of day):
   - Click Export Run JSON.
   - Optionally click Export Summary for shareable report.

## 4) Button Guide (Top Header)

- Save:
  - Creates manual checkpoint snapshot and updates local state.
- Restore Last Snapshot:
  - Restores latest valid snapshot if newer.
- Export Run JSON:
  - Exports full run for continuation/recovery.
- Load Run File:
  - Loads exported JSON or exported HTML with embedded run data.
- Upload Excel/CSV:
  - Imports test scenarios from spreadsheet.
- Download Excel Template:
  - Downloads spreadsheet template.
- Export Media:
  - Exports media package.
- Review Mode:
  - Switches UI for review workflows.
- New Test Run:
  - Clears current local run and starts fresh.
- Export Summary:
  - Generates shareable HTML report.
- Export Summary Excel:
  - Generates spreadsheet summary.

## 5) Buttons in Main Content

- Add Test:
  - Adds a new test card.
- Update (inside Consolidated Test Summary section):
  - Refreshes consolidated summary from latest test card content.
- Floating summary Update:
  - Refreshes floating navigation summary panel.
- Bottom floating buttons:
  - Save: quick manual save.
  - Add Test: quick create test.

## 6) How To Add Screenshots (Detailed)

### A) Add Screenshot to SetUp Section

1. In sidebar SetUp section, click Add Screenshot.
2. Pick one or multiple image files.
3. Thumbnails appear in SetUp -> Screenshots.
4. Click `x` on a thumbnail to remove it.
5. Press Ctrl+S or click Save.

### B) Add Screenshot to a Specific Test

1. Click Attach Screenshot inside that test card.
2. Select image files.
3. Images appear under that test.
4. Add annotation/comment if needed.
5. Press Ctrl+S or click Save.

### C) Paste Screenshot from Clipboard

1. Copy image to clipboard.
2. Click Paste Screenshot in the test card or use sidebar Paste Screenshot.
3. Verify image appears under the target test.
4. Save.

### D) Manual Screenshot Import Shortcut

- Ctrl+Shift+S opens screenshot import for latest test target.

Tip:
- If no test exists yet, create one first with Add Test.

## 7) How To Record Video (Detailed)

### A) Record Video for a Specific Test (Recommended)

1. In the target test card, click Record for this test.
2. Browser asks for screen/window/tab permission: choose what to record.
3. Recording starts; use Pause and Stop Recording buttons in that test card.
4. You can also use floating controls (Pause/Stop) while recording.
5. After stopping, video appears in that test card under videos.
6. Add notes if needed, then Save.

### B) Global Recording from Sidebar

1. Click Start Recording in sidebar.
2. Select screen/window/tab to capture.
3. Use Pause or Stop in sidebar/floating controls.
4. Verify recording is attached where expected.
5. Save.

Shortcuts:
- Ctrl+Shift+P: pause/resume recording.
- Ctrl+Shift+X: stop recording.

## 8) How To Avoid Data Loss (Detailed)

### A) During Active Testing

1. Enter run metadata early (run name, owner, version/build).
2. Save after each major scenario or attachment batch.
3. Check save status text after save.

### B) Before Closing Browser / End of Day

1. Click Export Run JSON and keep the file in a safe folder.
2. Optionally click Export Summary for a shareable report snapshot.
3. If needed, also export summary Excel.

### C) Before Clicking New Test Run

1. Export Run JSON first.
2. Confirm the file exists in Downloads/chosen folder.
3. Then click New Test Run.

## 9) How To Export Summary (Detailed)

1. Make sure all tests are updated.
2. Click Update in Consolidated Test Summary.
3. Click Export Summary.
4. Open exported HTML and verify:
   - Consolidated Test Summary
   - Floating summary navigation
   - Test details and attachments
5. Share this HTML for review.

Tip:
- Export Summary = best for sharing.
- Export Run JSON = best for resuming work.

## 10) How To Load a File and Continue an Old Test

1. Open TestCanvas.
2. Click Load Run File.
3. Select either:
   - Export Run JSON (recommended), or
   - Exported HTML summary with embedded run data.
4. Wait for UI re-render.
5. Click Save immediately to create a fresh checkpoint.
6. Continue editing.

If data looks missing:
1. Click Restore Last Snapshot.
2. If still missing, load latest Export Run JSON backup.

## 11) Keyboard Shortcuts

- Ctrl+S: full manual save/checkpoint.
- Ctrl+Shift+S: manual screenshot import for latest test target.
- Ctrl+Shift+V: voice annotation on latest screenshot context.
- Ctrl+Shift+P: pause/resume recording.
- Ctrl+Shift+X: stop recording.

## 12) Team Sharing Best Practice

Use two-file strategy for reliability:

1. For continuation: keep Export Run JSON.
2. For reading/review: share Export Summary HTML.

For GitHub Pages usage:
- Share the site URL for tool access.
- Still export JSON regularly because runtime data remains browser-local unless backend sync is added.
