# TestFlow Runner

TestFlow Runner is a local-first desktop automation tool for recording and running UI test flows with datasets, variables, screenshots, and HTML reports.

## How to run
1. Create/activate a Python 3.10+ environment.
2. Install dependencies:
   - `python3 -m pip install pyautogui pynput`
   - Optional for Excel import: `python3 -m pip install openpyxl`
3. Start the app:
   - `python3 autorunner.py`

## How to create targets
1. Open `Targets` page.
2. Use `Capture` to record current mouse coordinates, or `Manual Add` for explicit X/Y.
3. Use clear names like `login_button`, `username_field`.

## How to record a flow
1. Open `Recorder` page.
2. Click `Start Recording`.
3. Perform the workflow in the target app.
4. Press the configured stop hotkey (default `F8`) or click `Stop Recording`.
5. Review the preview and save as:
   - Reusable Flow
   - Test Case
   - Append to existing Flow

## How to import CSV test cases
1. Open `Test Cases` page.
2. Click `Import`.
3. Select `.csv` or `.xlsx`.
4. Review import summary and confirm.

Expected columns:
- `TestCaseId`
- `TestCaseName`
- `Suite`
- `StepNo`
- `ActionType`
- `Target`
- `Value`
- `Seconds`
- `Description`
- `Enabled`

## How to run tests
1. Open `Test Cases` or `Reusable Flows`.
2. Select the item.
3. Use:
   - `Run` for live execution
   - `Dry Run` for validation and plan expansion without clicks
4. For dataset test cases, choose a row index or use `Run All Rows`.
5. Open `Run Center` to inspect historical runs and reports.

## Keyboard shortcuts
- `Ctrl+S` / `Cmd+S`: Save project
- `F5`: Run selected
- `Shift+F5`: Dry run selected
- `F8`: Stop runner/recorder
- `Ctrl+D` / `Cmd+D`: Duplicate selected step
- `Delete`: Delete selected step

## Known limitations
- Live UI automation depends on desktop permissions (especially on macOS Accessibility/Screen Recording).
- Recorder requires `pynput`.
- XLSX import requires `openpyxl`.
- Coordinate-based automation can be sensitive to resolution, scaling, and UI layout changes.

## Project files
- `testflow_project.json`: project model
- `runs/`: run artifacts (`run.json`, `report.html`, screenshots)
- `logs/app_errors.log`: error tracebacks
- `testflow_project.backup_YYYYMMDD_HHMMSS.json`: migration/recovery backups
