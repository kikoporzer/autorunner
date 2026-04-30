# Codex Implementation Plan — Lightweight Visual Test Flow Runner

## Goal

Upgrade the current Python desktop automation prototype into a finished but intentionally lightweight product for testers.

The finished app should support:

- manual recording of flows
- reusable smaller flows/actions
- named click targets
- editable test cases
- Excel/CSV test case import
- variable substitution
- dataset-driven execution
- run logs
- screenshots/evidence
- basic assertions
- professional but simple UI
- local project storage

Do **not** over-engineer this into a cloud platform. Keep it local-first, file-based, and maintainable.

---

## Product Name Placeholder

Use this working name in code/UI unless another name already exists:

```text
TestFlow Runner
```

---

## Current Assumption

The project already has a Python/Tkinter GUI using:

```text
tkinter
pyautogui
keyboard
json
threading
time
ctypes
```

The existing app already supports:

- named positions
- flows
- adding/editing/deleting actions
- running flows
- recording clicks with timing
- JSON persistence

If some of this is missing, implement it as part of Pass 1.

---

# Target App Model

Use this product model:

```text
Project
├── Targets / Named Positions
├── Reusable Flows
├── Test Cases
├── Datasets
├── Runs
└── Reports / Evidence
```

Conceptually:

```text
Atomic Actions → Reusable Flows → Test Cases → Test Suites → Run Results
```

Keep this simple in implementation. Avoid building complex inheritance or plugin systems.

---

# Data Model

Store project data in a single local JSON file first.

Default file:

```text
testflow_project.json
```

Recommended structure:

```json
{
  "appVersion": "0.1.0",
  "settings": {
    "startupDelaySeconds": 3,
    "defaultActionPauseSeconds": 0.1,
    "screenshotOnFailure": true,
    "screenshotAfterEachStep": false,
    "stopHotkey": "f8",
    "recordingStopHotkey": "f8"
  },
  "environment": {
    "name": "Default",
    "expectedResolution": "",
    "notes": ""
  },
  "targets": {
    "login_button": {
      "x": 812,
      "y": 433,
      "description": "Main login button",
      "createdAt": "2026-04-30T12:00:00"
    }
  },
  "flows": {
    "Login_POS": {
      "name": "Login_POS",
      "description": "Reusable login flow",
      "parameters": ["username", "password"],
      "steps": []
    }
  },
  "testCases": {
    "POS001": {
      "id": "POS001",
      "name": "Valid login",
      "suite": "Smoke",
      "description": "",
      "dataset": "",
      "enabled": true,
      "steps": []
    }
  },
  "datasets": {},
  "runs": []
}
```

For now, keep datasets either:

1. imported into JSON as rows, or
2. referenced by path with parsed preview.

Prefer importing data into JSON for simplicity.

---

# Supported Step Types

Implement these step types.

## Core Actions

```json
{ "type": "click", "target": "login_button" }
{ "type": "click_xy", "x": 812, "y": 433 }
{ "type": "double_click", "target": "row_1" }
{ "type": "right_click", "target": "context_menu_area" }
{ "type": "type_text", "value": "${username}" }
{ "type": "press_key", "key": "enter" }
{ "type": "hotkey", "keys": ["ctrl", "s"] }
{ "type": "wait", "seconds": 2 }
{ "type": "screenshot", "name": "after_login" }
{ "type": "comment", "text": "Manual note" }
```

## Composition

```json
{ "type": "run_flow", "flow": "Login_POS" }
```

## Basic Assertions

```json
{ "type": "assert_window_title_contains", "value": "Dashboard" }
{ "type": "assert_clipboard_contains", "value": "${expectedValue}" }
{ "type": "assert_file_exists", "path": "C:/Temp/export.csv" }
```

Keep assertions basic. Do not implement OCR in this version.

---

# Variable System

Implement simple variable substitution.

Syntax:

```text
${username}
${password}
${productCode}
${expectedTotal}
```

Variable sources, in priority order:

1. current dataset row
2. test case variables
3. flow input variables
4. environment variables
5. built-in variables

Built-in variables:

```text
${today}
${now}
${timestamp}
${runId}
```

Examples:

```json
{ "type": "type_text", "value": "${username}" }
{ "type": "screenshot", "name": "login_${timestamp}" }
```

If a variable is unresolved, the app should fail pre-run validation before executing.

---

# Recommended Additional Features

Implement these because they will materially improve usability without making the app too complex.

## 1. Pre-Run Validation

Before running a test case or flow, validate:

- missing target names
- missing reusable flows
- unresolved variables
- missing dataset columns
- invalid step type
- invalid coordinates
- screen resolution mismatch, warning only
- empty flow/test case
- circular flow references

Show a clear validation dialog:

```text
Cannot run POS001:

- Step 3 references missing target: login_button
- Step 5 uses unresolved variable: ${password}
- Step 8 calls missing flow: Open_POS
```

## 2. Run Log Panel

Add a bottom run log panel showing:

```text
[12:01:03] Starting test case POS001
[12:01:06] Step 1 passed: run_flow Login_POS
[12:01:08] Step 2 passed: click login_button
[12:01:10] Step 3 failed: assert_window_title_contains Dashboard
```

Each run should store:

- run id
- start time
- end time
- status
- test case name
- step results
- error messages
- screenshot paths

## 3. Evidence Screenshots

Add:

- screenshot action
- screenshot on failure
- optional screenshot after every step
- evidence folder per run

Folder format:

```text
runs/
└── 2026-04-30_120103_POS001/
    ├── run.json
    ├── report.html
    └── screenshots/
        ├── step_001.png
        └── failure_step_006.png
```

## 4. Dry Run

Dry run should execute no clicks or keyboard input.

It should only:

- resolve variables
- expand reusable flows
- validate steps
- show the final execution plan

This is essential for trust.

## 5. Step-by-Step Mode

Allow the runner to execute one step at a time.

Buttons:

```text
Run
Dry Run
Run From Selected Step
Step Once
Stop
```

Step-by-step mode should be simple. Do not build a full debugger.

## 6. Duplicate and Disable Steps

Testers need fast editing.

Each step should support:

- duplicate
- delete
- enable/disable
- move up/down
- edit

Disabled step format:

```json
{
  "type": "click",
  "target": "login_button",
  "enabled": false
}
```

Disabled steps are skipped but shown in the UI.

## 7. Export HTML Report

Generate a simple local HTML report after each test case/suite run.

Include:

- test status
- duration
- step list
- errors
- screenshots as relative links
- dataset row used

No PDF export in this version.

---

# UI Design Requirements

Use Tkinter, but make it feel organized and product-like.

Do not spend time on custom themes beyond a clean layout.

## Main Layout

Use this structure:

```text
┌──────────────────────────────────────────────────────────────┐
│ Top Bar: Project name | Environment | Run status              │
├───────────────┬──────────────────────────────┬───────────────┤
│ Navigation    │ Main Workspace               │ Inspector     │
│               │                              │               │
│ Dashboard     │ Step list / table            │ Selected step │
│ Test Cases    │                              │ properties    │
│ Flows         │                              │               │
│ Targets       │                              │               │
│ Recorder      │                              │               │
│ Datasets      │                              │               │
│ Run Center    │                              │               │
│ Settings      │                              │               │
├───────────────┴──────────────────────────────┴───────────────┤
│ Run Log                                                       │
└──────────────────────────────────────────────────────────────┘
```

## Navigation Sections

Implement pages:

1. Dashboard
2. Test Cases
3. Reusable Flows
4. Targets
5. Recorder
6. Datasets
7. Run Center
8. Settings

Do not create separate windows for everything. Prefer one main window with panels.

## Step List

Use a Treeview table with columns:

```text
#
Enabled
Type
Target / Flow
Value
Wait
Description
```

## Inspector

When a step is selected, show editable fields on the right.

Minimum inspector fields:

```text
Type
Target
Value
Seconds
Description
Enabled
```

Only show relevant fields for the selected step type if easy. Otherwise show common fields.

---

# Excel / CSV Import

To avoid dependency issues, implement CSV import first.

If `openpyxl` is available, support `.xlsx`. If not, show a clear message.

## CSV Import for Test Cases

Supported columns:

```text
TestCaseId
TestCaseName
Suite
StepNo
ActionType
Target
Value
Seconds
Description
Enabled
```

Example:

```csv
TestCaseId,TestCaseName,Suite,StepNo,ActionType,Target,Value,Seconds,Description,Enabled
POS001,Valid login,Smoke,1,run_flow,Login_POS,,,Login to POS,TRUE
POS001,Valid login,Smoke,2,assert_window_title_contains,,Dashboard,,Validate dashboard,TRUE
```

Map rows into `testCases`.

## CSV Import for Datasets

Any CSV can be imported as a named dataset.

Example:

```csv
username,password,productCode,expectedTotal
user01,Pass123,P1001,12.50
user02,Pass456,P1002,18.90
```

Run test case once per dataset row.

## Import Validation

Before importing, show:

```text
Rows found: 48
Test cases found: 6
Invalid rows: 2
```

For invalid rows, show row number and reason.

---

# Recorder Requirements

The recorder should have its own page.

## Recorder Options

Checkboxes:

```text
[x] Record left clicks
[x] Record double clicks
[x] Record right clicks
[x] Record typing
[x] Record hotkeys
[x] Record timing gaps
[ ] Screenshot after each click
[ ] Convert repeated coordinates into named targets
```

Minimum implementation can support only clicks/double-clicks/right-clicks/timing first. But the UI should already show these options, with unsupported ones disabled or marked as "coming later".

## Recorder Output

After stopping recording, show a preview table:

```text
1 wait 1.25
2 click_xy 812 433
3 wait 0.80
4 double_click 912 540
```

Buttons:

```text
Save as Reusable Flow
Save as Test Case
Discard
```

Do not automatically save without preview.

---

# Runner Engine Requirements

Implement a clean runner module/class.

Suggested files:

```text
app.py
storage.py
models.py
runner.py
recorder.py
importer.py
reporting.py
validation.py
ui/
```

If the existing app is one file, refactor carefully. Do not rewrite all logic in one risky pass unless it is still small.

## Runner Responsibilities

Runner should:

1. load selected test case or flow
2. expand nested `run_flow` steps
3. resolve variables
4. validate
5. execute steps
6. log results
7. save evidence
8. generate report

## Stop Handling

Implement a stop flag.

The UI Stop button should set:

```python
runner.stop_requested = True
```

The runner should check before each step.

Also retain emergency mouse fail-safe if using pyautogui.

---

# Implementation Passes

Use these long targeted passes. Do not create many tiny commits unless necessary.

---

## Pass 1 — Stabilize Data Model and Runner

### Objective

Create a clean internal data model and execution engine that supports flows, test cases, reusable flow calls, variables, validation, run logging, screenshots, and dry run.

### Tasks

1. Create or refactor into these modules if reasonable:

```text
storage.py
runner.py
validation.py
reporting.py
```

2. Implement project JSON schema using the structure in this document.

3. Implement migration from the current `automation_config.json` if it exists:
   - old `positions` becomes `targets`
   - old flat `flows` becomes new flow objects
   - raw click actions remain supported as `click_xy`

4. Implement step execution for:

```text
click
click_xy
double_click
right_click
type_text
press_key
hotkey
wait
screenshot
comment
run_flow
assert_window_title_contains
assert_clipboard_contains
assert_file_exists
```

5. Implement variable substitution.

6. Implement reusable flow expansion.

7. Implement circular flow detection.

8. Implement pre-run validation.

9. Implement run result object.

10. Implement screenshot storage per run.

11. Implement basic HTML report generation.

12. Implement dry-run mode.

### Acceptance Criteria

- Existing flows can still run after migration.
- A test case can call a reusable flow.
- A test case can use `${username}` from a dataset row.
- Dry run shows expanded steps without clicking.
- Missing targets/variables are caught before execution.
- Failed assertions create a failed run result.
- Failure screenshot is saved if enabled.
- HTML report is generated for a run.

---

## Pass 2 — Build Product UI Shell and Editors

### Objective

Replace the basic GUI layout with a clean product-style shell and add dedicated pages for test cases, reusable flows, targets, recorder, datasets, run center, and settings.

### Tasks

1. Create a single-window Tkinter layout:

```text
left navigation
top bar
main workspace
right inspector
bottom run log
```

2. Add pages:

```text
Dashboard
Test Cases
Reusable Flows
Targets
Recorder
Datasets
Run Center
Settings
```

3. Implement reusable components:

```text
StepTable
StepInspector
RunLogPanel
TargetList
```

4. Test Cases page:
   - list test cases
   - create/delete/duplicate test case
   - edit metadata
   - edit steps
   - run selected test case
   - dry run selected test case

5. Reusable Flows page:
   - list flows
   - create/delete/duplicate flow
   - edit parameters
   - edit steps
   - run selected flow
   - dry run selected flow

6. Targets page:
   - list named targets
   - capture target
   - manual add target
   - rename/delete target
   - test click target
   - show coordinates

7. Settings page:
   - startup delay
   - default action pause
   - screenshot on failure
   - screenshot after each step
   - expected screen resolution
   - stop hotkey

8. Bottom run log:
   - append runner events
   - clear log button
   - save log button optional

### Acceptance Criteria

- User can manage test cases and reusable flows separately.
- User can edit steps through table + inspector.
- User can run/dry-run from the UI.
- Run log updates during execution.
- UI does not freeze while running.
- Stop button works before the next step starts.
- The app feels organized and not like a script menu.

---

## Pass 3 — Recorder Upgrade and Recording Cleanup

### Objective

Make recording usable as a product feature, not just a hidden utility.

### Tasks

1. Create a Recorder page.

2. Add recorder options:
   - record left clicks
   - record double clicks
   - record right clicks
   - record timing gaps
   - screenshot after each click
   - ignore clicks inside app window

3. Implement Start Recording and Stop Recording.

4. Display recording state clearly:
   - top bar status
   - red "Recording" label
   - stop hotkey hint

5. After stopping, show recorded step preview.

6. Add buttons:
   - Save as Reusable Flow
   - Save as Test Case
   - Append to Existing Flow
   - Discard

7. Add cleanup actions:
   - remove waits below X seconds
   - round waits to 0.1 seconds
   - convert repeated coordinates to suggested targets
   - replace raw coordinates with existing nearby targets if within tolerance

8. Implement "suggest targets" logic:
   - group repeated click coordinates within 8 px
   - propose names like `target_001`, `target_002`
   - allow user to accept/reject

### Acceptance Criteria

- User can record a manual workflow.
- User sees preview before saving.
- User can save recording as flow or test case.
- Repeated clicks can be converted to named targets.
- Existing nearby targets are reused where possible.
- Recorder ignores clicks on the app itself, if feasible.

---

## Pass 4 — CSV/Excel Import and Dataset Runs

### Objective

Let testers import real test cases and data from files.

### Tasks

1. Implement CSV dataset import.

2. Implement CSV test case import.

3. Add optional XLSX support:
   - try importing `openpyxl`
   - if missing, show message: "XLSX import requires openpyxl. Use CSV or install openpyxl."

4. Add Datasets page:
   - import CSV
   - preview rows
   - delete dataset
   - rename dataset
   - show columns

5. Add Test Case import:
   - choose file
   - validate columns
   - preview grouped test cases
   - import or cancel

6. Implement dataset-driven execution:
   - run selected test case once
   - run selected test case for all dataset rows
   - store dataset row index in run result

7. Add variable preview:
   - select test case
   - select dataset row
   - show resolved step values

### Acceptance Criteria

- CSV dataset import works without external libraries.
- CSV test case import works.
- XLSX import works only if openpyxl is installed.
- Test case can run once per dataset row.
- Reports identify which dataset row was used.
- Unresolved dataset variables are caught before execution.

---

## Pass 5 — Polish, Reliability, and Packaging Readiness

### Objective

Finish the product experience and make the app usable by testers without developer support.

### Tasks

1. Add startup checks:
   - Python version
   - pyautogui installed
   - keyboard installed
   - screen resolution
   - project file loaded
   - write permissions for project folder

2. Add command/status bar:
   - current project
   - current page
   - recording status
   - running status
   - last saved time

3. Add autosave after meaningful changes.

4. Add backup before migration:
   - `testflow_project.backup_YYYYMMDD_HHMMSS.json`

5. Add project import/export:
   - export project folder as zip if using only standard library
   - import project from folder or JSON

6. Add error handling:
   - friendly message boxes
   - detailed traceback saved to `logs/app_errors.log`

7. Add keyboard shortcuts:
   - Ctrl+S save
   - F5 run selected
   - Shift+F5 dry run
   - F8 stop
   - Ctrl+D duplicate step
   - Delete delete selected step

8. Add simple visual polish:
   - consistent spacing
   - consistent button names
   - status chips
   - colored pass/fail text
   - icons optional, text is enough
   - avoid clutter

9. Add sample project:
   - one dataset
   - one login flow
   - one demo test case
   - one generated report example

10. Add README:

```text
How to run
How to create targets
How to record a flow
How to import CSV test cases
How to run tests
Known limitations
```

### Acceptance Criteria

- App starts cleanly with empty project.
- App can load sample project.
- App can recover from invalid JSON with backup notice.
- User can export/import project.
- Error logs are written.
- README explains the tester workflow.
- Product is usable without reading code.

---

# Important Implementation Rules

## Keep It Simple

Do not implement:

- cloud sync
- users/roles
- web dashboard
- OCR
- AI features
- database server
- plugin marketplace
- CI/CD integrations
- Jira/TestRail integrations

These can come later.

## Avoid Fragile Rewrites

If the current app is stable, refactor incrementally.

Acceptable:

- move storage logic to `storage.py`
- move runner logic to `runner.py`
- keep UI in one file temporarily

Not acceptable:

- rewrite everything into a complex framework
- introduce unnecessary dependency injection
- introduce a database before needed

## Prefer File-Based Local Project

Use:

```text
project JSON
runs folder
screenshots folder
reports folder
logs folder
```

This is enough for the first product version.

## Maintain Backward Compatibility

Old recorded actions may look like:

```json
{ "type": "click", "x": 100, "y": 200 }
```

Support them as raw coordinate actions or migrate to:

```json
{ "type": "click_xy", "x": 100, "y": 200 }
```

---

# Definition of Done

The product is considered finished for this version when a tester can do this end-to-end:

1. Open the app.
2. Capture named targets.
3. Record a login sequence.
4. Save it as reusable flow `Login_POS`.
5. Create a test case that uses `Login_POS`.
6. Import a CSV dataset with usernames/passwords.
7. Run the test case for each dataset row.
8. See live run logs.
9. Get screenshots on failure.
10. Open an HTML report showing pass/fail evidence.
11. Edit steps and rerun without touching code.

---

# Suggested README Summary

Use this wording in the README:

```text
TestFlow Runner is a lightweight local desktop tool for testers who need to automate repetitive UI workflows across legacy apps, RDP sessions, POS systems, internal web apps, and other environments where code-based automation is not practical.

It records and runs editable test flows based on clicks, keyboard input, waits, reusable flows, variables, datasets, and simple assertions. It stores everything locally and produces evidence reports with screenshots and step logs.
```

---

# Final Note for Codex

Implement this as a practical desktop tool, not a perfect automation framework.

Prioritize:

1. reliability
2. editability
3. clear logs
4. reusable flows
5. CSV-driven tests
6. evidence reports
7. simple polished UI

Avoid scope creep.
