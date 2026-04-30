import copy
import os
import platform
import subprocess
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Callable

try:
    import pyautogui
except ModuleNotFoundError:  # pragma: no cover - runtime dependency may be absent
    pyautogui = None

from reporting import ensure_run_folders, generate_html_report, save_run_json
from validation import resolve_text_variables, validate_before_run


class RunnerExecutionError(Exception):
    pass


class TestFlowRunner:
    def __init__(
        self,
        project: dict[str, Any],
        *,
        runs_dir: Path | str = "runs",
        log_callback: Callable[[str], None] | None = None,
    ):
        self.project = project
        self.runs_dir = Path(runs_dir)
        self.log_callback = log_callback
        self.stop_requested = False

    def request_stop(self) -> None:
        self.stop_requested = True

    def reset_stop(self) -> None:
        self.stop_requested = False

    def run_flow(
        self,
        flow_name: str,
        *,
        dry_run: bool = False,
        flow_variables: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        flow = self.project.get("flows", {}).get(flow_name)
        if not flow:
            raise RunnerExecutionError(f"Flow not found: {flow_name}")
        return self._run(
            kind="flow",
            name=flow_name,
            raw_steps=flow.get("steps", []),
            dry_run=dry_run,
            dataset_row_index=None,
            flow_variables=flow_variables or {},
            test_case_variables={},
        )

    def run_test_case(
        self,
        test_case_id: str,
        *,
        dry_run: bool = False,
        dataset_row_index: int | None = None,
        flow_variables: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        test_case = self.project.get("testCases", {}).get(test_case_id)
        if not test_case:
            raise RunnerExecutionError(f"Test case not found: {test_case_id}")
        if not bool(test_case.get("enabled", True)):
            raise RunnerExecutionError(f"Test case is disabled: {test_case_id}")

        dataset_row, resolved_row_index = self._resolve_dataset_row(test_case, dataset_row_index)

        return self._run(
            kind="test_case",
            name=test_case_id,
            raw_steps=test_case.get("steps", []),
            dry_run=dry_run,
            dataset_row_index=resolved_row_index,
            flow_variables=flow_variables or {},
            test_case_variables=test_case.get("variables", {}),
            dataset_row=dataset_row,
        )

    def preview_test_case_execution(
        self,
        test_case_id: str,
        *,
        dataset_row_index: int | None = None,
        flow_variables: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        test_case = self.project.get("testCases", {}).get(test_case_id)
        if not test_case:
            raise RunnerExecutionError(f"Test case not found: {test_case_id}")

        dataset_row, resolved_row_index = self._resolve_dataset_row(test_case, dataset_row_index)

        prepared = self._prepare_execution(
            kind="test_case",
            name=test_case_id,
            raw_steps=test_case.get("steps", []),
            dataset_row=dataset_row,
            test_case_variables=test_case.get("variables", {}),
            flow_variables=flow_variables or {},
        )
        prepared["datasetRowIndex"] = resolved_row_index
        return prepared

    def _log(self, message: str) -> None:
        if self.log_callback:
            self.log_callback(message)

    def _resolve_dataset_row(
        self,
        test_case: dict[str, Any],
        dataset_row_index: int | None,
    ) -> tuple[dict[str, Any], int | None]:
        dataset_name = str(test_case.get("dataset", "")).strip()
        if not dataset_name:
            return {}, None

        dataset = self.project.get("datasets", {}).get(dataset_name)
        if not dataset:
            raise RunnerExecutionError(f"Dataset '{dataset_name}' not found.")

        rows: list[dict[str, Any]]
        if isinstance(dataset, dict):
            rows_raw = dataset.get("rows", [])
            if not isinstance(rows_raw, list):
                raise RunnerExecutionError(f"Dataset '{dataset_name}' rows are invalid.")
            rows = [r for r in rows_raw if isinstance(r, dict)]
        elif isinstance(dataset, list):
            rows = [r for r in dataset if isinstance(r, dict)]
        else:
            raise RunnerExecutionError(f"Dataset '{dataset_name}' format is invalid.")

        if not rows:
            raise RunnerExecutionError(f"Dataset '{dataset_name}' has no rows.")

        idx = 0 if dataset_row_index is None else dataset_row_index
        if idx < 0 or idx >= len(rows):
            raise RunnerExecutionError(f"Dataset row index {idx} is out of range for dataset '{dataset_name}'.")
        return rows[idx], idx

    def _resolve_variables(
        self,
        *,
        run_id: str,
        dataset_row: dict[str, Any],
        test_case_variables: dict[str, Any],
        flow_variables: dict[str, Any],
    ) -> dict[str, Any]:
        now = datetime.now()
        builtin = {
            "today": now.date().isoformat(),
            "now": now.isoformat(timespec="seconds"),
            "timestamp": now.strftime("%Y%m%d_%H%M%S"),
            "runId": run_id,
        }
        env_vars = self.project.get("environment", {}).get("variables", {})
        if not isinstance(env_vars, dict):
            env_vars = {}

        # Priority order (highest last update): dataset row, test case vars, flow vars, env vars, built-ins
        variables: dict[str, Any] = {}
        variables.update(builtin)
        variables.update(env_vars)
        variables.update(flow_variables if isinstance(flow_variables, dict) else {})
        variables.update(test_case_variables if isinstance(test_case_variables, dict) else {})
        variables.update(dataset_row if isinstance(dataset_row, dict) else {})
        return variables

    def _expand_steps(self, steps: list[dict[str, Any]], stack: list[str]) -> list[dict[str, Any]]:
        expanded: list[dict[str, Any]] = []

        for step in steps:
            if not isinstance(step, dict):
                continue
            if step.get("enabled", True) is False:
                continue

            step_type = step.get("type")
            if step_type == "run_flow":
                flow_name = str(step.get("flow", "")).strip()
                if not flow_name:
                    expanded.append(copy.deepcopy(step))
                    continue
                if flow_name in stack:
                    cycle = " -> ".join(stack + [flow_name])
                    raise RunnerExecutionError(f"Circular flow reference: {cycle}")

                flow = self.project.get("flows", {}).get(flow_name)
                if not flow:
                    expanded.append(copy.deepcopy(step))
                    continue

                nested_steps = flow.get("steps", []) if isinstance(flow, dict) else []
                if not isinstance(nested_steps, list):
                    nested_steps = []

                expanded.extend(self._expand_steps(nested_steps, stack + [flow_name]))
            else:
                expanded.append(copy.deepcopy(step))

        return expanded

    def _resolve_step(self, step: dict[str, Any], variables: dict[str, Any]) -> tuple[dict[str, Any], list[str]]:
        resolved = copy.deepcopy(step)
        unresolved: list[str] = []

        for field in ("target", "value", "name", "path", "text", "flow", "key"):
            if field in resolved:
                val, missing = resolve_text_variables(resolved[field], variables)
                resolved[field] = val
                unresolved.extend(missing)

        if "keys" in resolved and isinstance(resolved.get("keys"), list):
            out_keys: list[str] = []
            for key in resolved["keys"]:
                key_text, missing = resolve_text_variables(key, variables)
                out_keys.append(key_text)
                unresolved.extend(missing)
            resolved["keys"] = out_keys

        for numeric_field in ("x", "y", "seconds"):
            if numeric_field in resolved and isinstance(resolved[numeric_field], str):
                rendered, missing = resolve_text_variables(resolved[numeric_field], variables)
                resolved[numeric_field] = rendered
                unresolved.extend(missing)

        return resolved, sorted(set(unresolved))

    def _run(
        self,
        *,
        kind: str,
        name: str,
        raw_steps: list[dict[str, Any]],
        dry_run: bool,
        dataset_row_index: int | None,
        flow_variables: dict[str, Any],
        test_case_variables: dict[str, Any],
        dataset_row: dict[str, Any] | None = None,
    ) -> dict[str, Any]:
        started = datetime.now()
        run_id = f"{started.strftime('%Y%m%d_%H%M%S')}_{kind}_{name}"

        prepared = self._prepare_execution(
            kind=kind,
            name=name,
            raw_steps=raw_steps,
            dataset_row=dataset_row or {},
            test_case_variables=test_case_variables,
            flow_variables=flow_variables,
            run_id=run_id,
        )
        resolved_steps = prepared["executionPlan"]
        variables = prepared["variables"]
        validation = prepared["validation"]
        validation_errors = prepared["validationErrors"]
        if validation_errors:
            raise RunnerExecutionError("Pre-run validation failed:\n- " + "\n- ".join(validation_errors))

        run_root = self.runs_dir / f"{started.strftime('%Y-%m-%d_%H%M%S')}_{name}"
        ensure_run_folders(run_root)

        run_result: dict[str, Any] = {
            "runId": run_id,
            "kind": kind,
            "name": name,
            "startedAt": started.isoformat(timespec="seconds"),
            "endedAt": "",
            "durationSeconds": 0,
            "status": "running",
            "dryRun": dry_run,
            "datasetRowIndex": dataset_row_index,
            "variables": variables,
            "warnings": validation.warnings,
            "stepResults": [],
            "errors": [],
            "artifacts": {
                "runJson": "run.json",
                "reportHtml": "report.html",
                "screenshotsDir": "screenshots",
            },
            "executionPlan": resolved_steps,
            "runFolder": str(run_root),
        }

        self._log(f"[{started.strftime('%H:%M:%S')}] Starting {kind.replace('_', ' ')} {name} (dry_run={dry_run})")

        screenshot_after_each_step = bool(self.project.get("settings", {}).get("screenshotAfterEachStep", False))
        screenshot_on_failure = bool(self.project.get("settings", {}).get("screenshotOnFailure", True))
        post_action_delay = max(0.0, float(self.project.get("settings", {}).get("postActionDelaySeconds", 0.0)))
        default_pause = float(self.project.get("settings", {}).get("defaultActionPauseSeconds", 0.1))
        if pyautogui is not None:
            pyautogui.PAUSE = default_pause

        try:
            for index, step in enumerate(resolved_steps, start=1):
                if self.stop_requested:
                    run_result["status"] = "stopped"
                    run_result["errors"].append("Stop requested by user.")
                    self._log(f"[{datetime.now().strftime('%H:%M:%S')}] Stopped before step {index}")
                    break

                step_result = {
                    "stepIndex": index,
                    "stepType": str(step.get("type", "")),
                    "status": "passed",
                    "message": "",
                    "screenshot": "",
                }

                try:
                    self._execute_step(step, dry_run=dry_run, run_root=run_root, step_index=index)
                    step_result["message"] = "Dry run step validated." if dry_run else "Step executed."
                    if screenshot_after_each_step and not dry_run:
                        step_result["screenshot"] = self._take_step_screenshot(run_root, index, "step")
                    if not dry_run and post_action_delay > 0 and step_result["stepType"] != "wait":
                        time.sleep(post_action_delay)
                    self._log(
                        f"[{datetime.now().strftime('%H:%M:%S')}] Step {index} passed: {step_result['stepType']}"
                    )
                except Exception as exc:
                    step_result["status"] = "failed"
                    step_result["message"] = str(exc)
                    run_result["errors"].append(f"Step {index}: {exc}")
                    self._log(
                        f"[{datetime.now().strftime('%H:%M:%S')}] Step {index} failed: {step_result['stepType']} ({exc})"
                    )
                    if screenshot_on_failure and not dry_run:
                        step_result["screenshot"] = self._take_step_screenshot(run_root, index, "failure")
                    run_result["stepResults"].append(step_result)
                    run_result["status"] = "failed"
                    break

                run_result["stepResults"].append(step_result)

            if run_result["status"] == "running":
                run_result["status"] = "passed"

        except Exception as exc:
            if pyautogui is not None and isinstance(exc, pyautogui.FailSafeException):
                run_result["status"] = "stopped"
                run_result["errors"].append("PyAutoGUI fail-safe triggered.")
                self._log(f"[{datetime.now().strftime('%H:%M:%S')}] Stopped by PyAutoGUI fail-safe")
            else:
                raise

        finally:
            ended = datetime.now()
            run_result["endedAt"] = ended.isoformat(timespec="seconds")
            run_result["durationSeconds"] = round((ended - started).total_seconds(), 3)

            save_run_json(run_root, run_result)
            generate_html_report(run_root, run_result)

            runs = self.project.get("runs", [])
            if not isinstance(runs, list):
                runs = []
            runs.append(
                {
                    "runId": run_result["runId"],
                    "kind": run_result["kind"],
                    "name": run_result["name"],
                    "status": run_result["status"],
                    "startedAt": run_result["startedAt"],
                    "endedAt": run_result["endedAt"],
                    "durationSeconds": run_result["durationSeconds"],
                    "runFolder": run_result["runFolder"],
                }
            )
            self.project["runs"] = runs

        return run_result

    def _prepare_execution(
        self,
        *,
        kind: str,
        name: str,
        raw_steps: list[dict[str, Any]],
        dataset_row: dict[str, Any],
        test_case_variables: dict[str, Any],
        flow_variables: dict[str, Any],
        run_id: str | None = None,
    ) -> dict[str, Any]:
        prep_run_id = run_id or f"preview_{kind}_{name}"
        expanded_steps = self._expand_steps(raw_steps if isinstance(raw_steps, list) else [], [name])
        variables = self._resolve_variables(
            run_id=prep_run_id,
            dataset_row=dataset_row or {},
            test_case_variables=test_case_variables,
            flow_variables=flow_variables,
        )

        resolved_steps: list[dict[str, Any]] = []
        unresolved_errors: list[str] = []
        for i, step in enumerate(expanded_steps, start=1):
            step_resolved, missing = self._resolve_step(step, variables)
            resolved_steps.append(step_resolved)
            for var_name in missing:
                unresolved_errors.append(f"Step {i} uses unresolved variable: ${{{var_name}}}")

        validation = validate_before_run(
            project=self.project,
            run_kind=kind,
            run_name=name,
            resolved_variables=variables,
            expanded_steps=resolved_steps,
        )
        validation_errors = sorted(set(validation.errors + unresolved_errors))

        return {
            "variables": variables,
            "executionPlan": resolved_steps,
            "validation": validation,
            "validationErrors": validation_errors,
        }

    def _target_xy(self, target_name: str) -> tuple[int, int]:
        target = self.project.get("targets", {}).get(target_name)
        if not target:
            raise RunnerExecutionError(f"Missing target: {target_name}")
        try:
            return int(target["x"]), int(target["y"])
        except (TypeError, ValueError, KeyError):
            raise RunnerExecutionError(f"Invalid target coordinates for: {target_name}")

    def _take_step_screenshot(self, run_root: Path, step_index: int, label: str) -> str:
        _require_pyautogui()
        file_name = f"{label}_step_{step_index:03d}.png"
        screenshot_path = run_root / "screenshots" / file_name
        pyautogui.screenshot(str(screenshot_path))
        return f"screenshots/{file_name}"

    def _execute_step(self, step: dict[str, Any], *, dry_run: bool, run_root: Path, step_index: int) -> None:
        step_type = str(step.get("type", "")).strip()

        if step_type == "comment":
            return

        if dry_run:
            return
        _require_pyautogui()

        if step_type == "click":
            x, y = self._xy_from_step(step)
            pyautogui.click(x, y)
            return

        if step_type == "click_xy":
            try:
                x = int(float(step.get("x")))
                y = int(float(step.get("y")))
            except (TypeError, ValueError):
                raise RunnerExecutionError("click_xy requires numeric x and y")
            pyautogui.click(x, y)
            return

        if step_type == "double_click":
            x, y = self._xy_from_step(step)
            pyautogui.doubleClick(x, y)
            return

        if step_type == "right_click":
            x, y = self._xy_from_step(step)
            pyautogui.rightClick(x, y)
            return

        if step_type == "type_text":
            pyautogui.write(str(step.get("value", "")), interval=0.02)
            return

        if step_type == "press_key":
            pyautogui.press(str(step.get("key", "")))
            return

        if step_type == "hotkey":
            keys = step.get("keys", [])
            if not isinstance(keys, list) or not keys:
                raise RunnerExecutionError("hotkey requires a non-empty keys list")
            pyautogui.hotkey(*[str(k) for k in keys])
            return

        if step_type == "wait":
            try:
                seconds = float(step.get("seconds", 0))
            except (TypeError, ValueError):
                raise RunnerExecutionError("wait requires numeric seconds")
            if seconds < 0:
                raise RunnerExecutionError("wait seconds must be >= 0")
            time.sleep(seconds)
            return

        if step_type == "screenshot":
            screenshot_delay = max(
                0.0,
                float(
                    step.get(
                        "delay_before_seconds",
                        self.project.get("settings", {}).get("screenshotDelayBeforeSeconds", 0.0),
                    )
                ),
            )
            if screenshot_delay > 0:
                time.sleep(screenshot_delay)
            custom_name = str(step.get("name", f"step_{step_index:03d}"))
            safe_name = "".join(c for c in custom_name if c.isalnum() or c in ("-", "_")) or f"step_{step_index:03d}"
            out_path = run_root / "screenshots" / f"{safe_name}.png"
            pyautogui.screenshot(str(out_path))
            return

        if step_type == "assert_window_title_contains":
            expected = str(step.get("value", ""))
            actual = _get_active_window_title()
            if expected not in actual:
                raise RunnerExecutionError(f"Window title assertion failed. Expected '{expected}' in '{actual}'.")
            return

        if step_type == "assert_clipboard_contains":
            expected = str(step.get("value", ""))
            clipboard = _read_clipboard()
            if expected not in clipboard:
                raise RunnerExecutionError("Clipboard assertion failed.")
            return

        if step_type == "assert_file_exists":
            file_path = str(step.get("path", "")).strip()
            if not file_path:
                raise RunnerExecutionError("assert_file_exists requires a file path")
            if not Path(file_path).exists():
                raise RunnerExecutionError(f"File does not exist: {file_path}")
            return

        if step_type == "run_flow":
            raise RunnerExecutionError("run_flow steps should be expanded before execution")

        raise RunnerExecutionError(f"Unsupported step type: {step_type}")

    def _xy_from_step(self, step: dict[str, Any]) -> tuple[int, int]:
        target_name = str(step.get("target", "")).strip()
        if target_name:
            return self._target_xy(target_name)
        try:
            return int(float(step.get("x"))), int(float(step.get("y")))
        except (TypeError, ValueError):
            raise RunnerExecutionError("Step requires either target or numeric x/y coordinates.")


def _get_active_window_title() -> str:
    system = platform.system().lower()

    if system == "windows":
        if pyautogui is None:
            raise RunnerExecutionError("PyAutoGUI is required for window title assertion on Windows.")
        title = pyautogui.getActiveWindowTitle()
        return title or ""

    if system == "darwin":
        script = (
            "tell application \"System Events\" to "
            "tell (first application process whose frontmost is true) to "
            "name of front window"
        )
        try:
            output = subprocess.check_output(["osascript", "-e", script], text=True, stderr=subprocess.DEVNULL)
            return output.strip()
        except Exception as exc:
            raise RunnerExecutionError("Unable to read active window title on macOS. Check Accessibility permissions.") from exc

    # linux fallback
    try:
        output = subprocess.check_output(["xdotool", "getactivewindow", "getwindowname"], text=True)
        return output.strip()
    except Exception as exc:
        raise RunnerExecutionError("Unable to read active window title on this OS.") from exc


def _read_clipboard() -> str:
    system = platform.system().lower()

    if system == "darwin":
        try:
            return subprocess.check_output(["pbpaste"], text=True)
        except Exception as exc:
            raise RunnerExecutionError("Unable to read clipboard via pbpaste.") from exc

    if system == "windows":
        try:
            return subprocess.check_output(
                ["powershell", "-NoProfile", "-Command", "Get-Clipboard"], text=True
            )
        except Exception as exc:
            raise RunnerExecutionError("Unable to read clipboard via PowerShell.") from exc

    # linux fallbacks
    for cmd in (["wl-paste"], ["xclip", "-selection", "clipboard", "-o"], ["xsel", "--clipboard", "--output"]):
        try:
            return subprocess.check_output(cmd, text=True)
        except Exception:
            continue

    raise RunnerExecutionError("Unable to read clipboard on this OS.")


def _require_pyautogui() -> None:
    if pyautogui is None:
        raise RunnerExecutionError(
            "pyautogui is not installed. Install it to run live automation (dry-run works without it)."
        )
