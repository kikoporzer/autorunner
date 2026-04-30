import copy
import json
import shutil
from datetime import datetime
from pathlib import Path
from typing import Any

PROJECT_FILE = Path("testflow_project.json")
LEGACY_FILE = Path("automation_config.json")

DEFAULT_PROJECT: dict[str, Any] = {
    "appVersion": "0.1.0",
    "settings": {
        "startupDelaySeconds": 3,
        "defaultActionPauseSeconds": 0.1,
        "screenshotOnFailure": True,
        "screenshotAfterEachStep": False,
        "stopHotkey": "f8",
        "recordingStopHotkey": "f8",
    },
    "environment": {
        "name": "Default",
        "expectedResolution": "",
        "notes": "",
        "variables": {},
    },
    "targets": {},
    "flows": {},
    "testCases": {},
    "datasets": {},
    "runs": [],
}


def _deepcopy_default() -> dict[str, Any]:
    return copy.deepcopy(DEFAULT_PROJECT)


def _iso_now() -> str:
    return datetime.now().isoformat(timespec="seconds")


def _timestamp_compact() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def _backup_file(src: Path, *, prefix: str = "testflow_project.backup") -> Path:
    backup_path = src.parent / f"{prefix}_{_timestamp_compact()}.json"
    shutil.copy2(src, backup_path)
    return backup_path


def _normalize_target(raw: Any) -> dict[str, Any] | None:
    if not isinstance(raw, dict):
        return None
    if "x" not in raw or "y" not in raw:
        return None
    try:
        x = int(raw["x"])
        y = int(raw["y"])
    except (TypeError, ValueError):
        return None
    return {
        "x": x,
        "y": y,
        "description": str(raw.get("description", "")).strip(),
        "createdAt": str(raw.get("createdAt", _iso_now())),
    }


def _convert_legacy_step(step: dict[str, Any]) -> dict[str, Any]:
    step_type = str(step.get("type", "")).strip()

    if step_type in {"click", "double_click", "right_click"}:
        if "position" in step and step.get("position"):
            return {"type": step_type, "target": str(step["position"])}
        if "x" in step and "y" in step:
            return {"type": "click_xy", "x": int(step["x"]), "y": int(step["y"])}

    if step_type == "move":
        if "position" in step and step.get("position"):
            return {
                "type": "comment",
                "text": f"Legacy move step to target '{step['position']}' was kept as comment.",
            }
        if "x" in step and "y" in step:
            return {"type": "comment", "text": f"Legacy move step to x={step['x']}, y={step['y']} was kept as comment."}

    if step_type == "type":
        return {"type": "type_text", "value": str(step.get("text", ""))}

    if step_type == "press":
        return {"type": "press_key", "key": str(step.get("key", ""))}

    if step_type == "hotkey":
        keys = step.get("keys", [])
        if isinstance(keys, str):
            keys = [k.strip() for k in keys.split("+") if k.strip()]
        if not isinstance(keys, list):
            keys = []
        return {"type": "hotkey", "keys": [str(k) for k in keys]}

    if step_type == "wait":
        return {"type": "wait", "seconds": float(step.get("seconds", 0))}

    if step_type == "screenshot":
        return {"type": "screenshot", "name": str(step.get("name", "screenshot"))}

    return {"type": "comment", "text": f"Unmapped legacy step retained: {step}"}


def _normalize_flow(name: str, raw: Any) -> dict[str, Any]:
    if isinstance(raw, list):
        steps = [_convert_legacy_step(s) for s in raw if isinstance(s, dict)]
        return {
            "name": name,
            "description": "",
            "parameters": [],
            "steps": steps,
        }

    if not isinstance(raw, dict):
        return {
            "name": name,
            "description": "",
            "parameters": [],
            "steps": [],
        }

    parameters = raw.get("parameters", [])
    if not isinstance(parameters, list):
        parameters = []

    steps = raw.get("steps", raw.get("actions", []))
    if not isinstance(steps, list):
        steps = []

    normalized_steps: list[dict[str, Any]] = []
    for step in steps:
        if not isinstance(step, dict):
            continue

        step_type = str(step.get("type", "")).strip()
        is_legacy = (
            "position" in step
            or step_type in {"type", "press", "move"}
            or (step_type == "click" and "target" not in step)
            or (step_type == "double_click" and "target" not in step)
            or (step_type == "right_click" and "target" not in step)
        )
        normalized_steps.append(_convert_legacy_step(step) if is_legacy else step)

    return {
        "name": str(raw.get("name", name)),
        "description": str(raw.get("description", "")),
        "parameters": [str(p) for p in parameters],
        "steps": normalized_steps,
    }


def _normalize_test_case(case_id: str, raw: Any) -> dict[str, Any]:
    if not isinstance(raw, dict):
        raw = {}

    steps = raw.get("steps", [])
    if not isinstance(steps, list):
        steps = []

    return {
        "id": str(raw.get("id", case_id)),
        "name": str(raw.get("name", case_id)),
        "suite": str(raw.get("suite", "")),
        "description": str(raw.get("description", "")),
        "dataset": str(raw.get("dataset", "")),
        "enabled": bool(raw.get("enabled", True)),
        "variables": raw.get("variables", {}) if isinstance(raw.get("variables"), dict) else {},
        "steps": [s for s in steps if isinstance(s, dict)],
    }


def normalize_project_data(data: Any) -> dict[str, Any]:
    project = _deepcopy_default()
    if not isinstance(data, dict):
        return project

    project["appVersion"] = str(data.get("appVersion", project["appVersion"]))

    settings = data.get("settings", {}) if isinstance(data.get("settings"), dict) else {}
    project["settings"].update(
        {
            "startupDelaySeconds": float(settings.get("startupDelaySeconds", settings.get("startup_delay", project["settings"]["startupDelaySeconds"]))),
            "defaultActionPauseSeconds": float(settings.get("defaultActionPauseSeconds", settings.get("default_pause", project["settings"]["defaultActionPauseSeconds"]))),
            "screenshotOnFailure": bool(settings.get("screenshotOnFailure", True)),
            "screenshotAfterEachStep": bool(settings.get("screenshotAfterEachStep", False)),
            "stopHotkey": str(settings.get("stopHotkey", "f8")),
            "recordingStopHotkey": str(settings.get("recordingStopHotkey", "f8")),
        }
    )

    env = data.get("environment", {}) if isinstance(data.get("environment"), dict) else {}
    project["environment"] = {
        "name": str(env.get("name", "Default")),
        "expectedResolution": str(env.get("expectedResolution", "")),
        "notes": str(env.get("notes", "")),
        "variables": env.get("variables", {}) if isinstance(env.get("variables"), dict) else {},
    }

    targets = data.get("targets", {}) if isinstance(data.get("targets"), dict) else {}
    for target_name, raw_target in targets.items():
        normalized = _normalize_target(raw_target)
        if normalized:
            project["targets"][str(target_name)] = normalized

    flows = data.get("flows", {}) if isinstance(data.get("flows"), dict) else {}
    for flow_name, raw_flow in flows.items():
        project["flows"][str(flow_name)] = _normalize_flow(str(flow_name), raw_flow)

    test_cases = data.get("testCases", {}) if isinstance(data.get("testCases"), dict) else {}
    for case_id, raw_case in test_cases.items():
        project["testCases"][str(case_id)] = _normalize_test_case(str(case_id), raw_case)

    datasets = data.get("datasets", {}) if isinstance(data.get("datasets"), dict) else {}
    project["datasets"] = datasets

    runs = data.get("runs", [])
    if isinstance(runs, list):
        project["runs"] = runs

    return project


def migrate_legacy_data(legacy_data: dict[str, Any]) -> dict[str, Any]:
    project = _deepcopy_default()

    legacy_settings = legacy_data.get("settings", {}) if isinstance(legacy_data.get("settings"), dict) else {}
    project["settings"]["startupDelaySeconds"] = float(legacy_settings.get("startup_delay", 3))
    project["settings"]["defaultActionPauseSeconds"] = float(legacy_settings.get("default_pause", 0.1))

    positions = legacy_data.get("positions", {}) if isinstance(legacy_data.get("positions"), dict) else {}
    for name, raw in positions.items():
        normalized = _normalize_target(raw)
        if normalized:
            project["targets"][str(name)] = normalized

    legacy_flows = legacy_data.get("flows", {}) if isinstance(legacy_data.get("flows"), dict) else {}
    for flow_name, raw_steps in legacy_flows.items():
        flow_name = str(flow_name)
        project["flows"][flow_name] = {
            "name": flow_name,
            "description": "Migrated from legacy automation_config.json",
            "parameters": [],
            "steps": [_convert_legacy_step(s) for s in raw_steps if isinstance(s, dict)],
        }

    return project


def load_project(path: Path | None = None) -> tuple[dict[str, Any], list[str]]:
    project_path = path or PROJECT_FILE
    messages: list[str] = []

    if project_path.exists():
        try:
            with project_path.open("r", encoding="utf-8") as f:
                raw = json.load(f)
            project = normalize_project_data(raw)
            return project, messages
        except Exception as exc:
            backup_path = _backup_file(project_path)
            project = _deepcopy_default()
            save_project(project, project_path)
            messages.append(
                f"Project JSON was invalid and was backed up to {backup_path}. A new default project was created."
            )
            messages.append(f"Original load error: {exc}")
            return project, messages

    if LEGACY_FILE.exists():
        legacy_backup = _backup_file(LEGACY_FILE)
        with LEGACY_FILE.open("r", encoding="utf-8") as f:
            legacy = json.load(f)
        project = migrate_legacy_data(legacy)
        save_project(project, project_path)
        messages.append(f"Created backup before migration: {legacy_backup}")
        messages.append(f"Migrated legacy config from {LEGACY_FILE} to {project_path}.")
        return project, messages

    project = _deepcopy_default()
    save_project(project, project_path)
    messages.append(f"Created new project file at {project_path}.")
    return project, messages


def save_project(project: dict[str, Any], path: Path | None = None) -> None:
    project_path = path or PROJECT_FILE
    normalized = normalize_project_data(project)
    with project_path.open("w", encoding="utf-8") as f:
        json.dump(normalized, f, indent=2)
