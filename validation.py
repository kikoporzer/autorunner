import re
from collections import defaultdict
from dataclasses import dataclass
from typing import Any

VAR_PATTERN = re.compile(r"\$\{([A-Za-z_][A-Za-z0-9_]*)\}")

SUPPORTED_STEP_TYPES = {
    "click",
    "click_xy",
    "double_click",
    "right_click",
    "type_text",
    "press_key",
    "hotkey",
    "wait",
    "screenshot",
    "comment",
    "run_flow",
    "assert_window_title_contains",
    "assert_clipboard_contains",
    "assert_file_exists",
}


@dataclass
class ValidationIssue:
    level: str  # error | warning
    message: str


@dataclass
class ValidationResult:
    errors: list[str]
    warnings: list[str]

    @property
    def ok(self) -> bool:
        return len(self.errors) == 0


def extract_variables(value: Any) -> list[str]:
    if value is None:
        return []
    return VAR_PATTERN.findall(str(value))


def resolve_text_variables(value: Any, variables: dict[str, Any]) -> tuple[str, list[str]]:
    unresolved: list[str] = []

    def _replace(match: re.Match[str]) -> str:
        key = match.group(1)
        if key in variables:
            return str(variables[key])
        unresolved.append(key)
        return match.group(0)

    rendered = VAR_PATTERN.sub(_replace, str(value))
    return rendered, unresolved


def detect_circular_flows(project: dict[str, Any]) -> list[str]:
    flows = project.get("flows", {}) if isinstance(project.get("flows"), dict) else {}
    graph: dict[str, list[str]] = defaultdict(list)

    for flow_name, flow in flows.items():
        steps = flow.get("steps", []) if isinstance(flow, dict) else []
        for step in steps:
            if not isinstance(step, dict):
                continue
            if step.get("type") == "run_flow" and step.get("flow"):
                graph[str(flow_name)].append(str(step.get("flow")))

    cycles: list[str] = []
    visiting: set[str] = set()
    visited: set[str] = set()

    def dfs(node: str, path: list[str]) -> None:
        if node in visiting:
            start = path.index(node) if node in path else 0
            cycle_path = path[start:] + [node]
            cycles.append(" -> ".join(cycle_path))
            return
        if node in visited:
            return

        visiting.add(node)
        for nxt in graph.get(node, []):
            dfs(nxt, path + [nxt])
        visiting.remove(node)
        visited.add(node)

    for flow_name in flows.keys():
        dfs(str(flow_name), [str(flow_name)])

    # keep stable unique list
    return sorted(set(cycles))


def _step_field_for_target(step: dict[str, Any]) -> str | None:
    step_type = step.get("type")
    if step_type in {"click", "double_click", "right_click"}:
        return "target"
    return None


def _collect_steps(project: dict[str, Any], run_kind: str, run_name: str) -> tuple[list[dict[str, Any]], list[str]]:
    errors: list[str] = []
    if run_kind == "flow":
        flow = project.get("flows", {}).get(run_name)
        if not flow:
            return [], [f"Flow not found: {run_name}"]
        steps = flow.get("steps", [])
        if not isinstance(steps, list):
            return [], [f"Flow steps are invalid for: {run_name}"]
        return steps, errors

    test_case = project.get("testCases", {}).get(run_name)
    if not test_case:
        return [], [f"Test case not found: {run_name}"]
    steps = test_case.get("steps", [])
    if not isinstance(steps, list):
        return [], [f"Test case steps are invalid for: {run_name}"]
    return steps, errors


def validate_before_run(
    *,
    project: dict[str, Any],
    run_kind: str,
    run_name: str,
    resolved_variables: dict[str, Any],
    expanded_steps: list[dict[str, Any]],
) -> ValidationResult:
    errors: list[str] = []
    warnings: list[str] = []

    if run_kind not in {"flow", "test_case"}:
        return ValidationResult(errors=[f"Invalid run kind: {run_kind}"], warnings=[])

    base_steps, base_errors = _collect_steps(project, run_kind, run_name)
    errors.extend(base_errors)

    if len(base_steps) == 0:
        errors.append(f"{run_kind.replace('_', ' ').title()} '{run_name}' has no steps.")

    targets = project.get("targets", {}) if isinstance(project.get("targets"), dict) else {}
    flows = project.get("flows", {}) if isinstance(project.get("flows"), dict) else {}

    for i, step in enumerate(expanded_steps, start=1):
        if not isinstance(step, dict):
            errors.append(f"Step {i} is not an object.")
            continue

        step_type = step.get("type")
        if step_type not in SUPPORTED_STEP_TYPES:
            errors.append(f"Step {i} has invalid step type: {step_type}")
            continue

        target_field = _step_field_for_target(step)
        if target_field:
            target_name = step.get(target_field)
            if target_name:
                if str(target_name) not in targets:
                    errors.append(f"Step {i} references missing target: {target_name}")
            else:
                # Cross-platform recorder may emit x/y for these actions.
                has_xy = "x" in step and "y" in step
                if not has_xy:
                    errors.append(f"Step {i} is missing target name or x/y coordinates.")
                else:
                    try:
                        int(step.get("x"))
                        int(step.get("y"))
                    except (TypeError, ValueError):
                        errors.append(f"Step {i} has invalid x/y coordinates.")

        if step_type == "click_xy":
            try:
                int(step.get("x"))
                int(step.get("y"))
            except (TypeError, ValueError):
                errors.append(f"Step {i} has invalid coordinates for click_xy.")

        if step_type == "wait":
            try:
                seconds = float(step.get("seconds"))
                if seconds < 0:
                    errors.append(f"Step {i} wait seconds must be >= 0.")
            except (TypeError, ValueError):
                errors.append(f"Step {i} has invalid wait seconds.")

        if step_type == "run_flow":
            flow_name = str(step.get("flow", "")).strip()
            if not flow_name:
                errors.append(f"Step {i} run_flow is missing flow name.")
            elif flow_name not in flows:
                errors.append(f"Step {i} calls missing flow: {flow_name}")

        for key in ("value", "name", "path", "text"):
            if key in step:
                _, unresolved = resolve_text_variables(step.get(key), resolved_variables)
                for var_name in unresolved:
                    errors.append(f"Step {i} uses unresolved variable: ${{{var_name}}}")

    cycles = detect_circular_flows(project)
    if cycles:
        errors.append("Circular flow references detected: " + " | ".join(cycles))

    expected_resolution = str(project.get("environment", {}).get("expectedResolution", "")).strip()
    if expected_resolution:
        warnings.append(f"Expected resolution configured: {expected_resolution}")

    return ValidationResult(errors=sorted(set(errors)), warnings=sorted(set(warnings)))
