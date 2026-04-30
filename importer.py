import csv
from pathlib import Path
from typing import Any


class ImporterError(Exception):
    pass


def parse_bool(value: Any, default: bool = True) -> bool:
    if value is None:
        return default
    text = str(value).strip().lower()
    if text == "":
        return default
    if text in {"1", "true", "yes", "y", "on"}:
        return True
    if text in {"0", "false", "no", "n", "off"}:
        return False
    return default


def read_table_rows(path: str | Path) -> tuple[list[str], list[dict[str, Any]]]:
    p = Path(path)
    ext = p.suffix.lower()

    if ext == ".csv":
        with p.open("r", encoding="utf-8-sig", newline="") as f:
            reader = csv.DictReader(f)
            headers = [h.strip() for h in (reader.fieldnames or [])]
            rows = []
            for row in reader:
                rows.append({str(k).strip(): ("" if v is None else str(v)) for k, v in row.items() if k is not None})
        return headers, rows

    if ext == ".xlsx":
        try:
            import openpyxl  # type: ignore
        except Exception as exc:
            raise ImporterError("XLSX import requires openpyxl. Use CSV or install openpyxl.") from exc

        wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
        ws = wb.active
        rows_iter = ws.iter_rows(values_only=True)
        try:
            header_row = next(rows_iter)
        except StopIteration:
            return [], []

        headers = ["" if h is None else str(h).strip() for h in header_row]
        out_rows: list[dict[str, Any]] = []
        for raw in rows_iter:
            row_dict: dict[str, Any] = {}
            for idx, cell in enumerate(raw):
                key = headers[idx] if idx < len(headers) else f"col_{idx+1}"
                if not key:
                    key = f"col_{idx+1}"
                row_dict[key] = "" if cell is None else str(cell)
            out_rows.append(row_dict)
        return headers, out_rows

    raise ImporterError(f"Unsupported import file format: {ext}")


def _parse_xy(row: dict[str, Any]) -> tuple[int, int] | None:
    x_raw = str(row.get("X", "")).strip() or str(row.get("x", "")).strip()
    y_raw = str(row.get("Y", "")).strip() or str(row.get("y", "")).strip()
    if x_raw and y_raw:
        try:
            return int(float(x_raw)), int(float(y_raw))
        except Exception:
            return None

    value = str(row.get("Value", "")).strip()
    if "," in value:
        parts = [p.strip() for p in value.split(",", 1)]
        if len(parts) == 2:
            try:
                return int(float(parts[0])), int(float(parts[1]))
            except Exception:
                return None

    target = str(row.get("Target", "")).strip()
    if "," in target:
        parts = [p.strip() for p in target.split(",", 1)]
        if len(parts) == 2:
            try:
                return int(float(parts[0])), int(float(parts[1]))
            except Exception:
                return None

    return None


def _build_step_from_row(row: dict[str, Any], errors: list[str], row_no: int) -> dict[str, Any] | None:
    action = str(row.get("ActionType", "")).strip().lower()
    target = str(row.get("Target", "")).strip()
    value = str(row.get("Value", ""))
    seconds_raw = str(row.get("Seconds", "")).strip()
    description = str(row.get("Description", "")).strip()
    enabled = parse_bool(row.get("Enabled"), default=True)

    if not action:
        errors.append(f"Row {row_no}: missing ActionType")
        return None

    step: dict[str, Any] = {"type": action, "enabled": enabled}

    if action in {"click", "double_click", "right_click"}:
        if target:
            step["target"] = target
        else:
            xy = _parse_xy(row)
            if xy is None:
                errors.append(f"Row {row_no}: {action} requires Target or x/y coordinates")
                return None
            step["x"], step["y"] = xy

    elif action == "click_xy":
        xy = _parse_xy(row)
        if xy is None:
            errors.append(f"Row {row_no}: click_xy requires x/y coordinates in X/Y or Value")
            return None
        step["x"], step["y"] = xy

    elif action in {"type_text", "type"}:
        step["type"] = "type_text"
        step["value"] = value

    elif action in {"press_key", "press"}:
        step["type"] = "press_key"
        key_text = value.strip() or target
        if not key_text:
            errors.append(f"Row {row_no}: press_key requires Value or Target as key")
            return None
        step["key"] = key_text

    elif action == "hotkey":
        raw = value.strip() or target
        keys = [k.strip() for k in raw.split("+") if k.strip()]
        if not keys:
            errors.append(f"Row {row_no}: hotkey requires Value or Target like ctrl+s")
            return None
        step["keys"] = keys

    elif action == "wait":
        raw = seconds_raw or value.strip()
        if not raw:
            errors.append(f"Row {row_no}: wait requires Seconds or Value")
            return None
        try:
            step["seconds"] = float(raw)
        except Exception:
            errors.append(f"Row {row_no}: invalid wait seconds '{raw}'")
            return None

    elif action == "screenshot":
        step["name"] = value.strip() or "screenshot"

    elif action == "comment":
        step["text"] = description or value

    elif action == "run_flow":
        flow_name = target or value.strip()
        if not flow_name:
            errors.append(f"Row {row_no}: run_flow requires Target or Value as flow name")
            return None
        step["flow"] = flow_name

    elif action in {"assert_window_title_contains", "assert_clipboard_contains"}:
        if not value.strip():
            errors.append(f"Row {row_no}: {action} requires Value")
            return None
        step["value"] = value

    elif action == "assert_file_exists":
        if not value.strip():
            errors.append(f"Row {row_no}: assert_file_exists requires Value path")
            return None
        step["path"] = value.strip()

    else:
        errors.append(f"Row {row_no}: unsupported ActionType '{action}'")
        return None

    if description and action != "comment":
        step["description"] = description

    return step


def parse_test_case_rows(rows: list[dict[str, Any]]) -> tuple[dict[str, dict[str, Any]], list[str], list[dict[str, Any]]]:
    required = {"TestCaseId", "ActionType"}
    if not rows:
        return {}, [], []

    normalized_rows = []
    for r in rows:
        normalized_rows.append({str(k).strip(): v for k, v in r.items()})

    keys = set()
    for r in normalized_rows:
        keys.update(r.keys())

    missing_cols = sorted(required - keys)
    if missing_cols:
        return {}, [f"Missing required columns: {', '.join(missing_cols)}"], []

    grouped: dict[str, dict[str, Any]] = {}
    errors: list[str] = []
    invalid_rows: list[dict[str, Any]] = []

    for idx, row in enumerate(normalized_rows, start=2):
        case_id = str(row.get("TestCaseId", "")).strip()
        if not case_id:
            errors.append(f"Row {idx}: missing TestCaseId")
            invalid_rows.append({"row": idx, "reason": "missing TestCaseId"})
            continue

        case = grouped.setdefault(
            case_id,
            {
                "id": case_id,
                "name": str(row.get("TestCaseName", case_id)).strip() or case_id,
                "suite": str(row.get("Suite", "")).strip(),
                "description": "",
                "dataset": str(row.get("Dataset", "")).strip(),
                "enabled": parse_bool(row.get("CaseEnabled"), default=True),
                "variables": {},
                "steps": [],
                "_order": [],
            },
        )

        step = _build_step_from_row(row, errors, idx)
        if step is None:
            invalid_rows.append({"row": idx, "reason": "invalid step"})
            continue

        step_no = str(row.get("StepNo", "")).strip()
        try:
            order_key = int(step_no) if step_no else len(case["steps"]) + 1
        except Exception:
            order_key = len(case["steps"]) + 1

        case["steps"].append(step)
        case["_order"].append(order_key)

    # Sort each test case by step number while preserving association.
    for case in grouped.values():
        paired = list(zip(case["_order"], case["steps"]))
        paired.sort(key=lambda x: x[0])
        case["steps"] = [s for _, s in paired]
        case.pop("_order", None)

    return grouped, errors, invalid_rows
