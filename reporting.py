import html
import json
from pathlib import Path
from typing import Any


def ensure_run_folders(run_root: Path) -> tuple[Path, Path]:
    screenshots_dir = run_root / "screenshots"
    screenshots_dir.mkdir(parents=True, exist_ok=True)
    return run_root, screenshots_dir


def save_run_json(run_root: Path, run_result: dict[str, Any]) -> Path:
    out_path = run_root / "run.json"
    with out_path.open("w", encoding="utf-8") as f:
        json.dump(run_result, f, indent=2)
    return out_path


def _status_color(status: str) -> str:
    status = status.lower()
    if status == "passed":
        return "#0f7a29"
    if status in {"failed", "error"}:
        return "#b42318"
    if status == "stopped":
        return "#8b6f00"
    return "#475467"


def generate_html_report(run_root: Path, run_result: dict[str, Any]) -> Path:
    report_path = run_root / "report.html"
    status = str(run_result.get("status", "unknown"))
    header_color = _status_color(status)

    rows = []
    for step in run_result.get("stepResults", []):
        step_index = html.escape(str(step.get("stepIndex", "")))
        step_status = html.escape(str(step.get("status", "")))
        step_name = html.escape(str(step.get("stepType", "")))
        message = html.escape(str(step.get("message", "")))

        screenshot_html = ""
        screenshot_rel = step.get("screenshot")
        if screenshot_rel:
            safe_link = html.escape(str(screenshot_rel))
            screenshot_html = f'<a href="{safe_link}" target="_blank">view</a>'

        rows.append(
            "<tr>"
            f"<td>{step_index}</td>"
            f"<td>{step_name}</td>"
            f"<td>{step_status}</td>"
            f"<td>{message}</td>"
            f"<td>{screenshot_html}</td>"
            "</tr>"
        )

    rows_html = "\n".join(rows) if rows else "<tr><td colspan='5'>No steps executed.</td></tr>"

    html_doc = f"""<!DOCTYPE html>
<html lang=\"en\">
<head>
  <meta charset=\"UTF-8\" />
  <title>TestFlow Runner Report</title>
  <style>
    body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Helvetica, Arial, sans-serif; margin: 24px; color: #101828; }}
    .badge {{ display: inline-block; padding: 6px 10px; border-radius: 999px; color: #fff; background: {header_color}; font-weight: 600; }}
    table {{ width: 100%; border-collapse: collapse; margin-top: 16px; }}
    th, td {{ border: 1px solid #eaecf0; padding: 8px; text-align: left; vertical-align: top; }}
    th {{ background: #f9fafb; }}
    .meta {{ margin-top: 12px; color: #344054; }}
  </style>
</head>
<body>
  <h1>TestFlow Runner - Run Report</h1>
  <div class=\"badge\">{html.escape(status.upper())}</div>
  <div class=\"meta\"><b>Run ID:</b> {html.escape(str(run_result.get('runId', '')))}</div>
  <div class=\"meta\"><b>Name:</b> {html.escape(str(run_result.get('name', '')))}</div>
  <div class=\"meta\"><b>Kind:</b> {html.escape(str(run_result.get('kind', '')))}</div>
  <div class=\"meta\"><b>Started:</b> {html.escape(str(run_result.get('startedAt', '')))}</div>
  <div class=\"meta\"><b>Ended:</b> {html.escape(str(run_result.get('endedAt', '')))}</div>
  <div class=\"meta\"><b>Duration (seconds):</b> {html.escape(str(run_result.get('durationSeconds', '')))}</div>
  <div class=\"meta\"><b>Dataset Row:</b> {html.escape(str(run_result.get('datasetRowIndex', '')))}</div>

  <h2>Steps</h2>
  <table>
    <thead>
      <tr>
        <th>#</th>
        <th>Type</th>
        <th>Status</th>
        <th>Message</th>
        <th>Screenshot</th>
      </tr>
    </thead>
    <tbody>
      {rows_html}
    </tbody>
  </table>
</body>
</html>
"""

    with report_path.open("w", encoding="utf-8") as f:
        f.write(html_doc)

    return report_path
