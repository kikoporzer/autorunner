import copy
import json
import math
import os
import shutil
import sys
import threading
import time
import traceback
import tkinter as tk
import webbrowser
import zipfile
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk

try:
    import pyautogui
except ModuleNotFoundError:
    pyautogui = None

from recorder import GlobalClickRecorder
from runner import RunnerExecutionError, TestFlowRunner
from importer import ImporterError, parse_test_case_rows, read_table_rows
from storage import PROJECT_FILE, load_project, normalize_project_data, save_project as storage_save_project
from ui_components import RunLogPanel, StepInspector, StepTable


class AutomationApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("TestFlow Runner")
        self.root.geometry("1400x860")
        self.root.minsize(1150, 760)
        self.project_path = PROJECT_FILE.resolve()
        self.logs_dir = Path("logs")
        self.logs_dir.mkdir(parents=True, exist_ok=True)

        self.data, messages = load_project()

        self.current_page = "Dashboard"
        self.current_entity_kind: str | None = None  # flow | test_case
        self.current_entity_name: str | None = None
        self.current_step_index: int | None = None
        self.active_runner: TestFlowRunner | None = None

        self.recorder = GlobalClickRecorder()
        self.recorded_preview_steps: list[dict] = []
        self.recording_mode_active = False
        self._recording_window_bounds: tuple[int, int, int, int] | None = None
        self.last_saved_at: datetime | None = None

        if pyautogui is not None:
            pyautogui.FAILSAFE = True
            pyautogui.PAUSE = float(self.data["settings"].get("defaultActionPauseSeconds", 0.1))

        self._build_shell()
        self._setup_exception_handler()
        self._bind_shortcuts()
        self.show_page("Dashboard")
        self._run_startup_checks()

        for msg in messages:
            self.append_log(msg)

    # =====================================================
    # Shell
    # =====================================================

    def _build_shell(self) -> None:
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(1, weight=1)

        self.top_bar = ttk.Frame(self.root, padding=(12, 8))
        self.top_bar.grid(row=0, column=0, sticky="ew")
        self.top_bar.columnconfigure(1, weight=1)

        self.title_var = tk.StringVar(value="TestFlow Runner")
        self.page_var = tk.StringVar(value="Dashboard")
        self.status_var = tk.StringVar(value="Idle")
        self.recording_badge_var = tk.StringVar(value="")
        self.project_var = tk.StringVar(value=f"Project: {self.project_path}")
        self.last_saved_var = tk.StringVar(value="Last saved: not yet")
        env_name = self.data.get("environment", {}).get("name", "Default")

        ttk.Label(self.top_bar, textvariable=self.title_var, font=("Segoe UI", 14, "bold")).grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(self.top_bar, text=f"Environment: {env_name}", foreground="#475467").grid(
            row=0, column=1, sticky="w", padx=12
        )
        ttk.Label(self.top_bar, textvariable=self.status_var, foreground="#1d2939").grid(
            row=0, column=2, sticky="e", padx=8
        )
        ttk.Label(self.top_bar, textvariable=self.recording_badge_var, foreground="#b42318").grid(
            row=0, column=3, sticky="e", padx=8
        )
        ttk.Button(self.top_bar, text="Stop", command=self.stop_run).grid(row=0, column=4, sticky="e", padx=4)
        ttk.Button(self.top_bar, text="Save", command=self.save_project).grid(row=0, column=5, sticky="e", padx=4)
        ttk.Label(self.top_bar, textvariable=self.project_var, foreground="#475467").grid(
            row=1, column=0, columnspan=2, sticky="w", pady=(2, 0)
        )
        ttk.Label(self.top_bar, textvariable=self.page_var, foreground="#475467").grid(
            row=1, column=2, sticky="e", pady=(2, 0)
        )
        ttk.Label(self.top_bar, textvariable=self.last_saved_var, foreground="#475467").grid(
            row=1, column=3, columnspan=3, sticky="e", pady=(2, 0)
        )

        self.center = ttk.PanedWindow(self.root, orient="horizontal")
        self.center.grid(row=1, column=0, sticky="nsew")

        self.nav_frame = ttk.Frame(self.center, padding=10)
        self.main_frame = ttk.Frame(self.center, padding=10)
        self.inspector_frame = ttk.Frame(self.center, padding=10)

        self.center.add(self.nav_frame, weight=1)
        self.center.add(self.main_frame, weight=4)
        self.center.add(self.inspector_frame, weight=2)

        self._build_nav()
        self._build_inspector()

        self.log_panel = RunLogPanel(self.root)
        self.log_panel.grid(row=2, column=0, sticky="nsew")

    def _build_nav(self) -> None:
        ttk.Label(self.nav_frame, text="Navigation", font=("Segoe UI", 11, "bold")).pack(anchor="w", pady=(0, 8))

        for page in [
            "Dashboard",
            "Test Cases",
            "Reusable Flows",
            "Targets",
            "Recorder",
            "Datasets",
            "Run Center",
            "Settings",
        ]:
            ttk.Button(self.nav_frame, text=page, width=20, command=lambda p=page: self.show_page(p)).pack(
                anchor="w", pady=3
            )

    def _build_inspector(self) -> None:
        self.inspector = StepInspector(self.inspector_frame, on_apply=self.apply_step_from_inspector)
        self.inspector.pack(fill="both", expand=True)

    def _clear_main(self) -> None:
        for child in self.main_frame.winfo_children():
            child.destroy()

    def append_log(self, message: str) -> None:
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.root.after(0, lambda: self.log_panel.append(f"[{timestamp}] {message}"))

    def save_project(self) -> None:
        storage_save_project(self.data)
        self.last_saved_at = datetime.now()
        self.last_saved_var.set(f"Last saved: {self.last_saved_at.strftime('%Y-%m-%d %H:%M:%S')}")
        self.append_log("Project saved.")

    def set_status(self, text: str) -> None:
        self.root.after(0, lambda: self.status_var.set(text))

    def set_recording_badge(self, text: str) -> None:
        self.root.after(0, lambda: self.recording_badge_var.set(text))

    def _run_startup_checks(self) -> None:
        checks: list[str] = []
        warnings: list[str] = []

        if sys.version_info < (3, 10):
            warnings.append(f"Python version is {sys.version.split()[0]}. Recommended: 3.10+")
        else:
            checks.append(f"Python {sys.version.split()[0]} detected")

        if pyautogui is None:
            warnings.append("pyautogui not installed (live execution disabled, dry-run still works).")
        else:
            checks.append("pyautogui available")

        try:
            import keyboard as _kbd  # noqa: F401
            checks.append("keyboard package available")
        except Exception:
            warnings.append("keyboard package not installed (optional; recorder uses pynput).")

        recorder_status = self.recorder.availability()
        if recorder_status.available:
            checks.append("Recorder backend available (pynput)")
        else:
            warnings.append(recorder_status.message)

        if os.access(self.project_path.parent, os.W_OK):
            checks.append(f"Write access OK: {self.project_path.parent}")
        else:
            warnings.append(f"No write access to project folder: {self.project_path.parent}")

        expected = str(self.data.get("environment", {}).get("expectedResolution", "")).strip()
        screen = f"{self.root.winfo_screenwidth()}x{self.root.winfo_screenheight()}"
        checks.append(f"Screen resolution: {screen}")
        if expected and expected != screen:
            warnings.append(f"Expected resolution '{expected}' differs from current '{screen}'")

        checks.append(f"Project loaded: {self.project_path}")
        for line in checks:
            self.append_log(f"Startup check: {line}")
        for line in warnings:
            self.append_log(f"Startup warning: {line}")

    def _setup_exception_handler(self) -> None:
        def _handle(exc, val, tb):
            self._log_exception("Tk callback exception", (exc, val, tb))
            messagebox.showerror("Application Error", f"{val}\n\nDetails were saved to logs/app_errors.log")

        self.root.report_callback_exception = _handle

    def _log_exception(self, context: str, exc_info=None) -> None:
        if exc_info is None:
            exc_info = sys.exc_info()
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        path = self.logs_dir / "app_errors.log"
        with path.open("a", encoding="utf-8") as f:
            f.write(f"[{ts}] {context}\n")
            f.write("".join(traceback.format_exception(*exc_info)))
            f.write("\n")

    def _bind_shortcuts(self) -> None:
        self.root.bind("<Control-s>", lambda e: self.save_project())
        self.root.bind("<Command-s>", lambda e: self.save_project())
        self.root.bind("<F5>", lambda e: self._shortcut_run(False))
        self.root.bind("<Shift-F5>", lambda e: self._shortcut_run(True))
        self.root.bind("<F8>", lambda e: self.stop_run())
        self.root.bind("<Control-d>", lambda e: self._shortcut_duplicate_step())
        self.root.bind("<Command-d>", lambda e: self._shortcut_duplicate_step())
        self.root.bind("<Delete>", lambda e: self._shortcut_delete_step())

    def _shortcut_run(self, dry_run: bool) -> None:
        if self.current_entity_kind in {"flow", "test_case"}:
            self.run_selected(self.current_entity_kind, dry_run)

    def _shortcut_duplicate_step(self) -> None:
        if self.current_entity_kind in {"flow", "test_case"}:
            self._duplicate_step()

    def _shortcut_delete_step(self) -> None:
        if self.current_entity_kind in {"flow", "test_case"}:
            self._delete_step()

    # =====================================================
    # Page router
    # =====================================================

    def show_page(self, page: str) -> None:
        self.current_page = page
        self.title_var.set(f"TestFlow Runner - {page}")
        self.page_var.set(f"Page: {page}")
        self._clear_main()

        if page == "Dashboard":
            self._build_dashboard_page()
        elif page == "Test Cases":
            self._build_test_cases_page()
        elif page == "Reusable Flows":
            self._build_flows_page()
        elif page == "Targets":
            self._build_targets_page()
        elif page == "Recorder":
            self._build_recorder_page()
        elif page == "Datasets":
            self._build_datasets_page()
        elif page == "Run Center":
            self._build_run_center_page()
        elif page == "Settings":
            self._build_settings_page()

        self.inspector.load_step(None)
        self.current_entity_kind = None
        self.current_entity_name = None
        self.current_step_index = None

    # =====================================================
    # Dashboard
    # =====================================================

    def _build_dashboard_page(self) -> None:
        frame = ttk.Frame(self.main_frame)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="Dashboard", font=("Segoe UI", 16, "bold")).pack(anchor="w")

        flows = len(self.data.get("flows", {}))
        test_cases = len(self.data.get("testCases", {}))
        targets = len(self.data.get("targets", {}))
        runs = len(self.data.get("runs", []))

        cards = ttk.Frame(frame)
        cards.pack(fill="x", pady=10)
        for text in [f"Flows: {flows}", f"Test Cases: {test_cases}", f"Targets: {targets}", f"Runs: {runs}"]:
            card = ttk.LabelFrame(cards, text=text, padding=10)
            card.pack(side="left", padx=6, fill="x", expand=True)
            ttk.Label(card, text="Ready").pack(anchor="w")

        recent = ttk.LabelFrame(frame, text="Recent Runs", padding=10)
        recent.pack(fill="both", expand=True, pady=10)

        cols = ("started", "kind", "name", "status", "duration")
        tree = ttk.Treeview(recent, columns=cols, show="headings")
        for col, text in zip(cols, ["Started", "Kind", "Name", "Status", "Duration"]):
            tree.heading(col, text=text)
        tree.column("started", width=180)
        tree.column("kind", width=100)
        tree.column("name", width=220)
        tree.column("status", width=100)
        tree.column("duration", width=100)
        tree.pack(fill="both", expand=True)

        for run in list(self.data.get("runs", []))[-50:][::-1]:
            tree.insert(
                "",
                tk.END,
                values=(
                    run.get("startedAt", ""),
                    run.get("kind", ""),
                    run.get("name", ""),
                    run.get("status", ""),
                    run.get("durationSeconds", ""),
                ),
            )

    # =====================================================
    # Shared entity helpers
    # =====================================================

    def _get_entity_map(self, kind: str) -> dict:
        return self.data["flows"] if kind == "flow" else self.data["testCases"]

    def _get_steps(self, kind: str, name: str) -> list[dict]:
        entity = self._get_entity_map(kind).get(name, {})
        steps = entity.get("steps", [])
        return steps if isinstance(steps, list) else []

    def _set_active_context(self, kind: str, name: str) -> None:
        self.current_entity_kind = kind
        self.current_entity_name = name
        self.current_step_index = None
        self.inspector.load_step(None)

    def _refresh_active_step_table(self) -> None:
        if not hasattr(self, "active_step_table"):
            return
        if not self.current_entity_kind or not self.current_entity_name:
            self.active_step_table.set_steps([])
            return
        steps = self._get_steps(self.current_entity_kind, self.current_entity_name)
        self.active_step_table.set_steps(steps)

    def _handle_step_selection(self, index: int | None) -> None:
        self.current_step_index = index
        if index is None or not self.current_entity_kind or not self.current_entity_name:
            self.inspector.load_step(None)
            return
        steps = self._get_steps(self.current_entity_kind, self.current_entity_name)
        if 0 <= index < len(steps):
            self.inspector.load_step(steps[index])
        else:
            self.inspector.load_step(None)

    def apply_step_from_inspector(self, updated_step: dict) -> None:
        if self.current_entity_kind is None or self.current_entity_name is None:
            return
        if self.current_step_index is None:
            return

        steps = self._get_steps(self.current_entity_kind, self.current_entity_name)
        if not (0 <= self.current_step_index < len(steps)):
            return

        steps[self.current_step_index] = updated_step
        self.save_project()
        self._refresh_active_step_table()
        self._handle_step_selection(self.current_step_index)

    def _add_step(self) -> None:
        if self.current_entity_kind is None or self.current_entity_name is None:
            return
        steps = self._get_steps(self.current_entity_kind, self.current_entity_name)
        steps.append({"type": "comment", "text": "new step", "enabled": True})
        self.save_project()
        self._refresh_active_step_table()

    def _delete_step(self) -> None:
        if self.current_entity_kind is None or self.current_entity_name is None:
            return
        if self.current_step_index is None:
            return
        steps = self._get_steps(self.current_entity_kind, self.current_entity_name)
        if 0 <= self.current_step_index < len(steps):
            steps.pop(self.current_step_index)
            self.current_step_index = None
            self.inspector.load_step(None)
            self.save_project()
            self._refresh_active_step_table()

    def _duplicate_step(self) -> None:
        if self.current_entity_kind is None or self.current_entity_name is None:
            return
        if self.current_step_index is None:
            return
        steps = self._get_steps(self.current_entity_kind, self.current_entity_name)
        if 0 <= self.current_step_index < len(steps):
            steps.insert(self.current_step_index + 1, copy.deepcopy(steps[self.current_step_index]))
            self.save_project()
            self._refresh_active_step_table()

    def _move_step(self, direction: int) -> None:
        if self.current_entity_kind is None or self.current_entity_name is None:
            return
        if self.current_step_index is None:
            return
        steps = self._get_steps(self.current_entity_kind, self.current_entity_name)
        i = self.current_step_index
        j = i + direction
        if 0 <= i < len(steps) and 0 <= j < len(steps):
            steps[i], steps[j] = steps[j], steps[i]
            self.current_step_index = j
            self.save_project()
            self._refresh_active_step_table()
            self._handle_step_selection(j)

    # =====================================================
    # Test Cases Page
    # =====================================================

    def _build_test_cases_page(self) -> None:
        container = ttk.Frame(self.main_frame)
        container.pack(fill="both", expand=True)
        container.columnconfigure(1, weight=1)
        container.rowconfigure(1, weight=1)

        ttk.Label(container, text="Test Cases", font=("Segoe UI", 16, "bold")).grid(row=0, column=0, sticky="w")

        header_btns = ttk.Frame(container)
        header_btns.grid(row=0, column=1, sticky="e")
        ttk.Button(header_btns, text="Run", command=lambda: self.run_selected("test_case", False)).pack(side="left", padx=3)
        ttk.Button(header_btns, text="Run All Rows", command=self.run_selected_test_case_all_rows).pack(side="left", padx=3)
        ttk.Button(header_btns, text="Dry Run", command=lambda: self.run_selected("test_case", True)).pack(side="left", padx=3)
        ttk.Button(header_btns, text="Import", command=self.import_test_cases_file).pack(side="left", padx=3)
        ttk.Button(header_btns, text="New", command=self.new_test_case).pack(side="left", padx=3)
        ttk.Button(header_btns, text="Duplicate", command=self.duplicate_test_case).pack(side="left", padx=3)
        ttk.Button(header_btns, text="Delete", command=self.delete_test_case).pack(side="left", padx=3)

        body = ttk.PanedWindow(container, orient="horizontal")
        body.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(10, 0))

        left = ttk.Frame(body, padding=8)
        right = ttk.Frame(body, padding=8)
        body.add(left, weight=1)
        body.add(right, weight=3)
        right.columnconfigure(0, weight=1)
        right.rowconfigure(2, weight=1)

        ttk.Label(left, text="Test Case List", font=("Segoe UI", 11, "bold")).pack(anchor="w")
        self.tc_listbox = tk.Listbox(left, height=24)
        self.tc_listbox.pack(fill="both", expand=True, pady=8)
        self.tc_listbox.bind("<<ListboxSelect>>", self._on_test_case_select)

        meta = ttk.LabelFrame(right, text="Metadata", padding=8)
        meta.grid(row=0, column=0, sticky="ew")
        for i in range(8):
            meta.columnconfigure(i, weight=1)

        self.tc_id_var = tk.StringVar()
        self.tc_name_var = tk.StringVar()
        self.tc_suite_var = tk.StringVar()
        self.tc_dataset_var = tk.StringVar()
        self.tc_enabled_var = tk.BooleanVar(value=True)

        ttk.Label(meta, text="ID").grid(row=0, column=0, sticky="w")
        ttk.Entry(meta, textvariable=self.tc_id_var, state="disabled").grid(row=0, column=1, sticky="ew", padx=4)
        ttk.Label(meta, text="Name").grid(row=0, column=2, sticky="w")
        ttk.Entry(meta, textvariable=self.tc_name_var).grid(row=0, column=3, sticky="ew", padx=4)
        ttk.Label(meta, text="Suite").grid(row=0, column=4, sticky="w")
        ttk.Entry(meta, textvariable=self.tc_suite_var).grid(row=0, column=5, sticky="ew", padx=4)
        ttk.Label(meta, text="Dataset").grid(row=0, column=6, sticky="w")
        ttk.Entry(meta, textvariable=self.tc_dataset_var).grid(row=0, column=7, sticky="ew", padx=4)
        ttk.Checkbutton(meta, text="Enabled", variable=self.tc_enabled_var).grid(row=1, column=0, sticky="w", pady=6)
        ttk.Button(meta, text="Preview Variables", command=self.preview_selected_test_case_variables).grid(
            row=1, column=6, sticky="e"
        )
        ttk.Button(meta, text="Save Metadata", command=self.save_test_case_metadata).grid(row=1, column=7, sticky="e")

        self.tc_var_preview = tk.Text(meta, height=8)
        self.tc_var_preview.grid(row=2, column=0, columnspan=8, sticky="ew", pady=(6, 0))

        step_btns = ttk.Frame(right)
        step_btns.grid(row=1, column=0, sticky="ew", pady=(8, 4))
        ttk.Button(step_btns, text="Add Step", command=self._add_step).pack(side="left", padx=2)
        ttk.Button(step_btns, text="Duplicate Step", command=self._duplicate_step).pack(side="left", padx=2)
        ttk.Button(step_btns, text="Delete Step", command=self._delete_step).pack(side="left", padx=2)
        ttk.Button(step_btns, text="Up", command=lambda: self._move_step(-1)).pack(side="left", padx=2)
        ttk.Button(step_btns, text="Down", command=lambda: self._move_step(1)).pack(side="left", padx=2)

        self.active_step_table = StepTable(right, on_select=self._handle_step_selection)
        self.active_step_table.grid(row=2, column=0, sticky="nsew")

        self.refresh_test_case_list()

    def refresh_test_case_list(self) -> None:
        if not hasattr(self, "tc_listbox"):
            return
        self.tc_listbox.delete(0, tk.END)
        for case_id, tc in sorted(self.data["testCases"].items()):
            self.tc_listbox.insert(tk.END, f"{case_id} - {tc.get('name', case_id)}")

    def _on_test_case_select(self, _event=None) -> None:
        if not self.tc_listbox.curselection():
            return
        raw = self.tc_listbox.get(self.tc_listbox.curselection()[0])
        case_id = raw.split(" - ", 1)[0]

        self._set_active_context("test_case", case_id)
        tc = self.data["testCases"][case_id]
        self.tc_id_var.set(tc.get("id", case_id))
        self.tc_name_var.set(tc.get("name", ""))
        self.tc_suite_var.set(tc.get("suite", ""))
        self.tc_dataset_var.set(tc.get("dataset", ""))
        self.tc_enabled_var.set(bool(tc.get("enabled", True)))
        self._refresh_active_step_table()

    def new_test_case(self) -> None:
        case_id = simpledialog.askstring("New Test Case", "Test case ID (e.g. POS001):")
        if not case_id:
            return
        case_id = case_id.strip()
        if not case_id:
            return
        if case_id in self.data["testCases"]:
            messagebox.showerror("Error", "Test case already exists.")
            return

        self.data["testCases"][case_id] = {
            "id": case_id,
            "name": case_id,
            "suite": "",
            "description": "",
            "dataset": "",
            "enabled": True,
            "variables": {},
            "steps": [],
        }
        self.save_project()
        self.refresh_test_case_list()

    def duplicate_test_case(self) -> None:
        if self.current_entity_kind != "test_case" or not self.current_entity_name:
            return
        src = self.current_entity_name
        dst = simpledialog.askstring("Duplicate Test Case", "New test case ID:", initialvalue=f"{src}_COPY")
        if not dst:
            return
        dst = dst.strip()
        if not dst or dst in self.data["testCases"]:
            messagebox.showerror("Error", "Invalid or existing test case ID.")
            return
        duplicated = copy.deepcopy(self.data["testCases"][src])
        duplicated["id"] = dst
        duplicated["name"] = duplicated.get("name", dst) + " (Copy)"
        self.data["testCases"][dst] = duplicated
        self.save_project()
        self.refresh_test_case_list()

    def delete_test_case(self) -> None:
        if self.current_entity_kind != "test_case" or not self.current_entity_name:
            return
        case_id = self.current_entity_name
        if not messagebox.askyesno("Delete", f"Delete test case '{case_id}'?"):
            return
        self.data["testCases"].pop(case_id, None)
        self.current_entity_name = None
        self.save_project()
        self.refresh_test_case_list()
        self._refresh_active_step_table()

    def save_test_case_metadata(self) -> None:
        if self.current_entity_kind != "test_case" or not self.current_entity_name:
            return
        tc = self.data["testCases"][self.current_entity_name]
        tc["name"] = self.tc_name_var.get().strip()
        tc["suite"] = self.tc_suite_var.get().strip()
        tc["dataset"] = self.tc_dataset_var.get().strip()
        tc["enabled"] = bool(self.tc_enabled_var.get())
        self.save_project()
        self.refresh_test_case_list()

    def import_test_cases_file(self) -> None:
        path = filedialog.askopenfilename(
            title="Import test cases",
            filetypes=[("CSV/XLSX", "*.csv *.xlsx"), ("CSV", "*.csv"), ("Excel", "*.xlsx"), ("All", "*.*")],
        )
        if not path:
            return

        try:
            _, rows = read_table_rows(path)
            grouped, errors, invalid_rows = parse_test_case_rows(rows)
        except ImporterError as exc:
            messagebox.showerror("Import Error", str(exc))
            return
        except Exception as exc:
            self._log_exception("Import test cases failed")
            messagebox.showerror("Import Error", f"Failed to parse import file: {exc}")
            return

        if not grouped:
            extra = ""
            if errors:
                extra = "\n\n" + "\n".join(errors[:10])
            messagebox.showerror("Import Error", "No valid test cases were found." + extra)
            return

        summary_lines = [
            f"Rows found: {len(rows)}",
            f"Test cases found: {len(grouped)}",
            f"Invalid rows: {len(invalid_rows)}",
            "",
            "Preview:",
        ]
        for case_id, case in list(sorted(grouped.items()))[:20]:
            summary_lines.append(f"- {case_id}: {len(case.get('steps', []))} steps")

        if invalid_rows:
            summary_lines.append("")
            summary_lines.append("Invalid row details (first 10):")
            for item in invalid_rows[:10]:
                summary_lines.append(f"- Row {item.get('row')}: {item.get('reason')}")
        if errors:
            summary_lines.append("")
            summary_lines.append("Validation messages (first 15):")
            for err in errors[:15]:
                summary_lines.append(f"- {err}")

        summary_text = "\n".join(summary_lines)
        proceed = messagebox.askyesno("Import Test Cases", summary_text + "\n\nImport now?")
        if not proceed:
            return

        replaced = 0
        created = 0
        for case_id, case_obj in grouped.items():
            if case_id in self.data["testCases"]:
                replaced += 1
            else:
                created += 1
            self.data["testCases"][case_id] = case_obj

        self.save_project()
        self.refresh_test_case_list()
        self.append_log(f"Imported test cases from {Path(path).name}: created={created}, replaced={replaced}")
        messagebox.showinfo(
            "Import Complete",
            f"Imported {len(grouped)} test case(s).\nCreated: {created}\nReplaced: {replaced}\nInvalid rows: {len(invalid_rows)}",
        )

    def run_selected_test_case_all_rows(self) -> None:
        if self.current_entity_kind != "test_case" or not self.current_entity_name:
            messagebox.showwarning("No Test Case", "Select a test case first.")
            return
        case_id = self.current_entity_name
        test_case = self.data["testCases"].get(case_id, {})
        dataset_name = str(test_case.get("dataset", "")).strip()
        if not dataset_name:
            messagebox.showwarning("No Dataset", "Selected test case does not have a dataset configured.")
            return

        dataset = self.data.get("datasets", {}).get(dataset_name)
        if isinstance(dataset, dict):
            rows = dataset.get("rows", []) if isinstance(dataset.get("rows"), list) else []
        elif isinstance(dataset, list):
            rows = dataset
        else:
            rows = []

        if not rows:
            messagebox.showwarning("Empty Dataset", f"Dataset '{dataset_name}' has no rows.")
            return

        delay = float(self.data.get("settings", {}).get("startupDelaySeconds", 3))
        if pyautogui is None:
            messagebox.showerror(
                "Missing Dependency",
                "pyautogui is not installed. Install pyautogui for live execution.",
            )
            return

        proceed = messagebox.askyesno(
            "Run All Dataset Rows",
            f"Run test case '{case_id}' for all {len(rows)} dataset rows after {delay:.1f}s startup delay?",
        )
        if not proceed:
            return

        thread = threading.Thread(
            target=self._run_test_case_all_rows_thread,
            args=(case_id, delay, len(rows)),
            daemon=True,
        )
        thread.start()

    def _run_test_case_all_rows_thread(self, case_id: str, startup_delay: float, row_count: int) -> None:
        self.set_status(f"Running test_case:{case_id} all rows")
        self.append_log(f"Starting test case '{case_id}' for {row_count} dataset rows")
        try:
            if startup_delay > 0:
                self.append_log(f"Waiting {startup_delay:.2f}s before first row...")
                time.sleep(startup_delay)

            self.active_runner = TestFlowRunner(self.data, log_callback=self.append_log)
            self.active_runner.reset_stop()

            pass_count = 0
            fail_count = 0
            stop_count = 0
            for idx in range(row_count):
                if self.active_runner.stop_requested:
                    self.append_log("Stopped before remaining dataset rows.")
                    break
                result = self.active_runner.run_test_case(case_id, dry_run=False, dataset_row_index=idx)
                status = str(result.get("status", "")).lower()
                if status == "passed":
                    pass_count += 1
                elif status == "stopped":
                    stop_count += 1
                    break
                else:
                    fail_count += 1

            self.save_project()
            self.append_log(
                f"Dataset run complete for '{case_id}': passed={pass_count}, failed={fail_count}, stopped={stop_count}"
            )
            self.root.after(
                0,
                lambda: messagebox.showinfo(
                    "Dataset Run Complete",
                    f"Test case: {case_id}\nRows: {row_count}\nPassed: {pass_count}\nFailed: {fail_count}\nStopped: {stop_count}",
                ),
            )
        except RunnerExecutionError as exc:
            self.append_log(f"Dataset run error: {exc}")
            self.root.after(0, lambda: messagebox.showerror("Run Error", str(exc)))
        except Exception as exc:
            self._log_exception("Dataset run thread failure")
            self.append_log(f"Dataset run failure: {exc}")
            self.root.after(0, lambda: messagebox.showerror("Run Error", str(exc)))
        finally:
            self.active_runner = None
            self.set_status("Idle")

    def preview_selected_test_case_variables(self) -> None:
        if self.current_entity_kind != "test_case" or not self.current_entity_name:
            return
        case_id = self.current_entity_name
        tc = self.data["testCases"].get(case_id, {})
        dataset_name = str(tc.get("dataset", "")).strip()
        dataset_row_index = None
        if dataset_name:
            idx = simpledialog.askinteger(
                "Variable Preview",
                f"Dataset '{dataset_name}' row index (0-based):",
                initialvalue=0,
                minvalue=0,
            )
            if idx is None:
                return
            dataset_row_index = idx

        runner = TestFlowRunner(self.data)
        try:
            preview = runner.preview_test_case_execution(case_id, dataset_row_index=dataset_row_index)
        except RunnerExecutionError as exc:
            if hasattr(self, "tc_var_preview"):
                self.tc_var_preview.delete("1.0", tk.END)
                self.tc_var_preview.insert(tk.END, str(exc))
            return

        lines = []
        lines.append("Variables:")
        for k, v in sorted(preview.get("variables", {}).items()):
            lines.append(f"- {k} = {v}")
        lines.append("")
        lines.append("Resolved Steps:")
        for i, step in enumerate(preview.get("executionPlan", []), start=1):
            lines.append(f"{i}. {step}")
        if preview.get("validationErrors"):
            lines.append("")
            lines.append("Validation Errors:")
            for err in preview["validationErrors"]:
                lines.append(f"- {err}")

        if hasattr(self, "tc_var_preview"):
            self.tc_var_preview.delete("1.0", tk.END)
            self.tc_var_preview.insert(tk.END, "\n".join(lines))

    # =====================================================
    # Flows Page
    # =====================================================

    def _build_flows_page(self) -> None:
        container = ttk.Frame(self.main_frame)
        container.pack(fill="both", expand=True)
        container.columnconfigure(1, weight=1)
        container.rowconfigure(1, weight=1)

        ttk.Label(container, text="Reusable Flows", font=("Segoe UI", 16, "bold")).grid(row=0, column=0, sticky="w")

        header_btns = ttk.Frame(container)
        header_btns.grid(row=0, column=1, sticky="e")
        ttk.Button(header_btns, text="Run", command=lambda: self.run_selected("flow", False)).pack(side="left", padx=3)
        ttk.Button(header_btns, text="Dry Run", command=lambda: self.run_selected("flow", True)).pack(side="left", padx=3)
        ttk.Button(header_btns, text="New", command=self.new_flow).pack(side="left", padx=3)
        ttk.Button(header_btns, text="Duplicate", command=self.duplicate_flow).pack(side="left", padx=3)
        ttk.Button(header_btns, text="Delete", command=self.delete_flow).pack(side="left", padx=3)

        body = ttk.PanedWindow(container, orient="horizontal")
        body.grid(row=1, column=0, columnspan=2, sticky="nsew", pady=(10, 0))

        left = ttk.Frame(body, padding=8)
        right = ttk.Frame(body, padding=8)
        body.add(left, weight=1)
        body.add(right, weight=3)
        right.columnconfigure(0, weight=1)
        right.rowconfigure(2, weight=1)

        ttk.Label(left, text="Flow List", font=("Segoe UI", 11, "bold")).pack(anchor="w")
        self.flow_listbox = tk.Listbox(left, height=24)
        self.flow_listbox.pack(fill="both", expand=True, pady=8)
        self.flow_listbox.bind("<<ListboxSelect>>", self._on_flow_select)

        meta = ttk.LabelFrame(right, text="Metadata", padding=8)
        meta.grid(row=0, column=0, sticky="ew")
        meta.columnconfigure(3, weight=1)

        self.flow_name_var = tk.StringVar()
        self.flow_desc_var = tk.StringVar()
        self.flow_params_var = tk.StringVar()

        ttk.Label(meta, text="Name").grid(row=0, column=0, sticky="w")
        ttk.Entry(meta, textvariable=self.flow_name_var, state="disabled").grid(row=0, column=1, sticky="ew", padx=4)
        ttk.Label(meta, text="Description").grid(row=0, column=2, sticky="w")
        ttk.Entry(meta, textvariable=self.flow_desc_var).grid(row=0, column=3, sticky="ew", padx=4)
        ttk.Label(meta, text="Parameters (comma separated)").grid(row=1, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(meta, textvariable=self.flow_params_var).grid(row=1, column=1, columnspan=2, sticky="ew", padx=4, pady=(6, 0))
        ttk.Button(meta, text="Save Metadata", command=self.save_flow_metadata).grid(row=1, column=3, sticky="e", pady=(6, 0))

        step_btns = ttk.Frame(right)
        step_btns.grid(row=1, column=0, sticky="ew", pady=(8, 4))
        ttk.Button(step_btns, text="Add Step", command=self._add_step).pack(side="left", padx=2)
        ttk.Button(step_btns, text="Duplicate Step", command=self._duplicate_step).pack(side="left", padx=2)
        ttk.Button(step_btns, text="Delete Step", command=self._delete_step).pack(side="left", padx=2)
        ttk.Button(step_btns, text="Up", command=lambda: self._move_step(-1)).pack(side="left", padx=2)
        ttk.Button(step_btns, text="Down", command=lambda: self._move_step(1)).pack(side="left", padx=2)

        self.active_step_table = StepTable(right, on_select=self._handle_step_selection)
        self.active_step_table.grid(row=2, column=0, sticky="nsew")

        self.refresh_flow_list()

    def refresh_flow_list(self) -> None:
        if not hasattr(self, "flow_listbox"):
            return
        self.flow_listbox.delete(0, tk.END)
        for flow_name in sorted(self.data["flows"].keys()):
            self.flow_listbox.insert(tk.END, flow_name)

    def _on_flow_select(self, _event=None) -> None:
        if not self.flow_listbox.curselection():
            return
        flow_name = self.flow_listbox.get(self.flow_listbox.curselection()[0])

        self._set_active_context("flow", flow_name)
        flow = self.data["flows"][flow_name]
        self.flow_name_var.set(flow.get("name", flow_name))
        self.flow_desc_var.set(flow.get("description", ""))
        self.flow_params_var.set(", ".join([str(p) for p in flow.get("parameters", [])]))
        self._refresh_active_step_table()

    def new_flow(self) -> None:
        name = simpledialog.askstring("New Flow", "Flow name:")
        if not name:
            return
        name = name.strip()
        if not name:
            return
        if name in self.data["flows"]:
            messagebox.showerror("Error", "Flow already exists.")
            return

        self.data["flows"][name] = {
            "name": name,
            "description": "",
            "parameters": [],
            "steps": [],
        }
        self.save_project()
        self.refresh_flow_list()

    def duplicate_flow(self) -> None:
        if self.current_entity_kind != "flow" or not self.current_entity_name:
            return
        src = self.current_entity_name
        dst = simpledialog.askstring("Duplicate Flow", "New flow name:", initialvalue=f"{src}_COPY")
        if not dst:
            return
        dst = dst.strip()
        if not dst or dst in self.data["flows"]:
            messagebox.showerror("Error", "Invalid or existing flow name.")
            return
        duplicated = copy.deepcopy(self.data["flows"][src])
        duplicated["name"] = dst
        self.data["flows"][dst] = duplicated
        self.save_project()
        self.refresh_flow_list()

    def delete_flow(self) -> None:
        if self.current_entity_kind != "flow" or not self.current_entity_name:
            return
        name = self.current_entity_name
        if not messagebox.askyesno("Delete", f"Delete flow '{name}'?"):
            return
        self.data["flows"].pop(name, None)
        self.current_entity_name = None
        self.save_project()
        self.refresh_flow_list()
        self._refresh_active_step_table()

    def save_flow_metadata(self) -> None:
        if self.current_entity_kind != "flow" or not self.current_entity_name:
            return
        flow = self.data["flows"][self.current_entity_name]
        flow["description"] = self.flow_desc_var.get().strip()
        flow["parameters"] = [p.strip() for p in self.flow_params_var.get().split(",") if p.strip()]
        self.save_project()

    # =====================================================
    # Targets
    # =====================================================

    def _build_targets_page(self) -> None:
        container = ttk.Frame(self.main_frame)
        container.pack(fill="both", expand=True)
        container.rowconfigure(1, weight=1)
        container.columnconfigure(0, weight=1)

        header = ttk.Frame(container)
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)

        ttk.Label(header, text="Targets", font=("Segoe UI", 16, "bold")).grid(row=0, column=0, sticky="w")
        ttk.Button(header, text="Capture", command=self.capture_target).grid(row=0, column=1, padx=3)
        ttk.Button(header, text="Manual Add", command=self.manual_add_target).grid(row=0, column=2, padx=3)
        ttk.Button(header, text="Rename", command=self.rename_target).grid(row=0, column=3, padx=3)
        ttk.Button(header, text="Delete", command=self.delete_target).grid(row=0, column=4, padx=3)
        ttk.Button(header, text="Test Click", command=self.test_target).grid(row=0, column=5, padx=3)

        cols = ("name", "x", "y", "description")
        self.targets_tree = ttk.Treeview(container, columns=cols, show="headings")
        for col, label in zip(cols, ["Name", "X", "Y", "Description"]):
            self.targets_tree.heading(col, text=label)
        self.targets_tree.column("name", width=240)
        self.targets_tree.column("x", width=80, anchor="center")
        self.targets_tree.column("y", width=80, anchor="center")
        self.targets_tree.column("description", width=500)
        self.targets_tree.grid(row=1, column=0, sticky="nsew", pady=(10, 0))

        self.refresh_targets_tree()

    def refresh_targets_tree(self) -> None:
        if not hasattr(self, "targets_tree"):
            return
        self.targets_tree.delete(*self.targets_tree.get_children())
        for name, t in sorted(self.data.get("targets", {}).items()):
            self.targets_tree.insert(
                "",
                tk.END,
                iid=name,
                values=(name, t.get("x", ""), t.get("y", ""), t.get("description", "")),
            )

    def selected_target_name(self) -> str | None:
        if not hasattr(self, "targets_tree"):
            return None
        sel = self.targets_tree.selection()
        return sel[0] if sel else None

    def capture_target(self) -> None:
        if pyautogui is None:
            messagebox.showerror("Missing Dependency", "pyautogui is required to capture target coordinates.")
            return
        name = simpledialog.askstring("Target Name", "Target name:")
        if not name:
            return
        name = name.strip()
        if not name:
            return
        messagebox.showinfo("Capture", "Move mouse to target position, then press OK.")
        x, y = pyautogui.position()
        self.data["targets"][name] = {
            "x": int(x),
            "y": int(y),
            "description": "",
            "createdAt": datetime.now().isoformat(timespec="seconds"),
        }
        self.save_project()
        self.refresh_targets_tree()

    def manual_add_target(self) -> None:
        name = simpledialog.askstring("Target Name", "Target name:")
        if not name:
            return
        name = name.strip()
        if not name:
            return

        x = simpledialog.askinteger("X", "X coordinate:")
        y = simpledialog.askinteger("Y", "Y coordinate:")
        if x is None or y is None:
            return
        desc = simpledialog.askstring("Description", "Description (optional):", initialvalue="") or ""

        self.data["targets"][name] = {
            "x": int(x),
            "y": int(y),
            "description": desc.strip(),
            "createdAt": datetime.now().isoformat(timespec="seconds"),
        }
        self.save_project()
        self.refresh_targets_tree()

    def rename_target(self) -> None:
        name = self.selected_target_name()
        if not name:
            return
        new_name = simpledialog.askstring("Rename Target", "New name:", initialvalue=name)
        if not new_name:
            return
        new_name = new_name.strip()
        if not new_name or (new_name in self.data["targets"] and new_name != name):
            messagebox.showerror("Error", "Invalid target name.")
            return

        self.data["targets"][new_name] = self.data["targets"].pop(name)

        # Rename references in steps
        for flow in self.data.get("flows", {}).values():
            for step in flow.get("steps", []):
                if isinstance(step, dict) and step.get("target") == name:
                    step["target"] = new_name
        for tc in self.data.get("testCases", {}).values():
            for step in tc.get("steps", []):
                if isinstance(step, dict) and step.get("target") == name:
                    step["target"] = new_name

        self.save_project()
        self.refresh_targets_tree()

    def delete_target(self) -> None:
        name = self.selected_target_name()
        if not name:
            return
        if not messagebox.askyesno("Delete", f"Delete target '{name}'?"):
            return

        # quick in-use check
        used_by = []
        for flow_name, flow in self.data.get("flows", {}).items():
            if any(isinstance(s, dict) and s.get("target") == name for s in flow.get("steps", [])):
                used_by.append(f"flow:{flow_name}")
        for tc_name, tc in self.data.get("testCases", {}).items():
            if any(isinstance(s, dict) and s.get("target") == name for s in tc.get("steps", [])):
                used_by.append(f"test_case:{tc_name}")
        if used_by:
            messagebox.showerror("In Use", "Target is used by:\n" + "\n".join(used_by))
            return

        self.data["targets"].pop(name, None)
        self.save_project()
        self.refresh_targets_tree()

    def test_target(self) -> None:
        if pyautogui is None:
            messagebox.showerror("Missing Dependency", "pyautogui is required for target click testing.")
            return
        name = self.selected_target_name()
        if not name:
            return
        t = self.data["targets"][name]
        delay = float(self.data["settings"].get("startupDelaySeconds", 3))

        if not messagebox.askyesno(
            "Test Target", f"Click '{name}' after {delay:.1f}s at x={t['x']} y={t['y']}?"
        ):
            return

        def do_click():
            time.sleep(delay)
            pyautogui.click(int(t["x"]), int(t["y"]))

        threading.Thread(target=do_click, daemon=True).start()

    # =====================================================
    # Recorder
    # =====================================================

    def _build_recorder_page(self) -> None:
        container = ttk.Frame(self.main_frame)
        container.pack(fill="both", expand=True)
        container.rowconfigure(3, weight=1)
        container.columnconfigure(0, weight=1)

        ttk.Label(container, text="Recorder", font=("Segoe UI", 16, "bold")).grid(row=0, column=0, sticky="w")

        options = ttk.LabelFrame(container, text="Options", padding=8)
        options.grid(row=1, column=0, sticky="ew", pady=(8, 8))

        self.rec_opt_left = tk.BooleanVar(value=True)
        self.rec_opt_double = tk.BooleanVar(value=True)
        self.rec_opt_right = tk.BooleanVar(value=True)
        self.rec_opt_timing = tk.BooleanVar(value=True)
        self.rec_opt_ignore_app = tk.BooleanVar(value=True)
        self.rec_opt_typing = tk.BooleanVar(value=False)
        self.rec_opt_hotkeys = tk.BooleanVar(value=False)
        self.rec_opt_screenshot = tk.BooleanVar(value=False)
        self.rec_cleanup_min_wait = tk.StringVar(value="0.05")
        self.rec_cleanup_round_wait = tk.BooleanVar(value=True)
        self.rec_cleanup_round_step = tk.StringVar(value="0.1")
        self.rec_cleanup_target_tolerance = tk.StringVar(value="8")
        self.rec_cleanup_reuse_existing = tk.BooleanVar(value=True)
        self.rec_cleanup_suggest_targets = tk.BooleanVar(value=True)

        ttk.Checkbutton(options, text="Record left clicks", variable=self.rec_opt_left).grid(row=0, column=0, sticky="w")
        ttk.Checkbutton(options, text="Record double clicks", variable=self.rec_opt_double).grid(row=0, column=1, sticky="w")
        ttk.Checkbutton(options, text="Record right clicks", variable=self.rec_opt_right).grid(row=0, column=2, sticky="w")
        ttk.Checkbutton(options, text="Record timing gaps", variable=self.rec_opt_timing).grid(row=1, column=0, sticky="w")
        ttk.Checkbutton(options, text="Ignore clicks inside app window", variable=self.rec_opt_ignore_app).grid(
            row=1, column=1, columnspan=2, sticky="w"
        )
        ttk.Checkbutton(options, text="Record typing (coming later)", variable=self.rec_opt_typing, state="disabled").grid(
            row=2, column=0, sticky="w"
        )
        ttk.Checkbutton(options, text="Record hotkeys (coming later)", variable=self.rec_opt_hotkeys, state="disabled").grid(
            row=2, column=1, sticky="w"
        )
        ttk.Checkbutton(
            options,
            text="Screenshot after each click (coming later)",
            variable=self.rec_opt_screenshot,
            state="disabled",
        ).grid(
            row=2, column=2, sticky="w"
        )

        controls = ttk.Frame(options)
        controls.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(8, 0))

        self.rec_state_var = tk.StringVar(value="Idle")
        ttk.Label(controls, textvariable=self.rec_state_var).pack(side="left")
        ttk.Button(controls, text="Start Recording", command=self.start_recording).pack(side="left", padx=6)
        ttk.Button(controls, text="Stop Recording", command=self.stop_recording).pack(side="left", padx=6)
        ttk.Label(controls, text="Stop hotkey:", foreground="#475467").pack(side="left", padx=(8, 2))
        ttk.Label(
            controls,
            text=str(self.data.get("settings", {}).get("recordingStopHotkey", "f8")).upper(),
            foreground="#475467",
        ).pack(side="left")

        cleanup = ttk.LabelFrame(container, text="Cleanup and Target Suggestions", padding=8)
        cleanup.grid(row=2, column=0, sticky="ew", pady=(0, 8))
        cleanup.columnconfigure(5, weight=1)

        ttk.Label(cleanup, text="Remove waits below (s)").grid(row=0, column=0, sticky="w")
        ttk.Entry(cleanup, textvariable=self.rec_cleanup_min_wait, width=8).grid(row=0, column=1, sticky="w", padx=4)
        ttk.Checkbutton(cleanup, text="Round waits", variable=self.rec_cleanup_round_wait).grid(row=0, column=2, sticky="w")
        ttk.Label(cleanup, text="to").grid(row=0, column=3, sticky="e")
        ttk.Entry(cleanup, textvariable=self.rec_cleanup_round_step, width=8).grid(row=0, column=4, sticky="w", padx=4)

        ttk.Label(cleanup, text="Target tolerance (px)").grid(row=1, column=0, sticky="w", pady=(6, 0))
        ttk.Entry(cleanup, textvariable=self.rec_cleanup_target_tolerance, width=8).grid(
            row=1, column=1, sticky="w", padx=4, pady=(6, 0)
        )
        ttk.Checkbutton(cleanup, text="Reuse nearby existing targets", variable=self.rec_cleanup_reuse_existing).grid(
            row=1, column=2, columnspan=2, sticky="w", pady=(6, 0)
        )
        ttk.Checkbutton(cleanup, text="Suggest targets for repeated clicks", variable=self.rec_cleanup_suggest_targets).grid(
            row=1, column=4, columnspan=2, sticky="w", pady=(6, 0)
        )
        ttk.Button(cleanup, text="Apply Cleanup + Suggestions", command=self.apply_recording_cleanup).grid(
            row=0, column=5, rowspan=2, sticky="e"
        )

        preview = ttk.LabelFrame(container, text="Recorded Step Preview", padding=8)
        preview.grid(row=3, column=0, sticky="nsew")
        preview.rowconfigure(0, weight=1)
        preview.columnconfigure(0, weight=1)

        cols = ("index", "type", "details")
        self.rec_preview_tree = ttk.Treeview(preview, columns=cols, show="headings")
        self.rec_preview_tree.heading("index", text="#")
        self.rec_preview_tree.heading("type", text="Type")
        self.rec_preview_tree.heading("details", text="Details")
        self.rec_preview_tree.column("index", width=50, anchor="center")
        self.rec_preview_tree.column("type", width=180)
        self.rec_preview_tree.column("details", width=620)
        self.rec_preview_tree.grid(row=0, column=0, sticky="nsew")

        actions = ttk.Frame(preview)
        actions.grid(row=1, column=0, sticky="ew", pady=(8, 0))
        ttk.Button(actions, text="Save as Reusable Flow", command=self.save_recording_as_flow).pack(side="left", padx=3)
        ttk.Button(actions, text="Save as Test Case", command=self.save_recording_as_test_case).pack(side="left", padx=3)
        ttk.Button(actions, text="Append to Existing Flow", command=self.append_recording_to_flow).pack(side="left", padx=3)
        ttk.Button(actions, text="Discard", command=self.discard_recording).pack(side="left", padx=3)

        status = self.recorder.availability()
        if not status.available:
            self.rec_state_var.set(status.message)

    def start_recording(self) -> None:
        stop_key = str(self.data.get("settings", {}).get("recordingStopHotkey", "f8"))
        self.root.update_idletasks()
        self._recording_window_bounds = (
            self.root.winfo_rootx(),
            self.root.winfo_rooty(),
            self.root.winfo_width(),
            self.root.winfo_height(),
        )
        self.recorder.set_click_filter(self._recording_click_filter)
        status = self.recorder.start(stop_key_name=stop_key, on_finished=self._on_recording_finished)
        if not status.available:
            messagebox.showerror("Recorder Unavailable", status.message)
            self.rec_state_var.set(status.message)
            return

        self.recording_mode_active = True
        self.set_recording_badge("RECORDING")
        self.set_status("Recording")
        self.rec_state_var.set(f"Recording... press {stop_key.upper()} to stop")
        self.append_log("Recording started.")

    def stop_recording(self) -> None:
        if not self.recorder.is_running:
            return
        self.recorder.stop()
        self.rec_state_var.set("Stopping...")
        self.set_recording_badge("")
        self.set_status("Stopping recorder")

    def _on_recording_finished(self, steps: list[dict], error_text: str) -> None:
        def ui_update():
            self.recording_mode_active = False
            self.set_recording_badge("")
            self.set_status("Idle")
            self.recorded_preview_steps = self._cleanup_recorded_steps(steps)
            self.refresh_record_preview()
            if error_text:
                self.rec_state_var.set(f"Recorder error: {error_text}")
                self.append_log(f"Recorder error: {error_text}")
            else:
                self.rec_state_var.set(f"Idle ({len(self.recorded_preview_steps)} steps captured)")
                self.append_log(f"Recording stopped. Captured {len(self.recorded_preview_steps)} steps.")

        self.root.after(0, ui_update)

    def _cleanup_recorded_steps(self, steps: list[dict]) -> list[dict]:
        # Respect enabled options and apply wait cleanup.
        try:
            min_wait = max(0.0, float(self.rec_cleanup_min_wait.get()))
        except Exception:
            min_wait = 0.05
        try:
            round_step = max(0.001, float(self.rec_cleanup_round_step.get()))
        except Exception:
            round_step = 0.1
        do_round = bool(self.rec_cleanup_round_wait.get())

        out: list[dict] = []
        for step in steps:
            t = step.get("type")
            if t == "wait":
                if not self.rec_opt_timing.get():
                    continue
                try:
                    sec = float(step.get("seconds", 0))
                except Exception:
                    continue
                if do_round:
                    sec = round(sec / round_step) * round_step
                sec = round(sec, 3)
                if sec < min_wait:
                    continue
                out.append({"type": "wait", "seconds": sec, "enabled": True})
                continue

            if t == "click_xy":
                if not self.rec_opt_left.get():
                    continue
                out.append({"type": "click_xy", "x": int(step.get("x", 0)), "y": int(step.get("y", 0)), "enabled": True})
                continue

            if t == "double_click":
                if not self.rec_opt_double.get():
                    continue
                out.append(
                    {
                        "type": "double_click",
                        "x": int(step.get("x", 0)),
                        "y": int(step.get("y", 0)),
                        "enabled": True,
                    }
                )
                continue

            if t == "right_click":
                if not self.rec_opt_right.get():
                    continue
                out.append(
                    {
                        "type": "right_click",
                        "x": int(step.get("x", 0)),
                        "y": int(step.get("y", 0)),
                        "enabled": True,
                    }
                )
                continue

            out.append(step)

        return out

    def apply_recording_cleanup(self) -> None:
        if not self.recorded_preview_steps:
            return
        self.recorded_preview_steps = self._cleanup_recorded_steps(self.recorded_preview_steps)
        self.recorded_preview_steps = self._apply_target_suggestions(self.recorded_preview_steps)
        self.refresh_record_preview()
        self.append_log("Applied recording cleanup and target suggestions.")

    def refresh_record_preview(self) -> None:
        if not hasattr(self, "rec_preview_tree"):
            return
        self.rec_preview_tree.delete(*self.rec_preview_tree.get_children())
        for i, step in enumerate(self.recorded_preview_steps, start=1):
            details = ", ".join([f"{k}={v}" for k, v in step.items() if k != "type"])[:220]
            self.rec_preview_tree.insert("", tk.END, values=(i, step.get("type", ""), details))

    def save_recording_as_flow(self) -> None:
        if not self.recorded_preview_steps:
            messagebox.showwarning("No Recording", "No recorded steps to save.")
            return
        self.recorded_preview_steps = self._apply_target_suggestions(self.recorded_preview_steps)
        name = simpledialog.askstring("Save Flow", "Flow name:")
        if not name:
            return
        name = name.strip()
        if not name:
            return
        self.data["flows"][name] = {
            "name": name,
            "description": "Recorded flow",
            "parameters": [],
            "steps": copy.deepcopy(self.recorded_preview_steps),
        }
        self.save_project()
        self.append_log(f"Saved recording as flow '{name}'.")
        messagebox.showinfo("Saved", f"Flow '{name}' saved.")

    def save_recording_as_test_case(self) -> None:
        if not self.recorded_preview_steps:
            messagebox.showwarning("No Recording", "No recorded steps to save.")
            return
        self.recorded_preview_steps = self._apply_target_suggestions(self.recorded_preview_steps)
        case_id = simpledialog.askstring("Save Test Case", "Test case ID:")
        if not case_id:
            return
        case_id = case_id.strip()
        if not case_id:
            return
        self.data["testCases"][case_id] = {
            "id": case_id,
            "name": case_id,
            "suite": "",
            "description": "Recorded test case",
            "dataset": "",
            "enabled": True,
            "variables": {},
            "steps": copy.deepcopy(self.recorded_preview_steps),
        }
        self.save_project()
        self.append_log(f"Saved recording as test case '{case_id}'.")
        messagebox.showinfo("Saved", f"Test case '{case_id}' saved.")

    def append_recording_to_flow(self) -> None:
        if not self.recorded_preview_steps:
            messagebox.showwarning("No Recording", "No recorded steps to append.")
            return
        flow_names = sorted(self.data.get("flows", {}).keys())
        if not flow_names:
            messagebox.showwarning("No Flows", "Create a flow first.")
            return

        flow_name = simpledialog.askstring(
            "Append to Flow",
            "Flow name to append steps to:",
            initialvalue=flow_names[0],
        )
        if not flow_name:
            return
        flow_name = flow_name.strip()
        if flow_name not in self.data.get("flows", {}):
            messagebox.showerror("Missing Flow", f"Flow not found: {flow_name}")
            return

        cleaned = self._apply_target_suggestions(self.recorded_preview_steps)
        self.data["flows"][flow_name].setdefault("steps", []).extend(copy.deepcopy(cleaned))
        self.save_project()
        self.append_log(f"Appended {len(cleaned)} recorded steps to flow '{flow_name}'.")
        messagebox.showinfo("Appended", f"Added {len(cleaned)} steps to '{flow_name}'.")

    def discard_recording(self) -> None:
        self.recorded_preview_steps = []
        self.refresh_record_preview()
        self.rec_state_var.set("Idle")
        self.set_recording_badge("")

    def _recording_click_filter(self, x: int, y: int, _button_name: str) -> bool:
        if not bool(self.rec_opt_ignore_app.get()):
            return True
        bounds = self._recording_window_bounds
        if bounds is None:
            return True
        rx, ry, rw, rh = bounds
        inside = rx <= x <= rx + rw and ry <= y <= ry + rh
        return not inside

    def _apply_target_suggestions(self, steps: list[dict]) -> list[dict]:
        if not steps:
            return steps
        try:
            tolerance = max(1.0, float(self.rec_cleanup_target_tolerance.get()))
        except Exception:
            tolerance = 8.0

        click_types = {"click_xy", "click", "double_click", "right_click"}
        points: list[tuple[int, int, int]] = []
        for idx, step in enumerate(steps):
            if step.get("type") not in click_types:
                continue
            if "x" in step and "y" in step:
                try:
                    points.append((idx, int(step["x"]), int(step["y"])))
                except Exception:
                    pass

        if not points:
            return steps

        clusters: list[dict] = []
        for step_idx, x, y in points:
            assigned = None
            for cluster in clusters:
                if self._distance((x, y), (cluster["cx"], cluster["cy"])) <= tolerance:
                    assigned = cluster
                    break
            if assigned is None:
                clusters.append({"cx": x, "cy": y, "items": [(step_idx, x, y)]})
            else:
                assigned["items"].append((step_idx, x, y))
                xs = [p[1] for p in assigned["items"]]
                ys = [p[2] for p in assigned["items"]]
                assigned["cx"] = sum(xs) / len(xs)
                assigned["cy"] = sum(ys) / len(ys)

        existing = self.data.get("targets", {})
        used_names = set(existing.keys())
        created_names: dict[int, str] = {}
        cluster_target: dict[int, str] = {}
        pending_new_targets: list[tuple[str, int, int]] = []

        for ci, cluster in enumerate(clusters):
            cx, cy = int(round(cluster["cx"])), int(round(cluster["cy"]))
            chosen = ""

            if self.rec_cleanup_reuse_existing.get():
                nearest = self._nearest_target_name(cx, cy, tolerance)
                if nearest:
                    chosen = nearest

            if not chosen and self.rec_cleanup_suggest_targets.get() and len(cluster["items"]) >= 2:
                name = self._next_target_name(used_names)
                used_names.add(name)
                chosen = name
                created_names[ci] = name
                pending_new_targets.append((name, cx, cy))

            if chosen:
                cluster_target[ci] = chosen

        if pending_new_targets:
            lines = [f"{n}: x={x}, y={y}" for (n, x, y) in pending_new_targets]
            accepted = messagebox.askyesno(
                "Suggested Targets",
                "Apply these suggested targets?\n\n" + "\n".join(lines),
            )
            if not accepted:
                # Remove only newly generated names from mapping; keep reuse of existing targets.
                for ci in list(created_names.keys()):
                    cluster_target.pop(ci, None)
            else:
                for n, x, y in pending_new_targets:
                    self.data["targets"][n] = {
                        "x": int(x),
                        "y": int(y),
                        "description": "Auto-suggested from recording",
                        "createdAt": datetime.now().isoformat(timespec="seconds"),
                    }

        if not cluster_target:
            return steps

        new_steps = copy.deepcopy(steps)
        point_cluster: dict[int, int] = {}
        for ci, cluster in enumerate(clusters):
            for step_idx, _, _ in cluster["items"]:
                point_cluster[step_idx] = ci

        for idx, step in enumerate(new_steps):
            ci = point_cluster.get(idx)
            if ci is None:
                continue
            target_name = cluster_target.get(ci)
            if not target_name:
                continue
            step["target"] = target_name
            step.pop("x", None)
            step.pop("y", None)
            if step.get("type") == "click_xy":
                step["type"] = "click"

        self.save_project()
        return new_steps

    def _nearest_target_name(self, x: int, y: int, tolerance: float) -> str:
        nearest_name = ""
        nearest_dist = tolerance + 1.0
        for name, target in self.data.get("targets", {}).items():
            try:
                tx, ty = int(target.get("x")), int(target.get("y"))
            except Exception:
                continue
            d = self._distance((x, y), (tx, ty))
            if d <= tolerance and d < nearest_dist:
                nearest_dist = d
                nearest_name = name
        return nearest_name

    def _next_target_name(self, used_names: set[str]) -> str:
        i = 1
        while True:
            candidate = f"target_{i:03d}"
            if candidate not in used_names:
                return candidate
            i += 1

    @staticmethod
    def _distance(p1: tuple[float, float], p2: tuple[float, float]) -> float:
        return math.sqrt((p1[0] - p2[0]) ** 2 + (p1[1] - p2[1]) ** 2)

    # =====================================================
    # Datasets
    # =====================================================

    def _build_datasets_page(self) -> None:
        container = ttk.Frame(self.main_frame)
        container.pack(fill="both", expand=True)
        container.columnconfigure(1, weight=1)
        container.rowconfigure(1, weight=1)

        ttk.Label(container, text="Datasets", font=("Segoe UI", 16, "bold")).grid(row=0, column=0, sticky="w")
        actions = ttk.Frame(container)
        actions.grid(row=0, column=1, sticky="e")
        ttk.Button(actions, text="Import CSV/XLSX", command=self.import_dataset_file).pack(side="left", padx=3)
        ttk.Button(actions, text="Rename", command=self.rename_dataset).pack(side="left", padx=3)
        ttk.Button(actions, text="Delete", command=self.delete_dataset).pack(side="left", padx=3)

        left = ttk.Frame(container, padding=8)
        left.grid(row=1, column=0, sticky="nsw")
        right = ttk.Frame(container, padding=8)
        right.grid(row=1, column=1, sticky="nsew")
        right.columnconfigure(0, weight=1)
        right.rowconfigure(1, weight=1)

        ttk.Label(left, text="Dataset Names", font=("Segoe UI", 11, "bold")).pack(anchor="w")
        self.dataset_listbox = tk.Listbox(left, width=28, height=24)
        self.dataset_listbox.pack(fill="y", expand=True, pady=8)
        self.dataset_listbox.bind("<<ListboxSelect>>", self.on_dataset_select)

        self.dataset_info_var = tk.StringVar(value="Select a dataset to preview")
        ttk.Label(right, textvariable=self.dataset_info_var).grid(row=0, column=0, sticky="w")

        self.dataset_preview_text = tk.Text(right, height=28)
        self.dataset_preview_text.grid(row=1, column=0, sticky="nsew", pady=(8, 0))

        self.refresh_dataset_list()

    def refresh_dataset_list(self) -> None:
        if not hasattr(self, "dataset_listbox"):
            return
        self.dataset_listbox.delete(0, tk.END)
        for name in sorted(self.data.get("datasets", {}).keys()):
            self.dataset_listbox.insert(tk.END, name)

    def selected_dataset_name(self) -> str | None:
        if not hasattr(self, "dataset_listbox"):
            return None
        sel = self.dataset_listbox.curselection()
        if not sel:
            return None
        return self.dataset_listbox.get(sel[0])

    def on_dataset_select(self, _event=None) -> None:
        name = self.selected_dataset_name()
        if not name:
            return
        ds = self.data.get("datasets", {}).get(name)
        rows = []
        if isinstance(ds, dict):
            rows = ds.get("rows", []) if isinstance(ds.get("rows"), list) else []
        elif isinstance(ds, list):
            rows = ds

        cols = []
        if rows and isinstance(rows[0], dict):
            cols = list(rows[0].keys())

        self.dataset_info_var.set(f"Rows: {len(rows)} | Columns: {', '.join(cols)}")
        self.dataset_preview_text.delete("1.0", tk.END)
        for i, row in enumerate(rows[:100], start=1):
            self.dataset_preview_text.insert(tk.END, f"{i}. {row}\n")

    def import_dataset_file(self) -> None:
        path = filedialog.askopenfilename(
            title="Import dataset (CSV/XLSX)",
            filetypes=[("CSV/XLSX", "*.csv *.xlsx"), ("CSV", "*.csv"), ("Excel", "*.xlsx"), ("All", "*.*")],
        )
        if not path:
            return
        dataset_name = simpledialog.askstring("Dataset Name", "Dataset name:", initialvalue=Path(path).stem)
        if not dataset_name:
            return
        dataset_name = dataset_name.strip()
        if not dataset_name:
            return

        try:
            _, rows = read_table_rows(path)
        except ImporterError as exc:
            messagebox.showerror("Import Error", str(exc))
            return
        except Exception as exc:
            self._log_exception("Import dataset failed")
            messagebox.showerror("Import Error", f"Failed to import dataset: {exc}")
            return

        self.data.setdefault("datasets", {})[dataset_name] = {"rows": rows, "source": path}
        self.save_project()
        self.refresh_dataset_list()
        self.append_log(f"Imported dataset '{dataset_name}' with {len(rows)} rows.")

    def rename_dataset(self) -> None:
        name = self.selected_dataset_name()
        if not name:
            return
        new_name = simpledialog.askstring("Rename Dataset", "New name:", initialvalue=name)
        if not new_name:
            return
        new_name = new_name.strip()
        if not new_name or (new_name in self.data.get("datasets", {}) and new_name != name):
            messagebox.showerror("Error", "Invalid dataset name.")
            return
        self.data["datasets"][new_name] = self.data["datasets"].pop(name)
        self.save_project()
        self.refresh_dataset_list()

    def delete_dataset(self) -> None:
        name = self.selected_dataset_name()
        if not name:
            return
        if not messagebox.askyesno("Delete", f"Delete dataset '{name}'?"):
            return
        self.data.get("datasets", {}).pop(name, None)
        self.save_project()
        self.refresh_dataset_list()

    # =====================================================
    # Run center
    # =====================================================

    def _build_run_center_page(self) -> None:
        container = ttk.Frame(self.main_frame)
        container.pack(fill="both", expand=True)
        container.rowconfigure(1, weight=1)
        container.columnconfigure(0, weight=1)

        header = ttk.Frame(container)
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)

        ttk.Label(header, text="Run Center", font=("Segoe UI", 16, "bold")).grid(row=0, column=0, sticky="w")
        ttk.Button(header, text="Refresh", command=self.refresh_runs_tree).grid(row=0, column=1, padx=3)
        ttk.Button(header, text="Open Selected Report", command=self.open_selected_report).grid(row=0, column=2, padx=3)

        cols = ("run_id", "kind", "name", "status", "started", "duration", "folder")
        self.runs_tree = ttk.Treeview(container, columns=cols, show="headings")
        headers = ["Run ID", "Kind", "Name", "Status", "Started", "Duration", "Folder"]
        widths = [260, 90, 180, 90, 180, 90, 360]
        for c, h, w in zip(cols, headers, widths):
            self.runs_tree.heading(c, text=h)
            self.runs_tree.column(c, width=w)

        self.runs_tree.grid(row=1, column=0, sticky="nsew", pady=(10, 0))
        self.refresh_runs_tree()

    def refresh_runs_tree(self) -> None:
        if not hasattr(self, "runs_tree"):
            return
        self.runs_tree.delete(*self.runs_tree.get_children())
        for run in list(self.data.get("runs", []))[::-1]:
            folder = run.get("runFolder", "")
            self.runs_tree.insert(
                "",
                tk.END,
                values=(
                    run.get("runId", ""),
                    run.get("kind", ""),
                    run.get("name", ""),
                    run.get("status", ""),
                    run.get("startedAt", ""),
                    run.get("durationSeconds", ""),
                    folder,
                ),
            )

    def open_selected_report(self) -> None:
        if not hasattr(self, "runs_tree"):
            return
        sel = self.runs_tree.selection()
        if not sel:
            return
        vals = self.runs_tree.item(sel[0]).get("values", [])
        if len(vals) < 7:
            return
        folder = str(vals[6])
        report_path = Path(folder) / "report.html"
        if not report_path.exists():
            messagebox.showerror("Missing Report", f"Report not found:\n{report_path}")
            return
        webbrowser.open(report_path.resolve().as_uri())

    # =====================================================
    # Settings
    # =====================================================

    def _build_settings_page(self) -> None:
        container = ttk.Frame(self.main_frame)
        container.pack(fill="both", expand=True)

        ttk.Label(container, text="Settings", font=("Segoe UI", 16, "bold")).pack(anchor="w")

        form = ttk.LabelFrame(container, text="Runner Settings", padding=10)
        form.pack(fill="x", pady=10)

        settings = self.data.setdefault("settings", {})
        env = self.data.setdefault("environment", {})

        self.set_start_delay = tk.StringVar(value=str(settings.get("startupDelaySeconds", 3)))
        self.set_action_pause = tk.StringVar(value=str(settings.get("defaultActionPauseSeconds", 0.1)))
        self.set_shot_fail = tk.BooleanVar(value=bool(settings.get("screenshotOnFailure", True)))
        self.set_shot_each = tk.BooleanVar(value=bool(settings.get("screenshotAfterEachStep", False)))
        self.set_stop_hotkey = tk.StringVar(value=str(settings.get("stopHotkey", "f8")))
        self.set_record_stop_hotkey = tk.StringVar(value=str(settings.get("recordingStopHotkey", "f8")))
        self.set_expected_res = tk.StringVar(value=str(env.get("expectedResolution", "")))

        row = 0
        self._settings_row(form, row, "Startup delay (s)", self.set_start_delay)
        row += 1
        self._settings_row(form, row, "Default action pause (s)", self.set_action_pause)
        row += 1
        self._settings_row(form, row, "Stop hotkey", self.set_stop_hotkey)
        row += 1
        self._settings_row(form, row, "Recording stop hotkey", self.set_record_stop_hotkey)
        row += 1
        self._settings_row(form, row, "Expected resolution", self.set_expected_res)
        row += 1

        ttk.Checkbutton(form, text="Screenshot on failure", variable=self.set_shot_fail).grid(
            row=row, column=0, columnspan=2, sticky="w", pady=4
        )
        row += 1
        ttk.Checkbutton(form, text="Screenshot after each step", variable=self.set_shot_each).grid(
            row=row, column=0, columnspan=2, sticky="w", pady=4
        )
        row += 1

        ttk.Button(form, text="Save Settings", command=self.save_settings).grid(row=row, column=1, sticky="e", pady=8)

        notes = ttk.LabelFrame(container, text="System Notes", padding=10)
        notes.pack(fill="both", expand=True)
        msg = (
            "Cross-platform behavior:\n"
            "- Recorder uses pynput for macOS/Windows/Linux global input hooks.\n"
            "- Live execution uses pyautogui.\n"
            "- On macOS, grant Accessibility + Screen Recording permissions to your Python app/terminal.\n"
            "- If pyautogui is missing, dry-run remains available."
        )
        ttk.Label(notes, text=msg, justify="left").pack(anchor="w")

        tools = ttk.LabelFrame(container, text="Project Tools", padding=10)
        tools.pack(fill="x", pady=(10, 0))
        ttk.Button(tools, text="Export Project ZIP", command=self.export_project_zip).pack(side="left", padx=4)
        ttk.Button(tools, text="Import Project JSON", command=self.import_project_json).pack(side="left", padx=4)
        ttk.Button(tools, text="Import Project ZIP", command=self.import_project_zip).pack(side="left", padx=4)
        ttk.Button(tools, text="Import Project Folder", command=self.import_project_folder).pack(side="left", padx=4)
        ttk.Button(tools, text="Create Sample Project", command=self.create_sample_project).pack(side="left", padx=4)
        ttk.Button(tools, text="Run Startup Checks", command=self._run_startup_checks).pack(side="left", padx=4)

    def _settings_row(self, parent, row: int, label: str, var: tk.StringVar) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=4)
        ttk.Entry(parent, textvariable=var, width=30).grid(row=row, column=1, sticky="ew", padx=6, pady=4)

    def save_settings(self) -> None:
        settings = self.data.setdefault("settings", {})
        env = self.data.setdefault("environment", {})

        try:
            settings["startupDelaySeconds"] = float(self.set_start_delay.get())
            settings["defaultActionPauseSeconds"] = float(self.set_action_pause.get())
        except ValueError:
            messagebox.showerror("Invalid", "Startup delay and action pause must be numeric.")
            return

        settings["screenshotOnFailure"] = bool(self.set_shot_fail.get())
        settings["screenshotAfterEachStep"] = bool(self.set_shot_each.get())
        settings["stopHotkey"] = self.set_stop_hotkey.get().strip() or "f8"
        settings["recordingStopHotkey"] = self.set_record_stop_hotkey.get().strip() or "f8"
        env["expectedResolution"] = self.set_expected_res.get().strip()

        if pyautogui is not None:
            pyautogui.PAUSE = float(settings.get("defaultActionPauseSeconds", 0.1))

        self.save_project()
        self.append_log("Settings saved.")
        messagebox.showinfo("Saved", "Settings saved.")

    def export_project_zip(self) -> None:
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        suggested = f"testflow_project_export_{stamp}.zip"
        out_path = filedialog.asksaveasfilename(
            title="Export project as ZIP",
            defaultextension=".zip",
            initialfile=suggested,
            filetypes=[("ZIP", "*.zip")],
        )
        if not out_path:
            return
        out = Path(out_path)

        self.save_project()
        root_dir = Path.cwd()
        items = [PROJECT_FILE, Path("runs"), Path("logs")]
        with zipfile.ZipFile(out, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for item in items:
                full = root_dir / item
                if not full.exists():
                    continue
                if full.is_file():
                    zf.write(full, arcname=str(item))
                else:
                    for p in full.rglob("*"):
                        if p.is_file():
                            zf.write(p, arcname=str(p.relative_to(root_dir)))
        self.append_log(f"Exported project ZIP: {out}")
        messagebox.showinfo("Export Complete", f"Project exported to:\n{out}")

    def import_project_json(self) -> None:
        in_path = filedialog.askopenfilename(
            title="Import project JSON",
            filetypes=[("JSON", "*.json"), ("All", "*.*")],
        )
        if not in_path:
            return
        path = Path(in_path)
        try:
            raw = json.loads(path.read_text(encoding="utf-8"))
            project = normalize_project_data(raw)
        except Exception as exc:
            self._log_exception("Import project JSON failed")
            messagebox.showerror("Import Error", f"Unable to parse project JSON: {exc}")
            return

        self.data = project
        self.save_project()
        self.append_log(f"Imported project from JSON: {path}")
        self.show_page(self.current_page)
        messagebox.showinfo("Import Complete", f"Loaded project from:\n{path}")

    def import_project_zip(self) -> None:
        in_path = filedialog.askopenfilename(
            title="Import project ZIP",
            filetypes=[("ZIP", "*.zip"), ("All", "*.*")],
        )
        if not in_path:
            return
        path = Path(in_path)
        root_dir = Path.cwd()

        if not messagebox.askyesno(
            "Import Project ZIP",
            "This will replace current project JSON and may overwrite runs/logs with files from the ZIP.\n\nContinue?",
        ):
            return

        try:
            with zipfile.ZipFile(path, "r") as zf:
                names = zf.namelist()
                if "testflow_project.json" not in names:
                    raise ValueError("ZIP does not contain testflow_project.json")
                zf.extractall(root_dir)
            loaded, msgs = load_project()
            self.data = loaded
            for m in msgs:
                self.append_log(m)
        except Exception as exc:
            self._log_exception("Import project ZIP failed")
            messagebox.showerror("Import Error", f"Unable to import ZIP: {exc}")
            return

        self.save_project()
        self.show_page(self.current_page)
        self.append_log(f"Imported project ZIP: {path}")
        messagebox.showinfo("Import Complete", f"Loaded project from ZIP:\n{path}")

    def import_project_folder(self) -> None:
        folder = filedialog.askdirectory(title="Import project folder")
        if not folder:
            return
        source = Path(folder)
        project_file = source / "testflow_project.json"
        if not project_file.exists():
            messagebox.showerror("Import Error", "Selected folder does not contain testflow_project.json")
            return

        if not messagebox.askyesno(
            "Import Project Folder",
            "This will replace current project JSON and sync runs/logs from selected folder.\n\nContinue?",
        ):
            return

        try:
            shutil.copy2(project_file, PROJECT_FILE)
            for sub in ("runs", "logs"):
                src_sub = source / sub
                dst_sub = Path(sub)
                if src_sub.exists():
                    if dst_sub.exists():
                        shutil.rmtree(dst_sub)
                    shutil.copytree(src_sub, dst_sub)
            loaded, msgs = load_project()
            self.data = loaded
            for m in msgs:
                self.append_log(m)
        except Exception as exc:
            self._log_exception("Import project folder failed")
            messagebox.showerror("Import Error", f"Unable to import project folder: {exc}")
            return

        self.save_project()
        self.show_page(self.current_page)
        self.append_log(f"Imported project folder: {source}")
        messagebox.showinfo("Import Complete", f"Loaded project from folder:\n{source}")

    def create_sample_project(self) -> None:
        if not messagebox.askyesno(
            "Create Sample Project",
            "This will merge sample data into the current project (it will not delete your existing entries).\n\nContinue?",
        ):
            return

        self.data.setdefault("targets", {}).update(
            {
                "username_field": {
                    "x": 640,
                    "y": 360,
                    "description": "Sample username input",
                    "createdAt": datetime.now().isoformat(timespec="seconds"),
                },
                "password_field": {
                    "x": 640,
                    "y": 400,
                    "description": "Sample password input",
                    "createdAt": datetime.now().isoformat(timespec="seconds"),
                },
                "login_button": {
                    "x": 640,
                    "y": 450,
                    "description": "Sample login button",
                    "createdAt": datetime.now().isoformat(timespec="seconds"),
                },
            }
        )

        self.data.setdefault("flows", {})["Login_POS"] = {
            "name": "Login_POS",
            "description": "Sample reusable login flow",
            "parameters": [],
            "steps": [
                {"type": "click", "target": "username_field", "enabled": True},
                {"type": "type_text", "value": "${username}", "enabled": True},
                {"type": "click", "target": "password_field", "enabled": True},
                {"type": "type_text", "value": "${password}", "enabled": True},
                {"type": "click", "target": "login_button", "enabled": True},
                {"type": "press_key", "key": "enter", "enabled": True},
            ],
        }

        self.data.setdefault("datasets", {})["demo_users"] = {
            "rows": [
                {"username": "demo_user_1", "password": "DemoPass123", "expectedTitle": "Dashboard"},
                {"username": "demo_user_2", "password": "DemoPass456", "expectedTitle": "Dashboard"},
            ]
        }

        self.data.setdefault("testCases", {})["POS_DEMO_001"] = {
            "id": "POS_DEMO_001",
            "name": "Sample login with dataset",
            "suite": "Smoke",
            "description": "Sample test case generated by Pass 5 helper",
            "dataset": "demo_users",
            "enabled": True,
            "variables": {},
            "steps": [
                {"type": "run_flow", "flow": "Login_POS", "enabled": True},
                {"type": "assert_window_title_contains", "value": "${expectedTitle}", "enabled": True},
            ],
        }

        sample_run_dir = Path("runs") / "sample_report_example"
        sample_run_dir.mkdir(parents=True, exist_ok=True)
        (sample_run_dir / "screenshots").mkdir(exist_ok=True)
        sample_run = {
            "runId": "sample_report_example",
            "kind": "test_case",
            "name": "POS_DEMO_001",
            "status": "passed",
            "startedAt": datetime.now().isoformat(timespec="seconds"),
            "endedAt": datetime.now().isoformat(timespec="seconds"),
            "durationSeconds": 1.0,
            "datasetRowIndex": 0,
            "stepResults": [
                {"stepIndex": 1, "stepType": "run_flow", "status": "passed", "message": "Sample", "screenshot": ""},
                {
                    "stepIndex": 2,
                    "stepType": "assert_window_title_contains",
                    "status": "passed",
                    "message": "Sample",
                    "screenshot": "",
                },
            ],
        }
        (sample_run_dir / "run.json").write_text(json.dumps(sample_run, indent=2), encoding="utf-8")
        (sample_run_dir / "report.html").write_text(
            "<html><body><h1>Sample Report</h1><p>Status: PASSED</p><p>This is a sample artifact.</p></body></html>",
            encoding="utf-8",
        )
        runs = self.data.setdefault("runs", [])
        runs = [r for r in runs if str(r.get("runId")) != sample_run["runId"]]
        runs.append(
            {
                "runId": sample_run["runId"],
                "kind": sample_run["kind"],
                "name": sample_run["name"],
                "status": sample_run["status"],
                "startedAt": sample_run["startedAt"],
                "endedAt": sample_run["endedAt"],
                "durationSeconds": sample_run["durationSeconds"],
                "runFolder": str(sample_run_dir),
            }
        )
        self.data["runs"] = runs

        self.save_project()
        self.show_page(self.current_page)
        self.append_log("Sample project content added.")
        messagebox.showinfo("Sample Created", "Sample targets, flow, dataset, and test case were added.")

    # =====================================================
    # Run actions
    # =====================================================

    def run_selected(self, kind: str, dry_run: bool) -> None:
        if kind == "flow":
            if self.current_entity_kind != "flow" or not self.current_entity_name:
                messagebox.showwarning("No Flow", "Select a flow first.")
                return
            name = self.current_entity_name
            dataset_idx = None
        else:
            if self.current_entity_kind != "test_case" or not self.current_entity_name:
                messagebox.showwarning("No Test Case", "Select a test case first.")
                return
            name = self.current_entity_name
            dataset_idx = None
            dataset_name = str(self.data["testCases"][name].get("dataset", "")).strip()
            if dataset_name:
                idx = simpledialog.askinteger(
                    "Dataset Row",
                    f"Dataset '{dataset_name}' row index (0-based):",
                    initialvalue=0,
                    minvalue=0,
                )
                if idx is None:
                    return
                dataset_idx = idx

        delay = float(self.data.get("settings", {}).get("startupDelaySeconds", 3))
        if not dry_run:
            if pyautogui is None:
                messagebox.showerror(
                    "Missing Dependency",
                    "pyautogui is not installed. Use Dry Run or install pyautogui for live execution.",
                )
                return
            if not messagebox.askyesno(
                "Run",
                f"Run {kind.replace('_', ' ')} '{name}' after {delay:.1f}s?\n\n"
                "Move mouse to top-left to trigger PyAutoGUI fail-safe.",
            ):
                return

        thread = threading.Thread(
            target=self._run_thread,
            args=(kind, name, dry_run, delay, dataset_idx),
            daemon=True,
        )
        thread.start()

    def _run_thread(self, kind: str, name: str, dry_run: bool, delay: float, dataset_idx: int | None) -> None:
        self.set_status(f"Running {kind}:{name}")
        self.append_log(f"Starting {kind} '{name}' (dry_run={dry_run})")

        try:
            if not dry_run and delay > 0:
                self.append_log(f"Waiting {delay:.2f}s before execution...")
                time.sleep(delay)

            self.active_runner = TestFlowRunner(self.data, log_callback=self.append_log)
            self.active_runner.reset_stop()

            if kind == "flow":
                result = self.active_runner.run_flow(name, dry_run=dry_run)
            else:
                result = self.active_runner.run_test_case(name, dry_run=dry_run, dataset_row_index=dataset_idx)

            self.save_project()
            self.append_log(
                f"Run finished with status={result.get('status')} folder={result.get('runFolder', '')}"
            )

            run_folder = result.get("runFolder", "")
            report = Path(run_folder) / "report.html"
            if report.exists():
                self.root.after(
                    0,
                    lambda: messagebox.showinfo(
                        "Run Completed",
                        f"Status: {result.get('status')}\nRun folder:\n{run_folder}",
                    ),
                )

            if self.current_page == "Run Center":
                self.root.after(0, self.refresh_runs_tree)

        except RunnerExecutionError as exc:
            self.append_log(f"Validation/Run error: {exc}")
            self.root.after(0, lambda: messagebox.showerror("Run Error", str(exc)))
        except Exception as exc:
            self._log_exception("Run thread failure")
            self.append_log(f"Run failure: {exc}")
            self.root.after(0, lambda: messagebox.showerror("Run Error", str(exc)))
        finally:
            self.active_runner = None
            self.set_status("Idle")

    def stop_run(self) -> None:
        if self.recorder.is_running:
            self.stop_recording()
            self.append_log("Stop requested for recorder.")
            return
        if self.active_runner is None:
            self.append_log("No active run.")
            return
        self.active_runner.request_stop()
        self.append_log("Stop requested. Runner will stop before next step.")


def main() -> None:
    root = tk.Tk()
    AutomationApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
