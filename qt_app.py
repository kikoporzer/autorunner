import copy
import json
import os
import importlib
import subprocess
import sys
import threading
import time
import traceback
import webbrowser
import zipfile
from datetime import datetime
from pathlib import Path

from PySide6.QtCore import QObject, Qt, Signal
from PySide6.QtGui import QAction, QKeySequence
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFormLayout,
    QFrame,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QLineEdit,
    QListWidget,
    QMainWindow,
    QMessageBox,
    QDialog,
    QDialogButtonBox,
    QPushButton,
    QPlainTextEdit,
    QSplitter,
    QStackedWidget,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QToolBar,
    QVBoxLayout,
    QWidget,
    QSpinBox,
    QDoubleSpinBox,
)

try:
    import pyautogui
except ModuleNotFoundError:
    pyautogui = None

from importer import ImporterError, parse_test_case_rows, read_table_rows
from recorder import GlobalClickRecorder
from runner import RunnerExecutionError, TestFlowRunner
from storage import PROJECT_FILE, load_project, normalize_project_data, save_project as storage_save_project


class UiBridge(QObject):
    log = Signal(str)
    status = Signal(str)
    run_finished = Signal(dict)
    recorder_finished = Signal(list, str)


class TestFlowRunnerQt(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("TestFlow Runner")
        self.resize(1620, 980)

        self.logs_dir = Path("logs")
        self.logs_dir.mkdir(parents=True, exist_ok=True)
        self.project_path = PROJECT_FILE.resolve()
        self.data, messages = load_project()
        self.last_saved_at: datetime | None = None

        self.current_entity_kind: str | None = None  # flow | test_case
        self.current_entity_name: str | None = None
        self.current_step_index: int | None = None
        self.active_runner: TestFlowRunner | None = None
        self.recording_steps: list[dict] = []
        self.recorder = GlobalClickRecorder()

        self.bridge = UiBridge()
        self.bridge.log.connect(self.append_log)
        self.bridge.status.connect(self.set_status)
        self.bridge.run_finished.connect(self._on_run_finished_ui)
        self.bridge.recorder_finished.connect(self._on_recorder_finished_ui)

        if pyautogui is not None:
            pyautogui.FAILSAFE = True
            pyautogui.PAUSE = float(self.data.get("settings", {}).get("defaultActionPauseSeconds", 0.1))

        self._build_ui()
        self._apply_theme()
        self._bind_shortcuts()
        self._refresh_all()
        self._run_dependency_guard()
        self._run_startup_checks()
        for msg in messages:
            self.append_log(msg)

    # ---------------------------
    # UI setup
    # ---------------------------
    def _apply_theme(self) -> None:
        self.setStyleSheet(
            """
            QMainWindow { background: #f3f6fb; color: #1f2937; }
            QWidget {
                background: #f3f6fb; color: #1f2937; font-size: 14px;
                font-family: Inter, "Segoe UI Variable", "Segoe UI", "SF Pro Text", "Helvetica Neue", Arial, sans-serif;
                line-height: 1.4;
            }
            QFrame#panel, QGroupBox {
                background: #ffffff; border: 1px solid #d8e1ef; border-radius: 10px;
            }
            QGroupBox { margin-top: 8px; padding-top: 12px; }
            QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 4px; color: #60708f; }
            QPushButton {
                background: #f8faff; border: 1px solid #d5dff0; border-radius: 8px; padding: 7px 12px;
                color: #1f2937;
            }
            QPushButton:hover { background: #eef4ff; border-color: #bfd0ee; }
            QPushButton:pressed { background: #e3edff; }
            QLineEdit, QTextEdit, QPlainTextEdit, QComboBox, QListWidget, QTableWidget {
                background: #ffffff; border: 1px solid #d5dff0; border-radius: 8px;
                selection-background-color: #d8e6ff;
                selection-color: #1f2937;
                color: #1f2937;
            }
            QHeaderView::section {
                background: #f4f7fc; border: 0; border-bottom: 1px solid #d8e1ef;
                padding: 7px; color: #475569; font-weight: 600;
            }
            QTableWidget::item { padding: 4px; }
            QLabel#title { font-size: 22px; font-weight: 700; letter-spacing: 0px; }
            QLabel#muted { color: #64748b; }
            QToolBar {
                background: #ffffff; border-bottom: 1px solid #d8e1ef;
                spacing: 8px; padding: 8px;
            }
            QListWidget::item { padding: 8px; border-radius: 6px; margin: 2px 4px; }
            QListWidget::item:selected { background: #e8f0ff; color: #1f2937; }
            QScrollBar:vertical {
                background: #f4f7fc; width: 10px; margin: 0px; border-radius: 5px;
            }
            QScrollBar::handle:vertical {
                background: #c6d4ea; min-height: 24px; border-radius: 5px;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                height: 0px; background: none; border: none;
            }
            """
        )

    def _build_ui(self) -> None:
        self._build_toolbar()

        root = QWidget()
        root_layout = QVBoxLayout(root)
        root_layout.setContentsMargins(10, 10, 10, 10)
        root_layout.setSpacing(10)
        self.setCentralWidget(root)

        # top status row
        top = QFrame()
        top.setObjectName("panel")
        top_layout = QHBoxLayout(top)
        self.title_label = QLabel("TestFlow Runner")
        self.title_label.setObjectName("title")
        self.env_label = QLabel("Environment: Default")
        self.env_label.setObjectName("muted")
        self.status_label = QLabel("Idle")
        self.status_label.setObjectName("muted")
        self.page_label = QLabel("Home")
        self.page_label.setObjectName("muted")
        self.project_label = QLabel(f"Project: {self.project_path}")
        self.project_label.setObjectName("muted")
        self.saved_label = QLabel("Last saved: not yet")
        self.saved_label.setObjectName("muted")
        top_layout.addWidget(self.title_label)
        top_layout.addSpacing(12)
        top_layout.addWidget(self.env_label)
        top_layout.addStretch(1)
        top_layout.addWidget(self.status_label)
        top_layout.addSpacing(8)
        top_layout.addWidget(self.page_label)
        top_layout.addSpacing(8)
        top_layout.addWidget(self.saved_label)
        root_layout.addWidget(top)

        center_split = QSplitter(Qt.Horizontal)
        root_layout.addWidget(center_split, stretch=1)

        # nav
        nav = QFrame()
        nav.setObjectName("panel")
        nav_layout = QVBoxLayout(nav)
        nav_layout.addWidget(QLabel("Quick Start"))
        self.nav_list = QListWidget()
        for page in [
            "Home",
            "Record Flow",
            "Import Excel",
            "Flow Builder",
            "Exports",
            "Advanced",
        ]:
            self.nav_list.addItem(self._label_with_icon(page))
        self.nav_list.currentTextChanged.connect(self._switch_page)
        nav_layout.addWidget(self.nav_list)
        center_split.addWidget(nav)

        # main + inspector
        right_split = QSplitter(Qt.Horizontal)
        center_split.addWidget(right_split)
        center_split.setSizes([260, 1320])

        self.pages = QStackedWidget()
        right_split.addWidget(self.pages)
        self.inspector = self._build_inspector()
        right_split.addWidget(self.inspector)
        right_split.setSizes([980, 340])

        self._build_pages()

        # run log
        log_panel = QFrame()
        log_panel.setObjectName("panel")
        log_layout = QVBoxLayout(log_panel)
        row = QHBoxLayout()
        row.addWidget(QLabel("Run Log"))
        row.addStretch(1)
        b_clear = QPushButton("Clear")
        b_clear.clicked.connect(lambda: self.log_text.setPlainText(""))
        b_save = QPushButton("Save")
        b_save.clicked.connect(self._save_log)
        row.addWidget(b_clear)
        row.addWidget(b_save)
        log_layout.addLayout(row)
        self.log_text = QPlainTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumBlockCount(5000)
        log_layout.addWidget(self.log_text)
        root_layout.addWidget(log_panel, stretch=0)

        self.nav_list.setCurrentRow(0)

    def _build_toolbar(self) -> None:
        bar = QToolBar("Main")
        bar.setMovable(False)
        self.addToolBar(Qt.TopToolBarArea, bar)

        a_save = QAction("💾 Save", self)
        a_save.triggered.connect(self.save_project)
        bar.addAction(a_save)

        a_record = QAction("⏺ Record Flow", self)
        a_record.triggered.connect(lambda: self._switch_page("Record Flow"))
        bar.addAction(a_record)

        a_import = QAction("📥 Import Excel/CSV", self)
        a_import.triggered.connect(self._quick_import_excel)
        bar.addAction(a_import)

        a_flows = QAction("🧩 Flows", self)
        a_flows.triggered.connect(lambda: self._switch_page("Flow Builder"))
        bar.addAction(a_flows)

        a_runs = QAction("▶ Runs", self)
        a_runs.triggered.connect(lambda: self._switch_page("Exports"))
        bar.addAction(a_runs)

        a_stop = QAction("⏹ Stop", self)
        a_stop.triggered.connect(self.stop_run)
        bar.addAction(a_stop)

    def _build_pages(self) -> None:
        self.page_home = self._page_home()
        self.page_import = self._page_import()
        self.page_test_cases = self._page_test_cases()
        self.page_flows = self._page_flows()
        self.page_targets = self._page_targets()
        self.page_recorder = self._page_recorder()
        self.page_datasets = self._page_datasets()
        self.page_runs = self._page_run_center()
        self.page_settings = self._page_settings()
        self.page_advanced = self._page_advanced()
        self.page_exports = self._page_exports()

        for p in [
            self.page_home,
            self.page_recorder,
            self.page_import,
            self.page_flows,
            self.page_exports,
            self.page_advanced,
            self.page_test_cases,
            self.page_targets,
            self.page_datasets,
            self.page_runs,
            self.page_settings,
        ]:
            self.pages.addWidget(p)

    def _build_inspector(self) -> QFrame:
        panel = QFrame()
        panel.setObjectName("panel")
        layout = QVBoxLayout(panel)
        layout.addWidget(QLabel("Step Inspector"))

        form = QFormLayout()
        self.ins_type = QLineEdit()
        self.ins_enabled = QCheckBox("Enabled")
        self.ins_target = QLineEdit()
        self.ins_value = QLineEdit()
        self.ins_seconds = QLineEdit()
        self.ins_desc = QLineEdit()
        form.addRow("Type", self.ins_type)
        form.addRow("", self.ins_enabled)
        form.addRow("Target/Flow", self.ins_target)
        form.addRow("Value/Path/Key", self.ins_value)
        form.addRow("Seconds", self.ins_seconds)
        form.addRow("Description", self.ins_desc)
        layout.addLayout(form)

        b_apply = QPushButton("Apply")
        b_apply.clicked.connect(self.apply_step_from_inspector)
        layout.addWidget(b_apply)

        layout.addWidget(QLabel("Raw JSON"))
        self.ins_raw = QTextEdit()
        layout.addWidget(self.ins_raw, stretch=1)
        return panel

    # ---------------------------
    # Pages
    # ---------------------------
    def _page_home(self) -> QWidget:
        w = QWidget()
        l = QVBoxLayout(w)
        title = QLabel("Welcome")
        title.setObjectName("title")
        l.addWidget(title)
        subtitle = QLabel("Create or import a flow in under a minute")
        subtitle.setObjectName("muted")
        l.addWidget(subtitle)

        quick = QFrame()
        quick.setObjectName("panel")
        ql = QGridLayout(quick)
        b_record = QPushButton("⏺ Record Flow")
        b_record.setMinimumHeight(56)
        b_record.clicked.connect(lambda: self._switch_page("Record Flow"))
        b_import = QPushButton("📥 Import Excel/CSV")
        b_import.setMinimumHeight(56)
        b_import.clicked.connect(self._quick_import_excel)
        b_flows = QPushButton("🧩 Open Flows")
        b_flows.setMinimumHeight(56)
        b_flows.clicked.connect(lambda: self._switch_page("Flow Builder"))
        b_runs = QPushButton("▶ Open Runs")
        b_runs.setMinimumHeight(56)
        b_runs.clicked.connect(lambda: self._switch_page("Exports"))
        ql.addWidget(b_record, 0, 0)
        ql.addWidget(b_import, 0, 1)
        ql.addWidget(b_flows, 1, 0)
        ql.addWidget(b_runs, 1, 1)
        l.addWidget(quick)

        cards = QFrame()
        cards.setObjectName("panel")
        gl = QGridLayout(cards)
        self.card_flows = QLabel("Flows: 0")
        self.card_cases = QLabel("Test Cases: 0")
        self.card_targets = QLabel("Targets: 0")
        self.card_runs = QLabel("Runs: 0")
        gl.addWidget(self.card_flows, 0, 0)
        gl.addWidget(self.card_cases, 0, 1)
        gl.addWidget(self.card_targets, 0, 2)
        gl.addWidget(self.card_runs, 0, 3)
        l.addWidget(cards)

        group = QGroupBox("Recent Runs")
        vg = QVBoxLayout(group)
        self.dash_runs_table = QTableWidget(0, 5)
        self.dash_runs_table.setHorizontalHeaderLabels(["Started", "Kind", "Name", "Status", "Duration"])
        self.dash_runs_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.dash_runs_table.setSelectionBehavior(QTableWidget.SelectRows)
        vg.addWidget(self.dash_runs_table)
        l.addWidget(group, stretch=1)
        return w

    def _page_import(self) -> QWidget:
        w = QWidget()
        l = QVBoxLayout(w)
        title = QLabel("Import Excel/CSV")
        title.setObjectName("title")
        l.addWidget(title)
        subtitle = QLabel("Bring in test steps from Excel/CSV, or load a dataset.")
        subtitle.setObjectName("muted")
        l.addWidget(subtitle)

        panel = QFrame()
        panel.setObjectName("panel")
        gl = QGridLayout(panel)
        b_cases = QPushButton("Import Test Cases")
        b_cases.setMinimumHeight(56)
        b_cases.clicked.connect(self.import_test_cases_file)
        b_data = QPushButton("Import Dataset")
        b_data.setMinimumHeight(56)
        b_data.clicked.connect(self.import_dataset_file)
        b_preview = QPushButton("Open Datasets (Advanced)")
        b_preview.setMinimumHeight(48)
        b_preview.clicked.connect(lambda: self._switch_page("Advanced"))
        gl.addWidget(b_cases, 0, 0)
        gl.addWidget(b_data, 0, 1)
        gl.addWidget(b_preview, 1, 0, 1, 2)
        l.addWidget(panel)
        l.addStretch(1)
        return w

    def _page_exports(self) -> QWidget:
        w = QWidget()
        l = QVBoxLayout(w)
        title = QLabel("Exports")
        title.setObjectName("title")
        l.addWidget(title)
        subtitle = QLabel("Open HTML reports quickly, or export your full project ZIP.")
        subtitle.setObjectName("muted")
        l.addWidget(subtitle)

        row = QHBoxLayout()
        b_refresh = QPushButton("Refresh Runs")
        b_refresh.clicked.connect(self.refresh_runs_table)
        b_open = QPushButton("Open Selected HTML Report")
        b_open.clicked.connect(self.open_selected_report)
        b_zip = QPushButton("Export Project ZIP")
        b_zip.clicked.connect(self.export_project_zip)
        row.addWidget(b_refresh)
        row.addWidget(b_open)
        row.addWidget(b_zip)
        row.addStretch(1)
        l.addLayout(row)

        l.addWidget(self._page_run_center(), stretch=1)
        return w

    def _page_advanced(self) -> QWidget:
        w = QWidget()
        l = QVBoxLayout(w)
        title = QLabel("Advanced")
        title.setObjectName("title")
        l.addWidget(title)
        subtitle = QLabel("Everything beyond the day-to-day workflow")
        subtitle.setObjectName("muted")
        l.addWidget(subtitle)

        tools = QFrame()
        tools.setObjectName("panel")
        tl = QGridLayout(tools)
        actions = [
            ("Open Test Cases", lambda: self._switch_page("Test Cases")),
            ("Open Targets", lambda: self._switch_page("Targets")),
            ("Open Datasets", lambda: self._switch_page("Datasets")),
            ("Open Settings", lambda: self._switch_page("Settings")),
            ("Export Project ZIP", self.export_project_zip),
            ("Import Project JSON", self.import_project_json),
            ("Import Project ZIP", self.import_project_zip),
            ("Import Project Folder", self.import_project_folder),
            ("Create Sample Project", self.create_sample_project),
            ("Run Startup Checks", self._run_startup_checks),
        ]
        for i, (label, fn) in enumerate(actions):
            b = QPushButton(label)
            b.clicked.connect(fn)
            tl.addWidget(b, i // 2, i % 2)
        l.addWidget(tools)
        l.addStretch(1)
        return w

    def _build_entity_editor(self, kind: str) -> QWidget:
        w = QWidget()
        l = QVBoxLayout(w)
        header = QHBoxLayout()
        title = QLabel("Test Cases" if kind == "test_case" else "Flows")
        title.setObjectName("title")
        header.addWidget(title)
        header.addStretch(1)
        b_run = QPushButton("▶ Run")
        b_run.clicked.connect(lambda: self.run_selected(False))
        b_dry = QPushButton("🧪 Dry Run")
        b_dry.clicked.connect(lambda: self.run_selected(True))
        b_from = QPushButton("⏩ Run From Selected Step")
        b_from.clicked.connect(self.run_from_selected_step)
        b_once = QPushButton("👣 Step Once")
        b_once.clicked.connect(self.run_step_once)
        b_new = QPushButton("＋ New")
        b_new.clicked.connect(lambda: self._new_entity(kind))
        b_dup = QPushButton("⧉ Duplicate")
        b_dup.clicked.connect(lambda: self._duplicate_entity(kind))
        b_del = QPushButton("🗑 Delete")
        b_del.clicked.connect(lambda: self._delete_entity(kind))
        for b in [b_run, b_dry, b_from, b_once, b_new, b_dup, b_del]:
            header.addWidget(b)
        if kind == "test_case":
            b_import = QPushButton("📥 Import CSV/XLSX")
            b_import.clicked.connect(self.import_test_cases_file)
            b_all = QPushButton("▶ Run All Rows")
            b_all.clicked.connect(self.run_selected_test_case_all_rows)
            header.addWidget(b_import)
            header.addWidget(b_all)
        l.addLayout(header)

        split = QSplitter(Qt.Horizontal)
        l.addWidget(split, stretch=1)
        left = QFrame()
        left.setObjectName("panel")
        ll = QVBoxLayout(left)
        ll.addWidget(QLabel("Items"))
        list_widget = QListWidget()
        ll.addWidget(list_widget)
        split.addWidget(left)

        right = QFrame()
        right.setObjectName("panel")
        rl = QVBoxLayout(right)

        meta = QGroupBox("Metadata")
        mf = QFormLayout(meta)
        name = QLineEdit()
        suite = QLineEdit()
        dataset = QLineEdit()
        enabled = QCheckBox("Enabled")
        enabled.setChecked(True)
        mf.addRow("Name", name)
        if kind == "test_case":
            mf.addRow("Suite", suite)
            mf.addRow("Dataset", dataset)
        mf.addRow("", enabled)
        b_meta = QPushButton("Save Metadata")
        b_meta.clicked.connect(lambda: self._save_entity_meta(kind))
        mf.addRow("", b_meta)
        rl.addWidget(meta)

        steps_toolbar = QHBoxLayout()
        if kind == "flow":
            for label, step in [
                ("Click", {"type": "click", "target": "", "enabled": True}),
                ("Type", {"type": "type_text", "value": "", "enabled": True}),
                ("Wait", {"type": "wait", "seconds": 1.0, "enabled": True}),
                ("Screenshot", {"type": "screenshot", "name": "step", "enabled": True}),
                ("Subflow", {"type": "run_flow", "flow": "", "enabled": True}),
            ]:
                bq = QPushButton(label)
                bq.clicked.connect(lambda _=False, s=step: self._add_quick_step(s))
                steps_toolbar.addWidget(bq)
        for label, fn in [
            ("＋ Add Step", self._add_step),
            ("⧉ Duplicate Step", self._duplicate_step),
            ("🗑 Delete Step", self._delete_step),
            ("↑ Up", lambda: self._move_step(-1)),
            ("↓ Down", lambda: self._move_step(1)),
        ]:
            b = QPushButton(label)
            b.clicked.connect(fn)
            steps_toolbar.addWidget(b)
        steps_toolbar.addStretch(1)
        rl.addLayout(steps_toolbar)

        table = QTableWidget(0, 7)
        table.setHorizontalHeaderLabels(["#", "Enabled", "Type", "Target/Flow", "Value", "Wait", "Description"])
        table.setSelectionBehavior(QTableWidget.SelectRows)
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        table.itemSelectionChanged.connect(self._on_step_selected)
        if kind == "flow":
            table.itemDoubleClicked.connect(lambda _item: self._guided_edit_selected_step())
        rl.addWidget(table, stretch=1)
        split.addWidget(right)
        split.setSizes([280, 860])

        if kind == "test_case":
            preview_box = QGroupBox("Variable Preview")
            pv = QVBoxLayout(preview_box)
            b_prev = QPushButton("Preview Variables")
            b_prev.clicked.connect(self.preview_selected_test_case_variables)
            pv.addWidget(b_prev)
            preview = QPlainTextEdit()
            pv.addWidget(preview)
            rl.addWidget(preview_box)
            self.tc_var_preview = preview

        if kind == "test_case":
            self.tc_list = list_widget
            self.tc_name = name
            self.tc_suite = suite
            self.tc_dataset = dataset
            self.tc_enabled = enabled
            self.tc_steps_table = table
            self.tc_list.currentTextChanged.connect(lambda _: self._on_entity_select("test_case"))
        else:
            self.flow_list = list_widget
            self.flow_name = name
            self.flow_enabled = enabled
            self.flow_steps_table = table
            self.flow_list.currentTextChanged.connect(lambda _: self._on_entity_select("flow"))
        return w

    def _page_test_cases(self) -> QWidget:
        return self._build_entity_editor("test_case")

    def _page_flows(self) -> QWidget:
        return self._build_entity_editor("flow")

    def _page_targets(self) -> QWidget:
        w = QWidget()
        l = QVBoxLayout(w)
        row = QHBoxLayout()
        title = QLabel("Targets")
        title.setObjectName("title")
        row.addWidget(title)
        row.addStretch(1)
        for text, fn in [
            ("🎯 Capture", self.capture_target),
            ("＋ Manual Add", self.manual_add_target),
            ("✎ Rename", self.rename_target),
            ("🗑 Delete", self.delete_target),
            ("🖱 Test Click", self.test_target),
        ]:
            b = QPushButton(text)
            b.clicked.connect(fn)
            row.addWidget(b)
        l.addLayout(row)
        self.targets_table = QTableWidget(0, 4)
        self.targets_table.setHorizontalHeaderLabels(["Name", "X", "Y", "Description"])
        self.targets_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.targets_table.setEditTriggers(QTableWidget.NoEditTriggers)
        l.addWidget(self.targets_table, stretch=1)
        return w

    def _page_recorder(self) -> QWidget:
        w = QWidget()
        l = QVBoxLayout(w)
        title = QLabel("Recorder")
        title.setObjectName("title")
        l.addWidget(title)
        options = QGroupBox("Options")
        gl = QGridLayout(options)
        self.rec_opt_left = QCheckBox("Record left clicks")
        self.rec_opt_left.setChecked(True)
        self.rec_opt_double = QCheckBox("Record double clicks")
        self.rec_opt_double.setChecked(True)
        self.rec_opt_right = QCheckBox("Record right clicks")
        self.rec_opt_right.setChecked(True)
        self.rec_opt_timing = QCheckBox("Record timing gaps")
        self.rec_opt_timing.setChecked(True)
        self.rec_opt_typing = QCheckBox("Record typing")
        self.rec_opt_hotkeys = QCheckBox("Record hotkeys")
        for i, c in enumerate(
            [
                self.rec_opt_left,
                self.rec_opt_double,
                self.rec_opt_right,
                self.rec_opt_timing,
                self.rec_opt_typing,
                self.rec_opt_hotkeys,
            ]
        ):
            gl.addWidget(c, i // 3, i % 3)
        l.addWidget(options)

        controls = QHBoxLayout()
        self.rec_status = QLabel("Idle")
        controls.addWidget(self.rec_status)
        controls.addStretch(1)
        b_start = QPushButton("⏺ Start Recording")
        b_start.clicked.connect(self.start_recording)
        b_stop = QPushButton("⏹ Stop Recording")
        b_stop.clicked.connect(self.stop_recording)
        controls.addWidget(b_start)
        controls.addWidget(b_stop)
        l.addLayout(controls)

        self.rec_table = QTableWidget(0, 3)
        self.rec_table.setHorizontalHeaderLabels(["#", "Type", "Details"])
        self.rec_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.rec_table.setSelectionBehavior(QTableWidget.SelectRows)
        l.addWidget(self.rec_table, stretch=1)

        actions = QHBoxLayout()
        for text, fn in [
            ("💾 Save as Reusable Flow", self.save_recording_as_flow),
            ("💾 Save as Test Case", self.save_recording_as_test_case),
            ("➕ Append to Existing Flow", self.append_recording_to_flow),
            ("🗑 Discard", self.discard_recording),
        ]:
            b = QPushButton(text)
            b.clicked.connect(fn)
            actions.addWidget(b)
        actions.addStretch(1)
        l.addLayout(actions)
        return w

    def _page_datasets(self) -> QWidget:
        w = QWidget()
        l = QVBoxLayout(w)
        row = QHBoxLayout()
        title = QLabel("Datasets")
        title.setObjectName("title")
        row.addWidget(title)
        row.addStretch(1)
        for text, fn in [
            ("📥 Import CSV/XLSX", self.import_dataset_file),
            ("✎ Rename", self.rename_dataset),
            ("🗑 Delete", self.delete_dataset),
        ]:
            b = QPushButton(text)
            b.clicked.connect(fn)
            row.addWidget(b)
        l.addLayout(row)
        split = QSplitter(Qt.Horizontal)
        self.dataset_list = QListWidget()
        self.dataset_list.currentTextChanged.connect(self.on_dataset_select)
        split.addWidget(self.dataset_list)
        self.dataset_preview = QPlainTextEdit()
        split.addWidget(self.dataset_preview)
        split.setSizes([300, 900])
        l.addWidget(split, stretch=1)
        return w

    def _page_run_center(self) -> QWidget:
        w = QWidget()
        l = QVBoxLayout(w)
        row = QHBoxLayout()
        title = QLabel("Run Center")
        title.setObjectName("title")
        row.addWidget(title)
        row.addStretch(1)
        b_ref = QPushButton("↻ Refresh")
        b_ref.clicked.connect(self.refresh_runs_table)
        b_open = QPushButton("📄 Open Selected Report")
        b_open.clicked.connect(self.open_selected_report)
        row.addWidget(b_ref)
        row.addWidget(b_open)
        l.addLayout(row)
        self.runs_table = QTableWidget(0, 7)
        self.runs_table.setHorizontalHeaderLabels(
            ["Run ID", "Started", "Kind", "Name", "Status", "Duration", "Run Folder"]
        )
        self.runs_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.runs_table.setEditTriggers(QTableWidget.NoEditTriggers)
        l.addWidget(self.runs_table, stretch=1)
        return w

    def _page_settings(self) -> QWidget:
        w = QWidget()
        l = QVBoxLayout(w)
        title = QLabel("Settings")
        title.setObjectName("title")
        l.addWidget(title)
        form_box = QGroupBox("Runner Settings")
        form = QFormLayout(form_box)
        self.set_start_delay = QLineEdit()
        self.set_action_pause = QLineEdit()
        self.set_post_action_delay = QLineEdit()
        self.set_screenshot_delay = QLineEdit()
        self.set_stop_hotkey = QLineEdit()
        self.set_record_stop_hotkey = QLineEdit()
        self.set_expected_res = QLineEdit()
        self.set_shot_fail = QCheckBox("Screenshot on failure")
        self.set_shot_each = QCheckBox("Screenshot after each step")
        form.addRow("Startup delay (s)", self.set_start_delay)
        form.addRow("Default action pause (s)", self.set_action_pause)
        form.addRow("Post-action delay (s)", self.set_post_action_delay)
        form.addRow("Before screenshot delay (s)", self.set_screenshot_delay)
        form.addRow("Stop hotkey", self.set_stop_hotkey)
        form.addRow("Recording stop hotkey", self.set_record_stop_hotkey)
        form.addRow("Expected resolution", self.set_expected_res)
        form.addRow("", self.set_shot_fail)
        form.addRow("", self.set_shot_each)
        b_save = QPushButton("💾 Save Settings")
        b_save.clicked.connect(self.save_settings)
        form.addRow("", b_save)
        l.addWidget(form_box)

        tools = QGroupBox("Project Tools")
        tl = QHBoxLayout(tools)
        for text, fn in [
            ("📦 Export Project ZIP", self.export_project_zip),
            ("📥 Import Project JSON", self.import_project_json),
            ("📥 Import Project ZIP", self.import_project_zip),
            ("📁 Import Project Folder", self.import_project_folder),
            ("✨ Create Sample Project", self.create_sample_project),
            ("🩺 Run Startup Checks", self._run_startup_checks),
        ]:
            b = QPushButton(text)
            b.clicked.connect(fn)
            tl.addWidget(b)
        l.addWidget(tools)
        l.addStretch(1)
        return w

    # ---------------------------
    # General helpers
    # ---------------------------
    def _bind_shortcuts(self) -> None:
        for seq, fn in [
            (QKeySequence("Ctrl+S"), self.save_project),
            (QKeySequence("F5"), lambda: self.run_selected(False)),
            (QKeySequence("Shift+F5"), lambda: self.run_selected(True)),
            (QKeySequence("F8"), self.stop_run),
            (QKeySequence("Ctrl+D"), self._duplicate_step),
            (QKeySequence("Delete"), self._delete_step),
            (QKeySequence("F6"), self.run_step_once),
            (QKeySequence("Shift+F6"), self.run_from_selected_step),
        ]:
            action = QAction(self)
            action.setShortcut(seq)
            action.triggered.connect(fn)
            self.addAction(action)

    def _refresh_all(self) -> None:
        self._refresh_dashboard()
        self._refresh_test_case_list()
        self._refresh_flow_list()
        self._refresh_targets()
        self._refresh_datasets()
        self.refresh_runs_table()
        self._load_settings()
        self.env_label.setText(f"Environment: {self.data.get('environment', {}).get('name', 'Default')}")

    def _switch_page(self, page: str) -> None:
        page = self._normalize_nav_label(page)
        index_map = {
            "Home": 0,
            "Record Flow": 1,
            "Import Excel": 2,
            "Flow Builder": 3,
            "Exports": 4,
            "Advanced": 5,
            "Test Cases": 6,
            "Targets": 7,
            "Datasets": 8,
            "Runs": 9,
            "Settings": 10,
            "Recorder": 1,
            "Flows": 3,
            "Advanced Settings": 5,
        }
        i = index_map.get(page, 0)
        self.pages.setCurrentIndex(i)
        self.page_label.setText(f"Page: {page}")
        self.inspector.setVisible(page in {"Flow Builder", "Test Cases"})

    @staticmethod
    def _label_with_icon(page: str) -> str:
        icons = {
            "Home": "🏠",
            "Recorder": "⏺",
            "Flows": "🧩",
            "Runs": "▶",
            "Advanced Settings": "⚙",
            "Test Cases": "🧪",
            "Targets": "🎯",
            "Datasets": "🗂",
            "Settings": "⚙",
        }
        return f"{icons.get(page, '•')} {page}"

    @staticmethod
    def _normalize_nav_label(label: str) -> str:
        if not label:
            return ""
        parts = label.split(" ", 1)
        if len(parts) == 2:
            return parts[1].strip()
        return label.strip()

    def append_log(self, msg: str) -> None:
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_text.appendPlainText(f"[{ts}] {msg}")

    def set_status(self, text: str) -> None:
        self.status_label.setText(text)

    def _log_exception(self, context: str, exc_info=None) -> None:
        if exc_info is None:
            exc_info = sys.exc_info()
        path = self.logs_dir / "app_errors.log"
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with path.open("a", encoding="utf-8") as f:
            f.write(f"[{ts}] {context}\n")
            f.write("".join(traceback.format_exception(*exc_info)))
            f.write("\n")

    def _save_log(self) -> None:
        out, _ = QFileDialog.getSaveFileName(self, "Save run log", str(Path.cwd() / "run.log"), "Log Files (*.log)")
        if not out:
            return
        Path(out).write_text(self.log_text.toPlainText(), encoding="utf-8")

    def _run_startup_checks(self) -> None:
        checks, warns = [], []
        if sys.version_info < (3, 10):
            warns.append(f"Python {sys.version.split()[0]} detected, recommended 3.10+")
        else:
            checks.append(f"Python {sys.version.split()[0]} detected")
        if pyautogui is None:
            warns.append("pyautogui not installed (live execution disabled, dry run works).")
        else:
            checks.append("pyautogui available")
        status = self.recorder.availability()
        if status.available:
            checks.append("Recorder backend available (pynput)")
        else:
            warns.append(status.message)
        screen = QApplication.primaryScreen()
        if screen:
            size = screen.availableGeometry().size()
            current = f"{size.width()}x{size.height()}"
            checks.append(f"Screen resolution: {current}")
            expected = str(self.data.get("environment", {}).get("expectedResolution", "")).strip()
            if expected and expected != current:
                warns.append(f"Expected resolution '{expected}' differs from current '{current}'")
        checks.append(f"Project loaded: {self.project_path}")
        if os.access(self.project_path.parent, os.W_OK):
            checks.append(f"Write access OK: {self.project_path.parent}")
        else:
            warns.append(f"No write access to project folder: {self.project_path.parent}")
        for c in checks:
            self.append_log(f"Startup check: {c}")
        for w in warns:
            self.append_log(f"Startup warning: {w}")

    def _run_dependency_guard(self) -> None:
        checks = [
            ("Pillow", "PIL", "Pillow"),
            ("pyautogui", "pyautogui", "pyautogui"),
            ("pynput", "pynput", "pynput"),
        ]
        missing_core: list[tuple[str, str]] = []
        for display, module_name, pip_name in checks:
            try:
                importlib.import_module(module_name)
            except Exception:
                missing_core.append((display, pip_name))

        # Optional for Excel support only.
        try:
            importlib.import_module("openpyxl")
        except Exception:
            self.append_log("Optional dependency missing: openpyxl (Excel .xlsx import disabled; CSV still works).")

        if not missing_core:
            return

        names = ", ".join([n for n, _ in missing_core])
        pip_pkgs = " ".join([p for _, p in missing_core])
        msg = (
            "We need a few helpers before we can fully run.\n\n"
            f"Missing: {names}\n\n"
            "Click Yes to auto-install now.\n"
            "Click No if you want to install manually."
        )
        choice = QMessageBox.question(self, "Quick Setup", msg, QMessageBox.Yes | QMessageBox.No)
        if choice == QMessageBox.Yes:
            self.append_log(f"Installing missing dependencies: {pip_pkgs}")
            try:
                proc = subprocess.run(
                    [sys.executable, "-m", "pip", "install", *pip_pkgs.split()],
                    capture_output=True,
                    text=True,
                    timeout=300,
                )
                if proc.returncode == 0:
                    self.append_log("Dependency install succeeded. Restart the app to apply all packages cleanly.")
                    QMessageBox.information(
                        self,
                        "Setup Complete",
                        "Dependencies installed.\nPlease restart the app once so everything is loaded cleanly.",
                    )
                else:
                    self.append_log("Dependency install failed.")
                    details = (proc.stderr or proc.stdout or "").strip()[:1200]
                    QMessageBox.critical(
                        self,
                        "Setup Failed",
                        "Auto-install failed. Please run this command in terminal:\n\n"
                        f"{sys.executable} -m pip install {pip_pkgs}\n\n"
                        f"Error:\n{details}",
                    )
            except Exception as exc:
                self._log_exception("Dependency auto-install failed")
                QMessageBox.critical(
                    self,
                    "Setup Failed",
                    "Auto-install failed unexpectedly.\nRun this command manually:\n\n"
                    f"{sys.executable} -m pip install {pip_pkgs}\n\n{exc}",
                )
        else:
            QMessageBox.warning(
                self,
                "Manual Setup Needed",
                "Please run this command:\n\n"
                f"{sys.executable} -m pip install {pip_pkgs}\n\n"
                "Then restart the app.",
            )

    # ---------------------------
    # Dashboard + entity data
    # ---------------------------
    def _refresh_dashboard(self) -> None:
        self.card_flows.setText(f"Flows: {len(self.data.get('flows', {}))}")
        self.card_cases.setText(f"Test Cases: {len(self.data.get('testCases', {}))}")
        self.card_targets.setText(f"Targets: {len(self.data.get('targets', {}))}")
        runs = self.data.get("runs", [])
        self.card_runs.setText(f"Runs: {len(runs)}")
        recent = list(runs)[-50:][::-1]
        self.dash_runs_table.setRowCount(len(recent))
        for r, run in enumerate(recent):
            vals = [
                str(run.get("startedAt", "")),
                str(run.get("kind", "")),
                str(run.get("name", "")),
                str(run.get("status", "")),
                str(run.get("durationSeconds", "")),
            ]
            for c, v in enumerate(vals):
                self.dash_runs_table.setItem(r, c, QTableWidgetItem(v))

    def _entity_map(self, kind: str) -> dict:
        return self.data.get("flows", {}) if kind == "flow" else self.data.get("testCases", {})

    def _selected_entity_name(self, kind: str) -> str | None:
        item = self.flow_list.currentItem() if kind == "flow" else self.tc_list.currentItem()
        return item.text().split(" - ", 1)[0] if item else None

    def _refresh_test_case_list(self) -> None:
        self.tc_list.clear()
        for case_id, tc in sorted(self.data.get("testCases", {}).items()):
            self.tc_list.addItem(f"{case_id} - {tc.get('name', case_id)}")

    def _refresh_flow_list(self) -> None:
        self.flow_list.clear()
        for name, f in sorted(self.data.get("flows", {}).items()):
            self.flow_list.addItem(f"{name} - {f.get('description', '')}")

    def _on_entity_select(self, kind: str) -> None:
        name = self._selected_entity_name(kind)
        if not name:
            return
        self.current_entity_kind = kind
        self.current_entity_name = name
        self.current_step_index = None
        if kind == "test_case":
            tc = self.data["testCases"][name]
            self.tc_name.setText(str(tc.get("name", "")))
            self.tc_suite.setText(str(tc.get("suite", "")))
            self.tc_dataset.setText(str(tc.get("dataset", "")))
            self.tc_enabled.setChecked(bool(tc.get("enabled", True)))
            self._fill_steps_table(self.tc_steps_table, tc.get("steps", []))
        else:
            fl = self.data["flows"][name]
            self.flow_name.setText(str(fl.get("name", name)))
            self.flow_enabled.setChecked(True)
            self._fill_steps_table(self.flow_steps_table, fl.get("steps", []))
        self._load_inspector(None)

    def _steps_ref(self) -> list[dict]:
        if not self.current_entity_kind or not self.current_entity_name:
            return []
        entity = self._entity_map(self.current_entity_kind).get(self.current_entity_name, {})
        steps = entity.get("steps", [])
        return steps if isinstance(steps, list) else []

    def _fill_steps_table(self, table: QTableWidget, steps: list[dict]) -> None:
        table.setRowCount(len(steps))
        for i, step in enumerate(steps):
            vals = [
                str(i + 1),
                "Yes" if step.get("enabled", True) else "No",
                str(step.get("type", "")),
                str(step.get("target", "") or step.get("flow", "")),
                str(step.get("value", "") or step.get("path", "") or step.get("key", "")),
                str(step.get("seconds", "")) if "seconds" in step else "",
                str(step.get("description", "") or step.get("text", "")),
            ]
            for c, v in enumerate(vals):
                table.setItem(i, c, QTableWidgetItem(v))
        table.resizeColumnsToContents()

    def _active_table(self) -> QTableWidget | None:
        if self.current_entity_kind == "test_case":
            return self.tc_steps_table
        if self.current_entity_kind == "flow":
            return self.flow_steps_table
        return None

    def _on_step_selected(self) -> None:
        table = self._active_table()
        if table is None:
            return
        rows = table.selectionModel().selectedRows()
        if not rows:
            self.current_step_index = None
            self._load_inspector(None)
            return
        idx = rows[0].row()
        self.current_step_index = idx
        steps = self._steps_ref()
        if 0 <= idx < len(steps):
            # Keep flow builder visual-first: no raw inspector dependency for editing.
            if self.current_entity_kind == "flow":
                self._load_inspector(None)
            else:
                self._load_inspector(steps[idx])

    def _load_inspector(self, step: dict | None) -> None:
        if not step:
            self.ins_type.setText("")
            self.ins_enabled.setChecked(True)
            self.ins_target.setText("")
            self.ins_value.setText("")
            self.ins_seconds.setText("")
            self.ins_desc.setText("")
            self.ins_raw.setPlainText("")
            return
        self.ins_type.setText(str(step.get("type", "")))
        self.ins_enabled.setChecked(bool(step.get("enabled", True)))
        self.ins_target.setText(str(step.get("target", step.get("flow", ""))))
        value = step.get("value", "") or step.get("path", "") or step.get("key", "")
        self.ins_value.setText(str(value))
        self.ins_seconds.setText(str(step.get("seconds", "")))
        self.ins_desc.setText(str(step.get("description", "") or step.get("text", "")))
        self.ins_raw.setPlainText(json.dumps(step, indent=2))

    def apply_step_from_inspector(self) -> None:
        steps = self._steps_ref()
        if self.current_step_index is None or not (0 <= self.current_step_index < len(steps)):
            return
        try:
            parsed = json.loads(self.ins_raw.toPlainText().strip() or "{}")
            if not isinstance(parsed, dict):
                parsed = {}
        except Exception:
            parsed = copy.deepcopy(steps[self.current_step_index])
        parsed["type"] = self.ins_type.text().strip()
        parsed["enabled"] = bool(self.ins_enabled.isChecked())
        t = self.ins_target.text().strip()
        if parsed.get("type") == "run_flow":
            if t:
                parsed["flow"] = t
            parsed.pop("target", None)
        else:
            if t:
                parsed["target"] = t
            parsed.pop("flow", None)
        if self.ins_seconds.text().strip():
            try:
                parsed["seconds"] = float(self.ins_seconds.text().strip())
            except Exception:
                parsed["seconds"] = self.ins_seconds.text().strip()
        raw_val = self.ins_value.text()
        if raw_val:
            if parsed.get("type") == "press_key":
                parsed["key"] = raw_val
                parsed.pop("value", None)
                parsed.pop("path", None)
            elif parsed.get("type") == "assert_file_exists":
                parsed["path"] = raw_val
                parsed.pop("value", None)
                parsed.pop("key", None)
            else:
                parsed["value"] = raw_val
        d = self.ins_desc.text().strip()
        if d:
            if parsed.get("type") == "comment":
                parsed["text"] = d
            else:
                parsed["description"] = d
        steps[self.current_step_index] = parsed
        self.save_project()
        self._refresh_steps_after_change()

    def _refresh_steps_after_change(self) -> None:
        steps = self._steps_ref()
        table = self._active_table()
        if table is not None:
            self._fill_steps_table(table, steps)
            if self.current_step_index is not None and 0 <= self.current_step_index < len(steps):
                table.selectRow(self.current_step_index)

    def _new_entity(self, kind: str) -> None:
        if kind == "test_case":
            case_id, ok = QInputDialog.getText(self, "New Test Case", "Test case ID:")
            if not ok or not case_id.strip():
                return
            cid = case_id.strip()
            if cid in self.data.get("testCases", {}):
                QMessageBox.critical(self, "Error", "Test case already exists.")
                return
            self.data.setdefault("testCases", {})[cid] = {
                "id": cid,
                "name": cid,
                "suite": "",
                "description": "",
                "dataset": "",
                "enabled": True,
                "variables": {},
                "steps": [],
            }
            self._refresh_test_case_list()
        else:
            name, ok = QInputDialog.getText(self, "New Flow", "Flow name:")
            if not ok or not name.strip():
                return
            n = name.strip()
            if n in self.data.get("flows", {}):
                QMessageBox.critical(self, "Error", "Flow already exists.")
                return
            self.data.setdefault("flows", {})[n] = {"name": n, "description": "", "parameters": [], "steps": []}
            self._refresh_flow_list()
        self.save_project()

    def _duplicate_entity(self, kind: str) -> None:
        name = self._selected_entity_name(kind)
        if not name:
            return
        new_name, ok = QInputDialog.getText(self, "Duplicate", "New name/ID:", text=f"{name}_COPY")
        if not ok or not new_name.strip():
            return
        n = new_name.strip()
        m = self._entity_map(kind)
        if n in m:
            QMessageBox.critical(self, "Error", "Destination already exists.")
            return
        m[n] = copy.deepcopy(m[name])
        if kind == "test_case":
            m[n]["id"] = n
            self._refresh_test_case_list()
        else:
            m[n]["name"] = n
            self._refresh_flow_list()
        self.save_project()

    def _delete_entity(self, kind: str) -> None:
        name = self._selected_entity_name(kind)
        if not name:
            return
        if QMessageBox.question(self, "Delete", f"Delete '{name}'?") != QMessageBox.Yes:
            return
        self._entity_map(kind).pop(name, None)
        self.current_entity_name = None
        self.current_step_index = None
        self._load_inspector(None)
        if kind == "test_case":
            self._refresh_test_case_list()
            self.tc_steps_table.setRowCount(0)
        else:
            self._refresh_flow_list()
            self.flow_steps_table.setRowCount(0)
        self.save_project()

    def _save_entity_meta(self, kind: str) -> None:
        name = self._selected_entity_name(kind)
        if not name:
            return
        if kind == "test_case":
            tc = self.data["testCases"][name]
            tc["name"] = self.tc_name.text().strip() or name
            tc["suite"] = self.tc_suite.text().strip()
            tc["dataset"] = self.tc_dataset.text().strip()
            tc["enabled"] = bool(self.tc_enabled.isChecked())
            self._refresh_test_case_list()
        else:
            fl = self.data["flows"][name]
            fl["name"] = self.flow_name.text().strip() or name
            self._refresh_flow_list()
        self.save_project()

    # step actions
    def _add_step(self) -> None:
        steps = self._steps_ref()
        if steps is None:
            return
        steps.append({"type": "comment", "text": "new step", "enabled": True})
        self.current_step_index = len(steps) - 1
        self.save_project()
        self._refresh_steps_after_change()

    def _add_quick_step(self, step: dict) -> None:
        if self.current_entity_kind == "flow":
            suggested = str(step.get("type", "comment"))
            built = self._open_step_builder_dialog(initial_type=suggested, base_step=step)
            if built is None:
                return
            steps = self._steps_ref()
            if steps is None:
                return
            steps.append(built)
            self.current_step_index = len(steps) - 1
            self.save_project()
            self._refresh_steps_after_change()
            return
        steps = self._steps_ref()
        if steps is None:
            return
        steps.append(copy.deepcopy(step))
        self.current_step_index = len(steps) - 1
        self.save_project()
        self._refresh_steps_after_change()

    def _guided_edit_selected_step(self) -> None:
        if self.current_entity_kind != "flow":
            return
        steps = self._steps_ref()
        if self.current_step_index is None or not (0 <= self.current_step_index < len(steps)):
            return
        edited = self._open_step_builder_dialog(
            initial_type=str(steps[self.current_step_index].get("type", "comment")),
            base_step=steps[self.current_step_index],
        )
        if edited is None:
            return
        steps[self.current_step_index] = edited
        self.save_project()
        self._refresh_steps_after_change()

    def _open_step_builder_dialog(self, *, initial_type: str, base_step: dict | None = None) -> dict | None:
        dialog = QDialog(self)
        dialog.setWindowTitle("Step Builder")
        dialog.resize(520, 320)
        v = QVBoxLayout(dialog)
        form = QFormLayout()

        type_combo = QComboBox()
        supported = [
            "click",
            "type_text",
            "press_key",
            "wait",
            "screenshot",
            "run_flow",
            "assert_window_title_contains",
            "assert_file_exists",
            "comment",
        ]
        type_combo.addItems(supported)
        if initial_type in supported:
            type_combo.setCurrentText(initial_type)

        target_combo = QComboBox()
        target_combo.setEditable(True)
        target_combo.addItem("")
        for tname in sorted(self.data.get("targets", {}).keys()):
            target_combo.addItem(str(tname))

        flow_combo = QComboBox()
        flow_combo.setEditable(True)
        flow_combo.addItem("")
        for fname in sorted(self.data.get("flows", {}).keys()):
            flow_combo.addItem(str(fname))

        value_edit = QLineEdit()
        key_edit = QLineEdit()
        path_edit = QLineEdit()
        desc_edit = QLineEdit()
        seconds_spin = QDoubleSpinBox()
        seconds_spin.setRange(0, 3600)
        seconds_spin.setDecimals(3)
        seconds_spin.setSingleStep(0.1)
        enabled_check = QCheckBox("Enabled")
        enabled_check.setChecked(True)

        form.addRow("Action", type_combo)
        form.addRow("Target", target_combo)
        form.addRow("Subflow", flow_combo)
        form.addRow("Value/Text", value_edit)
        form.addRow("Key", key_edit)
        form.addRow("Path", path_edit)
        form.addRow("Seconds", seconds_spin)
        form.addRow("Description", desc_edit)
        form.addRow("", enabled_check)
        v.addLayout(form)

        if base_step:
            enabled_check.setChecked(bool(base_step.get("enabled", True)))
            target_combo.setCurrentText(str(base_step.get("target", "")))
            flow_combo.setCurrentText(str(base_step.get("flow", "")))
            value_edit.setText(str(base_step.get("value", base_step.get("text", ""))))
            key_edit.setText(str(base_step.get("key", "")))
            path_edit.setText(str(base_step.get("path", "")))
            try:
                seconds_spin.setValue(float(base_step.get("seconds", 0.0)))
            except Exception:
                seconds_spin.setValue(0.0)
            desc_edit.setText(str(base_step.get("description", base_step.get("text", ""))))

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        v.addWidget(buttons)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)

        if dialog.exec() != QDialog.Accepted:
            return None

        step_type = type_combo.currentText().strip()
        step: dict = {"type": step_type, "enabled": bool(enabled_check.isChecked())}
        target = target_combo.currentText().strip()
        flow = flow_combo.currentText().strip()
        value = value_edit.text()
        key = key_edit.text().strip()
        path = path_edit.text().strip()
        seconds = float(seconds_spin.value())
        desc = desc_edit.text().strip()

        if step_type == "click":
            if target:
                step["target"] = target
        elif step_type == "type_text":
            step["value"] = value
        elif step_type == "press_key":
            step["key"] = key or "enter"
        elif step_type == "wait":
            step["seconds"] = seconds
        elif step_type == "screenshot":
            step["name"] = value.strip() or "step"
        elif step_type == "run_flow":
            step["flow"] = flow
        elif step_type == "assert_window_title_contains":
            step["value"] = value
        elif step_type == "assert_file_exists":
            step["path"] = path or value
        elif step_type == "comment":
            step["text"] = value or desc or "note"

        if desc and step_type != "comment":
            step["description"] = desc
        return step

    def _duplicate_step(self) -> None:
        steps = self._steps_ref()
        if self.current_step_index is None or not (0 <= self.current_step_index < len(steps)):
            return
        steps.insert(self.current_step_index + 1, copy.deepcopy(steps[self.current_step_index]))
        self.current_step_index += 1
        self.save_project()
        self._refresh_steps_after_change()

    def _delete_step(self) -> None:
        steps = self._steps_ref()
        if self.current_step_index is None or not (0 <= self.current_step_index < len(steps)):
            return
        steps.pop(self.current_step_index)
        self.current_step_index = None
        self._load_inspector(None)
        self.save_project()
        self._refresh_steps_after_change()

    def _move_step(self, delta: int) -> None:
        steps = self._steps_ref()
        i = self.current_step_index
        if i is None:
            return
        j = i + delta
        if not (0 <= i < len(steps) and 0 <= j < len(steps)):
            return
        steps[i], steps[j] = steps[j], steps[i]
        self.current_step_index = j
        self.save_project()
        self._refresh_steps_after_change()

    # ---------------------------
    # Run actions
    # ---------------------------
    def _selected_kind_and_name(self) -> tuple[str | None, str | None]:
        nav_page = self._normalize_nav_label(self.nav_list.currentItem().text()) if self.nav_list.currentItem() else ""
        if nav_page == "Flow Builder":
            return "flow", self._selected_entity_name("flow")
        if nav_page == "Test Cases":
            return "test_case", self._selected_entity_name("test_case")
        if self.current_entity_kind in {"flow", "test_case"} and self.current_entity_name:
            return self.current_entity_kind, self.current_entity_name
        return None, None

    def _quick_import_excel(self) -> None:
        choice, ok = QInputDialog.getItem(
            self,
            "Import",
            "What do you want to import?",
            ["Test Cases (Excel/CSV)", "Dataset (Excel/CSV)"],
            0,
            False,
        )
        if not ok:
            return
        if choice.startswith("Test Cases"):
            self.import_test_cases_file()
        else:
            self.import_dataset_file()

    def run_selected(self, dry_run: bool) -> None:
        kind, name = self._selected_kind_and_name()
        if not kind or not name:
            QMessageBox.warning(self, "No Selection", "Select a flow or test case first.")
            return
        dataset_idx = None
        if kind == "test_case":
            ds = str(self.data["testCases"][name].get("dataset", "")).strip()
            if ds:
                idx, ok = QInputDialog.getInt(self, "Dataset Row", f"Dataset '{ds}' row index (0-based):", 0, 0)
                if not ok:
                    return
                dataset_idx = idx

        if not dry_run and pyautogui is None:
            QMessageBox.critical(self, "Missing Dependency", "pyautogui is not installed. Use Dry Run.")
            return
        delay = float(self.data.get("settings", {}).get("startupDelaySeconds", 3))
        if not dry_run and kind != "flow":
            if QMessageBox.question(
                self,
                "Run",
                f"Run {kind} '{name}' after {delay:.1f}s?\nMove mouse to top-left for fail-safe.",
            ) != QMessageBox.Yes:
                return
        thread = threading.Thread(
            target=self._run_thread,
            args=(kind, name, dry_run, delay, dataset_idx),
            daemon=True,
        )
        thread.start()

    def run_from_selected_step(self) -> None:
        self._run_partial("from_selected")

    def run_step_once(self) -> None:
        self._run_partial("step_once")

    def _run_partial(self, mode: str) -> None:
        kind, name = self._selected_kind_and_name()
        if not kind or not name:
            QMessageBox.warning(self, "No Selection", "Select a flow or test case first.")
            return
        if self.current_step_index is None:
            QMessageBox.warning(self, "No Step", "Select a step first.")
            return
        all_steps = self._steps_ref()
        if not (0 <= self.current_step_index < len(all_steps)):
            return
        subset = [copy.deepcopy(all_steps[self.current_step_index])] if mode == "step_once" else copy.deepcopy(
            all_steps[self.current_step_index :]
        )
        dataset_idx = None
        if kind == "test_case":
            ds = str(self.data["testCases"][name].get("dataset", "")).strip()
            if ds:
                idx, ok = QInputDialog.getInt(self, "Dataset Row", f"Dataset '{ds}' row index (0-based):", 0, 0)
                if not ok:
                    return
                dataset_idx = idx
        if pyautogui is None:
            QMessageBox.critical(self, "Missing Dependency", "pyautogui is required for live execution.")
            return
        delay = float(self.data.get("settings", {}).get("startupDelaySeconds", 3))
        if QMessageBox.question(
            self,
            "Confirm",
            f"{'Step Once' if mode == 'step_once' else 'Run From Selected Step'} after {delay:.1f}s?",
        ) != QMessageBox.Yes:
            return
        thread = threading.Thread(
            target=self._run_partial_thread,
            args=(kind, name, subset, mode, dataset_idx, delay, self.current_step_index + 1),
            daemon=True,
        )
        thread.start()

    def _run_thread(self, kind: str, name: str, dry_run: bool, delay: float, dataset_idx: int | None) -> None:
        self.bridge.status.emit(f"Running {kind}:{name}")
        self.bridge.log.emit(f"Starting {kind} '{name}' (dry_run={dry_run})")
        try:
            if not dry_run and delay > 0:
                self.bridge.log.emit(f"Waiting {delay:.2f}s before execution...")
                time.sleep(delay)
            self.active_runner = TestFlowRunner(self.data, log_callback=lambda m: self.bridge.log.emit(m))
            self.active_runner.reset_stop()
            if kind == "flow":
                result = self.active_runner.run_flow(name, dry_run=dry_run)
            else:
                result = self.active_runner.run_test_case(name, dry_run=dry_run, dataset_row_index=dataset_idx)
            self.bridge.run_finished.emit(result)
        except RunnerExecutionError as exc:
            self.bridge.log.emit(f"Run error: {exc}")
            self.bridge.run_finished.emit({"status": "failed", "error": str(exc)})
        except Exception as exc:
            self._log_exception("Run thread failure")
            self.bridge.log.emit(f"Run failure: {exc}")
            self.bridge.run_finished.emit({"status": "failed", "error": str(exc)})
        finally:
            self.active_runner = None
            self.bridge.status.emit("Idle")

    def _run_partial_thread(
        self,
        kind: str,
        name: str,
        subset_steps: list[dict],
        mode: str,
        dataset_idx: int | None,
        delay: float,
        selected_step_number: int,
    ) -> None:
        label = "step_once" if mode == "step_once" else "run_from_selected"
        self.bridge.status.emit(f"Running {label} {kind}:{name}")
        self.bridge.log.emit(f"Starting {label} for {kind} '{name}' from step {selected_step_number}")
        try:
            if delay > 0:
                self.bridge.log.emit(f"Waiting {delay:.2f}s before execution...")
                time.sleep(delay)
            temp_data = copy.deepcopy(self.data)
            existing_ids = {str(r.get("runId", "")) for r in self.data.get("runs", []) if isinstance(r, dict)}
            temp_name = f"{name}__{label}_{selected_step_number}"
            if kind == "flow":
                temp_data.setdefault("flows", {})[temp_name] = {
                    "name": temp_name,
                    "description": f"{label} from step {selected_step_number}",
                    "parameters": [],
                    "steps": subset_steps,
                }
            else:
                src = self.data["testCases"][name]
                temp_data.setdefault("testCases", {})[temp_name] = {
                    "id": temp_name,
                    "name": temp_name,
                    "suite": str(src.get("suite", "")),
                    "description": f"{label} from step {selected_step_number}",
                    "dataset": str(src.get("dataset", "")),
                    "enabled": True,
                    "variables": copy.deepcopy(src.get("variables", {})),
                    "steps": subset_steps,
                }
            self.active_runner = TestFlowRunner(temp_data, log_callback=lambda m: self.bridge.log.emit(m))
            self.active_runner.reset_stop()
            if kind == "flow":
                result = self.active_runner.run_flow(temp_name, dry_run=False)
            else:
                result = self.active_runner.run_test_case(temp_name, dry_run=False, dataset_row_index=dataset_idx)
            for run in temp_data.get("runs", []):
                if not isinstance(run, dict):
                    continue
                rid = str(run.get("runId", ""))
                if rid and rid not in existing_ids:
                    self.data.setdefault("runs", []).append(run)
            self.bridge.run_finished.emit(result)
        except Exception as exc:
            self._log_exception(f"{label} thread failure")
            self.bridge.run_finished.emit({"status": "failed", "error": str(exc)})
        finally:
            self.active_runner = None
            self.bridge.status.emit("Idle")

    def _on_run_finished_ui(self, result: dict) -> None:
        if result.get("error"):
            QMessageBox.critical(self, "Run Error", str(result.get("error")))
            return
        self.save_project()
        self._refresh_dashboard()
        self.refresh_runs_table()
        self.append_log(
            f"Run finished with status={result.get('status')} folder={result.get('runFolder', '')}"
        )

    def stop_run(self) -> None:
        if self.recorder.is_running:
            self.recorder.stop()
            self.append_log("Stop requested for recorder.")
            return
        if self.active_runner is None:
            self.append_log("No active run.")
            return
        self.active_runner.request_stop()
        self.append_log("Stop requested. Runner will stop before next step.")

    # ---------------------------
    # Recorder
    # ---------------------------
    def start_recording(self) -> None:
        stop_key = str(self.data.get("settings", {}).get("recordingStopHotkey", "f8"))
        status = self.recorder.start(
            stop_key_name=stop_key,
            on_finished=lambda steps, err: self.bridge.recorder_finished.emit(steps, err),
            record_typing=bool(self.rec_opt_typing.isChecked()),
            record_hotkeys=bool(self.rec_opt_hotkeys.isChecked()),
        )
        if not status.available:
            QMessageBox.critical(self, "Recorder", status.message)
            self.rec_status.setText(status.message)
            return
        self.rec_status.setText(f"Recording... press {stop_key.upper()} to stop")
        self.set_status("Recording")
        self.append_log("Recording started.")

    def stop_recording(self) -> None:
        if self.recorder.is_running:
            self.recorder.stop()
            self.rec_status.setText("Stopping...")

    def _on_recorder_finished_ui(self, steps: list, error_text: str) -> None:
        self.set_status("Idle")
        self.recording_steps = self._cleanup_recorded_steps(steps)
        self._refresh_recording_preview()
        if error_text:
            self.rec_status.setText(f"Recorder error: {error_text}")
            self.append_log(f"Recorder error: {error_text}")
        else:
            self.rec_status.setText(f"Idle ({len(self.recording_steps)} steps captured)")
            self.append_log(f"Recording stopped. Captured {len(self.recording_steps)} steps.")

    def _cleanup_recorded_steps(self, steps: list[dict]) -> list[dict]:
        out: list[dict] = []
        for step in steps:
            t = step.get("type")
            if t == "wait":
                if not self.rec_opt_timing.isChecked():
                    continue
                out.append({"type": "wait", "seconds": float(step.get("seconds", 0)), "enabled": True})
            elif t == "click_xy":
                if self.rec_opt_left.isChecked():
                    out.append({"type": "click_xy", "x": int(step.get("x", 0)), "y": int(step.get("y", 0)), "enabled": True})
            elif t == "double_click":
                if self.rec_opt_double.isChecked():
                    out.append(
                        {"type": "double_click", "x": int(step.get("x", 0)), "y": int(step.get("y", 0)), "enabled": True}
                    )
            elif t == "right_click":
                if self.rec_opt_right.isChecked():
                    out.append(
                        {"type": "right_click", "x": int(step.get("x", 0)), "y": int(step.get("y", 0)), "enabled": True}
                    )
            elif t == "type_text":
                if self.rec_opt_typing.isChecked():
                    out.append({"type": "type_text", "value": str(step.get("value", "")), "enabled": True})
            elif t == "hotkey":
                if self.rec_opt_hotkeys.isChecked():
                    out.append({"type": "hotkey", "keys": step.get("keys", []), "enabled": True})
            elif t == "press_key":
                if self.rec_opt_typing.isChecked():
                    out.append({"type": "press_key", "key": str(step.get("key", "")), "enabled": True})
        return out

    def _refresh_recording_preview(self) -> None:
        self.rec_table.setRowCount(len(self.recording_steps))
        for i, step in enumerate(self.recording_steps):
            self.rec_table.setItem(i, 0, QTableWidgetItem(str(i + 1)))
            self.rec_table.setItem(i, 1, QTableWidgetItem(str(step.get("type", ""))))
            details = ", ".join([f"{k}={v}" for k, v in step.items() if k != "type"])[:220]
            self.rec_table.setItem(i, 2, QTableWidgetItem(details))

    def save_recording_as_flow(self) -> None:
        if not self.recording_steps:
            QMessageBox.warning(self, "No Recording", "No recorded steps to save.")
            return
        name, ok = QInputDialog.getText(self, "Save Flow", "Flow name:")
        if not ok or not name.strip():
            return
        n = name.strip()
        self.data.setdefault("flows", {})[n] = {
            "name": n,
            "description": "Recorded flow",
            "parameters": [],
            "steps": copy.deepcopy(self.recording_steps),
        }
        self.save_project()
        self._refresh_flow_list()
        self.append_log(f"Saved recording as flow '{n}'.")

    def save_recording_as_test_case(self) -> None:
        if not self.recording_steps:
            QMessageBox.warning(self, "No Recording", "No recorded steps to save.")
            return
        case_id, ok = QInputDialog.getText(self, "Save Test Case", "Test case ID:")
        if not ok or not case_id.strip():
            return
        cid = case_id.strip()
        self.data.setdefault("testCases", {})[cid] = {
            "id": cid,
            "name": cid,
            "suite": "",
            "description": "Recorded test case",
            "dataset": "",
            "enabled": True,
            "variables": {},
            "steps": copy.deepcopy(self.recording_steps),
        }
        self.save_project()
        self._refresh_test_case_list()
        self.append_log(f"Saved recording as test case '{cid}'.")

    def append_recording_to_flow(self) -> None:
        if not self.recording_steps:
            QMessageBox.warning(self, "No Recording", "No recorded steps to append.")
            return
        flows = sorted(self.data.get("flows", {}).keys())
        if not flows:
            QMessageBox.warning(self, "No Flows", "Create a flow first.")
            return
        flow_name, ok = QInputDialog.getItem(self, "Append to Flow", "Flow:", flows, 0, False)
        if not ok or not flow_name:
            return
        self.data["flows"][flow_name].setdefault("steps", []).extend(copy.deepcopy(self.recording_steps))
        self.save_project()
        self.append_log(f"Appended {len(self.recording_steps)} steps to flow '{flow_name}'.")

    def discard_recording(self) -> None:
        self.recording_steps = []
        self._refresh_recording_preview()
        self.append_log("Discarded recorded preview.")

    # ---------------------------
    # Targets
    # ---------------------------
    def _refresh_targets(self) -> None:
        targets = self.data.get("targets", {})
        self.targets_table.setRowCount(len(targets))
        for r, (name, t) in enumerate(sorted(targets.items())):
            vals = [name, str(t.get("x", "")), str(t.get("y", "")), str(t.get("description", ""))]
            for c, v in enumerate(vals):
                self.targets_table.setItem(r, c, QTableWidgetItem(v))

    def _selected_target_name(self) -> str | None:
        rows = self.targets_table.selectionModel().selectedRows()
        if not rows:
            return None
        item = self.targets_table.item(rows[0].row(), 0)
        return item.text() if item else None

    def capture_target(self) -> None:
        if pyautogui is None:
            QMessageBox.critical(self, "Missing Dependency", "pyautogui is required to capture coordinates.")
            return
        name, ok = QInputDialog.getText(self, "Capture Target", "Target name:")
        if not ok or not name.strip():
            return
        n = name.strip()
        x, y = pyautogui.position()
        self.data.setdefault("targets", {})[n] = {
            "x": int(x),
            "y": int(y),
            "description": "",
            "createdAt": datetime.now().isoformat(timespec="seconds"),
        }
        self.save_project()
        self._refresh_targets()

    def manual_add_target(self) -> None:
        name, ok = QInputDialog.getText(self, "Add Target", "Target name:")
        if not ok or not name.strip():
            return
        x, okx = QInputDialog.getInt(self, "Add Target", "X:", 0)
        if not okx:
            return
        y, oky = QInputDialog.getInt(self, "Add Target", "Y:", 0)
        if not oky:
            return
        self.data.setdefault("targets", {})[name.strip()] = {
            "x": int(x),
            "y": int(y),
            "description": "",
            "createdAt": datetime.now().isoformat(timespec="seconds"),
        }
        self.save_project()
        self._refresh_targets()

    def rename_target(self) -> None:
        name = self._selected_target_name()
        if not name:
            return
        new_name, ok = QInputDialog.getText(self, "Rename Target", "New name:", text=name)
        if not ok or not new_name.strip():
            return
        n = new_name.strip()
        if n in self.data.get("targets", {}) and n != name:
            QMessageBox.critical(self, "Error", "Target name already exists.")
            return
        self.data["targets"][n] = self.data["targets"].pop(name)
        self.save_project()
        self._refresh_targets()

    def delete_target(self) -> None:
        name = self._selected_target_name()
        if not name:
            return
        if QMessageBox.question(self, "Delete", f"Delete target '{name}'?") != QMessageBox.Yes:
            return
        self.data.get("targets", {}).pop(name, None)
        self.save_project()
        self._refresh_targets()

    def test_target(self) -> None:
        if pyautogui is None:
            QMessageBox.critical(self, "Missing Dependency", "pyautogui is required for click testing.")
            return
        name = self._selected_target_name()
        if not name:
            return
        t = self.data["targets"][name]
        delay = float(self.data.get("settings", {}).get("startupDelaySeconds", 3))
        if QMessageBox.question(
            self, "Test Target", f"Click '{name}' after {delay:.1f}s at x={t['x']} y={t['y']}?"
        ) != QMessageBox.Yes:
            return

        def do_click():
            time.sleep(delay)
            pyautogui.click(int(t["x"]), int(t["y"]))

        threading.Thread(target=do_click, daemon=True).start()

    # ---------------------------
    # Datasets + Imports
    # ---------------------------
    def _refresh_datasets(self) -> None:
        self.dataset_list.clear()
        for name in sorted(self.data.get("datasets", {}).keys()):
            self.dataset_list.addItem(name)

    def on_dataset_select(self) -> None:
        item = self.dataset_list.currentItem()
        if not item:
            self.dataset_preview.setPlainText("")
            return
        name = item.text()
        ds = self.data.get("datasets", {}).get(name)
        rows = []
        if isinstance(ds, dict):
            rows = ds.get("rows", []) if isinstance(ds.get("rows", []), list) else []
        elif isinstance(ds, list):
            rows = ds
        cols = sorted({k for row in rows if isinstance(row, dict) for k in row.keys()})
        out = [f"Rows: {len(rows)} | Columns: {', '.join(cols)}", ""]
        for i, row in enumerate(rows[:500], start=1):
            out.append(f"{i}. {row}")
        self.dataset_preview.setPlainText("\n".join(out))

    def import_dataset_file(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "Import dataset", "", "CSV/XLSX (*.csv *.xlsx);;All files (*)")
        if not path:
            return
        name, ok = QInputDialog.getText(self, "Dataset Name", "Dataset name:", text=Path(path).stem)
        if not ok or not name.strip():
            return
        try:
            _, rows = read_table_rows(path)
        except ImporterError as exc:
            QMessageBox.critical(self, "Import Error", str(exc))
            return
        except Exception as exc:
            self._log_exception("Import dataset failed")
            QMessageBox.critical(self, "Import Error", f"Failed to import dataset: {exc}")
            return
        self.data.setdefault("datasets", {})[name.strip()] = {"rows": rows, "source": path}
        self.save_project()
        self._refresh_datasets()
        self.append_log(f"Imported dataset '{name.strip()}' with {len(rows)} rows.")

    def rename_dataset(self) -> None:
        item = self.dataset_list.currentItem()
        if not item:
            return
        name = item.text()
        new_name, ok = QInputDialog.getText(self, "Rename Dataset", "New name:", text=name)
        if not ok or not new_name.strip():
            return
        n = new_name.strip()
        if n in self.data.get("datasets", {}) and n != name:
            QMessageBox.critical(self, "Error", "Dataset name already exists.")
            return
        self.data["datasets"][n] = self.data["datasets"].pop(name)
        self.save_project()
        self._refresh_datasets()

    def delete_dataset(self) -> None:
        item = self.dataset_list.currentItem()
        if not item:
            return
        name = item.text()
        if QMessageBox.question(self, "Delete", f"Delete dataset '{name}'?") != QMessageBox.Yes:
            return
        self.data.get("datasets", {}).pop(name, None)
        self.save_project()
        self._refresh_datasets()

    def import_test_cases_file(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "Import test cases", "", "CSV/XLSX (*.csv *.xlsx);;All files (*)")
        if not path:
            return
        try:
            _, rows = read_table_rows(path)
            grouped, errors, invalid_rows = parse_test_case_rows(rows)
        except ImporterError as exc:
            QMessageBox.critical(self, "Import Error", str(exc))
            return
        except Exception as exc:
            self._log_exception("Import test cases failed")
            QMessageBox.critical(self, "Import Error", f"Failed to parse import file: {exc}")
            return
        if not grouped:
            QMessageBox.critical(self, "Import Error", "No valid test cases were found.")
            return
        lines = [
            f"Rows found: {len(rows)}",
            f"Test cases found: {len(grouped)}",
            f"Invalid rows: {len(invalid_rows)}",
            "",
            "Import now?",
        ]
        if invalid_rows:
            lines.append("")
            lines.append("Invalid row details (first 10):")
            for item in invalid_rows[:10]:
                lines.append(f"- Row {item.get('row')}: {item.get('reason')}")
        if errors:
            lines.append("")
            lines.append("Validation messages (first 15):")
            for err in errors[:15]:
                lines.append(f"- {err}")
        if QMessageBox.question(self, "Import Test Cases", "\n".join(lines)) != QMessageBox.Yes:
            return
        created, replaced = 0, 0
        for case_id, case_obj in grouped.items():
            if case_id in self.data.get("testCases", {}):
                replaced += 1
            else:
                created += 1
            self.data.setdefault("testCases", {})[case_id] = case_obj
        self.save_project()
        self._refresh_test_case_list()
        self.append_log(f"Imported test cases from {Path(path).name}: created={created}, replaced={replaced}")

    def preview_selected_test_case_variables(self) -> None:
        case_id = self._selected_entity_name("test_case")
        if not case_id:
            return
        tc = self.data["testCases"].get(case_id, {})
        dataset_name = str(tc.get("dataset", "")).strip()
        dataset_row_index = None
        if dataset_name:
            idx, ok = QInputDialog.getInt(self, "Variable Preview", f"Dataset '{dataset_name}' row index:", 0, 0)
            if not ok:
                return
            dataset_row_index = idx
        runner = TestFlowRunner(self.data)
        try:
            preview = runner.preview_test_case_execution(case_id, dataset_row_index=dataset_row_index)
        except RunnerExecutionError as exc:
            self.tc_var_preview.setPlainText(str(exc))
            return
        lines = ["Variables:"]
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
        self.tc_var_preview.setPlainText("\n".join(lines))

    def run_selected_test_case_all_rows(self) -> None:
        case_id = self._selected_entity_name("test_case")
        if not case_id:
            QMessageBox.warning(self, "No Test Case", "Select a test case first.")
            return
        test_case = self.data["testCases"].get(case_id, {})
        ds_name = str(test_case.get("dataset", "")).strip()
        if not ds_name:
            QMessageBox.warning(self, "No Dataset", "Selected test case does not have a dataset.")
            return
        ds = self.data.get("datasets", {}).get(ds_name)
        if isinstance(ds, dict):
            rows = ds.get("rows", []) if isinstance(ds.get("rows", []), list) else []
        elif isinstance(ds, list):
            rows = ds
        else:
            rows = []
        if not rows:
            QMessageBox.warning(self, "Empty Dataset", f"Dataset '{ds_name}' has no rows.")
            return
        if pyautogui is None:
            QMessageBox.critical(self, "Missing Dependency", "pyautogui is required for live execution.")
            return
        delay = float(self.data.get("settings", {}).get("startupDelaySeconds", 3))
        if QMessageBox.question(
            self,
            "Run All Dataset Rows",
            f"Run '{case_id}' for all {len(rows)} dataset rows after {delay:.1f}s delay?",
        ) != QMessageBox.Yes:
            return
        thread = threading.Thread(
            target=self._run_test_case_all_rows_thread,
            args=(case_id, delay, len(rows)),
            daemon=True,
        )
        thread.start()

    def _run_test_case_all_rows_thread(self, case_id: str, startup_delay: float, row_count: int) -> None:
        self.bridge.status.emit(f"Running test_case:{case_id} all rows")
        self.bridge.log.emit(f"Starting test case '{case_id}' for {row_count} dataset rows")
        try:
            if startup_delay > 0:
                time.sleep(startup_delay)
            self.active_runner = TestFlowRunner(self.data, log_callback=lambda m: self.bridge.log.emit(m))
            self.active_runner.reset_stop()
            pass_count, fail_count, stop_count = 0, 0, 0
            for idx in range(row_count):
                if self.active_runner.stop_requested:
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
            self.bridge.log.emit(
                f"Dataset run complete for '{case_id}': passed={pass_count}, failed={fail_count}, stopped={stop_count}"
            )
            self.bridge.run_finished.emit({"status": "done"})
        except Exception as exc:
            self._log_exception("Dataset run thread failure")
            self.bridge.run_finished.emit({"status": "failed", "error": str(exc)})
        finally:
            self.active_runner = None
            self.bridge.status.emit("Idle")

    # ---------------------------
    # Run center + settings + project ops
    # ---------------------------
    def refresh_runs_table(self) -> None:
        runs = list(self.data.get("runs", []))[::-1]
        self.runs_table.setRowCount(len(runs))
        for r, run in enumerate(runs):
            vals = [
                str(run.get("runId", "")),
                str(run.get("startedAt", "")),
                str(run.get("kind", "")),
                str(run.get("name", "")),
                str(run.get("status", "")),
                str(run.get("durationSeconds", "")),
                str(run.get("runFolder", "")),
            ]
            for c, v in enumerate(vals):
                self.runs_table.setItem(r, c, QTableWidgetItem(v))

    def open_selected_report(self) -> None:
        rows = self.runs_table.selectionModel().selectedRows()
        if not rows:
            return
        run_folder_item = self.runs_table.item(rows[0].row(), 6)
        if not run_folder_item:
            return
        folder = run_folder_item.text()
        report = Path(folder) / "report.html"
        if not report.exists():
            QMessageBox.critical(self, "Missing Report", f"Report not found:\n{report}")
            return
        webbrowser.open(report.resolve().as_uri())

    def _load_settings(self) -> None:
        settings = self.data.get("settings", {})
        env = self.data.get("environment", {})
        self.set_start_delay.setText(str(settings.get("startupDelaySeconds", 3)))
        self.set_action_pause.setText(str(settings.get("defaultActionPauseSeconds", 0.1)))
        self.set_post_action_delay.setText(str(settings.get("postActionDelaySeconds", 0.0)))
        self.set_screenshot_delay.setText(str(settings.get("screenshotDelayBeforeSeconds", 0.0)))
        self.set_stop_hotkey.setText(str(settings.get("stopHotkey", "f8")))
        self.set_record_stop_hotkey.setText(str(settings.get("recordingStopHotkey", "f8")))
        self.set_expected_res.setText(str(env.get("expectedResolution", "")))
        self.set_shot_fail.setChecked(bool(settings.get("screenshotOnFailure", True)))
        self.set_shot_each.setChecked(bool(settings.get("screenshotAfterEachStep", False)))

    def save_settings(self) -> None:
        settings = self.data.setdefault("settings", {})
        env = self.data.setdefault("environment", {})
        try:
            settings["startupDelaySeconds"] = float(self.set_start_delay.text())
            settings["defaultActionPauseSeconds"] = float(self.set_action_pause.text())
            settings["postActionDelaySeconds"] = float(self.set_post_action_delay.text())
            settings["screenshotDelayBeforeSeconds"] = float(self.set_screenshot_delay.text())
        except ValueError:
            QMessageBox.critical(self, "Invalid", "Delay values must be numeric.")
            return
        settings["screenshotOnFailure"] = bool(self.set_shot_fail.isChecked())
        settings["screenshotAfterEachStep"] = bool(self.set_shot_each.isChecked())
        settings["stopHotkey"] = self.set_stop_hotkey.text().strip() or "f8"
        settings["recordingStopHotkey"] = self.set_record_stop_hotkey.text().strip() or "f8"
        env["expectedResolution"] = self.set_expected_res.text().strip()
        if pyautogui is not None:
            pyautogui.PAUSE = float(settings.get("defaultActionPauseSeconds", 0.1))
        self.save_project()
        self.append_log("Settings saved.")

    def save_project(self) -> None:
        storage_save_project(self.data)
        self.last_saved_at = datetime.now()
        self.saved_label.setText(f"Last saved: {self.last_saved_at.strftime('%Y-%m-%d %H:%M:%S')}")

    def export_project_zip(self) -> None:
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        out, _ = QFileDialog.getSaveFileName(
            self, "Export project as ZIP", f"testflow_project_export_{stamp}.zip", "ZIP (*.zip)"
        )
        if not out:
            return
        out_path = Path(out)
        self.save_project()
        root_dir = Path.cwd()
        items = [PROJECT_FILE, Path("runs"), Path("logs")]
        with zipfile.ZipFile(out_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
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
        self.append_log(f"Exported project ZIP: {out_path}")

    def import_project_json(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "Import project JSON", "", "JSON (*.json);;All files (*)")
        if not path:
            return
        try:
            raw = json.loads(Path(path).read_text(encoding="utf-8"))
            self.data = normalize_project_data(raw)
        except Exception as exc:
            self._log_exception("Import project JSON failed")
            QMessageBox.critical(self, "Import Error", f"Unable to parse project JSON: {exc}")
            return
        self.save_project()
        self._refresh_all()
        self.append_log(f"Imported project from JSON: {path}")

    def import_project_zip(self) -> None:
        path, _ = QFileDialog.getOpenFileName(self, "Import project ZIP", "", "ZIP (*.zip);;All files (*)")
        if not path:
            return
        if QMessageBox.question(
            self,
            "Import Project ZIP",
            "This will replace current project JSON and may overwrite runs/logs.\nContinue?",
        ) != QMessageBox.Yes:
            return
        root_dir = Path.cwd()
        try:
            with zipfile.ZipFile(Path(path), "r") as zf:
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
            QMessageBox.critical(self, "Import Error", f"Unable to import ZIP: {exc}")
            return
        self.save_project()
        self._refresh_all()
        self.append_log(f"Imported project ZIP: {path}")

    def import_project_folder(self) -> None:
        folder = QFileDialog.getExistingDirectory(self, "Import project folder")
        if not folder:
            return
        source = Path(folder)
        project_file = source / "testflow_project.json"
        if not project_file.exists():
            QMessageBox.critical(self, "Import Error", "Selected folder does not contain testflow_project.json")
            return
        if QMessageBox.question(
            self, "Import Project Folder", "This will replace current project JSON and sync runs/logs.\nContinue?"
        ) != QMessageBox.Yes:
            return
        try:
            Path(PROJECT_FILE).write_text(project_file.read_text(encoding="utf-8"), encoding="utf-8")
            for sub in ("runs", "logs"):
                src_sub = source / sub
                dst_sub = Path(sub)
                if src_sub.exists():
                    if dst_sub.exists():
                        import shutil

                        shutil.rmtree(dst_sub)
                    import shutil

                    shutil.copytree(src_sub, dst_sub)
            loaded, msgs = load_project()
            self.data = loaded
            for m in msgs:
                self.append_log(m)
        except Exception as exc:
            self._log_exception("Import project folder failed")
            QMessageBox.critical(self, "Import Error", f"Unable to import folder: {exc}")
            return
        self.save_project()
        self._refresh_all()
        self.append_log(f"Imported project folder: {source}")

    def create_sample_project(self) -> None:
        if QMessageBox.question(
            self,
            "Create Sample Project",
            "This will merge sample data into the current project. Continue?",
        ) != QMessageBox.Yes:
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
            "description": "Sample test case generated by helper",
            "dataset": "demo_users",
            "enabled": True,
            "variables": {},
            "steps": [
                {"type": "run_flow", "flow": "Login_POS", "enabled": True},
                {"type": "assert_window_title_contains", "value": "${expectedTitle}", "enabled": True},
            ],
        }
        self.save_project()
        self._refresh_all()
        self.append_log("Sample project content added.")


def main() -> None:
    app = QApplication(sys.argv)
    win = TestFlowRunnerQt()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
