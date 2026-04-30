import copy
import json
import subprocess
import sys
import threading
import time
import webbrowser
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Any

from PySide6.QtCore import QSize, Qt, QTimer
from PySide6.QtGui import QAction, QColor, QKeySequence
from PySide6.QtWidgets import (
    QApplication,
    QAbstractItemView,
    QCheckBox,
    QComboBox,
    QDialog,
    QFileDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QInputDialog,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSpinBox,
    QSplitter,
    QStackedWidget,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QToolBar,
    QToolButton,
    QVBoxLayout,
    QWidget,
    QDoubleSpinBox,
)

from importer import ImporterError, read_table_rows
from recorder import GlobalClickRecorder
from runner import RunnerExecutionError, TestFlowRunner
from storage import PROJECT_FILE, load_project, save_project as storage_save_project


def _title_for_step_type(step_type: str) -> str:
    mapping = {
        "click": "Click",
        "click_xy": "Click",
        "double_click": "Double Click",
        "double_click_xy": "Double Click",
        "right_click": "Right Click",
        "type_text": "Type",
        "press_key": "Press Key",
        "hotkey": "Hotkey",
        "wait": "Wait",
        "screenshot": "Screenshot",
        "run_flow": "Reusable Action",
        "subflow": "Reusable Action",
        "comment": "Manual",
        "assert_window_title_contains": "Assert Window Title",
        "assert_clipboard_contains": "Assert Clipboard",
        "assert_file_exists": "Assert File Exists",
    }
    return mapping.get(step_type, step_type.replace("_", " ").title())


def _normalize_status(status: str, dry_run: bool) -> str:
    if dry_run:
        return "Preview"
    s = str(status or "").lower()
    if s == "passed":
        return "Passed"
    if s == "failed":
        return "Failed"
    if s in {"stopped", "cancelled", "canceled"}:
        return "Stopped"
    return status.title() if status else "Unknown"


MIN_EXPLICIT_WAIT_SECONDS = 1.5
MIN_RECORDED_WAIT_SECONDS = 0.15
PANEL_BG = "#F7F9FC"
BORDER = "#D7DEE8"
TEXT = "#172033"
MUTED_TEXT = "#607089"
SELECTED_BG = "#EAF2FF"

AUTOMATION_TYPES = [
    "Manual",
    "Click",
    "Double Click",
    "Type",
    "Wait",
    "Screenshot",
    "Reusable Action",
    "Press Key",
]


class MinimalTestFlowApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.base_window_title = "TestFlow Runner"
        self.setWindowTitle(self.base_window_title)
        self.resize(1500, 920)
        self.project_path = PROJECT_FILE.resolve()
        self.data, messages = load_project()
        self.recorder = GlobalClickRecorder()
        self.recorded_steps: list[dict] = []
        self.active_runner: TestFlowRunner | None = None
        self.current_flow_name: str | None = None
        self.current_step_index: int | None = None
        self._inspector_loading = False
        self._dirty = False

        self._build_ui()
        self._apply_theme()
        self._bind_shortcuts()
        self._check_dependencies()
        self.refresh_flows()
        self.refresh_runs()
        for m in messages:
            self.log(m)
        self._update_window_title()

    def _apply_theme(self):
        self.setStyleSheet(
            f"""
            QWidget {{ font-family: 'Segoe UI', Arial; font-size: 13px; color: {TEXT}; background: {PANEL_BG}; }}
            QToolBar {{ background: #ffffff; border-bottom: 1px solid {BORDER}; padding: 8px; spacing: 8px; }}
            QPushButton {{
                background: #ffffff; border: 1px solid {BORDER}; border-radius: 7px;
                padding: 6px 10px; min-height: 24px;
            }}
            QPushButton:hover {{ background: #eff4ff; }}
            QPushButton:pressed {{ background: #e7efff; }}
            QListWidget, QTableWidget, QTextEdit, QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox, QTabWidget::pane {{
                background: #ffffff; border: 1px solid {BORDER}; border-radius: 7px;
            }}
            QTabBar::tab {{
                background: #edf2fa; border: 1px solid {BORDER}; border-bottom: none; border-top-left-radius: 6px; border-top-right-radius: 6px;
                padding: 6px 10px; color: {MUTED_TEXT}; margin-right: 2px;
            }}
            QTabBar::tab:selected {{ background: #ffffff; color: {TEXT}; font-weight: 600; }}
            QHeaderView::section {{
                background: #eef3f9; border: 0; border-bottom: 1px solid {BORDER};
                padding: 7px; font-weight: 600; color: {MUTED_TEXT};
            }}
            QTableWidget::item {{ padding: 5px; border: none; }}
            QTableWidget::item:selected, QListWidget::item:selected {{
                background: {SELECTED_BG}; color: {TEXT};
            }}
            QListWidget::item {{ border-bottom: 1px solid #edf2f9; padding: 8px 10px; }}
            QLineEdit::placeholder, QTextEdit {{ color: {MUTED_TEXT}; }}
            QToolButton {{
                background: transparent; border: 1px solid {BORDER}; border-radius: 6px;
                padding: 5px 8px; text-align: left; color: {TEXT};
            }}
            QToolButton:checked {{ background: #f2f6fe; }}
            QToolButton[primary=\"true\"] {{
                background: #2f6fed; color: #ffffff; border: 1px solid #2a62d1;
            }}
            QToolButton[primary=\"true\"]:hover {{ background: #295fc8; }}
            QToolButton[secondary=\"true\"] {{
                background: #eef3ff; color: #1f3f7d; border: 1px solid #c7d7f6;
            }}
            QToolButton[recording_stop=\"true\"] {{
                background: #fef2f2; color: #9f1239; border: 1px solid #f3c4d0;
            }}
            QLabel#title {{ font-size: 20px; font-weight: 700; }}
            QLabel#paneTitle {{ font-size: 15px; font-weight: 700; color: #334155; }}
            QLabel#muted {{ color: {MUTED_TEXT}; }}
            QGroupBox {{ border: 1px solid {BORDER}; border-radius: 8px; margin-top: 8px; padding-top: 10px; background: #ffffff; }}
            QGroupBox::title {{ subcontrol-origin: margin; left: 10px; padding: 0 4px; color: #334155; }}
            """
        )

    def _build_ui(self):
        toolbar = QToolBar("Main")
        toolbar.setMovable(False)
        self.addToolBar(toolbar)
        self.act_start_recording = QAction("Start Recording", self)
        self.act_start_recording.triggered.connect(self.start_recording)
        toolbar.addAction(self.act_start_recording)
        self.act_stop_recording = QAction("Stop Recording", self)
        self.act_stop_recording.triggered.connect(self.stop_recording)
        self.act_stop_recording.setVisible(False)
        self.act_stop_recording.setEnabled(False)
        toolbar.addAction(self.act_stop_recording)
        toolbar.addSeparator()
        self.act_import = QAction("Import", self)
        self.act_import.triggered.connect(self.import_test_cases)
        toolbar.addAction(self.act_import)
        toolbar.addSeparator()
        self.act_preview_run = QAction("Preview Run", self)
        self.act_preview_run.triggered.connect(self.dry_run_selected_flow)
        toolbar.addAction(self.act_preview_run)
        self.act_run_test = QAction("Run Test", self)
        self.act_run_test.triggered.connect(self.run_selected_flow)
        toolbar.addAction(self.act_run_test)
        toolbar.addSeparator()
        self.act_save = QAction("Save", self)
        self.act_save.triggered.connect(lambda: self.save_project(clear_dirty=True))
        toolbar.addAction(self.act_save)
        run_btn = toolbar.widgetForAction(self.act_run_test)
        if run_btn is not None:
            run_btn.setProperty("primary", "true")
        preview_btn = toolbar.widgetForAction(self.act_preview_run)
        if preview_btn is not None:
            preview_btn.setProperty("secondary", "true")
        stop_btn = toolbar.widgetForAction(self.act_stop_recording)
        if stop_btn is not None:
            stop_btn.setProperty("recording_stop", "true")

        root = QWidget()
        self.setCentralWidget(root)
        v = QVBoxLayout(root)
        title = QLabel("TestFlow Runner")
        title.setObjectName("title")
        v.addWidget(title)
        self.status = QLabel("Ready")
        v.addWidget(self.status)
        split = QSplitter(Qt.Horizontal)
        v.addWidget(split, stretch=1)

        # Left pane: Test Cases
        left = QWidget()
        ll = QVBoxLayout(left)
        lt = QLabel("Test Cases")
        lt.setObjectName("paneTitle")
        ll.addWidget(lt)
        self.test_case_search = QLineEdit()
        self.test_case_search.setPlaceholderText("Search test cases...")
        self.test_case_search.textChanged.connect(self.refresh_flows)
        ll.addWidget(self.test_case_search)
        left_actions = QHBoxLayout()
        self.new_flow_btn = QPushButton("New Test")
        self.new_flow_btn.clicked.connect(self.new_flow)
        self.import_btn = QPushButton("Import Excel/CSV")
        self.import_btn.clicked.connect(self.import_test_cases)
        left_actions.addWidget(self.new_flow_btn)
        left_actions.addWidget(self.import_btn)
        ll.addLayout(left_actions)
        left_actions_2 = QHBoxLayout()
        self.del_flow_btn = QPushButton("Delete Test")
        self.del_flow_btn.clicked.connect(self.delete_flow)
        left_actions_2.addWidget(self.del_flow_btn)
        left_actions_2.addStretch(1)
        ll.addLayout(left_actions_2)
        self.flow_list = QListWidget()
        self.flow_list.currentItemChanged.connect(self.on_flow_select)
        ll.addWidget(self.flow_list, stretch=1)
        self.test_case_empty_hint = QLabel("")
        self.test_case_empty_hint.setObjectName("muted")
        self.test_case_empty_hint.setWordWrap(True)
        self.test_case_empty_hint.setVisible(False)
        ll.addWidget(self.test_case_empty_hint)
        split.addWidget(left)

        # Center pane: Step Builder + run activity
        center = QWidget()
        cl = QVBoxLayout(center)
        ct = QLabel("Step Builder")
        ct.setObjectName("paneTitle")
        cl.addWidget(ct)
        self.flow_name_edit = QLineEdit()
        self.flow_name_edit.setPlaceholderText("Test Case name")
        self.flow_name_edit.editingFinished.connect(self.rename_flow)
        cl.addWidget(self.flow_name_edit)
        step_toolbar = QHBoxLayout()
        for text, t in [
            ("Add Click", "click_xy"),
            ("Add Type", "type_text"),
            ("Add Wait", "wait"),
            ("Add Screenshot", "screenshot"),
            ("Add Reusable Action", "run_flow"),
            ("Add Manual", "comment"),
        ]:
            b = QPushButton(text)
            b.clicked.connect(lambda _=False, step_type=t: self.add_step(step_type))
            step_toolbar.addWidget(b)
        cl.addLayout(step_toolbar)
        step_ops = QHBoxLayout()
        self.dup_step_btn = QPushButton("Duplicate")
        self.dup_step_btn.clicked.connect(self.duplicate_step)
        self.del_step_btn = QPushButton("Delete")
        self.del_step_btn.clicked.connect(self.delete_step)
        self.move_up_btn = QPushButton("Move Up")
        self.move_up_btn.clicked.connect(lambda: self.move_step(-1))
        self.move_down_btn = QPushButton("Move Down")
        self.move_down_btn.clicked.connect(lambda: self.move_step(1))
        for b in [self.dup_step_btn, self.del_step_btn, self.move_up_btn, self.move_down_btn]:
            step_ops.addWidget(b)
        step_ops.addStretch(1)
        cl.addLayout(step_ops)
        self.step_content_stack = QStackedWidget()
        table_page = QWidget()
        table_layout = QVBoxLayout(table_page)
        table_layout.setContentsMargins(0, 0, 0, 0)
        self.steps_table = QTableWidget(0, 6)
        self.steps_table.setHorizontalHeaderLabels(["#", "Type", "Description", "Target / Action", "Input", "Wait"])
        self.steps_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.steps_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.steps_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.steps_table.setShowGrid(False)
        self.steps_table.verticalHeader().setVisible(False)
        self.steps_table.verticalHeader().setDefaultSectionSize(30)
        self.steps_table.setAlternatingRowColors(True)
        self.steps_table.itemSelectionChanged.connect(self.on_step_select)
        self.steps_table.itemDoubleClicked.connect(lambda _i: self._focus_inspector_for_selected_step())
        self.steps_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.steps_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.steps_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.steps_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        self.steps_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)
        self.steps_table.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeToContents)
        table_layout.addWidget(self.steps_table, stretch=1)
        self.step_content_stack.addWidget(table_page)
        empty_page = QWidget()
        empty_layout = QVBoxLayout(empty_page)
        empty_layout.setContentsMargins(12, 12, 12, 12)
        self.step_empty_hint = QLabel("")
        self.step_empty_hint.setObjectName("muted")
        self.step_empty_hint.setWordWrap(True)
        empty_layout.addWidget(self.step_empty_hint)
        empty_layout.addStretch(1)
        self.step_content_stack.addWidget(empty_page)
        cl.addWidget(self.step_content_stack, stretch=5)

        self.run_activity_box = QWidget()
        run_activity_layout = QVBoxLayout(self.run_activity_box)
        run_activity_layout.setContentsMargins(8, 8, 8, 8)
        run_activity_layout.setSpacing(6)
        self.run_activity_toggle = QToolButton()
        self.run_activity_toggle.setText("Run Activity >")
        self.run_activity_toggle.setToolButtonStyle(Qt.ToolButtonTextOnly)
        self.run_activity_toggle.setCheckable(True)
        self.run_activity_toggle.setChecked(False)
        self.run_activity_toggle.clicked.connect(self._on_run_activity_toggled)
        run_activity_layout.addWidget(self.run_activity_toggle)
        self.run_tabs = QTabWidget()
        run_activity_layout.addWidget(self.run_tabs)

        run_hist_tab = QWidget()
        run_hist_layout = QVBoxLayout(run_hist_tab)
        self.runs_table = QTableWidget(0, 5)
        self.runs_table.setHorizontalHeaderLabels(["Started", "Test Case", "Status", "Duration", "Report"])
        self.runs_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.runs_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.runs_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.runs_table.itemSelectionChanged.connect(self.on_run_selected)
        self.runs_table.itemDoubleClicked.connect(lambda _i: self.open_last_report())
        self.runs_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.runs_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.runs_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.runs_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.runs_table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
        run_hist_layout.addWidget(self.runs_table, stretch=1)
        self.run_tabs.addTab(run_hist_tab, "History")

        run_log_tab = QWidget()
        run_log_layout = QVBoxLayout(run_log_tab)
        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        run_log_layout.addWidget(self.log_box, stretch=1)
        self.run_tabs.addTab(run_log_tab, "Log")

        reports_tab = QWidget()
        reports_layout = QVBoxLayout(reports_tab)
        reports_hint = QLabel("Open selected run report or export the current project.")
        reports_hint.setWordWrap(True)
        reports_layout.addWidget(reports_hint)
        reports_actions = QHBoxLayout()
        self.open_report_btn = QPushButton("Open Report")
        self.open_report_btn.clicked.connect(self.open_last_report)
        self.export_zip_btn = QPushButton("Export ZIP")
        self.export_zip_btn.clicked.connect(self.export_zip)
        reports_actions.addWidget(self.open_report_btn)
        reports_actions.addWidget(self.export_zip_btn)
        reports_actions.addStretch(1)
        reports_layout.addLayout(reports_actions)
        reports_layout.addStretch(1)
        self.run_tabs.addTab(reports_tab, "Reports")

        cl.addWidget(self.run_activity_box, stretch=1)
        split.addWidget(center)

        # Right pane: Step + Run tabs
        right = QWidget()
        rl = QVBoxLayout(right)
        self.right_tabs = QTabWidget()
        rl.addWidget(self.right_tabs, stretch=1)

        step_tab = QWidget()
        step_layout = QVBoxLayout(step_tab)
        rt = QLabel("Step Inspector")
        rt.setObjectName("paneTitle")
        step_layout.addWidget(rt)
        self.inspector_step_title = QLabel("Step -")
        self.inspector_step_title.setStyleSheet("font-weight: 700; color: #334155;")
        step_layout.addWidget(self.inspector_step_title)
        self.inspector_step_badge = QLabel("")
        self.inspector_step_badge.setVisible(False)
        self.inspector_step_badge.setStyleSheet(
            f"background:{SELECTED_BG}; color:{TEXT}; border:1px solid {BORDER}; border-radius: 6px; padding:2px 8px; max-width: 180px;"
        )
        step_layout.addWidget(self.inspector_step_badge)
        self.inspector_hint = QLabel("No step selected\n\nSelect a step from the builder, or add a new step using the buttons above.")
        self.inspector_hint.setObjectName("muted")
        self.inspector_hint.setWordWrap(True)
        step_layout.addWidget(self.inspector_hint)
        self.step_editor_widget = QWidget()
        step_editor_layout = QVBoxLayout(self.step_editor_widget)
        step_editor_layout.setContentsMargins(0, 0, 0, 0)
        self.step_panels = QStackedWidget()
        step_editor_layout.addWidget(self.step_panels, stretch=1)
        self.panel_empty = QLabel("No editable fields for this step type.")
        self.step_panels.addWidget(self._wrap_panel(self.panel_empty))
        self.panel_click = self._build_click_panel()
        self.step_panels.addWidget(self.panel_click)
        self.panel_type = self._build_type_panel()
        self.step_panels.addWidget(self.panel_type)
        self.panel_wait = self._build_wait_panel()
        self.step_panels.addWidget(self.panel_wait)
        self.panel_screenshot = self._build_screenshot_panel()
        self.step_panels.addWidget(self.panel_screenshot)
        self.panel_reusable = self._build_reusable_panel()
        self.step_panels.addWidget(self.panel_reusable)
        self.panel_comment = self._build_comment_panel()
        self.step_panels.addWidget(self.panel_comment)
        self.panel_press_key = self._build_key_panel()
        self.step_panels.addWidget(self.panel_press_key)
        self.panel_assert_text = self._build_assert_text_panel()
        self.step_panels.addWidget(self.panel_assert_text)
        self.panel_assert_file = self._build_assert_file_panel()
        self.step_panels.addWidget(self.panel_assert_file)
        step_layout.addWidget(self.step_editor_widget, stretch=1)

        self.advanced_box = QWidget()
        adv_container = QVBoxLayout(self.advanced_box)
        adv_container.setContentsMargins(0, 0, 0, 0)
        adv_container.setSpacing(4)
        self.advanced_toggle = QToolButton()
        self.advanced_toggle.setText("Advanced >")
        self.advanced_toggle.setToolButtonStyle(Qt.ToolButtonTextOnly)
        self.advanced_toggle.setCheckable(True)
        self.advanced_toggle.setChecked(False)
        self.advanced_toggle.clicked.connect(self._on_advanced_toggled)
        adv_container.addWidget(self.advanced_toggle)
        self.advanced_content = QWidget()
        adv_layout = QFormLayout(self.advanced_content)
        self.step_type_change = QComboBox()
        for code in [
            "comment",
            "click_xy",
            "type_text",
            "wait",
            "screenshot",
            "run_flow",
            "click",
            "double_click",
            "right_click",
            "press_key",
            "assert_window_title_contains",
            "assert_clipboard_contains",
            "assert_file_exists",
        ]:
            self.step_type_change.addItem(_title_for_step_type(code), code)
        self.adv_target = QLineEdit()
        self.adv_value = QLineEdit()
        self.adv_key = QLineEdit()
        self.adv_path = QLineEdit()
        self.adv_timeout = QDoubleSpinBox()
        self.adv_timeout.setRange(0, 3600)
        self.adv_timeout.setDecimals(3)
        self.adv_retry = QSpinBox()
        self.adv_retry.setRange(0, 20)
        self.adv_internal_type = QLineEdit()
        self.adv_raw_json = QTextEdit()
        self.adv_raw_json.setReadOnly(True)
        adv_layout.addRow("Change type", self.step_type_change)
        adv_layout.addRow("Raw target", self.adv_target)
        adv_layout.addRow("Raw value", self.adv_value)
        adv_layout.addRow("Key", self.adv_key)
        adv_layout.addRow("Path", self.adv_path)
        adv_layout.addRow("Timeout", self.adv_timeout)
        adv_layout.addRow("Retry count", self.adv_retry)
        adv_layout.addRow("Internal action type", self.adv_internal_type)
        adv_layout.addRow("Raw JSON preview", self.adv_raw_json)
        adv_container.addWidget(self.advanced_content)
        self.advanced_content.setVisible(False)
        self.step_type_change.currentIndexChanged.connect(self._on_inspector_changed)
        for w in [self.adv_target, self.adv_value, self.adv_key, self.adv_path, self.adv_internal_type]:
            w.textEdited.connect(self._on_inspector_changed)
        self.adv_timeout.valueChanged.connect(self._on_inspector_changed)
        self.adv_retry.valueChanged.connect(self._on_inspector_changed)
        step_layout.addWidget(self.advanced_box)
        step_layout.addStretch(1)
        self.right_tabs.addTab(step_tab, "Step")

        run_tab = QWidget()
        run_layout = QVBoxLayout(run_tab)
        rd_title = QLabel("Run Details")
        rd_title.setObjectName("paneTitle")
        run_layout.addWidget(rd_title)
        self.run_details = QTextEdit()
        self.run_details.setReadOnly(True)
        run_layout.addWidget(self.run_details, stretch=1)
        run_open_row = QHBoxLayout()
        self.run_open_report_btn = QPushButton("Open Report")
        self.run_open_report_btn.clicked.connect(self.open_last_report)
        run_open_row.addWidget(self.run_open_report_btn)
        run_open_row.addStretch(1)
        run_layout.addLayout(run_open_row)
        self.right_tabs.addTab(run_tab, "Run")
        split.addWidget(right)
        split.setSizes([280, 820, 400])

        self.step_content_stack.setCurrentIndex(1)
        self.step_empty_hint.setText(
            "Select or create a test case\n\nChoose a test case from the left, create a new one, import Excel/CSV, or start recording."
        )
        self.right_tabs.setCurrentIndex(0)
        self._on_run_activity_toggled(False)
        self._set_inspector_enabled(False)

    def _wrap_panel(self, inner: QWidget) -> QWidget:
        w = QWidget()
        v = QVBoxLayout(w)
        v.setContentsMargins(0, 0, 0, 0)
        v.addWidget(inner)
        v.addStretch(1)
        return w

    def _on_run_activity_toggled(self, expanded: bool):
        self.run_activity_toggle.setText("Run Activity v" if expanded else "Run Activity >")
        self.run_tabs.setVisible(expanded)
        if expanded:
            self.run_activity_box.setMaximumHeight(360)
            self.run_activity_box.setMinimumHeight(270)
        else:
            self.run_activity_box.setMaximumHeight(50)
            self.run_activity_box.setMinimumHeight(50)

    def _on_advanced_toggled(self, expanded: bool):
        self.advanced_toggle.setText("Advanced v" if expanded else "Advanced >")
        self.advanced_content.setVisible(expanded)

    def _set_recording_ui_state(self, is_recording: bool):
        self.act_start_recording.setVisible(not is_recording)
        self.act_start_recording.setEnabled(not is_recording)
        self.act_stop_recording.setVisible(is_recording)
        self.act_stop_recording.setEnabled(is_recording)

    def _build_click_panel(self) -> QWidget:
        w = QWidget()
        form = QFormLayout(w)
        self.click_desc = QLineEdit()
        self.click_x = QSpinBox()
        self.click_x.setRange(-100000, 100000)
        self.click_y = QSpinBox()
        self.click_y.setRange(-100000, 100000)
        self.click_wait_after = QDoubleSpinBox()
        self.click_wait_after.setRange(0, 3600)
        self.click_wait_after.setDecimals(3)
        self.click_capture_after = QCheckBox("Take screenshot after click")
        self.click_desc.textEdited.connect(self._on_inspector_changed)
        self.click_x.valueChanged.connect(self._on_inspector_changed)
        self.click_y.valueChanged.connect(self._on_inspector_changed)
        self.click_wait_after.valueChanged.connect(self._on_inspector_changed)
        self.click_capture_after.toggled.connect(self._on_inspector_changed)
        form.addRow("Description", self.click_desc)
        xy_row = QWidget()
        xy_layout = QHBoxLayout(xy_row)
        xy_layout.setContentsMargins(0, 0, 0, 0)
        xy_layout.addWidget(QLabel("x:"))
        xy_layout.addWidget(self.click_x)
        xy_layout.addWidget(QLabel("y:"))
        xy_layout.addWidget(self.click_y)
        xy_layout.addStretch(1)
        form.addRow("Position", xy_row)
        form.addRow("Wait after (seconds)", self.click_wait_after)
        form.addRow("", self.click_capture_after)
        return w

    def _build_type_panel(self) -> QWidget:
        w = QWidget()
        form = QFormLayout(w)
        self.type_desc = QLineEdit()
        self.type_text_value = QLineEdit()
        self.type_press_enter = QCheckBox("Press Enter after typing")
        self.type_wait_after = QDoubleSpinBox()
        self.type_wait_after.setRange(0, 3600)
        self.type_wait_after.setDecimals(3)
        self.type_desc.textEdited.connect(self._on_inspector_changed)
        self.type_text_value.textEdited.connect(self._on_inspector_changed)
        self.type_press_enter.toggled.connect(self._on_inspector_changed)
        self.type_wait_after.valueChanged.connect(self._on_inspector_changed)
        form.addRow("Description", self.type_desc)
        form.addRow("Text to type", self.type_text_value)
        form.addRow("", self.type_press_enter)
        form.addRow("Wait after (seconds)", self.type_wait_after)
        return w

    def _build_wait_panel(self) -> QWidget:
        w = QWidget()
        form = QFormLayout(w)
        self.wait_desc = QLineEdit()
        self.wait_seconds = QDoubleSpinBox()
        self.wait_seconds.setRange(0, 3600)
        self.wait_seconds.setDecimals(3)
        self.wait_desc.textEdited.connect(self._on_inspector_changed)
        self.wait_seconds.valueChanged.connect(self._on_inspector_changed)
        form.addRow("Description", self.wait_desc)
        form.addRow("Duration (seconds)", self.wait_seconds)
        return w

    def _build_screenshot_panel(self) -> QWidget:
        w = QWidget()
        form = QFormLayout(w)
        self.screenshot_desc = QLineEdit()
        self.screenshot_name = QLineEdit()
        self.screenshot_expected = QTextEdit()
        self.screenshot_expected.setMaximumHeight(90)
        self.screenshot_desc.textEdited.connect(self._on_inspector_changed)
        self.screenshot_name.textEdited.connect(self._on_inspector_changed)
        self.screenshot_expected.textChanged.connect(self._on_inspector_changed)
        form.addRow("Description", self.screenshot_desc)
        form.addRow("Screenshot name", self.screenshot_name)
        form.addRow("Expected result / validation note", self.screenshot_expected)
        return w

    def _build_reusable_panel(self) -> QWidget:
        w = QWidget()
        v = QVBoxLayout(w)
        form = QFormLayout()
        self.reusable_desc = QLineEdit()
        self.reusable_action_select = QComboBox()
        self.reusable_action_select.setEditable(True)
        self.reusable_desc.textEdited.connect(self._on_inspector_changed)
        self.reusable_action_select.currentTextChanged.connect(self._on_inspector_changed)
        self.reusable_action_select.currentTextChanged.connect(self._refresh_reusable_preview)
        form.addRow("Description", self.reusable_desc)
        form.addRow("Select reusable action", self.reusable_action_select)
        v.addLayout(form)
        self.reusable_preview = QTextEdit()
        self.reusable_preview.setReadOnly(True)
        self.reusable_preview.setMaximumHeight(110)
        self.reusable_preview.setPlaceholderText("Preview of included steps")
        v.addWidget(self.reusable_preview)
        return w

    def _build_comment_panel(self) -> QWidget:
        w = QWidget()
        form = QFormLayout(w)
        self.comment_desc = QTextEdit()
        self.comment_desc.setMaximumHeight(110)
        self.comment_expected = QTextEdit()
        self.comment_expected.setMaximumHeight(90)
        self.comment_desc.textChanged.connect(self._on_inspector_changed)
        self.comment_expected.textChanged.connect(self._on_inspector_changed)
        form.addRow("Description", self.comment_desc)
        form.addRow("Expected result / note", self.comment_expected)
        return w

    def _build_key_panel(self) -> QWidget:
        w = QWidget()
        form = QFormLayout(w)
        self.press_key_value = QLineEdit()
        self.press_key_value.textEdited.connect(self._on_inspector_changed)
        form.addRow("Key", self.press_key_value)
        return w

    def _build_assert_text_panel(self) -> QWidget:
        w = QWidget()
        form = QFormLayout(w)
        self.assert_text = QLineEdit()
        self.assert_text.textEdited.connect(self._on_inspector_changed)
        form.addRow("Expected text", self.assert_text)
        return w

    def _build_assert_file_panel(self) -> QWidget:
        w = QWidget()
        form = QFormLayout(w)
        self.assert_file_path = QLineEdit()
        self.assert_file_path.textEdited.connect(self._on_inspector_changed)
        form.addRow("File path", self.assert_file_path)
        return w

    def _bind_shortcuts(self):
        for seq, fn in [
            (QKeySequence("Ctrl+S"), lambda: self.save_project(clear_dirty=True)),
            (QKeySequence("F5"), self.run_selected_flow),
            (QKeySequence("Shift+F5"), self.dry_run_selected_flow),
            (QKeySequence("Delete"), self.delete_step),
            (QKeySequence("F8"), self.stop_run),
        ]:
            a = QAction(self)
            a.setShortcut(seq)
            a.triggered.connect(fn)
            self.addAction(a)

    def _check_dependencies(self):
        missing = []
        for mod, pip_name in [
            ("PIL", "Pillow"),
            ("pyautogui", "pyautogui"),
            ("pynput", "pynput"),
            ("openpyxl", "openpyxl"),
        ]:
            try:
                __import__(mod)
            except Exception:
                missing.append(pip_name)
        if not missing:
            self.log("Dependencies OK.")
            return
        msg = "Missing required packages:\n- " + "\n- ".join(missing) + "\n\nInstall now?"
        if QMessageBox.question(self, "Setup", msg) == QMessageBox.Yes:
            cmd = [sys.executable, "-m", "pip", "install", *missing]
            proc = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
            if proc.returncode == 0:
                QMessageBox.information(self, "Setup", "Installed. Restart app.")
            else:
                QMessageBox.critical(self, "Setup Failed", (proc.stderr or proc.stdout or "")[:1500])

    def log(self, msg: str):
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_box.append(f"[{ts}] {msg}")

    def set_status(self, text: str):
        self.status.setText(text)

    def _update_window_title(self):
        suffix = " — Unsaved changes" if self._dirty else ""
        self.setWindowTitle(f"{self.base_window_title}{suffix}")

    def _set_dirty(self, dirty: bool = True):
        self._dirty = dirty
        self._update_window_title()

    def save_project(self, *, quiet: bool = False, clear_dirty: bool = False):
        storage_save_project(self.data)
        if clear_dirty:
            self._set_dirty(False)
        if not quiet:
            self.log("Project saved.")

    def _save_after_change(self):
        self._set_dirty(True)
        self.save_project(quiet=True)

    # Test Case CRUD (internally flow-backed)
    def refresh_flows(self):
        selected = self.current_flow_name
        query = self.test_case_search.text().strip().lower()
        self.flow_list.clear()
        total_cases = len(self.data.get("flows", {}))
        for name, flow in sorted(self.data.get("flows", {}).items()):
            if query and query not in name.lower() and query not in str(flow.get("description", "")).lower():
                continue
            steps_count = len(flow.get("steps", [])) if isinstance(flow.get("steps", []), list) else 0
            last_status = self._last_run_status_for(name)
            meta = f"{steps_count} steps · {last_status or 'Not run'}"
            item = QListWidgetItem(f"{name}\n{meta}")
            item.setSizeHint(QSize(100, 42))
            item.setToolTip(name)
            item.setData(Qt.UserRole, name)
            self.flow_list.addItem(item)

        if self.flow_list.count() == 0:
            self.test_case_empty_hint.setVisible(True)
            if total_cases == 0:
                self.test_case_empty_hint.setText(
                    "No test cases yet\n\nCreate a new test, import Excel/CSV, or start recording."
                )
            elif query:
                self.test_case_empty_hint.setText("No test cases found.")
            else:
                self.test_case_empty_hint.setText("No test cases available.")
        else:
            self.test_case_empty_hint.setVisible(False)

        if selected:
            for i in range(self.flow_list.count()):
                it = self.flow_list.item(i)
                if str(it.data(Qt.UserRole)) == selected:
                    self.flow_list.setCurrentRow(i)
                    break
        self._update_action_states()

    def _last_run_status_for(self, flow_name: str) -> str:
        for run in reversed(self.data.get("runs", [])):
            if str(run.get("name", "")).strip() == flow_name:
                return _normalize_status(str(run.get("status", "")), bool(run.get("dryRun", False)))
        return ""

    def selected_flow_name(self) -> str | None:
        item = self.flow_list.currentItem()
        if not item:
            return None
        return str(item.data(Qt.UserRole) or "").strip() or None

    def new_flow(self):
        name, ok = QInputDialog.getText(self, "New Test Case", "Test Case name:")
        if not ok or not name.strip():
            return
        n = name.strip()
        if n in self.data.get("flows", {}):
            QMessageBox.warning(self, "Exists", "Test Case already exists.")
            return
        self.data.setdefault("flows", {})[n] = {"name": n, "description": "", "parameters": [], "steps": []}
        self.current_flow_name = n
        self._save_after_change()
        self.refresh_flows()
        self.on_flow_select()

    def delete_flow(self):
        name = self.selected_flow_name()
        if not name:
            return
        if QMessageBox.question(self, "Delete Test", "Delete this test case?\nThis cannot be undone.") != QMessageBox.Yes:
            return
        self.data.get("flows", {}).pop(name, None)
        self.current_flow_name = None
        self.flow_name_edit.setText("")
        self.steps_table.setRowCount(0)
        self._set_inspector_enabled(False)
        self._save_after_change()
        self.refresh_flows()

    def rename_flow(self):
        if not self.current_flow_name:
            return
        new_name = self.flow_name_edit.text().strip()
        if not new_name or new_name == self.current_flow_name:
            return
        if new_name in self.data.get("flows", {}):
            QMessageBox.warning(self, "Exists", "Test Case name already exists.")
            self.flow_name_edit.setText(self.current_flow_name)
            return
        flows = self.data["flows"]
        flows[new_name] = flows.pop(self.current_flow_name)
        flows[new_name]["name"] = new_name
        self.current_flow_name = new_name
        self._save_after_change()
        self.refresh_flows()

    def on_flow_select(self):
        name = self.selected_flow_name()
        if not name:
            self.current_flow_name = None
            self.flow_name_edit.setText("")
            self.steps_table.setRowCount(0)
            self.step_content_stack.setCurrentIndex(1)
            self.step_empty_hint.setText(
                "Select or create a test case\n\nChoose a test case from the left, create a new one, import Excel/CSV, or start recording."
            )
            self._set_inspector_enabled(False)
            return
        self.current_flow_name = name
        self.flow_name_edit.setText(name)
        self.refresh_steps()

    def steps_ref(self) -> list[dict]:
        if not self.current_flow_name:
            return []
        flow = self.data.get("flows", {}).get(self.current_flow_name, {})
        steps = flow.get("steps", [])
        return steps if isinstance(steps, list) else []

    def _step_description(self, step: dict[str, Any]) -> str:
        if step.get("description"):
            return str(step.get("description"))
        t = str(step.get("type", ""))
        if t == "comment":
            return str(step.get("text", "Manual step"))
        if t in {"click", "click_xy", "double_click", "right_click"}:
            ta = self._step_target_action(step)
            return f"{_title_for_step_type(t)} at {ta}" if ta else _title_for_step_type(t)
        if t == "type_text":
            val = str(step.get("value", ""))
            return f"Type '{val[:30]}'" if val else "Type input"
        if t == "wait":
            return f"Wait {self._step_wait(step)}"
        if t == "screenshot":
            return f"Capture screenshot {step.get('name', 'step')}"
        if t == "run_flow":
            return f"Run reusable action {step.get('flow', '')}".strip()
        return _title_for_step_type(t)

    def _step_target_action(self, step: dict[str, Any]) -> str:
        t = str(step.get("type", ""))
        if t == "wait":
            return ""
        if t in {"run_flow", "subflow"}:
            return str(step.get("flow", ""))
        if t == "screenshot":
            return str(step.get("name", ""))
        if t in {"click", "double_click", "right_click"}:
            if step.get("target"):
                return str(step.get("target"))
            if "x" in step and "y" in step:
                return f"x:{step.get('x')} y:{step.get('y')}"
            return ""
        if t == "click_xy":
            return f"x:{step.get('x', '')} y:{step.get('y', '')}"
        if t == "press_key":
            return str(step.get("key", ""))
        if t == "hotkey":
            keys = step.get("keys", [])
            return " + ".join(str(k) for k in keys) if isinstance(keys, list) else ""
        if t == "assert_file_exists":
            return str(step.get("path", ""))
        if t == "assert_window_title_contains":
            return "Window title"
        if t == "assert_clipboard_contains":
            return "Clipboard"
        return str(step.get("target", ""))

    def _step_input(self, step: dict[str, Any]) -> str:
        t = str(step.get("type", ""))
        if t == "type_text":
            return str(step.get("value", ""))
        if t == "screenshot":
            return ""
        if t == "comment":
            return str(step.get("value", "")).strip()
        if t in {"wait", "click", "click_xy", "double_click", "right_click", "run_flow", "subflow"}:
            return ""
        if t in {"assert_window_title_contains", "assert_clipboard_contains"}:
            return str(step.get("value", ""))
        return str(step.get("value", step.get("path", "")))

    def _step_wait(self, step: dict[str, Any]) -> str:
        sec = step.get("seconds", "")
        if sec in (None, ""):
            sec = step.get("wait_after_seconds", "")
        if sec in (None, ""):
            return ""
        try:
            return f"{float(sec):g}s"
        except Exception:
            return str(sec)

    def refresh_steps(self):
        steps = self.steps_ref()
        self.steps_table.setRowCount(len(steps))
        for i, s in enumerate(steps):
            self._refresh_step_row(i, s)

        if not self.current_flow_name:
            self.step_content_stack.setCurrentIndex(1)
            self.step_empty_hint.setText(
                "Select or create a test case\n\nChoose a test case from the left, create a new one, import Excel/CSV, or start recording."
            )
            self.current_step_index = None
            self._set_inspector_enabled(False)
            return

        if steps:
            self.step_content_stack.setCurrentIndex(0)
            idx = self.current_step_index if self.current_step_index is not None else 0
            idx = max(0, min(idx, len(steps) - 1))
            self.steps_table.selectRow(idx)
            self.current_step_index = idx
            self._set_inspector_enabled(True)
            self._populate_step_inspector()
        else:
            self.step_content_stack.setCurrentIndex(1)
            self.step_empty_hint.setText(
                "No steps yet\n\nAdd a step manually, import steps from Excel/CSV, or start recording."
            )
            self.current_step_index = None
            self._set_inspector_enabled(False)

    def _refresh_step_row(self, row_index: int, step: dict[str, Any]):
        vals = [
            str(row_index + 1),
            _title_for_step_type(str(step.get("type", ""))),
            self._step_description(step),
            self._step_target_action(step),
            self._step_input(step),
            self._step_wait(step),
        ]
        for c, v in enumerate(vals):
            self.steps_table.setItem(row_index, c, QTableWidgetItem(v))

    def _default_step_for_type(self, step_type: str) -> dict[str, Any]:
        out: dict[str, Any] = {"type": step_type, "enabled": True}
        if step_type == "click_xy":
            out.update({"x": 0, "y": 0, "description": "Click at x:0 y:0"})
        elif step_type == "type_text":
            out.update({"value": "", "description": "Type value"})
        elif step_type == "wait":
            out.update({"seconds": 1.0, "description": "Wait 1 second"})
        elif step_type == "screenshot":
            out.update({"name": "step", "description": "Capture screenshot"})
        elif step_type == "run_flow":
            out.update({"flow": "", "description": "Run reusable action"})
        elif step_type == "comment":
            out.update({"text": "Manual step"})
        else:
            out.update({"description": _title_for_step_type(step_type)})
        return out

    def on_step_select(self):
        rows = self.steps_table.selectionModel().selectedRows()
        self.current_step_index = rows[0].row() if rows else None
        self._populate_step_inspector()
        self._update_action_states()

    def add_step(self, step_type: str):
        if not self.current_flow_name:
            QMessageBox.warning(self, "No Test Case", "Create/select a test case first.")
            return
        self.steps_ref().append(self._default_step_for_type(step_type))
        self.current_step_index = len(self.steps_ref()) - 1
        self._save_after_change()
        self.refresh_steps()

    def duplicate_step(self):
        steps = self.steps_ref()
        if self.current_step_index is None or not (0 <= self.current_step_index < len(steps)):
            return
        steps.insert(self.current_step_index + 1, copy.deepcopy(steps[self.current_step_index]))
        self.current_step_index += 1
        self._save_after_change()
        self.refresh_steps()

    def delete_step(self):
        steps = self.steps_ref()
        if self.current_step_index is None or not (0 <= self.current_step_index < len(steps)):
            return
        if QMessageBox.question(self, "Delete Step", "Delete selected step?\nThis cannot be undone.") != QMessageBox.Yes:
            return
        steps.pop(self.current_step_index)
        if steps:
            self.current_step_index = min(self.current_step_index, len(steps) - 1)
        else:
            self.current_step_index = None
        self._save_after_change()
        self.refresh_steps()

    def move_step(self, delta: int):
        steps = self.steps_ref()
        i = self.current_step_index
        if i is None:
            return
        j = i + delta
        if not (0 <= i < len(steps) and 0 <= j < len(steps)):
            return
        steps[i], steps[j] = steps[j], steps[i]
        self.current_step_index = j
        self._save_after_change()
        self.refresh_steps()

    def _set_inspector_enabled(self, enabled: bool):
        for w in [
            self.step_type_change,
            self.step_panels,
            self.advanced_toggle,
            self.advanced_content,
        ]:
            w.setEnabled(enabled)
        self.inspector_hint.setVisible(not enabled)
        self.step_editor_widget.setVisible(enabled)
        self.advanced_box.setVisible(enabled)
        self.inspector_step_badge.setVisible(enabled)
        if not enabled:
            self.inspector_step_title.setText("Step -")
            self.inspector_hint.setText(
                "No step selected\n\nSelect a step from the builder, or add a new step using the buttons above."
            )
            self.inspector_step_badge.setText("")
        self._update_action_states()

    def _update_action_states(self):
        has_flow = bool(self.current_flow_name)
        steps = self.steps_ref() if has_flow else []
        has_steps = bool(steps)
        i = self.current_step_index
        step_selected = i is not None and 0 <= i < len(steps)
        self.del_flow_btn.setEnabled(has_flow)
        self.act_run_test.setEnabled(has_flow and has_steps)
        self.act_preview_run.setEnabled(has_flow and has_steps)
        self.dup_step_btn.setEnabled(step_selected)
        self.del_step_btn.setEnabled(step_selected)
        self.move_up_btn.setEnabled(step_selected and i is not None and i > 0)
        self.move_down_btn.setEnabled(step_selected and i is not None and i < (len(steps) - 1))

    def _focus_inspector_for_selected_step(self):
        steps = self.steps_ref()
        if self.current_step_index is None or not (0 <= self.current_step_index < len(steps)):
            return
        t = str(steps[self.current_step_index].get("type", ""))
        if t in {"click", "click_xy", "double_click", "right_click"}:
            self.click_desc.setFocus()
        elif t == "type_text":
            self.type_text_value.setFocus()
        elif t == "wait":
            self.wait_seconds.setFocus()
        elif t == "screenshot":
            self.screenshot_name.setFocus()
        elif t in {"run_flow", "subflow"}:
            self.reusable_action_select.setFocus()
        elif t == "comment":
            self.comment_desc.setFocus()

    def _panel_index_for_type(self, step_type: str) -> int:
        mapping = {
            "click": 1,
            "click_xy": 1,
            "double_click": 1,
            "double_click_xy": 1,
            "right_click": 1,
            "type_text": 2,
            "wait": 3,
            "screenshot": 4,
            "run_flow": 5,
            "subflow": 5,
            "comment": 6,
            "press_key": 7,
            "assert_window_title_contains": 8,
            "assert_clipboard_contains": 8,
            "assert_file_exists": 9,
        }
        return mapping.get(step_type, 0)

    def _populate_step_inspector(self):
        steps = self.steps_ref()
        if self.current_step_index is None or not (0 <= self.current_step_index < len(steps)):
            self._set_inspector_enabled(False)
            return

        self._inspector_loading = True
        self._set_inspector_enabled(True)
        step = steps[self.current_step_index]
        st = str(step.get("type", "comment"))
        self.inspector_step_title.setText(f"Step {self.current_step_index + 1}")
        self.inspector_step_badge.setText(_title_for_step_type(st))
        self.step_panels.setCurrentIndex(self._panel_index_for_type(st))
        idx = self.step_type_change.findData("run_flow" if st == "subflow" else st)
        if idx >= 0:
            self.step_type_change.setCurrentIndex(idx)

        if st in {"click", "click_xy", "double_click", "right_click"}:
            self.click_desc.setText(str(step.get("description", f"Click at x:{step.get('x', 0)} y:{step.get('y', 0)}")))
            self.click_x.setValue(int(float(step.get("x", 0) or 0)))
            self.click_y.setValue(int(float(step.get("y", 0) or 0)))
            self.click_wait_after.setValue(float(step.get("wait_after_seconds", 0.0) or 0.0))
            self.click_capture_after.setChecked(bool(step.get("capture_after_click", False)))

        if st == "type_text":
            self.type_desc.setText(str(step.get("description", "Type")))
            self.type_text_value.setText(str(step.get("value", "")))
            self.type_press_enter.setChecked(bool(step.get("press_enter", False)))
            self.type_wait_after.setValue(float(step.get("wait_after_seconds", 0.0) or 0.0))

        if st == "wait":
            self.wait_desc.setText(str(step.get("description", f"Wait {float(step.get('seconds', 0.0) or 0.0):g}s")))
            self.wait_seconds.setValue(float(step.get("seconds", 0.0) or 0.0))

        if st == "screenshot":
            self.screenshot_desc.setText(str(step.get("description", "Screenshot")))
            self.screenshot_name.setText(str(step.get("name", "step")))
            self.screenshot_expected.setPlainText(str(step.get("expected_result", "")))

        if st in {"run_flow", "subflow"}:
            self.reusable_desc.setText(str(step.get("description", "Reusable action")))
            self._refresh_reusable_actions()
            self.reusable_action_select.setCurrentText(str(step.get("flow", "")))
            self._refresh_reusable_preview()

        if st == "comment":
            self.comment_desc.setPlainText(str(step.get("description", step.get("text", ""))))
            self.comment_expected.setPlainText(str(step.get("expected_result", "")))

        if st == "press_key":
            self.press_key_value.setText(str(step.get("key", "enter")))

        if st in {"assert_window_title_contains", "assert_clipboard_contains"}:
            self.assert_text.setText(str(step.get("value", "")))

        if st == "assert_file_exists":
            self.assert_file_path.setText(str(step.get("path", "")))

        self.adv_target.setText(str(step.get("target", "")))
        self.adv_value.setText(str(step.get("value", "")))
        self.adv_key.setText(str(step.get("key", "")))
        self.adv_path.setText(str(step.get("path", "")))
        self.adv_timeout.setValue(float(step.get("timeout", 0.0) or 0.0))
        self.adv_retry.setValue(int(step.get("retry_count", 0) or 0))
        self.adv_internal_type.setText(str(step.get("type", "")))
        self.adv_raw_json.setPlainText(json.dumps(step, indent=2, ensure_ascii=False))

        self._inspector_loading = False

    def _parse_click_position(self, raw: str) -> tuple[str, int | None, int | None]:
        txt = raw.strip()
        if not txt:
            return "", None, None
        lowered = txt.lower()
        if "x:" in lowered and "y:" in lowered:
            try:
                x_part = lowered.split("x:", 1)[1].split()[0]
                y_part = lowered.split("y:", 1)[1].split()[0]
                return "", int(float(x_part)), int(float(y_part))
            except Exception:
                return txt, None, None
        if "," in txt:
            parts = [p.strip() for p in txt.split(",", 1)]
            if len(parts) == 2:
                try:
                    return "", int(float(parts[0])), int(float(parts[1]))
                except Exception:
                    pass
        return txt, None, None

    def _on_inspector_changed(self):
        if self._inspector_loading:
            return
        steps = self.steps_ref()
        if self.current_step_index is None or not (0 <= self.current_step_index < len(steps)):
            return

        step = steps[self.current_step_index]
        new_type = str(self.step_type_change.currentData() or step.get("type", "comment")).strip()
        self.step_panels.setCurrentIndex(self._panel_index_for_type(new_type))
        base = copy.deepcopy(step)
        type_changed = str(base.get("type", "")) != new_type

        if type_changed:
            base = self._default_step_for_type(new_type)
            base["enabled"] = step.get("enabled", True)

        base["type"] = new_type

        # Type-specific friendly fields
        if new_type in {"click", "click_xy", "double_click", "right_click"}:
            base["description"] = self.click_desc.text().strip() or f"Click at x:{self.click_x.value()} y:{self.click_y.value()}"
            base["x"] = int(self.click_x.value())
            base["y"] = int(self.click_y.value())
            base.pop("target", None)
            if new_type == "click":
                base["type"] = "click_xy"
            wait_after = float(self.click_wait_after.value())
            if wait_after > 0:
                base["wait_after_seconds"] = wait_after
            else:
                base.pop("wait_after_seconds", None)
            base["capture_after_click"] = bool(self.click_capture_after.isChecked())

        elif new_type == "type_text":
            base["description"] = self.type_desc.text().strip() or "Type"
            base["value"] = self.type_text_value.text()
            base["press_enter"] = bool(self.type_press_enter.isChecked())
            wait_after = float(self.type_wait_after.value())
            if wait_after > 0:
                base["wait_after_seconds"] = wait_after
            else:
                base.pop("wait_after_seconds", None)

        elif new_type == "wait":
            base["description"] = self.wait_desc.text().strip() or f"Wait {float(self.wait_seconds.value()):g}s"
            base["seconds"] = float(self.wait_seconds.value())

        elif new_type == "screenshot":
            base["description"] = self.screenshot_desc.text().strip() or "Screenshot"
            base["name"] = self.screenshot_name.text().strip() or "step"
            exp = self.screenshot_expected.toPlainText().strip()
            if exp:
                base["expected_result"] = exp
            else:
                base.pop("expected_result", None)
            base.pop("delay_before_seconds", None)

        elif new_type == "run_flow":
            base["description"] = self.reusable_desc.text().strip() or "Reusable action"
            base["flow"] = self.reusable_action_select.currentText().strip()

        elif new_type == "comment":
            manual_desc = self.comment_desc.toPlainText().strip() or "Manual step"
            base["text"] = manual_desc
            base["description"] = manual_desc
            note = self.comment_expected.toPlainText().strip()
            if note:
                base["expected_result"] = note
            else:
                base.pop("expected_result", None)

        elif new_type == "press_key":
            base["key"] = self.press_key_value.text().strip() or "enter"
            base["description"] = base.get("description") or f"Press {base['key']}"

        elif new_type in {"assert_window_title_contains", "assert_clipboard_contains"}:
            base["value"] = self.assert_text.text().strip()

        elif new_type == "assert_file_exists":
            base["path"] = self.assert_file_path.text().strip()

        # Advanced overrides
        adv_target = self.adv_target.text().strip()
        adv_value = self.adv_value.text()
        adv_key = self.adv_key.text().strip()
        adv_path = self.adv_path.text().strip()

        if adv_target:
            base["target"] = adv_target
        if adv_value:
            base["value"] = adv_value
        if adv_key:
            base["key"] = adv_key
        if adv_path:
            base["path"] = adv_path

        timeout = float(self.adv_timeout.value())
        retry = int(self.adv_retry.value())
        if timeout > 0:
            base["timeout"] = timeout
        else:
            base.pop("timeout", None)
        if retry > 0:
            base["retry_count"] = retry
        else:
            base.pop("retry_count", None)

        steps[self.current_step_index] = base
        self.inspector_step_title.setText(f"Step {self.current_step_index + 1}")
        self.inspector_step_badge.setText(_title_for_step_type(str(base.get("type", ""))))
        self.adv_raw_json.setPlainText(json.dumps(base, indent=2, ensure_ascii=False))
        self._save_after_change()
        if type_changed:
            self.refresh_steps()
        else:
            self._refresh_step_row(self.current_step_index, base)

    def _refresh_reusable_actions(self):
        current = self.reusable_action_select.currentText().strip()
        self.reusable_action_select.blockSignals(True)
        self.reusable_action_select.clear()
        self.reusable_action_select.addItem("")
        for fname in sorted(self.data.get("flows", {}).keys()):
            if fname != (self.current_flow_name or ""):
                self.reusable_action_select.addItem(fname)
        self.reusable_action_select.setCurrentText(current)
        self.reusable_action_select.blockSignals(False)

    def _refresh_reusable_preview(self):
        action_name = self.reusable_action_select.currentText().strip()
        flow = self.data.get("flows", {}).get(action_name)
        if not action_name or not isinstance(flow, dict):
            self.reusable_preview.setPlainText("")
            return
        lines = []
        for idx, step in enumerate(flow.get("steps", []), start=1):
            if isinstance(step, dict):
                lines.append(f"{idx}. {_title_for_step_type(str(step.get('type', '')))} - {self._step_description(step)}")
        self.reusable_preview.setPlainText("\n".join(lines) if lines else "No steps in this reusable action.")

    # Recorder
    def start_recording(self):
        status = self.recorder.start(
            stop_key_name=str(self.data.get("settings", {}).get("recordingStopHotkey", "f8")),
            on_finished=self._on_recording_finished,
            record_typing=True,
            record_hotkeys=True,
        )
        if not status.available:
            QMessageBox.critical(self, "Recorder", status.message)
            return
        self._set_recording_ui_state(True)
        self.set_status("Recording...")
        self.log("Recording started.")

    def stop_recording(self):
        if self.recorder.is_running:
            self.recorder.stop()
            self.log("Stopping recorder...")
        self._set_recording_ui_state(False)

    def _normalize_recorded_steps(self, steps: list[dict]) -> list[dict]:
        normalized: list[dict[str, Any]] = []

        def _append_step(step: dict[str, Any]):
            if normalized and str(normalized[-1].get("type", "")) == "wait" and str(step.get("type", "")) == "wait":
                prev = float(normalized[-1].get("seconds", 0.0) or 0.0)
                curr = float(step.get("seconds", 0.0) or 0.0)
                merged = round(prev + curr, 3)
                normalized[-1]["seconds"] = merged
                normalized[-1]["description"] = f"Wait {merged:g}s"
                return
            normalized.append(step)

        for raw in steps:
            if not isinstance(raw, dict):
                continue
            s = copy.deepcopy(raw)
            t = str(s.get("type", "")).strip()

            if t == "wait":
                seconds = float(s.get("seconds", 0.0) or 0.0)
                if seconds < MIN_RECORDED_WAIT_SECONDS:
                    continue
                if seconds < MIN_EXPLICIT_WAIT_SECONDS and normalized:
                    prev = normalized[-1]
                    prev_wait = float(prev.get("wait_after_seconds", 0.0) or 0.0)
                    prev["wait_after_seconds"] = round(prev_wait + seconds, 3)
                    continue
                _append_step(
                    {
                        "type": "wait",
                        "seconds": round(seconds, 3),
                        "enabled": True,
                        "description": f"Wait {seconds:g}s",
                    }
                )
                continue

            if t == "type_text":
                txt = str(s.get("value", ""))
                if normalized and str(normalized[-1].get("type", "")) == "type_text":
                    normalized[-1]["value"] = str(normalized[-1].get("value", "")) + txt
                    merged = str(normalized[-1].get("value", ""))
                    normalized[-1]["description"] = f"Type '{merged[:40]}'" if merged else "Type"
                else:
                    _append_step(
                        {
                            "type": "type_text",
                            "value": txt,
                            "enabled": True,
                            "description": f"Type '{txt[:40]}'" if txt else "Type",
                        }
                    )
                continue

            if t == "click_xy":
                _append_step(
                    {
                        "type": "click_xy",
                        "x": int(float(s.get("x", 0) or 0)),
                        "y": int(float(s.get("y", 0) or 0)),
                        "enabled": True,
                        "description": f"Click at x:{int(float(s.get('x', 0) or 0))} y:{int(float(s.get('y', 0) or 0))}",
                    }
                )
                continue

            if t in {"double_click", "double_click_xy"}:
                x = int(float(s.get("x", 0) or 0))
                y = int(float(s.get("y", 0) or 0))
                _append_step(
                    {
                        "type": "double_click",
                        "x": x,
                        "y": y,
                        "enabled": True,
                        "description": f"Double click at x:{x} y:{y}",
                    }
                )
                continue

            if t == "right_click":
                x = int(float(s.get("x", 0) or 0))
                y = int(float(s.get("y", 0) or 0))
                _append_step(
                    {
                        "type": "right_click",
                        "x": x,
                        "y": y,
                        "enabled": True,
                        "description": f"Right click at x:{x} y:{y}",
                    }
                )
                continue

            if t == "screenshot":
                name = str(s.get("name", "step")).strip() or "step"
                _append_step({"type": "screenshot", "name": name, "enabled": True, "description": f"Capture screenshot {name}"})
                continue

            if t == "press_key":
                key = str(s.get("key", "enter")).strip() or "enter"
                _append_step({"type": "press_key", "key": key, "enabled": True, "description": f"Press {key}"})
                continue

            if t == "hotkey":
                keys = s.get("keys", [])
                if isinstance(keys, list) and keys:
                    _append_step({"type": "hotkey", "keys": keys, "enabled": True, "description": "Press " + " + ".join(str(k) for k in keys)})
                continue

            _append_step(s)

        # Merge consecutive waits created by edge cases in fallback branch.
        out: list[dict[str, Any]] = []
        for step in normalized:
            if out and str(out[-1].get("type", "")) == "wait" and str(step.get("type", "")) == "wait":
                merged = round(float(out[-1].get("seconds", 0.0) or 0.0) + float(step.get("seconds", 0.0) or 0.0), 3)
                out[-1]["seconds"] = merged
                out[-1]["description"] = f"Wait {merged:g}s"
            else:
                out.append(step)
        return out

    def _on_recording_finished(self, steps: list[dict], err: str):
        def ui():
            self._set_recording_ui_state(False)
            self.set_status("Ready")
            if err:
                self.log(f"Recorder error: {err}")
                QMessageBox.critical(self, "Recorder", err)
                return

            reviewed = self._open_recording_review_dialog(self._normalize_recorded_steps(steps))
            if reviewed is None:
                self.log("Recording discarded.")
                return

            action = reviewed["action"]
            reviewed_steps = reviewed["steps"]

            if action == "add_current":
                if not self.current_flow_name:
                    QMessageBox.warning(self, "Recording", "Select a test case first.")
                    return
                self.steps_ref().extend(copy.deepcopy(reviewed_steps))
                self._save_after_change()
                self.refresh_steps()
                self.log(f"Added {len(reviewed_steps)} recorded steps to '{self.current_flow_name}'.")
                return

            if action == "save_new":
                name = reviewed.get("name", "").strip()
                if not name:
                    return
                self.data.setdefault("flows", {})[name] = {
                    "name": name,
                    "description": "Recorded test case",
                    "parameters": [],
                    "steps": copy.deepcopy(reviewed_steps),
                }
                self.current_flow_name = name
                self._save_after_change()
                self.refresh_flows()
                self.refresh_steps()
                self.log(f"Saved recording as new test case '{name}'.")
                return

            if action == "save_reusable":
                name = reviewed.get("name", "").strip()
                if not name:
                    return
                self.data.setdefault("flows", {})[name] = {
                    "name": name,
                    "description": "Reusable action",
                    "parameters": [],
                    "steps": copy.deepcopy(reviewed_steps),
                }
                self._save_after_change()
                self.refresh_flows()
                self.log(f"Saved recording as reusable action '{name}'.")

        QTimer.singleShot(0, ui)

    def _open_recording_review_dialog(self, steps: list[dict]) -> dict[str, Any] | None:
        d = QDialog(self)
        d.setWindowTitle("Recording Review")
        d.resize(1000, 640)
        v = QVBoxLayout(d)

        v.addWidget(QLabel("Review recorded steps before saving."))
        table = QTableWidget(0, 6)
        table.setHorizontalHeaderLabels(["#", "Type", "Description", "Target / Action", "Input", "Wait"])
        table.setSelectionBehavior(QTableWidget.SelectRows)
        table.setEditTriggers(QTableWidget.NoEditTriggers)
        table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        table.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)
        table.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeToContents)
        v.addWidget(table, stretch=1)

        working_steps = [copy.deepcopy(s) for s in steps]

        def refresh_table():
            table.setRowCount(len(working_steps))
            for i, s in enumerate(working_steps):
                vals = [
                    str(i + 1),
                    _title_for_step_type(str(s.get("type", ""))),
                    str(s.get("description", s.get("text", ""))),
                    self._step_target_action(s),
                    self._step_input(s),
                    self._step_wait(s),
                ]
                for c, val in enumerate(vals):
                    item = QTableWidgetItem(val)
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                    table.setItem(i, c, item)

        actions = QHBoxLayout()
        b_rename = QPushButton("Rename Selected Step")
        b_delete = QPushButton("Delete Selected")
        b_group = QPushButton("Merge Selected Steps into Reusable Action")
        actions.addWidget(b_rename)
        actions.addWidget(b_delete)
        actions.addWidget(b_group)
        actions.addStretch(1)
        v.addLayout(actions)

        payload: dict[str, Any] = {}

        def rename_selected():
            rows = sorted({mi.row() for mi in table.selectionModel().selectedRows()})
            if not rows:
                QMessageBox.information(d, "Rename", "Select a step first.")
                return
            current = str(working_steps[rows[0]].get("description", working_steps[rows[0]].get("text", "")))
            new_desc, ok = QInputDialog.getText(d, "Rename Step Description", "Description:", text=current)
            if not ok:
                return
            txt = new_desc.strip()
            if txt:
                working_steps[rows[0]]["description"] = txt
                if str(working_steps[rows[0]].get("type", "")) == "comment":
                    working_steps[rows[0]]["text"] = txt
            else:
                working_steps[rows[0]].pop("description", None)
            refresh_table()

        def delete_selected():
            rows = sorted({mi.row() for mi in table.selectionModel().selectedRows()}, reverse=True)
            if not rows:
                return
            for r in rows:
                if 0 <= r < len(working_steps):
                    working_steps.pop(r)
            refresh_table()

        def group_selected():
            rows = sorted({mi.row() for mi in table.selectionModel().selectedRows()})
            if not rows:
                QMessageBox.information(d, "Reusable Action", "Select steps to group.")
                return
            name, ok = QInputDialog.getText(d, "Reusable Action", "Reusable action name:")
            if not ok or not name.strip():
                return
            n = name.strip()
            grouped = [copy.deepcopy(working_steps[r]) for r in rows if 0 <= r < len(working_steps)]
            self.data.setdefault("flows", {})[n] = {
                "name": n,
                "description": "Reusable action",
                "parameters": [],
                "steps": grouped,
            }
            first = rows[0]
            for r in reversed(rows):
                working_steps.pop(r)
            working_steps.insert(
                first,
                {
                    "type": "run_flow",
                    "flow": n,
                    "description": f"Run reusable action {n}",
                    "enabled": True,
                },
            )
            self._save_after_change()
            self.refresh_flows()
            refresh_table()

        b_rename.clicked.connect(rename_selected)
        b_delete.clicked.connect(delete_selected)
        b_group.clicked.connect(group_selected)

        button_row = QHBoxLayout()
        b_add_current = QPushButton("Add to Current Test Case")
        b_save_new = QPushButton("Save as New Test Case")
        b_save_reusable = QPushButton("Save as Reusable Action")
        b_discard = QPushButton("Discard")

        button_row.addWidget(b_add_current)
        button_row.addWidget(b_save_new)
        button_row.addWidget(b_save_reusable)
        button_row.addWidget(b_discard)
        v.addLayout(button_row)

        def done(action: str, name: str = ""):
            if not working_steps and action != "discard":
                QMessageBox.warning(d, "Recording", "No steps to save.")
                return
            payload["action"] = action
            payload["steps"] = working_steps
            payload["name"] = name
            d.accept()

        b_add_current.clicked.connect(lambda: done("add_current"))

        def _save_new():
            name, ok = QInputDialog.getText(d, "Save as New Test Case", "Test Case name:")
            if not ok or not name.strip():
                return
            done("save_new", name.strip())

        def _save_reusable():
            name, ok = QInputDialog.getText(d, "Save as Reusable Action", "Reusable action name:")
            if not ok or not name.strip():
                return
            done("save_reusable", name.strip())

        b_save_new.clicked.connect(_save_new)
        b_save_reusable.clicked.connect(_save_reusable)
        def _discard():
            if QMessageBox.question(d, "Discard Recording", "Discard this recording review?\nThis cannot be undone.") == QMessageBox.Yes:
                d.reject()

        b_discard.clicked.connect(_discard)

        refresh_table()
        if d.exec() != QDialog.Accepted:
            return None
        return payload

    # Import
    def import_test_cases(self):
        path, _ = QFileDialog.getOpenFileName(self, "Import test cases", "", "CSV/XLSX (*.csv *.xlsx);;All (*)")
        if not path:
            return
        try:
            headers, rows = read_table_rows(path)
        except ImporterError as exc:
            QMessageBox.critical(self, "Import", str(exc))
            return
        if not rows:
            QMessageBox.critical(self, "Import", "No rows found in file.")
            return

        config = self._open_import_wizard_dialog(path=path, headers=headers, rows=rows)
        if config is None:
            return
        flow_name = str(config.get("flow_name", "")).strip()
        built_steps = config.get("built_steps", [])
        if not flow_name or not isinstance(built_steps, list) or not built_steps:
            QMessageBox.warning(self, "Import", "No draft steps were generated.")
            return

        existing = self.data.setdefault("flows", {}).get(flow_name)
        if existing:
            existing.setdefault("steps", []).extend(built_steps)
        else:
            self.data.setdefault("flows", {})[flow_name] = {
                "name": flow_name,
                "description": f"Imported manual test case from {Path(path).name}",
                "parameters": [],
                "steps": built_steps,
            }
        self.current_flow_name = flow_name
        self._save_after_change()
        self.refresh_flows()
        if bool(config.get("edit_after_import", True)):
            self.refresh_steps()
        self.log(f"Created draft test case '{flow_name}' with {len(built_steps)} steps.")

    def _find_likely_column(self, headers: list[str], options: list[str]) -> str:
        lowered_map = {h.lower().strip(): h for h in headers}
        for opt in options:
            if opt.lower() in lowered_map:
                return lowered_map[opt.lower()]
        for h in headers:
            hl = h.lower()
            for opt in options:
                if opt.lower() in hl:
                    return h
        return ""

    def _parse_click_target_value(self, raw: str) -> tuple[str, int | None, int | None]:
        txt = str(raw or "").strip()
        if not txt:
            return "", None, None
        lowered = txt.lower()
        if "x:" in lowered and "y:" in lowered:
            try:
                x = int(float(lowered.split("x:", 1)[1].split()[0]))
                y = int(float(lowered.split("y:", 1)[1].split()[0]))
                return f"x:{x} y:{y}", x, y
            except Exception:
                return txt, None, None
        tokens = txt.replace(",", " ").split()
        if len(tokens) >= 2:
            try:
                x = int(float(tokens[0]))
                y = int(float(tokens[1]))
                return f"x:{x} y:{y}", x, y
            except Exception:
                pass
        return txt, None, None

    def _automation_instruction_to_step(self, row: dict[str, Any]) -> tuple[dict[str, Any] | None, list[str]]:
        warnings: list[str] = []
        manual_description = str(row.get("manual_description", "")).strip()
        expected_result = str(row.get("expected_result", "")).strip()
        notes = str(row.get("notes", "")).strip()
        source_screenshot = str(row.get("source_screenshot", "")).strip()
        automation_type = str(row.get("automation_type", "Manual") or "Manual").strip()
        target_action = str(row.get("target_action", "")).strip()
        input_value = str(row.get("input_value", "")).strip()
        wait_text = str(row.get("wait_text", "")).strip()
        screenshot_name = str(row.get("screenshot_name", "")).strip()

        wait_seconds: float | None = None
        if wait_text:
            try:
                wait_seconds = float(wait_text)
            except Exception:
                warnings.append("Invalid wait value; ignored.")

        meta = {}
        if expected_result:
            meta["expected_result"] = expected_result
        if notes:
            meta["notes"] = notes
        if source_screenshot:
            meta["source_screenshot"] = source_screenshot

        if automation_type == "Manual":
            step: dict[str, Any] = {"type": "comment", "description": manual_description or "Manual step", "enabled": True}
            step["text"] = manual_description or "Manual step"
            step.update(meta)
            return step, warnings

        if automation_type in {"Click", "Double Click"}:
            step = {"type": "click_xy" if automation_type == "Click" else "double_click", "enabled": True}
            if manual_description:
                step["description"] = manual_description
            target_display, x_val, y_val = self._parse_click_target_value(target_action)
            if x_val is not None and y_val is not None:
                step["x"] = x_val
                step["y"] = y_val
            elif target_action:
                step["target"] = target_action
            else:
                warnings.append("Click target is empty.")
            if wait_seconds is not None:
                step["wait_after_seconds"] = wait_seconds
            step.update(meta)
            return step, warnings

        if automation_type == "Type":
            step = {"type": "type_text", "enabled": True, "value": input_value}
            if manual_description:
                step["description"] = manual_description
            if not input_value:
                warnings.append("Type step has empty Input.")
            if wait_seconds is not None:
                step["wait_after_seconds"] = wait_seconds
            step.update(meta)
            return step, warnings

        if automation_type == "Wait":
            if wait_seconds is None:
                warnings.append("Wait step missing seconds.")
                wait_seconds = 0.0
            step = {
                "type": "wait",
                "enabled": True,
                "seconds": wait_seconds,
                "description": manual_description or f"Wait {wait_seconds:g}s",
            }
            step.update(meta)
            return step, warnings

        if automation_type == "Screenshot":
            name = screenshot_name or source_screenshot or f"screenshot_{int(time.time())}"
            step = {"type": "screenshot", "enabled": True, "name": name}
            if manual_description:
                step["description"] = manual_description
            step.update(meta)
            return step, warnings

        if automation_type == "Reusable Action":
            step = {"type": "run_flow", "enabled": True, "flow": target_action}
            if manual_description:
                step["description"] = manual_description
            if not target_action:
                warnings.append("Reusable Action is empty.")
            step.update(meta)
            return step, warnings

        if automation_type == "Press Key":
            step = {"type": "press_key", "enabled": True, "key": input_value or "enter"}
            if manual_description:
                step["description"] = manual_description
            if not input_value:
                warnings.append("Press Key missing input; defaulted to enter.")
            step.update(meta)
            return step, warnings

        step = {"type": "comment", "description": manual_description or "Manual step", "text": manual_description or "Manual step", "enabled": True}
        step.update(meta)
        warnings.append(f"Unknown automation type '{automation_type}', imported as Manual.")
        return step, warnings

    def _open_import_wizard_dialog(self, *, path: str, headers: list[str], rows: list[dict[str, Any]]) -> dict[str, Any] | None:
        d = QDialog(self)
        d.setWindowTitle("Import Manual Test Case")
        d.resize(1250, 800)
        v = QVBoxLayout(d)

        title = QLabel("Import Manual Test Case")
        title.setObjectName("paneTitle")
        v.addWidget(title)

        step_label = QLabel("Step 1 of 3: Select Manual Columns")
        helper_label = QLabel(
            "Convert manual test case rows into draft automation steps.\n"
            "Select manual columns first, then add automation instructions."
        )
        helper_label.setObjectName("muted")
        helper_label.setWordWrap(True)
        v.addWidget(step_label)
        v.addWidget(helper_label)

        pages = QStackedWidget()
        v.addWidget(pages, stretch=1)

        # Step 1: Select Manual Columns
        p1 = QWidget()
        p1v = QVBoxLayout(p1)
        p1v.addWidget(QLabel(f"File: {Path(path).name}"))
        p1v.addWidget(QLabel(f"Rows: {len(rows)}"))
        p1v.addWidget(QLabel(f"Detected columns: {', '.join(headers)}"))

        cols = [""] + headers

        def mk_combo(defaults: list[str]) -> QComboBox:
            cb = QComboBox()
            cb.addItems(cols)
            found = self._find_likely_column(headers, defaults)
            if found:
                cb.setCurrentText(found)
            return cb

        select_form = QFormLayout()
        desc_col = mk_combo(["description", "step description", "test step", "action", "instruction"])
        step_num_col = mk_combo(["step", "step no", "#", "number"])
        expected_col = mk_combo(["expected result", "result", "expected"])
        screenshot_col = mk_combo(["screenshot", "evidence", "attachment"])
        notes_col = mk_combo(["notes", "comment", "comments"])
        select_form.addRow("Step description column (required)", desc_col)
        select_form.addRow("Step number column (optional)", step_num_col)
        select_form.addRow("Expected result column (optional)", expected_col)
        select_form.addRow("Screenshot/reference column (optional)", screenshot_col)
        select_form.addRow("Notes column (optional)", notes_col)
        p1v.addLayout(select_form)

        preview = QTableWidget(0, len(headers))
        preview.setHorizontalHeaderLabels(headers)
        preview.setEditTriggers(QTableWidget.NoEditTriggers)
        preview.setSelectionBehavior(QTableWidget.SelectRows)
        max_rows = min(len(rows), 10)
        preview.setRowCount(max_rows)
        for r in range(max_rows):
            for c, h in enumerate(headers):
                preview.setItem(r, c, QTableWidgetItem(str(rows[r].get(h, ""))))
        preview.resizeColumnsToContents()
        p1v.addWidget(preview, stretch=1)
        pages.addWidget(p1)

        # Step 2: Add Automation Instructions
        p2 = QWidget()
        p2v = QVBoxLayout(p2)
        helper_row = QHBoxLayout()
        set_manual_btn = QPushButton("Set selected rows as Manual")
        set_click_btn = QPushButton("Set selected rows as Click")
        set_type_btn = QPushButton("Set selected rows as Type")
        set_action_btn = QPushButton("Set selected rows as Reusable Action")
        clear_sel_btn = QPushButton("Clear selected automation")
        apply_wait_btn = QPushButton("Apply wait to selected rows")
        helper_row.addWidget(set_manual_btn)
        helper_row.addWidget(set_click_btn)
        helper_row.addWidget(set_type_btn)
        helper_row.addWidget(set_action_btn)
        helper_row.addWidget(clear_sel_btn)
        helper_row.addWidget(apply_wait_btn)
        p2v.addLayout(helper_row)

        reusable_row = QHBoxLayout()
        reusable_row.addWidget(QLabel("Use selected reusable action for selected rows:"))
        reusable_pick = QComboBox()
        reusable_pick.addItem("")
        for fname in sorted(self.data.get("flows", {}).keys()):
            reusable_pick.addItem(fname)
        reusable_apply_btn = QPushButton("Apply")
        reusable_row.addWidget(reusable_pick)
        reusable_row.addWidget(reusable_apply_btn)
        reusable_row.addStretch(1)
        p2v.addLayout(reusable_row)

        instructions = QTableWidget(0, 9)
        instructions.setHorizontalHeaderLabels(
            ["#", "Manual Description", "Automation Type", "Target / Action", "Input", "Wait", "Expected Result", "Screenshot Name", "Notes"]
        )
        instructions.setSelectionBehavior(QTableWidget.SelectRows)
        instructions.setSelectionMode(QAbstractItemView.ExtendedSelection)
        instructions.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        instructions.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        instructions.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        instructions.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        instructions.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)
        instructions.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeToContents)
        instructions.horizontalHeader().setSectionResizeMode(6, QHeaderView.Stretch)
        instructions.horizontalHeader().setSectionResizeMode(7, QHeaderView.Stretch)
        instructions.horizontalHeader().setSectionResizeMode(8, QHeaderView.Stretch)
        p2v.addWidget(instructions, stretch=1)

        validation_label = QLabel("")
        validation_label.setObjectName("muted")
        validation_label.setWordWrap(True)
        p2v.addWidget(validation_label)
        pages.addWidget(p2)

        # Step 3: Preview Draft Test Case
        p3 = QWidget()
        p3v = QVBoxLayout(p3)
        p3v.addWidget(QLabel("Executable preview"))
        generated = QTableWidget(0, 6)
        generated.setHorizontalHeaderLabels(["#", "Type", "Description", "Target / Action", "Input", "Wait"])
        generated.setEditTriggers(QTableWidget.NoEditTriggers)
        generated.setSelectionBehavior(QTableWidget.SelectRows)
        generated.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        generated.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        generated.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        generated.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        generated.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)
        generated.horizontalHeader().setSectionResizeMode(5, QHeaderView.ResizeToContents)
        p3v.addWidget(generated, stretch=1)
        preview_validation_label = QLabel("")
        preview_validation_label.setObjectName("muted")
        preview_validation_label.setWordWrap(True)
        p3v.addWidget(preview_validation_label)

        flow_name_edit = QLineEdit(Path(path).stem)
        flow_form = QFormLayout()
        flow_form.addRow("Draft Test Case name", flow_name_edit)
        p3v.addLayout(flow_form)
        edit_after_import = QCheckBox("Edit generated steps after import")
        edit_after_import.setChecked(True)
        p3v.addWidget(edit_after_import)
        pages.addWidget(p3)

        nav = QHBoxLayout()
        back_btn = QPushButton("Back")
        next_btn = QPushButton("Next")
        import_btn = QPushButton("Create Draft Test Case")
        cancel_btn = QPushButton("Cancel")
        nav.addWidget(back_btn)
        nav.addWidget(next_btn)
        nav.addStretch(1)
        nav.addWidget(import_btn)
        nav.addWidget(cancel_btn)
        v.addLayout(nav)

        import_btn.setEnabled(False)

        prepared_rows: list[dict[str, Any]] = []
        built_steps_cache: list[dict[str, Any]] = []
        build_warnings: list[str] = []

        def _selected_rows() -> list[int]:
            return sorted({mi.row() for mi in instructions.selectionModel().selectedRows()})

        def _update_instruction_row(row_idx: int):
            if not (0 <= row_idx < len(prepared_rows)):
                return
            row = prepared_rows[row_idx]
            for c, key in enumerate(
                [
                    "step_no",
                    "manual_description",
                    "automation_type",
                    "target_action",
                    "input_value",
                    "wait_text",
                    "expected_result",
                    "screenshot_name",
                    "notes",
                ]
            ):
                if c == 2:
                    cb = QComboBox()
                    cb.addItems(AUTOMATION_TYPES)
                    cb.setCurrentText(str(row.get(key, "Manual") or "Manual"))

                    def _make_handler(r: int, combo: QComboBox):
                        def _h(_txt: str):
                            prepared_rows[r]["automation_type"] = combo.currentText()
                            _refresh_validation_summary()

                        return _h

                    cb.currentTextChanged.connect(_make_handler(row_idx, cb))
                    instructions.setCellWidget(row_idx, c, cb)
                else:
                    item = QTableWidgetItem(str(row.get(key, "")))
                    if c in (0, 1):
                        item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                    instructions.setItem(row_idx, c, item)

        def _populate_instruction_table():
            instructions.blockSignals(True)
            instructions.setRowCount(len(prepared_rows))
            for r in range(len(prepared_rows)):
                _update_instruction_row(r)
            instructions.blockSignals(False)
            _refresh_validation_summary()

        def _ingest_step1_rows():
            prepared_rows.clear()
            desc_name = desc_col.currentText().strip()
            step_name = step_num_col.currentText().strip()
            expected_name = expected_col.currentText().strip()
            screenshot_name_col = screenshot_col.currentText().strip()
            notes_name = notes_col.currentText().strip()
            if not desc_name:
                return
            for idx, src in enumerate(rows, start=1):
                desc_txt = str(src.get(desc_name, "")).strip()
                if not desc_txt:
                    continue
                step_no = str(src.get(step_name, "")).strip() if step_name else str(idx)
                prepared_rows.append(
                    {
                        "step_no": step_no or str(idx),
                        "manual_description": desc_txt,
                        "automation_type": "Manual",
                        "target_action": "",
                        "input_value": "",
                        "wait_text": "",
                        "expected_result": str(src.get(expected_name, "")).strip() if expected_name else "",
                        "screenshot_name": "",
                        "notes": str(src.get(notes_name, "")).strip() if notes_name else "",
                        "source_screenshot": str(src.get(screenshot_name_col, "")).strip() if screenshot_name_col else "",
                    }
                )
            _populate_instruction_table()

        def _refresh_validation_summary():
            warnings: list[str] = []
            for i, row in enumerate(prepared_rows, start=1):
                t = str(row.get("automation_type", "Manual"))
                if t == "Type" and not str(row.get("input_value", "")).strip():
                    warnings.append(f"Row {i}: Type requires Input.")
                if t == "Wait" and not str(row.get("wait_text", "")).strip():
                    warnings.append(f"Row {i}: Wait requires seconds.")
                if t in {"Click", "Double Click"} and not str(row.get("target_action", "")).strip():
                    warnings.append(f"Row {i}: Click target is empty.")
                if t == "Reusable Action" and not str(row.get("target_action", "")).strip():
                    warnings.append(f"Row {i}: Reusable Action is empty.")
            validation_label.setText("Warnings:\n- " + "\n- ".join(warnings) if warnings else "No validation warnings.")

        def _apply_type_to_selected(type_name: str):
            rows_sel = _selected_rows()
            if not rows_sel:
                return
            for r in rows_sel:
                if 0 <= r < len(prepared_rows):
                    prepared_rows[r]["automation_type"] = type_name
            _populate_instruction_table()

        def _clear_selected():
            rows_sel = _selected_rows()
            if not rows_sel:
                return
            if QMessageBox.question(d, "Clear Automation", "Clear selected automation instructions?\nThis cannot be undone.") != QMessageBox.Yes:
                return
            for r in rows_sel:
                if 0 <= r < len(prepared_rows):
                    prepared_rows[r]["automation_type"] = "Manual"
                    prepared_rows[r]["target_action"] = ""
                    prepared_rows[r]["input_value"] = ""
                    prepared_rows[r]["wait_text"] = ""
                    prepared_rows[r]["screenshot_name"] = ""
            _populate_instruction_table()

        def _apply_wait_to_selected():
            rows_sel = _selected_rows()
            if not rows_sel:
                return
            val, ok = QInputDialog.getDouble(d, "Apply Wait", "Wait seconds:", value=1.0, minValue=0.0, maxValue=3600.0, decimals=3)
            if not ok:
                return
            for r in rows_sel:
                if 0 <= r < len(prepared_rows):
                    prepared_rows[r]["wait_text"] = f"{val:g}"
            _populate_instruction_table()

        def _apply_reusable_to_selected():
            rows_sel = _selected_rows()
            selected_action = reusable_pick.currentText().strip()
            if not rows_sel or not selected_action:
                return
            for r in rows_sel:
                if 0 <= r < len(prepared_rows):
                    prepared_rows[r]["automation_type"] = "Reusable Action"
                    prepared_rows[r]["target_action"] = selected_action
            _populate_instruction_table()

        def _sync_instruction_edits():
            for r in range(instructions.rowCount()):
                if not (0 <= r < len(prepared_rows)):
                    continue
                cb = instructions.cellWidget(r, 2)
                if isinstance(cb, QComboBox):
                    prepared_rows[r]["automation_type"] = cb.currentText()
                for c, key in [
                    (3, "target_action"),
                    (4, "input_value"),
                    (5, "wait_text"),
                    (6, "expected_result"),
                    (7, "screenshot_name"),
                    (8, "notes"),
                ]:
                    it = instructions.item(r, c)
                    prepared_rows[r][key] = it.text().strip() if it else ""
            _refresh_validation_summary()

        def build_preview_steps() -> tuple[list[dict[str, Any]], list[str]]:
            _sync_instruction_edits()
            steps: list[dict[str, Any]] = []
            warnings: list[str] = []
            for row in prepared_rows:
                step, warn = self._automation_instruction_to_step(row)
                warnings.extend(warn)
                if step:
                    steps.append(step)
            return steps, warnings

        def refresh_generated_preview():
            nonlocal built_steps_cache
            nonlocal build_warnings
            built_steps_cache, build_warnings = build_preview_steps()
            generated.setRowCount(len(built_steps_cache))
            for i, s in enumerate(built_steps_cache):
                vals = [
                    str(i + 1),
                    _title_for_step_type(str(s.get("type", ""))),
                    self._step_description(s),
                    self._step_target_action(s),
                    self._step_input(s),
                    self._step_wait(s),
                ]
                for c, val in enumerate(vals):
                    generated.setItem(i, c, QTableWidgetItem(val))
            preview_validation_label.setText(
                "Warnings:\n- " + "\n- ".join(build_warnings[:20]) if build_warnings else "No conversion warnings."
            )

        def refresh_nav():
            idx = pages.currentIndex()
            back_btn.setEnabled(idx > 0)
            next_btn.setVisible(idx < 2)
            import_btn.setEnabled(idx == 2 and len(prepared_rows) > 0 and len(built_steps_cache) > 0)
            if idx == 0:
                step_label.setText("Step 1 of 3: Select Manual Columns")
            elif idx == 1:
                step_label.setText("Step 2 of 3: Add Automation Instructions")
            else:
                step_label.setText("Step 3 of 3: Preview Draft Test Case")
            if idx == 0:
                next_btn.setEnabled(bool(desc_col.currentText().strip()))

        def go_next():
            idx = pages.currentIndex()
            if idx == 0:
                if not desc_col.currentText().strip():
                    QMessageBox.warning(d, "Import Manual Test Case", "Step description column is required.")
                    return
                _ingest_step1_rows()
                if not prepared_rows:
                    QMessageBox.warning(d, "Import Manual Test Case", "No manual rows found from selected description column.")
                    return
            if idx == 1:
                refresh_generated_preview()
            pages.setCurrentIndex(min(2, idx + 1))
            refresh_nav()

        def go_back():
            idx = pages.currentIndex()
            pages.setCurrentIndex(max(0, idx - 1))
            refresh_nav()

        payload: dict[str, Any] = {}

        def finish_import():
            _sync_instruction_edits()
            refresh_generated_preview()
            name = flow_name_edit.text().strip()
            if not name:
                QMessageBox.warning(d, "Import Manual Test Case", "Draft Test Case name is required.")
                return
            if not built_steps_cache:
                QMessageBox.warning(d, "Import Manual Test Case", "No steps generated from current instructions.")
                return
            payload["flow_name"] = name
            payload["built_steps"] = built_steps_cache
            payload["edit_after_import"] = bool(edit_after_import.isChecked())
            d.accept()

        def _desc_changed(_txt: str):
            refresh_nav()

        desc_col.currentTextChanged.connect(_desc_changed)
        set_manual_btn.clicked.connect(lambda: _apply_type_to_selected("Manual"))
        set_click_btn.clicked.connect(lambda: _apply_type_to_selected("Click"))
        set_type_btn.clicked.connect(lambda: _apply_type_to_selected("Type"))
        set_action_btn.clicked.connect(lambda: _apply_type_to_selected("Reusable Action"))
        clear_sel_btn.clicked.connect(_clear_selected)
        apply_wait_btn.clicked.connect(_apply_wait_to_selected)
        reusable_apply_btn.clicked.connect(_apply_reusable_to_selected)
        instructions.itemChanged.connect(lambda _it: _refresh_validation_summary())

        back_btn.clicked.connect(go_back)
        next_btn.clicked.connect(go_next)
        import_btn.clicked.connect(finish_import)
        cancel_btn.clicked.connect(d.reject)

        refresh_nav()
        if d.exec() != QDialog.Accepted:
            return None
        return payload

    def _row_to_step(
        self,
        *,
        row: dict[str, Any],
        action_col: str,
        target_col: str,
        value_col: str,
        seconds_col: str,
        desc_col: str,
        flow_ref_col: str,
        key_col: str,
        path_col: str,
        expected_col: str,
        screenshot_col: str,
        override: dict[str, str] | None = None,
        allow_manual_comment: bool = False,
    ) -> dict[str, Any] | None:
        def _txt(col: str) -> str:
            return str(row.get(col, "")).strip() if col else ""

        override = override or {}
        subflow = _txt(flow_ref_col)
        if subflow:
            step = {"type": "run_flow", "flow": subflow, "enabled": True}
            d = _txt(desc_col)
            if d:
                step["description"] = d
            return step

        action = (override.get("action", "").strip() or _txt(action_col)).lower()
        target = override.get("target", "").strip() or _txt(target_col)
        value = override.get("value", "").strip() or _txt(value_col)
        key = _txt(key_col)
        path = _txt(path_col)
        desc = _txt(desc_col)
        seconds_raw = _txt(seconds_col)
        expected = _txt(expected_col)
        screenshot_name = _txt(screenshot_col)

        if not action:
            if allow_manual_comment:
                text = desc or "Manual step"
                step = {"type": "comment", "text": text, "description": text, "enabled": True}
                if value:
                    step["value"] = value
                return step
            return None

        step: dict[str, Any] = {"type": action, "enabled": True}

        if action in {"click", "double_click", "right_click"}:
            if target:
                step["target"] = target
            elif value:
                parts = [p.strip() for p in value.replace("x:", "").replace("y:", "").split(",")]
                if len(parts) == 2:
                    try:
                        step["x"] = int(float(parts[0]))
                        step["y"] = int(float(parts[1]))
                        if action == "click":
                            step["type"] = "click_xy"
                    except Exception:
                        if allow_manual_comment:
                            return {"type": "comment", "text": desc or value, "enabled": True}
                        return None
                elif allow_manual_comment:
                    return {"type": "comment", "text": desc or value or "Manual step", "enabled": True}
                else:
                    return None
            elif allow_manual_comment:
                return {"type": "comment", "text": desc or "Manual step", "enabled": True}
            else:
                return None
        elif action == "click_xy":
            parts = [p.strip() for p in (value or target).replace("x:", "").replace("y:", "").split(",")]
            if len(parts) != 2:
                if allow_manual_comment:
                    return {"type": "comment", "text": desc or value or target or "Manual step", "enabled": True}
                return None
            try:
                step["x"] = int(float(parts[0]))
                step["y"] = int(float(parts[1]))
            except Exception:
                if allow_manual_comment:
                    return {"type": "comment", "text": desc or value or target or "Manual step", "enabled": True}
                return None
        elif action in {"type_text", "type"}:
            step["type"] = "type_text"
            step["value"] = value
        elif action in {"press_key", "press"}:
            step["type"] = "press_key"
            step["key"] = key or value or target or "enter"
        elif action == "wait":
            try:
                step["seconds"] = float(seconds_raw or value or "0")
            except Exception:
                if allow_manual_comment:
                    return {"type": "comment", "text": desc or value or "Manual step", "enabled": True}
                return None
        elif action == "screenshot":
            step["name"] = screenshot_name or value or "step"
        elif action == "run_flow":
            step["flow"] = target or value
            if not step["flow"]:
                if allow_manual_comment:
                    return {"type": "comment", "text": desc or "Manual step", "enabled": True}
                return None
        elif action == "comment":
            step["text"] = desc or value or "Manual step"
            if value:
                step["value"] = value
        elif action == "assert_window_title_contains":
            if not value:
                if allow_manual_comment:
                    return {"type": "comment", "text": desc or "Manual step", "enabled": True}
                return None
            step["value"] = value
        elif action == "assert_file_exists":
            p = path or value
            if not p:
                if allow_manual_comment:
                    return {"type": "comment", "text": desc or "Manual step", "enabled": True}
                return None
            step["path"] = p
        else:
            if allow_manual_comment:
                return {"type": "comment", "text": desc or value or target or f"Manual step ({action})", "enabled": True}
            return None

        if expected:
            step["expected_result"] = expected
        if desc and action != "comment":
            step["description"] = desc
        return step

    # Runs/exports
    def run_selected_flow(self):
        self._run_flow(dry=False)

    def dry_run_selected_flow(self):
        self._run_flow(dry=True)

    def _run_flow(self, dry: bool):
        name = self.selected_flow_name()
        if not name:
            QMessageBox.warning(self, "Run Test", "Select a test case first.")
            return
        delay = float(self.data.get("settings", {}).get("startupDelaySeconds", 3))
        if not dry:
            if QMessageBox.question(self, "Run Test", f"Run test case '{name}' after {delay:.1f}s?") != QMessageBox.Yes:
                return
        thread = threading.Thread(target=self._run_flow_thread, args=(name, dry, delay), daemon=True)
        thread.start()

    def _run_flow_thread(self, name: str, dry: bool, delay: float):
        self.set_status(f"Running {name}")
        try:
            if not dry and delay > 0:
                time.sleep(delay)
            self.active_runner = TestFlowRunner(self.data, log_callback=self.log)
            result = self.active_runner.run_flow(name, dry_run=dry)
            self.save_project()
            self.refresh_runs()
            self.log(f"Run finished: status={result.get('status')}")
        except RunnerExecutionError as exc:
            self.log(f"Run error: {exc}")
            QMessageBox.critical(self, "Run Error", str(exc))
        finally:
            self.active_runner = None
            self.set_status("Ready")

    def stop_run(self):
        if self.recorder.is_running:
            self.recorder.stop()
            return
        if self.active_runner:
            self.active_runner.request_stop()
            self.log("Stop requested.")

    def _status_chip_color(self, status: str) -> QColor:
        s = status.lower()
        if s == "passed":
            return QColor("#d1fae5")
        if s == "failed":
            return QColor("#fee2e2")
        if s == "stopped":
            return QColor("#fef3c7")
        if s == "preview":
            return QColor("#dbeafe")
        return QColor("#e5e7eb")

    def refresh_runs(self):
        runs = list(self.data.get("runs", []))[::-1]
        self.runs_table.setRowCount(len(runs))
        for r, run in enumerate(runs):
            started = str(run.get("startedAt", ""))
            test_case = str(run.get("name", ""))
            status = _normalize_status(str(run.get("status", "")), bool(run.get("dryRun", False)))
            duration = str(run.get("durationSeconds", ""))
            report = "Open"
            vals = [started, test_case, status, duration, report]
            for c, v in enumerate(vals):
                item = QTableWidgetItem(v)
                if c == 2:
                    item.setBackground(self._status_chip_color(status))
                self.runs_table.setItem(r, c, item)
            self.runs_table.item(r, 0).setData(Qt.UserRole, run)

        if self.runs_table.rowCount() == 0:
            self.run_details.setPlainText("No run selected\n\nSelect a run from Run Activity to view details.")
        else:
            self.run_details.setPlainText("No run selected\n\nSelect a run from Run Activity to view details.")

    def on_run_selected(self):
        rows = self.runs_table.selectionModel().selectedRows()
        if not rows:
            self.run_details.setPlainText("No run selected\n\nSelect a run from Run Activity to view details.")
            return
        row = rows[0].row()
        item = self.runs_table.item(row, 0)
        run = item.data(Qt.UserRole) if item else None
        if not isinstance(run, dict):
            self.run_details.setPlainText("")
            return

        run_folder = str(run.get("runFolder", "")).strip()
        run_json_path = Path(run_folder) / "run.json" if run_folder else None
        report_path = Path(run_folder) / "report.html" if run_folder else Path("")
        lines = []
        step_total = 0
        step_passed = 0
        step_failed = 0

        if run_json_path and run_json_path.exists():
            try:
                payload = json.loads(run_json_path.read_text(encoding="utf-8"))
                step_results = [sr for sr in payload.get("stepResults", []) if isinstance(sr, dict)]
                step_total = len(step_results)
                step_passed = len([sr for sr in step_results if str(sr.get("status", "")).lower() == "passed"])
                step_failed = len([sr for sr in step_results if str(sr.get("status", "")).lower() == "failed"])
                lines.extend(
                    [
                        f"Status: {_normalize_status(str(run.get('status', '')), bool(run.get('dryRun', False)))}",
                        f"Test Case: {run.get('name', '')}",
                        f"Duration: {run.get('durationSeconds', '')}s",
                        f"Steps: {step_total} total · {step_passed} passed · {step_failed} failed",
                        f"Report: {report_path}",
                        "",
                        "Step results:",
                    ]
                )
                for sr in step_results:
                    if not isinstance(sr, dict):
                        continue
                    lines.append(
                        f"- #{sr.get('stepIndex', '')} {_title_for_step_type(str(sr.get('stepType', '')))}: {sr.get('status', '')}"
                        f" | {sr.get('message', '')}"
                    )
                    ss = str(sr.get("screenshot", "")).strip()
                    if ss:
                        lines.append(f"  screenshot: {Path(run_folder) / ss}")
                errors = payload.get("errors", [])
                if errors:
                    lines.append("")
                    lines.append("Errors:")
                    for e in errors:
                        lines.append(f"- {e}")
            except Exception as exc:
                lines.extend(
                    [
                        f"Status: {_normalize_status(str(run.get('status', '')), bool(run.get('dryRun', False)))}",
                        f"Test Case: {run.get('name', '')}",
                        f"Duration: {run.get('durationSeconds', '')}s",
                        f"Report: {report_path}",
                        "",
                    ]
                )
                lines.append("")
                lines.append(f"Unable to read run details: {exc}")
        else:
            lines.extend(
                [
                    f"Status: {_normalize_status(str(run.get('status', '')), bool(run.get('dryRun', False)))}",
                    f"Test Case: {run.get('name', '')}",
                    f"Duration: {run.get('durationSeconds', '')}s",
                    "Steps: 0 total · 0 passed · 0 failed",
                    f"Report: {report_path}",
                ]
            )

        self.run_details.setPlainText("\n".join(lines))

    def open_last_report(self):
        rows = self.runs_table.selectionModel().selectedRows()
        row = rows[0].row() if rows else 0
        if self.runs_table.rowCount() == 0:
            QMessageBox.warning(self, "Report", "No runs available.")
            return
        run_item = self.runs_table.item(row, 0)
        run = run_item.data(Qt.UserRole) if run_item else None
        if not isinstance(run, dict):
            return
        folder = Path(str(run.get("runFolder", "")))
        report = folder / "report.html"
        if not report.exists():
            QMessageBox.warning(self, "Report", f"Report not found:\n{report}")
            return
        webbrowser.open(report.resolve().as_uri())

    def export_zip(self):
        out, _ = QFileDialog.getSaveFileName(
            self, "Export ZIP", f"testflow_project_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip", "ZIP (*.zip)"
        )
        if not out:
            return
        self.save_project()
        root = Path.cwd()
        with zipfile.ZipFile(Path(out), "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for item in [PROJECT_FILE, Path("runs"), Path("logs")]:
                p = root / item
                if not p.exists():
                    continue
                if p.is_file():
                    zf.write(p, arcname=str(item))
                else:
                    for f in p.rglob("*"):
                        if f.is_file():
                            zf.write(f, arcname=str(f.relative_to(root)))
        self.log(f"Exported ZIP: {out}")


def main():
    app = QApplication(sys.argv)
    win = MinimalTestFlowApp()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
