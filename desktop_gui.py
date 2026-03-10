from __future__ import annotations

import csv
import json
import re
import shlex
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

from PySide6.QtCore import QProcess, Qt, QPoint
from PySide6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QFileDialog,
    QFormLayout,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QProgressBar,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

APP_DIR = Path(__file__).resolve().parent
CONFIG_PATH = APP_DIR / "gui_config.json"
RUN_HISTORY_PATH = APP_DIR / "run_history.csv"


I18N = {
    "zh": {
        "title": "棉花果订单工作台",
        "order": "订单查询",
        "wechat": "微信核对",
        "settings": "设置",
        "input_excel": "输入",
        "output_excel": "输出",
        "wechat_excel": "微信导出",
        "orders_excel": "订单结果",
        "wechat_log": "微信汇总库",
        "reconcile_output": "核对输出",
        "browse": "浏览",
        "run_order": "运行订单查询",
        "run_reconcile": "运行微信核对",
        "auto_mode": "自动模式",
        "cookie": "Cookie",
        "start_date": "开始日期",
        "end_date": "结束日期",
        "delay": "请求间隔(秒)",
        "save_config": "保存设置",
        "open_output": "打开输出目录",
        "cancel": "取消运行",
        "history": "运行历史",
        "logs": "运行日志",
        "retry": "重试",
        "replay": "复用所选记录",
        "enable_wechat_compare": "订单查询时启用微信比对",
        "status_ready": "就绪",
        "status_running": "运行中",
        "status_success": "成功",
        "status_failed": "失败",
        "lang": "Language",
        "theme": "Theme",
        "config_title": "文件配置",
        "features": "功能",
        "system": "系统",
    },
    "en": {
        "title": "MHG Order Station",
        "order": "Order Query",
        "wechat": "WeChat Reconcile",
        "settings": "Settings",
        "input_excel": "Input",
        "output_excel": "Output",
        "wechat_excel": "WeChat Export",
        "orders_excel": "Orders Result",
        "wechat_log": "WeChat Log",
        "reconcile_output": "Reconcile Output",
        "browse": "Browse",
        "run_order": "Run Order Query",
        "run_reconcile": "Run WeChat Reconcile",
        "auto_mode": "Auto Mode",
        "cookie": "Cookie",
        "start_date": "Start Date",
        "end_date": "End Date",
        "delay": "Delay(s)",
        "save_config": "Save Settings",
        "open_output": "Open Output",
        "cancel": "Cancel",
        "history": "Run History",
        "logs": "Logs",
        "retry": "Retry",
        "replay": "Replay Selected",
        "enable_wechat_compare": "Enable WeChat Compare",
        "status_ready": "Ready",
        "status_running": "Running",
        "status_success": "Success",
        "status_failed": "Failed",
        "lang": "Language",
        "theme": "Theme",
        "config_title": "FILE CONFIG",
        "features": "FEATURES",
        "system": "SYSTEM",
    },
}

DEFAULT_CONFIG = {
    "language": "zh",
    "theme": "dark",
    "order_input_excel": str(APP_DIR / "dist" / "orders_20260305_212724.xlsx"),
    "order_output_excel": str(APP_DIR / "data" / "orders_result_gui.xlsx"),
    "wechat_excel": str(APP_DIR / "data" / "群聊_单号群.xlsx"),
    "orders_excel": str(APP_DIR / "data" / "orders_result.xlsx"),
    "wechat_log": str(APP_DIR / "data" / "wechat_shipment_log.xlsx"),
    "reconcile_output": str(APP_DIR / "data" / "reconcile_result_gui.xlsx"),
    "cookie": "",
    "start_date": "2024-01-01",
    "end_date": "2026-12-31",
    "request_delay": "1.0",
    "reconcile_auto": False,
    "order_enable_wechat_compare": False,
}


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.config = self.load_config()
        self.process: QProcess | None = None
        self.running_task_name = ""
        self.running_args: list[str] = []
        self.task_started_ts: float | None = None
        self.last_task_payload: dict | None = None
        self.history_rows: list[dict] = []

        self.setWindowFlags(Qt.FramelessWindowHint)
        self.setMinimumSize(1200, 760)
        self._drag_pos: QPoint | None = None

        root = QWidget()
        outer = QVBoxLayout(root)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        outer.addWidget(self.build_title_bar())

        body = QHBoxLayout()
        body.setContentsMargins(0, 0, 0, 0)
        body.setSpacing(0)
        body.addWidget(self.build_sidebar())
        body.addLayout(self.build_content_area(), 1)

        outer.addLayout(body, 1)
        self.setCentralWidget(root)

        self.apply_config_to_ui()
        self.apply_styles()
        self.retranslate()
        self.refresh_history_view()

    def t(self, key: str) -> str:
        lang = self.config.get("language", "zh")
        return I18N.get(lang, I18N["zh"]).get(key, key)

    def load_config(self) -> dict:
        if CONFIG_PATH.exists():
            try:
                data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
                cfg = dict(DEFAULT_CONFIG)
                cfg.update(data)
                return cfg
            except Exception:
                pass
        return dict(DEFAULT_CONFIG)

    def save_config(self):
        self.collect_ui_to_config()
        CONFIG_PATH.write_text(json.dumps(self.config, ensure_ascii=False, indent=2), encoding="utf-8")
        self.append_log("配置已保存", "ok")

    def build_title_bar(self) -> QWidget:
        w = QWidget()
        w.setObjectName("TitleBar")
        w.setFixedHeight(46)
        lay = QHBoxLayout(w)
        lay.setContentsMargins(12, 0, 12, 0)

        for color in ("#ff6057", "#ffbd2e", "#28c840"):
            dot = QLabel("●")
            dot.setStyleSheet(f"color:{color}; font-size:30px;")
            dot.setFixedWidth(35)
            lay.addWidget(dot)

        self.title_label = QLabel()
        lay.addWidget(self.title_label)
        lay.addStretch(1)

        self.lang_label = QLabel()
        self.lang_combo = QComboBox()
        self.lang_combo.addItems(["中文", "English"])
        self.lang_combo.setFixedWidth(92)
        self.lang_combo.currentIndexChanged.connect(self.change_language)

        self.theme_label = QLabel()
        self.theme_combo = QComboBox()
        self.theme_combo.addItems(["Dark", "Light"])
        self.theme_combo.setFixedWidth(84)
        self.theme_combo.currentIndexChanged.connect(self.change_theme)

        self.btn_min = QPushButton("—")
        self.btn_min.setFixedWidth(34)
        self.btn_close = QPushButton("✕")
        self.btn_close.setFixedWidth(34)
        self.btn_min.clicked.connect(self.showMinimized)
        self.btn_close.clicked.connect(self.close)

        # 可拖动标题栏（排除控件区）
        w.mousePressEvent = self._title_mouse_press
        w.mouseMoveEvent = self._title_mouse_move
        w.mouseReleaseEvent = self._title_mouse_release

        lay.addWidget(self.lang_label)
        lay.addWidget(self.lang_combo)
        lay.addSpacing(12)
        lay.addWidget(self.theme_label)
        lay.addWidget(self.theme_combo)
        lay.addSpacing(8)
        lay.addWidget(self.btn_min)
        lay.addWidget(self.btn_close)
        return w

    def build_sidebar(self) -> QWidget:
        w = QWidget()
        w.setObjectName("Sidebar")
        w.setFixedWidth(222)
        lay = QVBoxLayout(w)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(0)

        logo = QWidget()
        logo_l = QVBoxLayout(logo)
        logo_l.setContentsMargins(16, 16, 16, 16)
        self.logo_main = QLabel("MHG DESK")
        self.logo_sub = QLabel("ORDER STATION")
        logo_l.addWidget(self.logo_main)
        logo_l.addWidget(self.logo_sub)

        self.lbl_features = QLabel()
        self.lbl_system = QLabel()
        self.btn_nav_order = QPushButton()
        self.btn_nav_wechat = QPushButton()
        self.btn_nav_settings = QPushButton()

        for b in (self.btn_nav_order, self.btn_nav_wechat, self.btn_nav_settings):
            b.setCheckable(True)
            b.setFixedHeight(46)

        self.btn_nav_order.clicked.connect(lambda: self.switch_page(0))
        self.btn_nav_wechat.clicked.connect(lambda: self.switch_page(1))
        self.btn_nav_settings.clicked.connect(lambda: self.switch_page(2))

        lay.addWidget(logo)
        lay.addSpacing(12)
        lay.addWidget(self.lbl_features)
        lay.addWidget(self.btn_nav_order)
        lay.addWidget(self.btn_nav_wechat)
        lay.addSpacing(10)
        lay.addWidget(self.lbl_system)
        lay.addWidget(self.btn_nav_settings)
        lay.addStretch(1)
        return w

    def build_content_area(self):
        main = QVBoxLayout()
        main.setContentsMargins(0, 0, 0, 0)
        main.setSpacing(0)

        self.status_bar = self.build_status_bar()
        main.addWidget(self.status_bar)

        content = QWidget()
        content_l = QVBoxLayout(content)
        content_l.setContentsMargins(16, 16, 16, 16)
        content_l.setSpacing(12)

        self.pages = [self.build_order_page(), self.build_wechat_page(), self.build_settings_page()]
        for p in self.pages:
            p.setVisible(False)
            content_l.addWidget(p)
        self.pages[0].setVisible(True)

        content_l.addLayout(self.build_bottom_panels(), 1)
        main.addWidget(content, 1)
        return main

    def build_status_bar(self) -> QWidget:
        w = QWidget()
        w.setObjectName("TopStatusBar")
        w.setFixedHeight(58)
        lay = QHBoxLayout(w)
        lay.setContentsMargins(16, 0, 16, 0)
        lay.setSpacing(12)

        self.lbl_total_name = QLabel("总任务")
        self.lbl_ok_name = QLabel("成功")
        self.lbl_fail_name = QLabel("失败")
        self.lbl_total = QLabel("0")
        self.lbl_ok = QLabel("0")
        self.lbl_fail = QLabel("0")

        for v in (self.lbl_total, self.lbl_ok, self.lbl_fail):
            v.setStyleSheet("font-family:'Consolas'; font-size:20px; font-weight:700; color:#E6EDF3;")
            v.setMinimumWidth(28)
            v.setAlignment(Qt.AlignRight | Qt.AlignVCenter)

        for name_label in (self.lbl_total_name, self.lbl_ok_name, self.lbl_fail_name):
            name_label.setStyleSheet("color:#7D8590;")

        for name_label, val in ((self.lbl_total_name, self.lbl_total), (self.lbl_ok_name, self.lbl_ok), (self.lbl_fail_name, self.lbl_fail)):
            holder = QWidget()
            box = QHBoxLayout(holder)
            box.setContentsMargins(0, 0, 0, 0)
            box.setSpacing(6)
            box.addWidget(name_label)
            box.addWidget(val)
            lay.addWidget(holder)

        lay.addSpacing(8)
        self.status_sep = QLabel("|")
        lay.addWidget(self.status_sep)
        lay.addStretch(1)

        self.progress = QProgressBar()
        self.progress.setFixedWidth(190)
        self.progress.setFixedHeight(7)
        self.progress.setTextVisible(False)
        self.progress_pct = QLabel("0%")
        self.progress_pct.setFixedWidth(36)

        self.btn_retry_last = QPushButton()
        self.btn_retry_last.clicked.connect(self.retry_last_task)
        self.btn_retry_last.setEnabled(False)

        self.btn_cancel = QPushButton()
        self.btn_cancel.clicked.connect(self.cancel_running)
        self.btn_cancel.setVisible(False)

        lay.addWidget(self.progress)
        lay.addWidget(self.progress_pct)
        lay.addSpacing(6)
        lay.addWidget(self.btn_retry_last)
        lay.addWidget(self.btn_cancel)
        return w

    def build_order_page(self) -> QWidget:
        p = QWidget()
        l = QVBoxLayout(p)
        l.setSpacing(12)
        l.setContentsMargins(0, 0, 0, 0)

        p.setObjectName("PageFrame")
        p.setStyleSheet("#PageFrame { background:#1C2128; border:1px solid #21262D; border-radius:6px; }")

        cfg = QFrame()
        grid = QGridLayout(cfg)
        grid.setContentsMargins(16, 16, 16, 16)
        grid.setHorizontalSpacing(12)
        grid.setVerticalSpacing(12)
        grid.setColumnStretch(1, 1)
        self.order_section_title = QLabel()
        self.order_section_title.setObjectName("PanelHeader")

        self.order_input_label = QLabel()
        self.order_input_edit = QLineEdit()
        self.order_input_edit.setMinimumHeight(34)
        self.btn_order_input_browse = QPushButton()
        self.btn_order_input_browse.setObjectName("BrowseButton")
        self.btn_order_input_browse.setFixedWidth(98)
        self.btn_order_input_browse.setMinimumHeight(34)
        self.btn_order_input_browse.clicked.connect(lambda: self.pick_file(self.order_input_edit))

        self.order_output_label = QLabel()
        self.order_output_edit = QLineEdit()
        self.order_output_edit.setMinimumHeight(34)
        self.btn_order_output_browse = QPushButton()
        self.btn_order_output_browse.setObjectName("BrowseButton")
        self.btn_order_output_browse.setFixedWidth(98)
        self.btn_order_output_browse.setMinimumHeight(34)
        self.btn_order_output_browse.clicked.connect(lambda: self.pick_save_file(self.order_output_edit))

        self.chk_order_wechat_compare = QCheckBox()
        self.chk_order_wechat_compare.stateChanged.connect(self.on_order_wechat_compare_changed)
        self.order_wechat_file_hint = QLabel()
        self.order_wechat_file_hint.setAlignment(Qt.AlignRight | Qt.AlignVCenter)

        grid.addWidget(self.order_section_title, 0, 0, 1, 3)
        grid.addWidget(self.order_input_label, 1, 0)
        grid.addWidget(self.order_input_edit, 1, 1)
        grid.addWidget(self.btn_order_input_browse, 1, 2)
        grid.addWidget(self.order_output_label, 2, 0)
        grid.addWidget(self.order_output_edit, 2, 1)
        grid.addWidget(self.btn_order_output_browse, 2, 2)
        grid.addWidget(self.chk_order_wechat_compare, 3, 0, 1, 2)
        grid.addWidget(self.order_wechat_file_hint, 3, 2)

        row = QHBoxLayout()
        row.setSpacing(12)
        self.btn_run_order = QPushButton()
        self.btn_run_order.setObjectName("PrimaryButton")
        self.btn_run_order.setMinimumHeight(36)
        self.btn_run_order.clicked.connect(self.run_order_query)
        self.btn_open_order_output = QPushButton()
        self.btn_open_order_output.setMinimumHeight(36)
        self.btn_open_order_output.clicked.connect(lambda: self.open_parent(self.order_output_edit.text()))
        row.addWidget(self.btn_run_order, 1)
        row.addWidget(self.btn_open_order_output)

        l.addWidget(cfg)
        l.addLayout(row)
        return p

    def build_wechat_page(self) -> QWidget:
        p = QWidget()
        l = QVBoxLayout(p)
        l.setContentsMargins(0, 0, 0, 0)
        l.setSpacing(12)

        p.setObjectName("PageFrame")
        p.setStyleSheet("#PageFrame { background:#1C2128; border:1px solid #30363D; border-radius:6px; }")

        grid = QGridLayout()
        grid.setHorizontalSpacing(12)
        grid.setVerticalSpacing(12)
        grid.setContentsMargins(16, 16, 16, 16)

        self.wechat_input_label = QLabel()
        self.wechat_input_edit = QLineEdit()
        self.btn_wechat_browse = QPushButton()
        self.btn_wechat_browse.setObjectName("BrowseButton")
        self.btn_wechat_browse.clicked.connect(lambda: self.pick_file(self.wechat_input_edit))

        self.orders_input_label = QLabel()
        self.orders_input_edit = QLineEdit()
        self.btn_orders_browse = QPushButton()
        self.btn_orders_browse.setObjectName("BrowseButton")
        self.btn_orders_browse.clicked.connect(lambda: self.pick_file(self.orders_input_edit))

        self.wechat_log_label = QLabel()
        self.wechat_log_edit = QLineEdit()
        self.btn_log_browse = QPushButton()
        self.btn_log_browse.setObjectName("BrowseButton")
        self.btn_log_browse.clicked.connect(lambda: self.pick_save_file(self.wechat_log_edit))

        self.reconcile_out_label = QLabel()
        self.reconcile_out_edit = QLineEdit()
        self.btn_reconcile_out_browse = QPushButton()
        self.btn_reconcile_out_browse.setObjectName("BrowseButton")
        self.btn_reconcile_out_browse.clicked.connect(lambda: self.pick_save_file(self.reconcile_out_edit))

        self.chk_auto_mode = QCheckBox()
        self.btn_run_reconcile = QPushButton()
        self.btn_run_reconcile.clicked.connect(self.run_wechat_reconcile)
        self.btn_open_reconcile_output = QPushButton()
        self.btn_open_reconcile_output.clicked.connect(lambda: self.open_parent(self.reconcile_out_edit.text()))

        row = 0
        for lab, edit, btn in (
            (self.wechat_input_label, self.wechat_input_edit, self.btn_wechat_browse),
            (self.orders_input_label, self.orders_input_edit, self.btn_orders_browse),
            (self.wechat_log_label, self.wechat_log_edit, self.btn_log_browse),
            (self.reconcile_out_label, self.reconcile_out_edit, self.btn_reconcile_out_browse),
        ):
            grid.addWidget(lab, row, 0)
            grid.addWidget(edit, row, 1)
            grid.addWidget(btn, row, 2)
            row += 1

        grid.addWidget(self.chk_auto_mode, row, 1)
        row += 1
        grid.addWidget(self.btn_run_reconcile, row, 1)
        grid.addWidget(self.btn_open_reconcile_output, row, 2)

        l.addLayout(grid)
        return p

    def build_settings_page(self) -> QWidget:
        p = QWidget()
        form = QFormLayout(p)
        form.setContentsMargins(16, 16, 16, 16)
        form.setHorizontalSpacing(12)
        form.setVerticalSpacing(12)

        p.setObjectName("PageFrame")
        p.setStyleSheet("#PageFrame { background:#1C2128; border:1px solid #30363D; border-radius:6px; }")
        self.settings_cookie_label = QLabel()
        self.settings_cookie_edit = QLineEdit()
        self.settings_start_label = QLabel()
        self.settings_start_edit = QLineEdit()
        self.settings_end_label = QLabel()
        self.settings_end_edit = QLineEdit()
        self.settings_delay_label = QLabel()
        self.settings_delay_edit = QLineEdit()
        self.btn_save_config = QPushButton()
        self.btn_save_config.clicked.connect(self.save_config)

        form.addRow(self.settings_cookie_label, self.settings_cookie_edit)
        form.addRow(self.settings_start_label, self.settings_start_edit)
        form.addRow(self.settings_end_label, self.settings_end_edit)
        form.addRow(self.settings_delay_label, self.settings_delay_edit)
        form.addRow(self.btn_save_config)
        return p

    def build_bottom_panels(self):
        row = QHBoxLayout()
        row.setSpacing(12)
        self.log_panel = QFrame()
        ll = QVBoxLayout(self.log_panel)
        ll.setContentsMargins(16, 16, 16, 16)
        ll.setSpacing(12)
        self.log_header = QLabel()
        self.log_header.setObjectName("PanelHeader")
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        ll.addWidget(self.log_header)
        ll.addWidget(self.log_text)

        self.history_panel = QFrame()
        self.history_panel.setFixedWidth(300)
        hl = QVBoxLayout(self.history_panel)
        hl.setContentsMargins(16, 16, 16, 16)
        hl.setSpacing(12)
        self.history_header = QLabel()
        self.history_header.setObjectName("PanelHeader")
        self.history_text = QTextEdit()
        self.history_text.setReadOnly(True)
        self.btn_replay_history = QPushButton()
        self.btn_replay_history.clicked.connect(self.replay_selected_history)
        hl.addWidget(self.history_header)
        hl.addWidget(self.history_text, 1)
        hl.addWidget(self.btn_replay_history)

        row.addWidget(self.log_panel, 1)
        row.addWidget(self.history_panel)
        return row

    def switch_page(self, idx: int):
        for i, p in enumerate(self.pages):
            p.setVisible(i == idx)
        self.btn_nav_order.setChecked(idx == 0)
        self.btn_nav_wechat.setChecked(idx == 1)
        self.btn_nav_settings.setChecked(idx == 2)

    def retranslate(self):
        self.setWindowTitle(self.t("title"))
        self.title_label.setText(f" {self.t('title')}  v2.0")
        self.lang_label.setText(self.t("lang"))
        self.theme_label.setText(self.t("theme"))

        self.lbl_features.setText(self.t("features"))
        self.lbl_system.setText(self.t("system"))
        self.btn_nav_order.setText("  ◈  " + self.t("order"))
        self.btn_nav_wechat.setText("  ◯  " + self.t("wechat"))
        self.btn_nav_settings.setText("  ⚙  " + self.t("settings"))

        self.order_section_title.setText(self.t("config_title"))
        self.order_input_label.setText(self.t("input_excel"))
        self.order_output_label.setText(self.t("output_excel"))
        self.btn_order_input_browse.setText(self.t("browse"))
        self.btn_order_output_browse.setText(self.t("browse"))
        self.btn_run_order.setText("▶  " + self.t("run_order"))
        self.btn_open_order_output.setText(self.t("open_output"))
        self.chk_order_wechat_compare.setText(self.t("enable_wechat_compare"))

        self.wechat_input_label.setText(self.t("wechat_excel"))
        self.orders_input_label.setText(self.t("orders_excel"))
        self.wechat_log_label.setText(self.t("wechat_log"))
        self.reconcile_out_label.setText(self.t("reconcile_output"))
        self.btn_wechat_browse.setText(self.t("browse"))
        self.btn_orders_browse.setText(self.t("browse"))
        self.btn_log_browse.setText(self.t("browse"))
        self.btn_reconcile_out_browse.setText(self.t("browse"))
        self.chk_auto_mode.setText(self.t("auto_mode"))
        self.btn_run_reconcile.setText(self.t("run_reconcile"))
        self.btn_open_reconcile_output.setText(self.t("open_output"))

        self.settings_cookie_label.setText(self.t("cookie"))
        self.settings_start_label.setText(self.t("start_date"))
        self.settings_end_label.setText(self.t("end_date"))
        self.settings_delay_label.setText(self.t("delay"))
        self.btn_save_config.setText(self.t("save_config"))

        self.log_header.setText(self.t("logs"))
        self.history_header.setText(self.t("history"))
        self.btn_retry_last.setText(self.t("retry"))
        self.btn_replay_history.setText(self.t("replay"))
        self.btn_cancel.setText(self.t("cancel"))

        self.lbl_total_name.setText("Total" if self.config.get("language") == "en" else "总任务")
        self.lbl_ok_name.setText("OK" if self.config.get("language") == "en" else "成功")
        self.lbl_fail_name.setText("Fail" if self.config.get("language") == "en" else "失败")
        self.refresh_order_wechat_file_hint()
        self.switch_page(0 if self.btn_nav_order.isChecked() or not any([self.btn_nav_order.isChecked(), self.btn_nav_wechat.isChecked(), self.btn_nav_settings.isChecked()]) else (1 if self.btn_nav_wechat.isChecked() else 2))

    def apply_config_to_ui(self):
        self.lang_combo.setCurrentIndex(0 if self.config.get("language", "zh") == "zh" else 1)
        self.theme_combo.setCurrentIndex(0 if self.config.get("theme", "dark") == "dark" else 1)
        self.order_input_edit.setText(self.config.get("order_input_excel", ""))
        self.order_output_edit.setText(self.config.get("order_output_excel", ""))
        self.wechat_input_edit.setText(self.config.get("wechat_excel", ""))
        self.orders_input_edit.setText(self.config.get("orders_excel", ""))
        self.wechat_log_edit.setText(self.config.get("wechat_log", ""))
        self.reconcile_out_edit.setText(self.config.get("reconcile_output", ""))
        self.settings_cookie_edit.setText(self.config.get("cookie", ""))
        self.settings_start_edit.setText(self.config.get("start_date", ""))
        self.settings_end_edit.setText(self.config.get("end_date", ""))
        self.settings_delay_edit.setText(self.config.get("request_delay", "1.0"))
        self.chk_auto_mode.setChecked(bool(self.config.get("reconcile_auto", False)))
        self.chk_order_wechat_compare.setChecked(bool(self.config.get("order_enable_wechat_compare", False)))
        self.btn_nav_order.setChecked(True)

    def collect_ui_to_config(self):
        self.config["order_input_excel"] = self.order_input_edit.text().strip()
        self.config["order_output_excel"] = self.order_output_edit.text().strip()
        self.config["wechat_excel"] = self.wechat_input_edit.text().strip()
        self.config["orders_excel"] = self.orders_input_edit.text().strip()
        self.config["wechat_log"] = self.wechat_log_edit.text().strip()
        self.config["reconcile_output"] = self.reconcile_out_edit.text().strip()
        self.config["cookie"] = self.settings_cookie_edit.text().strip()
        self.config["start_date"] = self.settings_start_edit.text().strip()
        self.config["end_date"] = self.settings_end_edit.text().strip()
        self.config["request_delay"] = self.settings_delay_edit.text().strip() or "1.0"
        self.config["reconcile_auto"] = self.chk_auto_mode.isChecked()
        self.config["order_enable_wechat_compare"] = self.chk_order_wechat_compare.isChecked()
        self.config["theme"] = "dark" if self.theme_combo.currentIndex() == 0 else "light"

    def apply_styles(self):
        dark = self.config.get("theme", "dark") == "dark"

        if dark:
            style = """
            QMainWindow, QWidget { background:#0D1117; color:#E6EDF3; font-family:'Segoe UI','Microsoft YaHei'; font-size:12px; }
            #TitleBar { background:#161B22; border-bottom:1px solid #30363D; }
            #Sidebar { background:#161B22; }
            QLabel { color:#E6EDF3; }
            #PanelHeader { color:#7D8590; font-weight:700; font-size:11px; letter-spacing:0.2px; }
            QFrame { background:#1C2128; border:1px solid #21262D; border-radius:6px; }
            QLineEdit, QComboBox, QTextEdit { background:#0D1117; border:1px solid #21262D; border-radius:6px; padding:6px 8px; color:#E6EDF3; }
            QTextEdit { line-height: 1.45; }
            QLineEdit { min-height: 22px; }
            QLineEdit:focus, QComboBox:focus, QTextEdit:focus { border:1px solid #2F81F7; }
            QPushButton { background:#1C2128; border:1px solid #21262D; border-radius:6px; color:#E6EDF3; padding:6px 10px; font-weight:400; }
            QPushButton:hover { background:#22262E; }
            QPushButton:checked { border-left:3px solid #2F81F7; background: rgba(47,129,247,0.1); }
            #PrimaryButton { background:rgba(31,111,235,0.12); border:1px solid #1F6FEB; color:#4D9EF8; }
            #PrimaryButton:hover { background:rgba(31,111,235,0.18); }
            #BrowseButton { background:transparent; border:1px solid #30363D; color:#7D8590; border-radius:6px; padding:4px 12px; }
            #BrowseButton:hover { background:rgba(255,255,255,0.04); }
            #TopStatusBar { background:#161B22; border-top:1px solid #21262D; border-bottom:1px solid #21262D; }
            #TopStatusBar QPushButton { background:#1C2128; border:1px solid #21262D; border-radius:6px; padding:5px 9px; }
            #TopStatusBar QPushButton:hover { background:#22262E; }
            #TopStatusBar QLabel { color:#7D8590; }
            QProgressBar { background:#21262D; border:none; border-radius:2px; }
            QProgressBar::chunk { background:#2F81F7; border-radius:2px; }
            QScrollBar:vertical { background:#0D1117; width:5px; }
            QScrollBar::handle:vertical { background:#21262D; }
            """
        else:
            style = """
            QMainWindow, QWidget { background:#f3f5f9; color:#2f3a4d; font-family:'Segoe UI','Microsoft YaHei'; font-size:12px; }
            #TitleBar { background:#ffffff; border-bottom:1px solid #d5dce9; }
            QLabel { color:#50607c; }
            QFrame { background:#ffffff; border:1px solid #d9e0ee; border-radius:8px; }
            QLineEdit, QComboBox, QTextEdit { background:#ffffff; border:1px solid #c7d2e7; border-radius:6px; padding:6px 8px; color:#2f3a4d; }
            QLineEdit:focus, QComboBox:focus, QTextEdit:focus { border:1px solid #4b79d8; }
            QPushButton { background:#eef2fa; border:1px solid #c7d2e7; border-radius:6px; color:#2f3a4d; padding:6px 10px; }
            QPushButton:hover { background:#e4ebf9; }
            QPushButton:checked { border-left:2px solid #4b79d8; background:#dde7fb; }
            #TopStatusBar { background:#f7f9fd; border-top:1px solid #d7dfef; border-bottom:1px solid #d7dfef; }
            #TopStatusBar QPushButton { background:#f4f7fd; border:1px solid #c8d3eb; border-radius:5px; padding:5px 9px; }
            #TopStatusBar QPushButton:hover { background:#ebf1fc; }
            #TopStatusBar QLabel { color:#51607a; }
            QProgressBar { background:#dde4f1; border:none; border-radius:2px; }
            QProgressBar::chunk { background:#4b79d8; border-radius:2px; }
            QScrollBar:vertical { background:#f3f5f9; width:5px; }
            QScrollBar::handle:vertical { background:#c5cfe2; }
            """

        self.setStyleSheet(style)

    def change_language(self, idx: int):
        self.config["language"] = "zh" if idx == 0 else "en"
        self.retranslate()

    def change_theme(self, idx: int):
        self.config["theme"] = "dark" if idx == 0 else "light"
        self.apply_styles()

    def refresh_order_wechat_file_hint(self):
        fp = self.wechat_input_edit.text().strip()
        if fp and Path(fp).exists():
            self.order_wechat_file_hint.setText(Path(fp).name)
            self.order_wechat_file_hint.setToolTip(fp)
        elif fp:
            self.order_wechat_file_hint.setText("INVALID")
            self.order_wechat_file_hint.setToolTip(fp)
        else:
            self.order_wechat_file_hint.setText("NO FILE")

    def pick_file(self, target: QLineEdit):
        fp, _ = QFileDialog.getOpenFileName(self, "Select file", str(APP_DIR), "Excel Files (*.xlsx *.xlsm);;All Files (*)")
        if fp:
            target.setText(fp)
            if target is self.wechat_input_edit:
                self.refresh_order_wechat_file_hint()

    def pick_save_file(self, target: QLineEdit):
        fp, _ = QFileDialog.getSaveFileName(self, "Save file", target.text() or str(APP_DIR / "data"), "Excel Files (*.xlsx)")
        if fp:
            target.setText(fp)

    def append_log(self, text: str, level: str = "info"):
        colors = {"info": "#c8d0e8", "ok": "#3ecf6e", "warn": "#e0a830", "err": "#e05555"}
        ts = datetime.now().strftime("%H:%M:%S")
        msg = text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        self.log_text.append(f'<span style="color:#3a4260">{ts}</span> <span style="color:{colors.get(level,"#c8d0e8")}">{msg}</span>')

    def try_update_progress_from_line(self, line: str):
        m = re.search(r"\[(?:STEP\s*)?(\d+)\/(\d+)\]", line, flags=re.IGNORECASE)
        if not m:
            return
        cur, total = int(m.group(1)), int(m.group(2))
        if total <= 0:
            return
        pct = int(cur * 100 / total)
        self.progress.setRange(0, 100)
        self.progress.setValue(pct)
        self.progress_pct.setText(f"{pct}%")

    def set_running_state(self, running: bool):
        self.btn_cancel.setVisible(running)
        self.btn_cancel.setEnabled(running)
        if running:
            self.append_log("任务开始", "info")
            self.progress.setRange(0, 0)
            self.progress_pct.setText("...")
        else:
            self.progress.setRange(0, 100)
            self.progress.setValue(0)
            self.progress_pct.setText("0%")

    def write_run_history(self, task_name: str, status: str, exit_code: int, elapsed_s: float, args: list[str]):
        exists = RUN_HISTORY_PATH.exists()
        with RUN_HISTORY_PATH.open("a", encoding="utf-8", newline="") as f:
            w = csv.writer(f)
            if not exists:
                w.writerow(["timestamp", "task", "status", "exit_code", "elapsed_s", "args"])
            w.writerow([datetime.now().strftime("%Y-%m-%d %H:%M:%S"), task_name, status, exit_code, f"{elapsed_s:.2f}", " ".join(args)])
        self.refresh_history_view()

    def refresh_history_view(self):
        if not RUN_HISTORY_PATH.exists():
            self.history_rows = []
            self.history_text.setPlainText("No history yet.")
            return
        rows = []
        with RUN_HISTORY_PATH.open("r", encoding="utf-8", newline="") as f:
            for row in csv.DictReader(f):
                rows.append(row)
        self.history_rows = list(reversed(rows[-10:]))
        self.history_text.setPlainText("\n".join([f"[{i+1}] {r.get('timestamp','')} | {r.get('task','')} | {r.get('status','')}" for i, r in enumerate(self.history_rows)]))

    def start_process(self, program: str, args: list[str], task_name: str):
        if self.process and self.process.state() != QProcess.NotRunning:
            QMessageBox.information(self, "Info", "A task is already running.")
            return
        self.process = QProcess(self)
        self.process.setProgram(program)
        self.process.setArguments(args)
        self.process.setWorkingDirectory(str(APP_DIR))
        self.process.readyReadStandardOutput.connect(self.on_stdout)
        self.process.readyReadStandardError.connect(self.on_stderr)
        self.process.finished.connect(self.on_finished)

        self.running_task_name = task_name
        self.running_args = list(args)
        self.task_started_ts = time.time()
        self.last_task_payload = {"program": program, "args": list(args), "task_name": task_name}
        self.btn_retry_last.setEnabled(True)

        self.append_log(f"$ {program} {' '.join(args)}")
        self.set_running_state(True)
        self.lbl_total.setText(str(int(self.lbl_total.text()) + 1))
        self.process.start()

    def on_stdout(self):
        if not self.process:
            return
        data = bytes(self.process.readAllStandardOutput()).decode("utf-8", errors="ignore")
        for line in data.splitlines():
            level = "info"
            low = line.lower()
            if "[ok]" in low:
                level = "ok"
            elif "[warn]" in low:
                level = "warn"
            elif "[err]" in low or "failed" in low:
                level = "err"
            self.append_log(line, level)
            self.try_update_progress_from_line(line)

    def on_stderr(self):
        if not self.process:
            return
        data = bytes(self.process.readAllStandardError()).decode("utf-8", errors="ignore")
        for line in data.splitlines():
            self.append_log(line, "err")

    def on_finished(self, exit_code: int, _status):
        elapsed = max(0.0, time.time() - self.task_started_ts) if self.task_started_ts else 0.0
        self.set_running_state(False)
        if exit_code == 0:
            self.lbl_ok.setText(str(int(self.lbl_ok.text()) + 1))
            self.append_log("执行完成", "ok")
            self.write_run_history(self.running_task_name, "success", exit_code, elapsed, self.running_args)
        elif exit_code == -1:
            self.append_log("已取消", "warn")
            self.write_run_history(self.running_task_name, "cancelled", exit_code, elapsed, self.running_args)
        else:
            self.lbl_fail.setText(str(int(self.lbl_fail.text()) + 1))
            self.append_log(f"执行失败: exit_code={exit_code}", "err")
            self.write_run_history(self.running_task_name, "failed", exit_code, elapsed, self.running_args)

    def cancel_running(self):
        if self.process and self.process.state() != QProcess.NotRunning:
            self.process.kill()

    def on_order_wechat_compare_changed(self, state: int):
        enabled = bool(state)
        self.config["order_enable_wechat_compare"] = enabled

        if not enabled:
            self.append_log("[INFO] 已关闭订单查询微信比对", "info")
            return

        current = self.wechat_input_edit.text().strip()
        if current and Path(current).exists():
            self.config["wechat_excel"] = current
            self.refresh_order_wechat_file_hint()
            self.append_log(f"[INFO] 已启用微信比对文件: {Path(current).name}", "info")
            return

        self.pick_file(self.wechat_input_edit)
        selected = self.wechat_input_edit.text().strip()
        if selected and Path(selected).exists():
            self.config["wechat_excel"] = selected
            self.refresh_order_wechat_file_hint()
            self.append_log(f"[INFO] 已选择微信比对文件: {Path(selected).name}", "ok")
            return

        self.chk_order_wechat_compare.blockSignals(True)
        self.chk_order_wechat_compare.setChecked(False)
        self.chk_order_wechat_compare.blockSignals(False)
        self.config["order_enable_wechat_compare"] = False
        self.append_log("[WARN] 未选择有效微信文件，已取消启用微信比对", "warn")

    def open_parent(self, file_path: str):
        p = Path(file_path).expanduser()
        folder = p.parent if p.suffix else p
        folder.mkdir(parents=True, exist_ok=True)
        if sys.platform.startswith("win"):
            subprocess.Popen(["explorer", str(folder)])

    def run_order_query(self):
        self.collect_ui_to_config()
        self.save_config()
        script = APP_DIR / "mhg_batch_query.py"
        if not script.exists():
            QMessageBox.critical(self, "Error", f"Script not found: {script}")
            return

        args = [
            str(script), "--no-interactive", "--input", self.config.get("order_input_excel", ""),
            "--output", self.config.get("order_output_excel", ""), "--cookie", self.config.get("cookie", ""),
            "--start-date", self.config.get("start_date", ""), "--end-date", self.config.get("end_date", ""),
            "--delay", str(self.config.get("request_delay", "1.0") or "1.0"),
        ]

        if self.config.get("order_enable_wechat_compare"):
            wechat_fp = self.config.get("wechat_excel", "").strip()
            wechat_log_fp = self.config.get("wechat_log", "").strip() or str(APP_DIR / "data" / "wechat_shipment_log.xlsx")
            if not wechat_fp or not Path(wechat_fp).exists():
                QMessageBox.warning(self, "Warning", "微信比对文件无效。")
                return
            args += ["--enable-wechat-compare", "--wechat", wechat_fp, "--wechat-log", wechat_log_fp]
            self.append_log(f"本次微信比对文件: {wechat_fp}")

        self.start_process(sys.executable, ["-u", *args], "Order Query")

    def run_wechat_reconcile(self):
        self.collect_ui_to_config()
        self.save_config()
        script = APP_DIR / "wechat_order_reconcile.py"
        if not script.exists():
            QMessageBox.critical(self, "Error", f"Script not found: {script}")
            return
        args = [str(script), "--log", self.config["wechat_log"], "--output", self.config["reconcile_output"]]
        if self.config.get("reconcile_auto"):
            args.append("--auto")
        else:
            wechat_fp = self.config.get("wechat_excel", "").strip()
            orders_fp = self.config.get("orders_excel", "").strip()
            if not wechat_fp or not orders_fp:
                QMessageBox.warning(self, "Warning", "WeChat/Orders path is empty.")
                return
            args += ["--wechat", wechat_fp, "--orders", orders_fp, "--no-interactive"]
        self.start_process(sys.executable, ["-u", *args], "WeChat Reconcile")

    def retry_last_task(self):
        if not self.last_task_payload:
            return
        p = self.last_task_payload
        self.start_process(p["program"], p["args"], p.get("task_name", "task"))

    def replay_selected_history(self):
        raw = self.history_text.textCursor().selectedText().strip()
        m = re.search(r"\[(\d+)\]", raw)
        if not m:
            return
        idx = int(m.group(1)) - 1
        if idx < 0 or idx >= len(self.history_rows):
            return
        row = self.history_rows[idx]
        arg_line = row.get("args", "").strip()
        if not arg_line:
            return
        args = shlex.split(arg_line, posix=False)
        if not args:
            return
        self.start_process(args[0], args[1:], f"{row.get('task','task')} (replay)")


    def _title_mouse_press(self, event):
        if event.button() == Qt.LeftButton:
            self._drag_pos = event.globalPosition().toPoint() - self.frameGeometry().topLeft()

    def _title_mouse_move(self, event):
        if self._drag_pos is not None and event.buttons() & Qt.LeftButton:
            self.move(event.globalPosition().toPoint() - self._drag_pos)

    def _title_mouse_release(self, _event):
        self._drag_pos = None


def main():
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
