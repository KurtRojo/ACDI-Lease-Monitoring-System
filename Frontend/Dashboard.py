import csv
import os
import sys
from copy import deepcopy
from datetime import date
from functools import partial

from PyQt6.QtCore import QSize, Qt, QSignalBlocker
from PyQt6.QtGui import QAction, QColor, QIcon, QKeySequence, QPainter, QPen, QPixmap
from PyQt6.QtWidgets import (
    QApplication,
    QAbstractItemView,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFrame,
    QFormLayout,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMenu,
    QMessageBox,
    QPlainTextEdit,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QStackedWidget,
    QStyledItemDelegate,
    QTableWidget,
    QTableWidgetItem,
    QToolButton,
    QVBoxLayout,
    QWidget,
)

from Add import AddBranchDialog

current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.dirname(current_dir)
if parent_dir not in sys.path:
    sys.path.insert(0, parent_dir)

from Backend.shared_data import store


class SheetTableWidget(QTableWidget):
    def keyPressEvent(self, event):
        if event.key() in (Qt.Key.Key_Delete, Qt.Key.Key_Backspace):
            selected_items = self.selectedItems()
            if selected_items:
                parent_window = self.window()
                if hasattr(parent_window, "handle_delete_selected_cells"):
                    parent_window.handle_delete_selected_cells()
                    event.accept()
                    return
                event.accept()
                return
        super().keyPressEvent(event)


class BranchDetailsDialog(QDialog):
    def __init__(self, headers, values=None, title="Add Branch", parent=None, read_only=False):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(760, 620)
        self.inputs = []

        layout = QVBoxLayout(self)
        form = QFormLayout()
        form.setSpacing(10)
        form.setContentsMargins(8, 8, 8, 8)

        values = values or [""] * len(headers)
        for index, header in enumerate(headers):
            clean_header = header.replace("\n", " ")
            if read_only:
                widget = QLabel(values[index] if index < len(values) and values[index] else "-")
                widget.setObjectName("detailValue")
                widget.setWordWrap(True)
            elif clean_header == "REMARKS":
                widget = QPlainTextEdit()
                widget.setPlainText(values[index] if index < len(values) else "")
                widget.setFixedHeight(90)
            else:
                widget = QLineEdit(values[index] if index < len(values) else "")

            if not read_only:
                widget.setReadOnly(read_only)
                widget.setObjectName("modalField")
            label = QLabel(clean_header)
            label.setObjectName("modalLabel")
            form.addRow(label, widget)
            self.inputs.append(widget)

        layout.addLayout(form)

        buttons = QDialogButtonBox(self)
        if read_only:
            self.delete_button = buttons.addButton("Delete", QDialogButtonBox.ButtonRole.DestructiveRole)
            self.cancel_button = buttons.addButton("Cancel", QDialogButtonBox.ButtonRole.RejectRole)
            self.delete_button.clicked.connect(self.accept)
            self.cancel_button.clicked.connect(self.reject)
        else:
            self.save_button = buttons.addButton("Add Branch", QDialogButtonBox.ButtonRole.AcceptRole)
            self.cancel_button = buttons.addButton("Cancel", QDialogButtonBox.ButtonRole.RejectRole)
            self.save_button.clicked.connect(self.accept)
            self.cancel_button.clicked.connect(self.reject)

        layout.addWidget(buttons)

    def values(self):
        data = []
        for widget in self.inputs:
            if isinstance(widget, QPlainTextEdit):
                data.append(widget.toPlainText().strip())
            elif isinstance(widget, QLabel):
                data.append(widget.text().strip())
            else:
                data.append(widget.text().strip())
        return data


class NotificationsDialog(QDialog):
    def __init__(self, total_contracts, notices, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Notifications")
        self.resize(760, 520)

        layout = QVBoxLayout(self)

        summary = QLabel(
            f"Contracts: {total_contracts}    Needs Action: {len(notices)}"
        )
        summary.setObjectName("sectionLabel")
        layout.addWidget(summary)

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setFrameShape(QFrame.Shape.NoFrame)

        container = QWidget()
        container_layout = QVBoxLayout(container)
        container_layout.setContentsMargins(0, 0, 0, 0)
        container_layout.setSpacing(10)

        if notices:
            for notice in notices:
                card = QFrame()
                card.setObjectName("notificationCard")
                card_layout = QVBoxLayout(card)
                card_layout.setContentsMargins(14, 12, 14, 12)
                card_layout.setSpacing(4)

                branch = QLabel(notice["branch"])
                branch.setObjectName("sectionLabel")

                details = QLabel(
                    f"Officer: {notice['officer']}\n"
                    f"Contact: {notice['contact']}\n"
                    f"Current Status: {notice.get('status', '')}\n"
                    f"Alert Window: {notice.get('window', '')}\n"
                    f"Reminder Date: {notice.get('reminder_date', '')}\n"
                    f"Contract End: {notice['end_date']}\n"
                    f"Action: {notice.get('action', 'Renew / Escalation or Increase')}"
                )
                details.setObjectName("subText")
                details.setWordWrap(True)

                card_layout.addWidget(branch)
                card_layout.addWidget(details)
                container_layout.addWidget(card)
        else:
            empty_label = QLabel("No contracts are currently inside the 2-month action window.")
            empty_label.setObjectName("subText")
            empty_label.setWordWrap(True)
            container_layout.addWidget(empty_label)

        container_layout.addStretch()
        scroll_area.setWidget(container)
        layout.addWidget(scroll_area)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Close, self)
        buttons.rejected.connect(self.reject)
        buttons.accepted.connect(self.accept)
        layout.addWidget(buttons)


class DashboardStageCellWidget(QWidget):
    def __init__(self, stage_name, state, on_status_change, on_date_change, parent=None):
        super().__init__(parent)
        self.stage_name = stage_name
        self.on_status_change = on_status_change
        self.on_date_change = on_date_change
        self.locked = False
        self.current_style = {
            "label": "New",
            "background": "#64748b",
            "foreground": "#334155",
            "border": "#475569",
            "tooltip": "",
        }

        self.setMinimumSize(104, 28)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(6, 4, 6, 4)
        layout.setSpacing(6)

        self.status_button = QToolButton(self)
        self.status_button.setPopupMode(QToolButton.ToolButtonPopupMode.InstantPopup)
        self.status_button.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonIconOnly)
        self.status_button.setAutoRaise(True)
        self.status_button.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)
        self.status_button.setFixedSize(16, 16)

        self.date_edit = QLineEdit(self)
        self.date_edit.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        self.date_edit.setMinimumWidth(72)
        self.date_edit.setPlaceholderText("dd-MMM-yy")
        self.date_edit.editingFinished.connect(self.handle_date_edit_finished)

        menu = QMenu(self.status_button)
        for status_key, status_label in LeaseMonitoringWindow.DASHBOARD_STATUS_OPTIONS:
            action = menu.addAction(status_label)
            color = LeaseMonitoringWindow.DASHBOARD_STATUS_STYLES.get(status_key, LeaseMonitoringWindow.DASHBOARD_STATUS_STYLES["new"])[1]
            action.setIcon(LeaseMonitoringWindow.status_icon(color))
            action.triggered.connect(lambda _checked=False, status=status_key: self.on_status_change(status))
        self.status_button.setMenu(menu)

        layout.addWidget(self.status_button, 0, Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        layout.addWidget(self.date_edit, 1)
        self.set_state(state)

    def sizeHint(self):
        return QSize(112, 28)

    def set_state(self, state):
        self.state = dict(state)
        completed_on = self.state.get("completed_on", "").strip()
        is_complete = self.state.get("status") == "complete"
        self.date_edit.setVisible(is_complete)
        self.date_edit.setEnabled(is_complete and not self.locked)
        self.date_edit.setText(completed_on if is_complete else "")
        self.date_edit.setToolTip(completed_on if completed_on else self.stage_name)

    def set_status_style(self, label, background, foreground, tooltip):
        self.current_style = {
            "label": label,
            "background": background,
            "foreground": foreground,
            "border": LeaseMonitoringWindow.DASHBOARD_STATUS_BORDERS.get(
                self.state.get("status", "new"), background
            ),
            "tooltip": tooltip,
        }
        if self.locked:
            return
        self.status_button.setText("")
        self.status_button.setIcon(LeaseMonitoringWindow.status_icon(background, 10))
        self.status_button.setIconSize(QSize(10, 10))
        self.status_button.setStyleSheet(
            f"""
            QToolButton {{
                background-color: transparent;
                color: {foreground};
                border: 1px solid {self.current_style['border']};
                border-radius: 8px;
                padding: 0px;
            }}
            QToolButton::menu-indicator {{
                image: none;
                width: 0px;
            }}
            """
        )
        self.status_button.setToolTip(tooltip)

    def set_date_style(self, foreground):
        self.date_edit.setStyleSheet(
            f"""
            QLineEdit {{
                background: transparent;
                color: {foreground};
                border: none;
                padding: 0px;
                margin: 0px;
                font-size: 10px;
                font-weight: 700;
            }}
            QLineEdit:disabled {{
                color: {foreground};
            }}
            """
        )

    def set_locked(self, locked, message=""):
        self.locked = bool(locked)
        if locked:
            self.status_button.setEnabled(False)
            self.status_button.setText("")
            self.status_button.setIcon(LeaseMonitoringWindow.status_icon("#94a3b8", 10))
            self.status_button.setIconSize(QSize(10, 10))
            self.status_button.setToolTip(message or f"{self.stage_name} is locked.")
            self.status_button.setStyleSheet(
                """
                QToolButton {
                    background-color: transparent;
                    color: #64748b;
                    border: 1px solid #475569;
                    border-radius: 8px;
                    padding: 0px;
                }
                QToolButton::menu-indicator {
                    image: none;
                    width: 0px;
                }
                """
            )
            self.date_edit.setVisible(False)
            self.date_edit.setEnabled(False)
            self.date_edit.setText("")
            self.set_date_style("#64748b")
        else:
            self.status_button.setEnabled(True)
            self.set_state(self.state)
            self.set_status_style(
                self.current_style["label"],
                self.current_style["background"],
                self.current_style["foreground"],
                self.current_style["tooltip"],
            )

    def handle_date_edit_finished(self):
        if self.locked or self.state.get("status") != "complete":
            return
        self.on_date_change(self.date_edit.text().strip())


class StatusComboDelegate(QStyledItemDelegate):
    def __init__(self, options, parent=None):
        super().__init__(parent)
        self.options = options

    def createEditor(self, parent, option, index):
        combo = QComboBox(parent)
        combo.setEditable(False)
        combo.addItems(self.options)
        return combo

    def setEditorData(self, editor, index):
        value = index.model().data(index, Qt.ItemDataRole.EditRole) or ""
        combo_index = editor.findText(value)
        if combo_index < 0:
            combo_index = 0
        editor.setCurrentIndex(combo_index)

    def setModelData(self, editor, model, index):
        model.setData(index, editor.currentText(), Qt.ItemDataRole.EditRole)


class PlainCellDelegate(QStyledItemDelegate):
    def createEditor(self, parent, option, index):
        editor = super().createEditor(parent, option, index)
        if hasattr(editor, "setFrame"):
            editor.setFrame(False)
        editor.setStyleSheet("border:none; background:transparent; padding:0px; margin:0px;")
        return editor


class LeaseMonitoringWindow(QMainWindow):
    DASHBOARD_STAGE_COLUMNS = {
        4: "Legal",
        5: "VLG H",
        6: "GSD",
        7: "AD",
        8: "OD",
        9: "VP-Assigned OTD",
        10: "EVPO-EVPA",
        11: "President Globodox",
    }
    DASHBOARD_STATUS_OPTIONS = [
        ("new", "New"),
        ("pending_action", "Pending Action"),
        ("in_progress", "In Progress"),
        ("complete", "Complete"),
    ]
    DASHBOARD_STATUS_STYLES = {
        "new": ("New", "#64748b", "#475569"),
        "pending_action": ("Pending", "#dc2626", "#b91c1c"),
        "in_progress": ("In Progress", "#d97706", "#92400e"),
        "complete": ("Complete", "#16a34a", "#166534"),
    }
    DASHBOARD_STATUS_BORDERS = {
        "new": "#475569",
        "pending_action": "#991b1b",
        "in_progress": "#92400e",
        "complete": "#166534",
    }
    STATUS_OPTIONS = [
        "",
        "For Legal Review",
        "For VLG Head Review",
        "For GSD Review",
        "For AD Review",
        "For OD Review",
        "For VP-Assigned OTD Review",
        "For EVPO-EVPA Review",
        "For President Approval",
        "For EVP Approval",
        "Approved",
        "Done",
    ]
    EXPIRY_HEADERS = [
        "BRANCH",
        "DATE RECEIVED",
        "DATE SENT\nTO HO",
        "TERM",
        "HEAD",
        "HEAD CONTACT NO.",
        "REMINDER\n(2 mos before deadline)",
        "FROM",
        "TO",
        "FLOOR AREA",
        "MEMO #",
        "DATE COMPLETED\n/SENT",
        "REMARKS",
    ]

    @staticmethod
    def status_icon(color, diameter=10):
        pixmap = QPixmap(diameter, diameter)
        pixmap.fill(Qt.GlobalColor.transparent)
        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.setBrush(QColor(color))
        painter.setPen(QPen(QColor(color)))
        painter.drawEllipse(1, 1, diameter - 2, diameter - 2)
        painter.end()
        return QIcon(pixmap)

    def __init__(self):
        super().__init__()
        self.setWindowTitle("ACDI Lease Monitoring System")
        self.resize(1800, 900)

        self.current_theme = store.get_theme()
        self.expiry_table_updating = False
        self.expiry_dirty = False
        self.expiry_undo_stack = []
        self.expiry_redo_stack = []
        self.expiry_last_snapshot = []
        self.legend_labels = []
        self.notification_cards = []
        self.dashboard_status_buttons = {}
        self.main_table_updating = False

        self.stacked = QStackedWidget()
        self.create_shortcuts()
        self.main_page = self.build_main_dashboard_page()
        self.expiry_page = self.build_contract_expiry_page()

        self.stacked.addWidget(self.main_page)
        self.stacked.addWidget(self.expiry_page)

        self.setCentralWidget(self.stacked)
        self.apply_theme(self.current_theme)
        self.refresh_main_dashboard_table()

    def create_shortcuts(self):
        self.copy_action = QAction("Copy", self)
        self.copy_action.setShortcut(QKeySequence.StandardKey.Copy)
        self.copy_action.triggered.connect(self.copy_selected_cells)
        self.addAction(self.copy_action)

        self.cut_action = QAction("Cut", self)
        self.cut_action.setShortcut(QKeySequence.StandardKey.Cut)
        self.cut_action.triggered.connect(self.cut_selected_cells)
        self.addAction(self.cut_action)

        self.paste_action = QAction("Paste", self)
        self.paste_action.setShortcut(QKeySequence.StandardKey.Paste)
        self.paste_action.triggered.connect(self.paste_cells)
        self.addAction(self.paste_action)

        self.save_action = QAction("Save", self)
        self.save_action.setShortcut(QKeySequence.StandardKey.Save)
        self.save_action.triggered.connect(self.save_expiry_sheet)
        self.addAction(self.save_action)

        self.undo_action = QAction("Undo", self)
        self.undo_action.setShortcut(QKeySequence.StandardKey.Undo)
        self.undo_action.triggered.connect(self.undo_expiry_change)
        self.addAction(self.undo_action)

        self.redo_action = QAction("Redo", self)
        self.redo_action.setShortcut(QKeySequence.StandardKey.Redo)
        self.redo_action.triggered.connect(self.redo_expiry_change)
        self.addAction(self.redo_action)

    def create_top_bar(self, left_button_text, left_button_handler):
        top_bar = QHBoxLayout()
        top_bar.setSpacing(10)

        nav_btn = QPushButton(left_button_text)
        nav_btn.setMinimumHeight(40)
        nav_btn.clicked.connect(left_button_handler)

        settings_btn = QToolButton()
        settings_btn.setText("Settings")
        settings_btn.setObjectName("settingsButton")
        settings_btn.setPopupMode(QToolButton.ToolButtonPopupMode.InstantPopup)

        settings_menu = QMenu(settings_btn)

        light_action = QAction("Light Mode", self)
        light_action.triggered.connect(lambda: self.apply_theme("light"))

        dark_action = QAction("Dark Mode", self)
        dark_action.triggered.connect(lambda: self.apply_theme("dark"))

        exit_action = QAction("Exit App", self)
        exit_action.triggered.connect(self.close)

        settings_menu.addAction(light_action)
        settings_menu.addAction(dark_action)
        settings_menu.addSeparator()
        settings_menu.addAction(exit_action)

        settings_btn.setMenu(settings_menu)

        top_bar.addWidget(nav_btn)
        top_bar.addStretch()
        top_bar.addWidget(settings_btn)

        return top_bar

    def build_main_dashboard_page(self):
        page = QWidget()
        root = QVBoxLayout(page)
        root.setContentsMargins(18, 18, 18, 18)
        root.setSpacing(14)

        root.addLayout(self.create_top_bar("Open Contract Expiry", self.show_expiry_page))

        header_card = QFrame()
        header_card.setObjectName("headerCard")
        header_layout = QVBoxLayout(header_card)
        header_layout.setContentsMargins(20, 18, 20, 18)
        header_layout.setSpacing(6)

        company = QLabel("VISMIN LENDING GROUP")
        company.setObjectName("companyLabel")

        subtitle = QLabel("STATUS OF LEASE CONTRACT")
        subtitle.setObjectName("pageTitle")

        description = QLabel(
            "Track lease contract movement, routing progress, and approval status across offices."
        )
        description.setObjectName("subText")
        description.setWordWrap(True)

        header_layout.addWidget(company)
        header_layout.addWidget(subtitle)
        header_layout.addWidget(description)

        root.addWidget(header_card)

        dashboard_summary = QFrame()
        dashboard_summary.setObjectName("summaryStrip")
        summary_layout = QHBoxLayout(dashboard_summary)
        summary_layout.setContentsMargins(12, 12, 12, 12)
        summary_layout.setSpacing(12)

        self.summary_expired = self.build_summary_card("Expired", "0")
        self.summary_expiring = self.build_summary_card("Expiring in 90 Days", "0")
        self.summary_pending = self.build_summary_card("Pending at HO", "0")
        self.summary_completed = self.build_summary_card("Completed", "0")

        summary_layout.addWidget(self.summary_expired)
        summary_layout.addWidget(self.summary_expiring)
        summary_layout.addWidget(self.summary_pending)
        summary_layout.addWidget(self.summary_completed)

        root.addWidget(dashboard_summary)

        self.main_table = QTableWidget()
        main_headers = [
            "DATE RECEIVED",
            "TITLE",
            "RS No.",
            "GSS MEMO No.",
            "LEGAL",
            "VLG H",
            "GSD",
            "AD",
            "OD",
            "VP-ASSIGNED OTD",
            "EVPO-EVPA",
            "PRESIDENT(GLOBODOX)",
            "REMARKS",
        ]
        main_widths = [130, 560, 80, 105, 90, 90, 120, 80, 80, 110, 100, 130, 450]

        self.configure_table(
            table=self.main_table,
            headers=main_headers,
            widths=main_widths,
            rows=store.get_main_dashboard_rows(),
            row_height=56,
            fixed_resize=False,
            editable=True,
        )
        self.main_table.setTextElideMode(Qt.TextElideMode.ElideNone)
        self.main_table.verticalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        self.main_table.horizontalHeader().setSectionsMovable(False)
        self.main_table.horizontalHeader().setStretchLastSection(False)
        self.main_table.horizontalHeader().setMinimumSectionSize(90)
        self.main_table.setItemDelegate(PlainCellDelegate(self.main_table))
        self.main_table.itemChanged.connect(self.handle_main_table_item_changed)

        root.addWidget(self.main_table)

        footer_box = QGroupBox("LEASE CONTRACT ROUTING PROCESS (HEAD OFFICE)")
        footer_box.setObjectName("infoBox")
        footer_layout = QVBoxLayout(footer_box)

        note_title = QLabel("NOTE")
        note_title.setObjectName("sectionLabel")

        note_text = QLabel(
            "Once the branch sends the request (renewal, extension, new contract, etc.), "
            "the document is routed to the Head Office for review and approval in sequence. "
            "Each office indicates completion before forwarding the document to the next office "
            "until final approval is obtained."
        )
        note_text.setWordWrap(True)
        note_text.setObjectName("subText")

        footer_layout.addWidget(note_title)
        footer_layout.addWidget(note_text)

        root.addWidget(footer_box)
        return page

    def build_summary_card(self, title, value):
        card = QFrame()
        card.setObjectName("summaryCard")
        layout = QVBoxLayout(card)
        layout.setContentsMargins(14, 12, 14, 12)
        layout.setSpacing(4)

        title_label = QLabel(title)
        title_label.setObjectName("summaryTitle")
        value_label = QLabel(value)
        value_label.setObjectName("summaryValue")

        layout.addWidget(title_label)
        layout.addWidget(value_label)
        card.value_label = value_label
        return card

    def refresh_main_dashboard_table(self, recalc_layout=True):
        rows = store.get_main_dashboard_rows()
        blocker = QSignalBlocker(self.main_table)
        self.main_table_updating = True
        self.main_table.setUpdatesEnabled(False)
        self.main_table.setRowCount(len(rows))
        self.dashboard_status_buttons = {}

        for row_index, row_data in enumerate(rows):
            for col_index, value in enumerate(row_data):
                item = self.main_table.item(row_index, col_index)
                if item is None:
                    item = QTableWidgetItem()
                    self.main_table.setItem(row_index, col_index, item)
                if col_index in self.DASHBOARD_STAGE_COLUMNS:
                    item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                else:
                    item.setTextAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
                item.setText("" if col_index in self.DASHBOARD_STAGE_COLUMNS else str(value))
                if col_index == 12:
                    item.setFlags(
                        Qt.ItemFlag.ItemIsSelectable
                        | Qt.ItemFlag.ItemIsEnabled
                        | Qt.ItemFlag.ItemIsEditable
                    )
                else:
                    item.setFlags(Qt.ItemFlag.ItemIsSelectable | Qt.ItemFlag.ItemIsEnabled)
            self.install_dashboard_status_widgets(row_index, row_data)
        del blocker
        self.main_table_updating = False
        if recalc_layout:
            self.auto_size_main_dashboard_columns()
            self.main_table.resizeRowsToContents()
            self.ensure_main_table_minimum_row_heights()
        else:
            self.ensure_main_table_minimum_row_heights()
        self.main_table.setUpdatesEnabled(True)
        self.main_table.viewport().update()
        self.refresh_dashboard_summary()

    def refresh_dashboard_stage_row(self, row_index, row_data):
        if not (0 <= row_index < self.main_table.rowCount()):
            return

        self.main_table.setUpdatesEnabled(False)
        for column_index in self.DASHBOARD_STAGE_COLUMNS:
            cell_widget = self.main_table.cellWidget(row_index, column_index)
            if cell_widget is None:
                continue
            state = store.get_dashboard_stage_state(row_data, column_index)
            is_unlocked = self.dashboard_stage_is_unlocked(row_data, column_index)
            self.apply_dashboard_status_button_style(cell_widget, state, is_unlocked)

        self.main_table.setRowHeight(row_index, max(self.main_table.rowHeight(row_index), 56))
        self.main_table.setUpdatesEnabled(True)
        self.main_table.viewport().update()
        self.refresh_dashboard_summary()

    def install_dashboard_status_widgets(self, row_index, row_data):
        for column_index in self.DASHBOARD_STAGE_COLUMNS:
            state = store.get_dashboard_stage_state(row_data, column_index)
            is_unlocked = self.dashboard_stage_is_unlocked(row_data, column_index)
            cell_widget = DashboardStageCellWidget(
                stage_name=self.DASHBOARD_STAGE_COLUMNS[column_index],
                state=state,
                on_status_change=partial(self.handle_dashboard_status_change, row_index, column_index),
                on_date_change=partial(self.handle_dashboard_date_change, row_index, column_index),
                parent=self.main_table,
            )
            self.apply_dashboard_status_button_style(cell_widget, state, is_unlocked)
            self.main_table.setCellWidget(row_index, column_index, cell_widget)
            self.dashboard_status_buttons[(row_index, column_index)] = cell_widget

    def apply_dashboard_status_button_style(self, cell_widget, state, is_unlocked=True):
        status_key = state.get("status", "new")
        label, background, foreground = self.DASHBOARD_STATUS_STYLES.get(
            status_key, self.DASHBOARD_STATUS_STYLES["new"]
        )
        tooltip = cell_widget.stage_name
        if status_key == "complete" and state.get("completed_on"):
            tooltip = f"{cell_widget.stage_name}: {label} on {state['completed_on']}"
        else:
            tooltip = f"{cell_widget.stage_name}: {label}"
        cell_widget.set_locked(not is_unlocked, self.dashboard_locked_tooltip(cell_widget.stage_name))
        if not is_unlocked:
            return
        cell_widget.set_state(state)
        cell_widget.set_status_style(label, background, foreground, tooltip)
        cell_widget.set_date_style(foreground)

    def dashboard_stage_is_unlocked(self, row_data, column_index):
        previous_column = self.dashboard_previous_stage_column(column_index)
        if previous_column is None:
            return True
        previous_state = store.get_dashboard_stage_state(row_data, previous_column)
        return previous_state.get("status") == "complete"

    def dashboard_previous_stage_column(self, column_index):
        ordered_columns = list(self.DASHBOARD_STAGE_COLUMNS.keys())
        current_index = ordered_columns.index(column_index)
        if current_index == 0:
            return None
        return ordered_columns[current_index - 1]

    def dashboard_later_stages_started(self, row_data, column_index):
        ordered_columns = list(self.DASHBOARD_STAGE_COLUMNS.keys())
        current_index = ordered_columns.index(column_index)
        for later_column in ordered_columns[current_index + 1 :]:
            later_state = store.get_dashboard_stage_state(row_data, later_column)
            if later_state.get("status") != "new":
                return later_column, later_state
        return None, None

    def can_proceed_to_dashboard_stage(self, row_data, column_index):
        previous_column = self.dashboard_previous_stage_column(column_index)
        if previous_column is None:
            return True, None, None
        previous_state = store.get_dashboard_stage_state(row_data, previous_column)
        if previous_state.get("status") == "complete":
            return True, previous_column, previous_state
        return False, previous_column, previous_state

    def dashboard_block_message(self, current_column, previous_column, previous_state):
        current_stage = self.DASHBOARD_STAGE_COLUMNS.get(current_column, "This stage")
        stage_name = self.DASHBOARD_STAGE_COLUMNS.get(previous_column, "Previous stage")
        status_label = self.DASHBOARD_STATUS_STYLES.get(
            previous_state.get("status", "new"), self.DASHBOARD_STATUS_STYLES["new"]
        )[0].lower()
        return f"{current_stage} cannot start yet. {stage_name} is still {status_label}. Please follow up."

    def dashboard_locked_tooltip(self, stage_name):
        return f"{stage_name} is waiting for the previous approval to be completed."

    def dashboard_status_feedback(self, stage_name, state):
        label = self.DASHBOARD_STATUS_STYLES.get(state.get("status", "new"), self.DASHBOARD_STATUS_STYLES["new"])[0]
        if state.get("status") == "complete" and state.get("completed_on"):
            return f"{stage_name} marked as {label} on {state['completed_on']}."
        return f"{stage_name} marked as {label}."

    def handle_dashboard_status_change(self, row_index, column_index, status_key):
        rows = store.get_main_dashboard_rows()
        if not (0 <= row_index < len(rows)):
            return

        row_data = rows[row_index]
        current_state = store.get_dashboard_stage_state(row_data, column_index)
        allowed, previous_column, previous_state = self.can_proceed_to_dashboard_stage(row_data, column_index)
        if not allowed:
            QMessageBox.warning(
                self,
                "Routing Locked",
                self.dashboard_block_message(column_index, previous_column, previous_state),
            )
            return

        is_reverting_completed_stage = current_state.get("status") == "complete" and status_key != "complete"
        if status_key != "complete" and not is_reverting_completed_stage:
            later_column, later_state = self.dashboard_later_stages_started(row_data, column_index)
            if later_column is not None:
                later_stage = self.DASHBOARD_STAGE_COLUMNS.get(later_column, "Later stage")
                later_status = self.DASHBOARD_STATUS_STYLES.get(
                    later_state.get("status", "new"), self.DASHBOARD_STATUS_STYLES["new"]
                )[0].lower()
                QMessageBox.warning(
                    self,
                    "Routing Locked",
                    f"{later_stage} is already {later_status}. Reset the later stage first before changing this one.",
                )
                return

        completed_on = ""
        if status_key == "complete":
            completed_on = current_state.get("completed_on", "").strip() or date.today().strftime("%d-%b-%y")

        store.set_dashboard_stage_state(row_data, column_index, status_key, completed_on)
        if current_state.get("status") == "complete" and status_key != "complete":
            store.clear_dashboard_stage_states_after(row_data, column_index)
        self.refresh_dashboard_stage_row(row_index, row_data)

    def handle_dashboard_date_change(self, row_index, column_index, date_value):
        rows = store.get_main_dashboard_rows()
        if not (0 <= row_index < len(rows)):
            return
        row_data = rows[row_index]
        current_state = store.get_dashboard_stage_state(row_data, column_index)
        current_status = current_state.get("status", "new")
        if current_status != "complete":
            self.refresh_dashboard_stage_row(row_index, row_data)
            return
        store.set_dashboard_stage_state(row_data, column_index, current_status, date_value)
        self.refresh_dashboard_stage_row(row_index, row_data)

    def handle_main_table_item_changed(self, item):
        if self.main_table_updating:
            return
        if item.column() != 12:
            return
        rows = store.get_main_dashboard_rows()
        if not (0 <= item.row() < len(rows)):
            return
        store.set_dashboard_remark(rows[item.row()], item.text())

    def refresh_dashboard_summary(self):
        summary = store.dashboard_summary()
        self.summary_expired.value_label.setText(str(summary["expired"]))
        self.summary_expiring.value_label.setText(str(summary["expiring_90"]))
        self.summary_pending.value_label.setText(str(summary["pending_ho"]))
        self.summary_completed.value_label.setText(str(summary["completed"]))

    def build_contract_expiry_page(self):
        page = QWidget()
        root = QVBoxLayout(page)
        root.setContentsMargins(18, 18, 18, 18)
        root.setSpacing(14)

        root.addLayout(self.create_top_bar("Back to Dashboard", self.show_main_page))

        info_row = QHBoxLayout()
        info_row.setSpacing(14)

        title_card = QFrame()
        title_card.setObjectName("headerCard")
        title_layout = QVBoxLayout(title_card)
        title_layout.setContentsMargins(20, 14, 20, 14)
        title_layout.setSpacing(4)

        title = QLabel("CONTRACT EXPIRY")
        title.setObjectName("pageTitle")

        desc = QLabel(
            "Monitor branches with active lease terms, deadline reminders, and completion updates."
        )
        desc.setObjectName("subText")
        desc.setWordWrap(True)

        title_layout.addWidget(title)
        title_layout.addWidget(desc)

        legend_group = QGroupBox("LEGEND")
        legend_group.setObjectName("legendBox")
        legend_layout = QGridLayout(legend_group)
        legend_layout.setContentsMargins(12, 8, 12, 10)
        legend_layout.setHorizontalSpacing(10)
        legend_layout.setVerticalSpacing(6)
        self.legend_labels = []

        for row_index, (count, text) in enumerate(store.get_legend_rows()):
            count_label = QLabel(str(count))
            count_label.setObjectName("legendCount")

            text_label = QLabel(text)
            text_label.setObjectName("legendText")

            legend_layout.addWidget(count_label, row_index, 0)
            legend_layout.addWidget(text_label, row_index, 1)
            self.legend_labels.append((count_label, text_label))

        info_row.addWidget(title_card, 3)
        info_row.addWidget(legend_group, 2)
        info_row.setStretch(0, 4)
        info_row.setStretch(1, 2)
        root.addLayout(info_row)

        action_bar = QFrame()
        action_bar.setObjectName("actionBar")
        action_layout = QHBoxLayout(action_bar)
        action_layout.setContentsMargins(14, 12, 14, 12)
        action_layout.setSpacing(12)

        self.add_row_btn = QPushButton("Add Branch")
        self.add_row_btn.clicked.connect(self.add_expiry_row)

        self.remove_row_btn = QPushButton("Delete Branch")
        self.remove_row_btn.clicked.connect(self.remove_selected_expiry_row)

        self.refresh_btn = QPushButton("Refresh Reminders")
        self.refresh_btn.clicked.connect(self.refresh_expiry_views)

        self.notifications_btn = QPushButton("Notifications")
        self.notifications_btn.clicked.connect(self.show_notifications_dialog)
        self.notifications_btn.setMinimumWidth(170)

        self.save_btn = QPushButton("Save")
        self.save_btn.clicked.connect(self.save_expiry_sheet)

        self.import_btn = QPushButton("Import CSV")
        self.import_btn.clicked.connect(self.import_csv_data)

        self.daily_report_btn = QPushButton("Daily Report")
        self.daily_report_btn.clicked.connect(self.export_daily_report)

        self.undo_btn = QPushButton("?")
        self.undo_btn.setToolTip("Undo")
        self.undo_btn.clicked.connect(self.undo_expiry_change)

        self.redo_btn = QPushButton("?")
        self.redo_btn.setToolTip("Redo")
        self.redo_btn.clicked.connect(self.redo_expiry_change)

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("Search branch name")
        self.search_input.setObjectName("ribbonInput")
        self.search_input.setMinimumWidth(180)
        self.search_input.textChanged.connect(self.apply_search_filter)

        self.sort_combo = QComboBox()
        self.sort_combo.setObjectName("ribbonInput")
        self.sort_combo.setMinimumWidth(170)
        self.sort_combo.addItems(["Sort by Urgency", "Sort by Branch A-Z", "Sort by Branch Z-A"])
        self.sort_combo.currentIndexChanged.connect(self.sort_expiry_rows)

        self.undo_btn.setText("Undo")
        self.redo_btn.setText("Redo")

        self.upload_pdf_btn = QPushButton("Upload PDF")
        self.upload_pdf_btn.clicked.connect(self.upload_contract_pdf)

        self.open_pdf_btn = QPushButton("Open PDF")
        self.open_pdf_btn.clicked.connect(self.open_contract_pdf)

        file_group = self.build_ribbon_group("File", [self.save_btn, self.import_btn, self.daily_report_btn])
        edit_group = self.build_ribbon_group("Edit", [self.undo_btn, self.redo_btn])
        rows_group = self.build_ribbon_group("Rows", [self.add_row_btn, self.remove_row_btn])
        docs_group = self.build_ribbon_group("Contract Files", [self.upload_pdf_btn, self.open_pdf_btn])
        search_group = self.build_ribbon_group("Search Contract", [self.search_input, self.sort_combo])

        action_layout.addWidget(file_group)
        action_layout.addWidget(self.build_ribbon_divider())
        action_layout.addWidget(edit_group)
        action_layout.addWidget(self.build_ribbon_divider())
        action_layout.addWidget(rows_group)
        action_layout.addWidget(self.build_ribbon_divider())
        action_layout.addWidget(docs_group)
        action_layout.addWidget(self.build_ribbon_divider())
        action_layout.addWidget(search_group)
        action_layout.addStretch()
        action_layout.addWidget(self.notifications_btn)

        root.addWidget(action_bar)

        self.expiry_table = SheetTableWidget()
        expiry_widths = [180, 100, 100, 140, 170, 145, 140, 95, 95, 90, 180, 120, 330]

        self.configure_table(
            table=self.expiry_table,
            headers=self.EXPIRY_HEADERS,
            widths=expiry_widths,
            rows=[],
            row_height=34,
            fixed_resize=True,
            editable=True,
        )
        self.expiry_table.setItemDelegate(PlainCellDelegate(self.expiry_table))
        self.expiry_table.setItemDelegateForColumn(12, StatusComboDelegate(self.STATUS_OPTIONS, self.expiry_table))

        self.expiry_table.itemChanged.connect(self.handle_expiry_item_changed)
        root.addWidget(self.expiry_table)

        self.populate_expiry_table()
        self.refresh_expiry_views()
        return page

    def build_ribbon_group(self, title, buttons):
        group = QFrame()
        group.setObjectName("ribbonGroup")
        layout = QVBoxLayout(group)
        layout.setContentsMargins(10, 8, 10, 8)
        layout.setSpacing(8)

        title_label = QLabel(title)
        title_label.setObjectName("ribbonTitle")

        button_row = QHBoxLayout()
        button_row.setSpacing(8)
        for button in buttons:
            button_row.addWidget(button)

        layout.addWidget(title_label)
        layout.addLayout(button_row)
        return group

    def build_ribbon_divider(self):
        divider = QFrame()
        divider.setObjectName("ribbonDivider")
        divider.setFixedWidth(1)
        return divider

    def configure_table(self, table, headers, widths, rows, row_height=32, fixed_resize=False, editable=False):
        table.setColumnCount(len(headers))
        table.setHorizontalHeaderLabels(headers)
        table.setRowCount(len(rows))

        table.verticalHeader().setVisible(False)
        table.setWordWrap(True)
        table.setAlternatingRowColors(True)
        table.setShowGrid(True)
        table.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        table.setFrameShape(QTableWidget.Shape.StyledPanel)
        table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        table.setGridStyle(Qt.PenStyle.SolidLine)
        table.setCornerButtonEnabled(False)

        if editable:
            table.setEditTriggers(
                QAbstractItemView.EditTrigger.DoubleClicked
                | QAbstractItemView.EditTrigger.EditKeyPressed
                | QAbstractItemView.EditTrigger.AnyKeyPressed
            )
            table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectItems)
        else:
            table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
            table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)

        resize_mode = QHeaderView.ResizeMode.Fixed if fixed_resize else QHeaderView.ResizeMode.Interactive
        table.horizontalHeader().setSectionResizeMode(resize_mode)

        for index, width in enumerate(widths):
            table.setColumnWidth(index, width)

        for row_index, row_data in enumerate(rows):
            table.setRowHeight(row_index, row_height)
            for col_index, value in enumerate(row_data):
                item = QTableWidgetItem(str(value))
                item.setTextAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
                table.setItem(row_index, col_index, item)

        table.horizontalHeader().setStretchLastSection(False)

    def populate_expiry_table(self):
        rows = store.get_expiry_rows()
        self.restore_expiry_rows_to_table(rows)
        self.expiry_undo_stack.clear()
        self.expiry_redo_stack.clear()
        self.expiry_dirty = False
        self.update_window_title()
        self.auto_size_main_dashboard_columns()

    def auto_size_main_dashboard_columns(self):
        header = self.main_table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        self.main_table.resizeColumnsToContents()

        max_widths = {
            0: 170,
            1: 900,
            2: 120,
            3: 220,
            4: 118,
            5: 118,
            6: 118,
            7: 118,
            8: 118,
            9: 128,
            10: 128,
            11: 128,
            12: 680,
        }
        min_widths = {
            0: 130,
            1: 500,
            2: 80,
            3: 140,
            4: 92,
            5: 92,
            6: 92,
            7: 92,
            8: 92,
            9: 100,
            10: 100,
            11: 100,
            12: 420,
        }

        for column in range(self.main_table.columnCount()):
            current = self.main_table.columnWidth(column)
            if column in self.DASHBOARD_STAGE_COLUMNS:
                current = max(current, self.dashboard_stage_column_width(column))
            current = max(current, min_widths.get(column, 80))
            current = min(current, max_widths.get(column, 520))
            self.main_table.setColumnWidth(column, current)

        header.setSectionResizeMode(QHeaderView.ResizeMode.Fixed)

    def dashboard_stage_column_width(self, column_index):
        width = 92
        for row_index in range(self.main_table.rowCount()):
            widget = self.main_table.cellWidget(row_index, column_index)
            if widget is not None:
                width = max(width, widget.sizeHint().width() + 10)
        return width

    def ensure_main_table_minimum_row_heights(self):
        for row_index in range(self.main_table.rowCount()):
            current_height = self.main_table.rowHeight(row_index)
            self.main_table.setRowHeight(row_index, max(current_height, 56))

    def collect_expiry_rows_from_table(self):
        rows = []
        for row_index in range(self.expiry_table.rowCount()):
            row_data = []
            for col_index in range(self.expiry_table.columnCount()):
                item = self.expiry_table.item(row_index, col_index)
                row_data.append(item.text().strip() if item else "")
            row_data[6] = store.reminder_date_for(row_data)
            rows.append(row_data)
        return rows

    def update_expiry_snapshot(self):
        if hasattr(self, "expiry_table"):
            self.expiry_last_snapshot = deepcopy(self.collect_expiry_rows_from_table())

    def sync_dirty_state(self):
        self.expiry_dirty = self.collect_expiry_rows_from_table() != store.get_expiry_rows()
        self.update_window_title()

    def record_expiry_change(self, previous_rows):
        current_rows = self.collect_expiry_rows_from_table()
        if current_rows == previous_rows:
            return
        self.expiry_undo_stack.append(deepcopy(previous_rows))
        if len(self.expiry_undo_stack) > 100:
            self.expiry_undo_stack.pop(0)
        self.expiry_redo_stack.clear()
        self.expiry_last_snapshot = deepcopy(current_rows)
        self.sync_dirty_state()
        self.refresh_expiry_views()

    def restore_expiry_rows_to_table(self, rows):
        blocker = QSignalBlocker(self.expiry_table)
        self.expiry_table_updating = True
        self.expiry_table.setRowCount(len(rows))

        for row_index, row_data in enumerate(rows):
            self.expiry_table.setRowHeight(row_index, 36)
            normalized = list(row_data[:13])
            if len(normalized) < 13:
                normalized.extend([""] * (13 - len(normalized)))
            normalized[6] = store.reminder_date_for(normalized)

            for col_index, value in enumerate(normalized):
                item = self.expiry_table.item(row_index, col_index)
                if item is None:
                    item = QTableWidgetItem()
                    item.setTextAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
                    self.expiry_table.setItem(row_index, col_index, item)
                item.setText(str(value))

        del blocker
        self.expiry_table_updating = False
        self.expiry_last_snapshot = deepcopy(self.collect_expiry_rows_from_table())
        self.apply_expiry_row_styles(self.expiry_last_snapshot)
        self.apply_search_filter()

    def undo_expiry_change(self):
        if not self.expiry_undo_stack:
            return
        current_rows = self.collect_expiry_rows_from_table()
        previous_rows = self.expiry_undo_stack.pop()
        self.expiry_redo_stack.append(deepcopy(current_rows))
        self.restore_expiry_rows_to_table(previous_rows)
        self.sync_dirty_state()

    def redo_expiry_change(self):
        if not self.expiry_redo_stack:
            return
        current_rows = self.collect_expiry_rows_from_table()
        next_rows = self.expiry_redo_stack.pop()
        self.expiry_undo_stack.append(deepcopy(current_rows))
        self.restore_expiry_rows_to_table(next_rows)
        self.sync_dirty_state()

    def handle_expiry_item_changed(self, _item):
        if self.expiry_table_updating:
            return
        previous_rows = deepcopy(self.expiry_last_snapshot)
        self.record_expiry_change(previous_rows)

    def mark_expiry_dirty(self):
        self.expiry_dirty = True
        self.update_window_title()

    def update_window_title(self):
        suffix = " *" if self.expiry_dirty else ""
        backend_suffix = " [Local SQL]" if getattr(store, "backend", "mysql") != "mysql" else ""
        self.setWindowTitle(f"ACDI Lease Monitoring System{backend_suffix}{suffix}")

    def save_expiry_sheet(self):
        if not hasattr(self, "expiry_table"):
            return
        rows = self.collect_expiry_rows_from_table()
        store.set_expiry_rows(rows)
        self.expiry_dirty = False
        self.expiry_undo_stack.clear()
        self.expiry_redo_stack.clear()
        self.refresh_main_dashboard_table()
        self.populate_expiry_table()
        self.refresh_expiry_views()
        self.update_window_title()

    def revert_expiry_sheet(self):
        if not hasattr(self, "expiry_table"):
            return
        if self.expiry_dirty and not self.confirm_revert_action():
            return
        store.load()
        self.refresh_main_dashboard_table()
        self.populate_expiry_table()
        self.refresh_expiry_views()
        self.update_window_title()

    def refresh_expiry_views(self):
        current_rows = self.collect_expiry_rows_from_table() if hasattr(self, "expiry_table") else store.get_expiry_rows()
        self.update_legend(current_rows)
        self.update_notifications(current_rows)
        self.apply_expiry_row_styles(current_rows)

    def summarize_legend_rows(self, rows):
        counts = {
            "Done Lease Contracts": 0,
            "For Reminder / Action": 0,
            "Active Contracts": 0,
            "Expired Contracts": 0,
        }

        for row in rows:
            status = store.contract_status_for(row)
            if status == "done":
                counts["Done Lease Contracts"] += 1
            elif status == "due":
                counts["For Reminder / Action"] += 1
            elif status == "expired":
                counts["Expired Contracts"] += 1
            elif status == "active":
                counts["Active Contracts"] += 1

        return [
            (str(counts["Done Lease Contracts"]), "Done Lease Contracts"),
            (str(counts["For Reminder / Action"]), "For Reminder / Action"),
            (str(counts["Active Contracts"]), "Active Contracts"),
            (str(counts["Expired Contracts"]), "Expired Contracts"),
        ]

    def update_legend(self, rows):
        legend_rows = self.summarize_legend_rows(rows)
        for index, (count, text) in enumerate(legend_rows):
            count_label, text_label = self.legend_labels[index]
            count_label.setText(str(count))
            text_label.setText(text)

    def build_notification_rows(self, rows):
        notices = []
        for row in rows:
            branch = row[0].strip()
            end_date = store.parse_date(row[8])
            reminder_windows = store.reminder_windows_for(row)
            reminder_date = store.parse_date(store.reminder_date_for(row))
            if not branch or not end_date or not reminder_windows:
                continue
            notices.append(
                {
                    "branch": branch,
                    "officer": row[4].strip() or "Officer not assigned",
                    "contact": row[5].strip() or "No contact number",
                    "status": store.manual_status_for(row) or store.pending_stage_for(row),
                    "window": ", ".join(reminder_windows),
                    "reminder_date": store.format_date(reminder_date),
                    "end_date": store.format_date(end_date),
                    "action": "Renew / Escalation or Increase",
                }
            )
        notices.sort(key=lambda item: store.parse_date(item["end_date"]))
        return notices

    def update_notifications(self, rows):
        notice_count = len(self.build_notification_rows(rows))
        self.notifications_btn.setText(f"Notifications ({notice_count})")

    def show_notifications_dialog(self):
        rows = self.collect_expiry_rows_from_table() if hasattr(self, "expiry_table") else store.get_expiry_rows()
        notices = self.build_notification_rows(rows)
        total_contracts = len([row for row in rows if row[0].strip()])
        dialog = NotificationsDialog(total_contracts, notices, self)
        dialog.exec()

    def selected_expiry_row_values(self):
        selected_rows = sorted({index.row() for index in self.expiry_table.selectedIndexes()})
        if not selected_rows:
            return None, None
        row_index = selected_rows[0]
        values = []
        for col_index in range(self.expiry_table.columnCount()):
            item = self.expiry_table.item(row_index, col_index)
            values.append(item.text() if item else "")
        return row_index, values

    def apply_expiry_row_styles(self, rows):
        for row_index, row_data in enumerate(rows):
            status = store.contract_status_for(row_data)

            if status == "done":
                background = QColor("#ecfdf3") if self.current_theme == "light" else QColor("#163528")
            elif status == "due":
                background = QColor("#fff7d6") if self.current_theme == "light" else QColor("#463315")
            elif status == "expired":
                background = QColor("#ffebeb") if self.current_theme == "light" else QColor("#402020")
            else:
                background = QColor("#ffffff") if self.current_theme == "light" else QColor("#0f172a")

            for col_index in range(self.expiry_table.columnCount()):
                item = self.expiry_table.item(row_index, col_index)
                if item:
                    item.setBackground(background)

    def add_expiry_row(self):
        default_row = ["", "", "", "1 YR", "", "", "", "", "", "", "", "", "FOR RENEW / ESCALATION"]
        dialog = AddBranchDialog(
            headers=self.EXPIRY_HEADERS,
            values=default_row,
            parent=self,
            title="Add Branch",
            theme=self.current_theme,
        )
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return

        row_values = dialog.values()
        previous_rows = self.collect_expiry_rows_from_table()
        self.expiry_table_updating = True
        self.expiry_table.insertRow(self.expiry_table.rowCount())
        new_row = self.expiry_table.rowCount() - 1
        self.expiry_table.setRowHeight(new_row, 36)
        for col_index, value in enumerate(row_values):
            item = QTableWidgetItem(str(value))
            item.setTextAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
            self.expiry_table.setItem(new_row, col_index, item)
        self.expiry_table_updating = False
        self.record_expiry_change(previous_rows)
        self.expiry_table.setCurrentCell(new_row, 0)
        self.expiry_table.scrollToBottom()

    def remove_selected_expiry_row(self):
        selected_rows = sorted({index.row() for index in self.expiry_table.selectedIndexes()}, reverse=True)
        if not selected_rows:
            QMessageBox.information(self, "Delete Branch", "Select at least one branch row to delete.")
            return

        if len(selected_rows) > 1:
            QMessageBox.information(self, "Delete Branch", "Please select one branch row at a time.")
            return

        row_index = selected_rows[0]
        row_values = []
        for col_index in range(self.expiry_table.columnCount()):
            item = self.expiry_table.item(row_index, col_index)
            row_values.append(item.text() if item else "")

        dialog = BranchDetailsDialog(
            self.EXPIRY_HEADERS,
            row_values,
            "Delete Branch",
            self,
            read_only=True,
        )
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return

        previous_rows = self.collect_expiry_rows_from_table()
        self.expiry_table_updating = True
        self.expiry_table.removeRow(row_index)
        self.expiry_table_updating = False
        self.record_expiry_change(previous_rows)

    def confirm_delete_action(self):
        result = QMessageBox.question(
            self,
            "Confirm Delete",
            "Do you confirm in deleting info?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        return result == QMessageBox.StandardButton.Yes

    def confirm_revert_action(self):
        result = QMessageBox.question(
            self,
            "Revert Changes",
            "Revert unsaved changes in the sheet?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        return result == QMessageBox.StandardButton.Yes

    def prompt_unsaved_changes(self):
        if not self.expiry_dirty:
            return "save"

        dialog = QMessageBox(self)
        dialog.setWindowTitle("Unsaved Changes")
        dialog.setText("This sheet has unsaved changes.")
        dialog.setInformativeText("Choose Save, Don't Save, or Cancel.")
        save_button = dialog.addButton("Save", QMessageBox.ButtonRole.AcceptRole)
        discard_button = dialog.addButton("Don't Save", QMessageBox.ButtonRole.DestructiveRole)
        cancel_button = dialog.addButton("Cancel", QMessageBox.ButtonRole.RejectRole)
        dialog.setDefaultButton(save_button)
        dialog.exec()

        clicked = dialog.clickedButton()
        if clicked == save_button:
            self.save_expiry_sheet()
            return "save"
        if clicked == discard_button:
            self.revert_expiry_sheet_silent()
            return "discard"
        return "cancel"

    def revert_expiry_sheet_silent(self):
        store.load()
        self.refresh_main_dashboard_table()
        self.populate_expiry_table()
        self.refresh_expiry_views()
        self.update_window_title()

    def copy_selected_cells(self):
        if not hasattr(self, "expiry_table") or self.stacked.currentWidget() != self.expiry_page:
            return
        selected_indexes = self.expiry_table.selectedIndexes()
        if not selected_indexes:
            return

        rows = sorted({index.row() for index in selected_indexes})
        cols = sorted({index.column() for index in selected_indexes})
        lines = []
        for row in rows:
            values = []
            for col in cols:
                item = self.expiry_table.item(row, col)
                values.append(item.text() if item else "")
            lines.append("\t".join(values))
        QApplication.clipboard().setText("\n".join(lines))

    def handle_delete_selected_cells(self):
        if not hasattr(self, "expiry_table") or self.stacked.currentWidget() != self.expiry_page:
            return
        selected_items = self.expiry_table.selectedItems()
        if not selected_items:
            return
        if not self.confirm_delete_action():
            return
        previous_rows = self.collect_expiry_rows_from_table()
        self.expiry_table_updating = True
        for item in selected_items:
            item.setText("")
        self.expiry_table_updating = False
        self.record_expiry_change(previous_rows)

    def cut_selected_cells(self):
        if not hasattr(self, "expiry_table") or self.stacked.currentWidget() != self.expiry_page:
            return
        selected_items = self.expiry_table.selectedItems()
        if not selected_items:
            return
        self.copy_selected_cells()
        if not self.confirm_delete_action():
            return
        previous_rows = self.collect_expiry_rows_from_table()
        self.expiry_table_updating = True
        for item in selected_items:
            item.setText("")
        self.expiry_table_updating = False
        self.record_expiry_change(previous_rows)

    def paste_cells(self):
        if not hasattr(self, "expiry_table") or self.stacked.currentWidget() != self.expiry_page:
            return
        start_row = self.expiry_table.currentRow()
        start_col = self.expiry_table.currentColumn()
        if start_row < 0 or start_col < 0:
            return
        clipboard_text = QApplication.clipboard().text()
        if not clipboard_text:
            return

        previous_rows = self.collect_expiry_rows_from_table()
        self.expiry_table_updating = True
        for row_offset, line in enumerate(clipboard_text.splitlines()):
            values = line.split("\t")
            target_row = start_row + row_offset
            while target_row >= self.expiry_table.rowCount():
                self.expiry_table.insertRow(self.expiry_table.rowCount())
                self.expiry_table.setRowHeight(self.expiry_table.rowCount() - 1, 36)
            for col_offset, value in enumerate(values):
                target_col = start_col + col_offset
                if target_col >= self.expiry_table.columnCount():
                    continue
                item = self.expiry_table.item(target_row, target_col)
                if item is None:
                    item = QTableWidgetItem()
                    item.setTextAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
                    self.expiry_table.setItem(target_row, target_col, item)
                item.setText(value)
        self.expiry_table_updating = False
        self.record_expiry_change(previous_rows)

    def import_csv_data(self):
        if self.stacked.currentWidget() != self.expiry_page:
            self.show_expiry_page()

        result = QMessageBox.warning(
            self,
            "Import CSV",
            "Importing a CSV file will override the current contract data in the system.",
            QMessageBox.StandardButton.Ok | QMessageBox.StandardButton.Cancel,
            QMessageBox.StandardButton.Cancel,
        )
        if result != QMessageBox.StandardButton.Ok:
            return

        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Import Contract CSV",
            "",
            "CSV Files (*.csv);;All Files (*.*)",
        )
        if not file_path:
            return

        try:
            with open(file_path, "r", encoding="utf-8-sig", newline="") as csv_file:
                reader = list(csv.reader(csv_file))
        except OSError as exc:
            QMessageBox.critical(self, "Import CSV", f"Could not open file:\n{exc}")
            return

        if not reader:
            QMessageBox.information(self, "Import CSV", "The selected CSV file is empty.")
            return

        normalized_headers = [header.replace("\n", " ").strip().lower() for header in self.EXPIRY_HEADERS]
        first_row = [cell.strip().lower() for cell in reader[0]]
        data_rows = reader[1:] if first_row[: len(normalized_headers)] == normalized_headers else reader

        imported_rows = []
        for row in data_rows:
            if not any(str(cell).strip() for cell in row):
                continue
            current = [str(value).strip() for value in row[:13]]
            if len(current) < 13:
                current.extend([""] * (13 - len(current)))
            current[6] = store.reminder_date_for(current)
            imported_rows.append(current)

        store.set_expiry_rows(imported_rows)
        self.refresh_main_dashboard_table()
        self.populate_expiry_table()
        self.refresh_expiry_views()
        self.search_input.clear()
        self.sort_combo.setCurrentIndex(0)
        QMessageBox.information(self, "Import CSV", "CSV file imported successfully. System data has been updated.")

    def export_daily_report(self):
        report_rows = store.daily_report_rows()
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Export Daily Report",
            "daily_report.txt",
            "Text Files (*.txt);;CSV Files (*.csv)",
        )
        if not file_path:
            return

        try:
            if file_path.lower().endswith(".csv"):
                with open(file_path, "w", encoding="utf-8", newline="") as report_file:
                    writer = csv.writer(report_file)
                    writer.writerow(["Branch", "Pending Stage", "Status", "Officer", "Contact", "Expiry Date", "Remarks"])
                    for row in report_rows:
                        writer.writerow(
                            [
                                row["branch"],
                                row["pending_stage"],
                                row["status"],
                                row["officer"],
                                row["contact"],
                                row["expiry_date"],
                                row["remarks"],
                            ]
                        )
            else:
                with open(file_path, "w", encoding="utf-8") as report_file:
                    report_file.write("ACDI Lease Monitoring Daily Report - 5:00 PM\n\n")
                    for row in report_rows:
                        report_file.write(
                            f"Branch: {row['branch']}\n"
                            f"Pending Stage: {row['pending_stage']}\n"
                            f"Status: {row['status']}\n"
                            f"Officer: {row['officer']}\n"
                            f"Contact: {row['contact']}\n"
                            f"Expiry Date: {row['expiry_date']}\n"
                            f"Remarks: {row['remarks']}\n\n"
                        )
        except OSError as exc:
            QMessageBox.critical(self, "Daily Report", f"Could not export report:\n{exc}")
            return

        QMessageBox.information(self, "Daily Report", "Daily report exported successfully.")

    def upload_contract_pdf(self):
        selected = self.selected_expiry_row_values()
        if selected[1] is None:
            QMessageBox.information(self, "Upload PDF", "Select a contract row first.")
            return
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Upload Final Notarized Contract",
            "",
            "PDF Files (*.pdf)",
        )
        if not file_path:
            return
        store.set_contract_document(selected[1], file_path)
        QMessageBox.information(self, "Upload PDF", "Contract PDF linked successfully.")

    def open_contract_pdf(self):
        selected = self.selected_expiry_row_values()
        if selected[1] is None:
            QMessageBox.information(self, "Open PDF", "Select a contract row first.")
            return
        file_path = store.get_contract_document(selected[1])
        if not file_path:
            QMessageBox.information(self, "Open PDF", "No PDF is linked to the selected contract yet.")
            return
        if not os.path.exists(file_path):
            QMessageBox.warning(self, "Open PDF", "The linked PDF file could not be found.")
            return
        os.startfile(file_path)

    def urgency_rank_for_row(self, row):
        status = store.contract_status_for(row)
        order = {"due": 0, "expired": 1, "active": 2, "done": 3, "blank": 4}
        return order.get(status, 5)

    def sort_expiry_rows(self, *_):
        if not hasattr(self, "expiry_table"):
            return
        current_rows = self.collect_expiry_rows_from_table()
        previous_rows = deepcopy(current_rows)

        if self.sort_combo.currentIndex() == 0:
            sorted_rows = sorted(
                current_rows,
                key=lambda row: (self.urgency_rank_for_row(row), row[0].strip().lower()),
            )
        elif self.sort_combo.currentIndex() == 1:
            sorted_rows = sorted(current_rows, key=lambda row: row[0].strip().lower())
        else:
            sorted_rows = sorted(current_rows, key=lambda row: row[0].strip().lower(), reverse=True)

        self.restore_expiry_rows_to_table(sorted_rows)
        self.record_expiry_change(previous_rows)

    def apply_search_filter(self, *_):
        if not hasattr(self, "expiry_table"):
            return
        query = self.search_input.text().strip().lower() if hasattr(self, "search_input") else ""
        for row_index in range(self.expiry_table.rowCount()):
            searchable_values = []
            for column in (0, 3, 8):
                item = self.expiry_table.item(row_index, column)
                searchable_values.append(item.text().strip().lower() if item else "")
            haystack = " ".join(searchable_values)
            self.expiry_table.setRowHidden(row_index, bool(query) and query not in haystack)

    def apply_theme(self, theme):
        self.current_theme = theme
        store.set_theme(theme)
        if theme == "dark":
            self.setStyleSheet(self.dark_stylesheet())
        else:
            self.setStyleSheet(self.light_stylesheet())
        if hasattr(self, "expiry_table"):
            self.apply_expiry_row_styles(self.collect_expiry_rows_from_table())

    def dark_stylesheet(self):
        return """
        QMainWindow, QWidget {
            background-color: #081120;
            color: #e2e8f0;
            font-family: Segoe UI, Arial, sans-serif;
            font-size: 13px;
        }

        QFrame#headerCard {
            background: qlineargradient(
                x1:0, y1:0, x2:1, y2:1,
                stop:0 #0f4c81,
                stop:0.55 #146c94,
                stop:1 #19a7a0
            );
            border: none;
            border-radius: 18px;
            min-height: 96px;
            max-height: 132px;
        }

        QFrame#actionBar,
        QFrame#notificationCard,
        QFrame#summaryStrip,
        QFrame#summaryCard {
            background-color: #0f1b2d;
            border: 1px solid #22324a;
            border-radius: 18px;
        }

        QFrame#ribbonGroup {
            background-color: transparent;
            border: none;
        }

        QFrame#ribbonDivider {
            background-color: #22324a;
            border: none;
            min-height: 58px;
            margin-top: 6px;
            margin-bottom: 6px;
        }

        QLabel#companyLabel {
            font-size: 14px;
            font-weight: 700;
            color: #dbeafe;
            letter-spacing: 1px;
            background: transparent;
        }

        QLabel#pageTitle {
            font-size: 18px;
            font-weight: 900;
            color: white;
            background: transparent;
        }

        QLabel#subText {
            font-size: 12px;
            color: #d7e4f3;
            background: transparent;
        }

        QLabel#sectionLabel {
            font-size: 14px;
            font-weight: 700;
            color: #93c5fd;
            background: transparent;
        }

        QLabel#ribbonTitle {
            font-size: 11px;
            font-weight: 800;
            color: #8db4e8;
            letter-spacing: 0.6px;
            text-transform: uppercase;
            background: transparent;
            padding-left: 2px;
        }

        QLabel#modalLabel {
            color: #c2ddff;
            font-weight: 700;
            background: transparent;
        }

        QLabel#detailValue {
            color: #f8fafc;
            background-color: #12243a;
            border: 1px solid #29405e;
            border-radius: 10px;
            padding: 8px 10px;
        }

        QLabel#summaryTitle {
            color: #8db4e8;
            font-size: 12px;
            font-weight: 700;
            background: transparent;
        }

        QLabel#summaryValue {
            color: #f8fafc;
            font-size: 24px;
            font-weight: 900;
            background: transparent;
        }

        QDialog {
            background-color: #0f1b2d;
            color: #e2e8f0;
        }

        QLineEdit#modalField,
        QPlainTextEdit#modalField {
            background-color: #0b1626;
            color: #f8fafc;
            border: 1px solid #29405e;
            border-radius: 10px;
            padding: 8px 10px;
        }

        QLineEdit#modalField:read-only,
        QPlainTextEdit#modalField:read-only {
            background-color: #12243a;
        }

        QLineEdit#ribbonInput,
        QComboBox#ribbonInput {
            background-color: #0b1626;
            color: #f8fafc;
            border: 1px solid #29405e;
            border-radius: 10px;
            padding: 8px 10px;
            min-height: 20px;
        }

        QComboBox#ribbonInput::drop-down {
            border: none;
            width: 20px;
        }

        QPushButton {
            background-color: #1e63d7;
            color: white;
            border: none;
            border-radius: 12px;
            padding: 11px 18px;
            font-weight: 700;
            min-height: 20px;
        }

        QPushButton:hover {
            background-color: #3277ea;
        }

        QPushButton:pressed {
            background-color: #184fb0;
        }

        QToolButton#settingsButton {
            background-color: #13243b;
            color: #f8fafc;
            border: 1px solid #29405e;
            border-radius: 18px;
            font-size: 13px;
            font-weight: 700;
            min-width: 84px;
            min-height: 38px;
            padding: 4px 12px;
        }

        QToolButton#settingsButton:hover {
            background-color: #1a3150;
        }

        QMenu {
            background-color: #0f1b2d;
            color: #f8fafc;
            border: 1px solid #22324a;
            border-radius: 10px;
            padding: 8px;
        }

        QMenu::item {
            padding: 8px 22px 8px 12px;
            border-radius: 6px;
        }

        QMenu::item:selected {
            background-color: #1e63d7;
        }

        QGroupBox {
            font-weight: 700;
            border: 1px solid #22324a;
            border-radius: 16px;
            margin-top: 10px;
            padding-top: 10px;
            background-color: #0f1b2d;
        }

        QGroupBox::title {
            subcontrol-origin: margin;
            left: 14px;
            padding: 0 6px 0 6px;
            color: #93c5fd;
        }

        QLabel#legendCount {
            background-color: #1e63d7;
            color: white;
            border-radius: 8px;
            padding: 4px 8px;
            font-weight: 800;
            min-width: 24px;
        }

        QLabel#legendText {
            color: #e5e7eb;
            padding-left: 2px;
            background: transparent;
        }

        QTableWidget {
            background-color: #0b1626;
            alternate-background-color: #102036;
            border: 1px solid #22324a;
            border-radius: 16px;
            color: #f8fafc;
            selection-background-color: #2a6adf;
            selection-color: white;
            padding: 10px;
            outline: 0;
            gridline-color: #243750;
        }

        QHeaderView::section {
            background-color: #13243b;
            color: #c2ddff;
            padding: 10px 12px;
            border: 1px solid #22324a;
            font-weight: 800;
            text-align: center;
        }

        QTableWidget::item {
            padding: 12px 10px;
            border-right: 1px solid #243750;
            border-bottom: 1px solid #243750;
        }

        QTableWidget QLineEdit {
            background-color: #0b1626;
            color: #f8fafc;
            border: none;
            border-radius: 0px;
            padding: 0px;
            margin: 0px;
            selection-background-color: #3277ea;
        }

        QTableWidget QLineEdit:focus {
            background-color: #0b1626;
            border: none;
        }

        QTableWidget QComboBox {
            background-color: #0b1626;
            color: #f8fafc;
            border: none;
            border-radius: 0px;
            padding: 0px;
        }

        QTableCornerButton::section {
            background-color: #13243b;
            border: 1px solid #22324a;
        }

        QScrollBar:vertical {
            background: transparent;
            width: 15px;
            margin: 4px;
            border: none;
        }

        QScrollBar::handle:vertical {
            background: #48627f;
            border-radius: 7px;
            min-height: 34px;
        }

        QScrollBar::add-line:vertical,
        QScrollBar::sub-line:vertical,
        QScrollBar::add-line:horizontal,
        QScrollBar::sub-line:horizontal {
            border: none;
            background: none;
            height: 0px;
            width: 0px;
        }

        QScrollBar:horizontal {
            background: transparent;
            height: 15px;
            margin: 4px;
            border: none;
        }

        QScrollBar::handle:horizontal {
            background: #48627f;
            border-radius: 7px;
            min-width: 34px;
        }
        """

    def light_stylesheet(self):
        return """
        QMainWindow, QWidget {
            background-color: #eef4fb;
            color: #0f172a;
            font-family: Segoe UI, Arial, sans-serif;
            font-size: 13px;
        }

        QFrame#headerCard {
            background: qlineargradient(
                x1:0, y1:0, x2:1, y2:1,
                stop:0 #d9ecff,
                stop:0.45 #a6d7ff,
                stop:1 #91f0dc
            );
            border: 1px solid #b9d7f2;
            border-radius: 18px;
            min-height: 96px;
            max-height: 132px;
        }

        QFrame#actionBar,
        QFrame#notificationCard,
        QFrame#summaryStrip,
        QFrame#summaryCard {
            background-color: #ffffff;
            border: 1px solid #d4dfec;
            border-radius: 18px;
        }

        QFrame#ribbonGroup {
            background-color: transparent;
            border: none;
        }

        QFrame#ribbonDivider {
            background-color: #d4dfec;
            border: none;
            min-height: 58px;
            margin-top: 6px;
            margin-bottom: 6px;
        }

        QLabel#companyLabel {
            font-size: 14px;
            font-weight: 700;
            color: #1e3a8a;
            letter-spacing: 1px;
            background: transparent;
        }

        QLabel#pageTitle {
            font-size: 18px;
            font-weight: 900;
            color: #0f3054;
            background: transparent;
        }

        QLabel#subText {
            font-size: 12px;
            color: #42556b;
            background: transparent;
        }

        QLabel#sectionLabel {
            font-size: 14px;
            font-weight: 700;
            color: #1d4ed8;
            background: transparent;
        }

        QLabel#ribbonTitle {
            font-size: 11px;
            font-weight: 800;
            color: #5b7da7;
            letter-spacing: 0.6px;
            text-transform: uppercase;
            background: transparent;
            padding-left: 2px;
        }

        QLabel#modalLabel {
            color: #27486f;
            font-weight: 700;
            background: transparent;
        }

        QLabel#detailValue {
            color: #17304d;
            background-color: #f5f9fd;
            border: 1px solid #c9d8e6;
            border-radius: 10px;
            padding: 8px 10px;
        }

        QLabel#summaryTitle {
            color: #5b7da7;
            font-size: 12px;
            font-weight: 700;
            background: transparent;
        }

        QLabel#summaryValue {
            color: #17304d;
            font-size: 24px;
            font-weight: 900;
            background: transparent;
        }

        QDialog {
            background-color: #eef4fb;
            color: #17304d;
        }

        QLineEdit#modalField,
        QPlainTextEdit#modalField {
            background-color: white;
            color: #17304d;
            border: 1px solid #c9d8e6;
            border-radius: 10px;
            padding: 8px 10px;
        }

        QLineEdit#modalField:read-only,
        QPlainTextEdit#modalField:read-only {
            background-color: #f5f9fd;
        }

        QLineEdit#ribbonInput,
        QComboBox#ribbonInput {
            background-color: white;
            color: #17304d;
            border: 1px solid #c9d8e6;
            border-radius: 10px;
            padding: 8px 10px;
            min-height: 20px;
        }

        QComboBox#ribbonInput::drop-down {
            border: none;
            width: 20px;
        }

        QPushButton {
            background-color: #1f6feb;
            color: white;
            border: none;
            border-radius: 12px;
            padding: 11px 18px;
            font-weight: 700;
            min-height: 20px;
        }

        QPushButton:hover {
            background-color: #3180fb;
        }

        QPushButton:pressed {
            background-color: #1859c5;
        }

        QToolButton#settingsButton {
            background-color: #ffffff;
            color: #0f172a;
            border: 1px solid #d4dfec;
            border-radius: 18px;
            font-size: 13px;
            font-weight: 700;
            min-width: 84px;
            min-height: 38px;
            padding: 4px 12px;
        }

        QToolButton#settingsButton:hover {
            background-color: #f1f6fb;
        }

        QMenu {
            background-color: white;
            color: #0f172a;
            border: 1px solid #d4dfec;
            border-radius: 10px;
            padding: 8px;
        }

        QMenu::item {
            padding: 8px 22px 8px 12px;
            border-radius: 6px;
        }

        QMenu::item:selected {
            background-color: #dbeafe;
        }

        QGroupBox {
            font-weight: 700;
            border: 1px solid #d4dfec;
            border-radius: 16px;
            margin-top: 10px;
            padding-top: 10px;
            background-color: white;
        }

        QGroupBox::title {
            subcontrol-origin: margin;
            left: 14px;
            padding: 0 6px 0 6px;
            color: #1d4ed8;
        }

        QLabel#legendCount {
            background-color: #dbeafe;
            color: #1d4ed8;
            border-radius: 8px;
            padding: 4px 8px;
            font-weight: 800;
            min-width: 24px;
        }

        QLabel#legendText {
            color: #334155;
            padding-left: 2px;
            background: transparent;
        }

        QTableWidget {
            background-color: white;
            alternate-background-color: #f7fbff;
            border: 1px solid #d4dfec;
            border-radius: 16px;
            color: #0f172a;
            selection-background-color: #bfdbfe;
            selection-color: #0f172a;
            padding: 10px;
            outline: 0;
            gridline-color: #d7e3ef;
        }

        QHeaderView::section {
            background-color: #edf4fb;
            color: #173b67;
            padding: 10px 12px;
            border: 1px solid #d4dfec;
            font-weight: 800;
            text-align: center;
        }

        QTableWidget::item {
            padding: 12px 10px;
            border-right: 1px solid #d7e3ef;
            border-bottom: 1px solid #d7e3ef;
        }

        QTableWidget QLineEdit {
            background-color: #ffffff;
            color: #17304d;
            border: none;
            border-radius: 0px;
            padding: 0px;
            margin: 0px;
            selection-background-color: #bfdbfe;
        }

        QTableWidget QLineEdit:focus {
            background-color: #ffffff;
            border: none;
        }

        QTableWidget QComboBox {
            background-color: #ffffff;
            color: #17304d;
            border: none;
            border-radius: 0px;
            padding: 0px;
        }

        QTableCornerButton::section {
            background-color: #edf4fb;
            border: 1px solid #d4dfec;
        }

        QScrollBar:vertical {
            background: transparent;
            width: 15px;
            margin: 4px;
            border: none;
        }

        QScrollBar::handle:vertical {
            background: #8aa3be;
            border-radius: 7px;
            min-height: 34px;
        }

        QScrollBar::add-line:vertical,
        QScrollBar::sub-line:vertical,
        QScrollBar::add-line:horizontal,
        QScrollBar::sub-line:horizontal {
            border: none;
            background: none;
            height: 0px;
            width: 0px;
        }

        QScrollBar:horizontal {
            background: transparent;
            height: 15px;
            margin: 4px;
            border: none;
        }

        QScrollBar::handle:horizontal {
            background: #8aa3be;
            border-radius: 7px;
            min-width: 34px;
        }
        """

    def show_expiry_page(self):
        self.populate_expiry_table()
        self.refresh_expiry_views()
        self.stacked.setCurrentWidget(self.expiry_page)

    def show_main_page(self):
        if self.stacked.currentWidget() == self.expiry_page:
            choice = self.prompt_unsaved_changes()
            if choice == "cancel":
                return
        self.refresh_main_dashboard_table()
        self.stacked.setCurrentWidget(self.main_page)

    def closeEvent(self, event):
        choice = self.prompt_unsaved_changes()
        if choice == "cancel":
            event.ignore()
            return
        event.accept()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = LeaseMonitoringWindow()
    window.show()
    sys.exit(app.exec())
