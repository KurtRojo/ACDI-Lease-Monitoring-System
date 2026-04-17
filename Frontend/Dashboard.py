import csv
import os
import sys
from copy import deepcopy

PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from PyQt6.QtCore import Qt, QItemSelectionModel, QSignalBlocker
from PyQt6.QtGui import QAction, QColor, QKeySequence
from PyQt6.QtWidgets import (
    QApplication,
    QAbstractItemView,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFrame,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMenu,
    QMessageBox,
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

try:
    from Backend.shared_data import store
except ModuleNotFoundError:
    import sys
    from pathlib import Path

    project_root = Path(__file__).resolve().parents[1]
    if str(project_root) not in sys.path:
        sys.path.insert(0, str(project_root))
    from Backend.shared_data import store

try:
    from Frontend.Add import AddBranchDialog
except ModuleNotFoundError:
    from Add import AddBranchDialog

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


class LeaseMonitoringWindow(QMainWindow):
    STATUS_OPTIONS = [
        "",
        "For Legal Review",
        "For GSD Review",
        "For AD Review",
        "For OD Review",
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
        "REMINDER",
        "FROM",
        "TO",
        "FLOOR AREA",
        "MEMO #",
        "DATE COMPLETED / SENT",
        "REMARKS",
    ]

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
        self.expiry_highlighted_column = -1
        self.legend_labels = []
        self.notification_cards = []

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
            row_height=32,
            fixed_resize=False,
            editable=False,
        )

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

    def refresh_main_dashboard_table(self):
        rows = store.get_main_dashboard_rows()
        blocker = QSignalBlocker(self.main_table)
        self.main_table.setRowCount(len(rows))

        for row_index, row_data in enumerate(rows):
            self.main_table.setRowHeight(row_index, 32)
            for col_index, value in enumerate(row_data):
                item = self.main_table.item(row_index, col_index)
                if item is None:
                    item = QTableWidgetItem()
                    item.setTextAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
                    self.main_table.setItem(row_index, col_index, item)
                item.setText(str(value))
        del blocker
        self.auto_size_main_dashboard_columns()
        self.refresh_dashboard_summary()

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

        self.undo_btn = QPushButton("↶")
        self.undo_btn.setToolTip("Undo")
        self.undo_btn.clicked.connect(self.undo_expiry_change)

        self.redo_btn = QPushButton("↷")
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
        alerts_group = self.build_ribbon_group("Alerts", [self.notifications_btn])

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
        action_layout.addWidget(alerts_group)

        root.addWidget(action_bar)

        self.expiry_table = SheetTableWidget()
        expiry_widths = [180, 130, 120, 100, 170, 160, 0, 100, 100, 100, 140, 200, 330]

        self.configure_table(
            table=self.expiry_table,
            headers=self.EXPIRY_HEADERS,
            widths=expiry_widths,
            rows=[],
            row_height=34,
            fixed_resize=True,
            editable=True,
        )
        self.expiry_table.setColumnHidden(6, True)

        self.expiry_table.setItemDelegateForColumn(12, StatusComboDelegate(self.STATUS_OPTIONS, self.expiry_table))

        self.expiry_table.itemChanged.connect(self.handle_expiry_item_changed)
        self.expiry_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectItems)
        self.expiry_table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.expiry_table.cellClicked.connect(self.handle_expiry_cell_clicked)
        self.expiry_table.currentCellChanged.connect(self.handle_expiry_current_cell_changed)
        self.expiry_table.horizontalHeader().setDefaultAlignment(Qt.AlignmentFlag.AlignCenter)
        self.expiry_table.horizontalHeader().setFixedHeight(75)
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
        table.setVerticalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        table.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
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
            table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
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
        self.main_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        self.main_table.resizeColumnsToContents()

        max_widths = {
            0: 170,
            1: 720,
            2: 120,
            3: 180,
            4: 150,
            5: 150,
            6: 170,
            7: 120,
            8: 120,
            9: 180,
            10: 170,
            11: 190,
            12: 520,
        }
        min_widths = {
            0: 130,
            1: 420,
            2: 80,
            3: 105,
            4: 100,
            5: 100,
            6: 120,
            7: 80,
            8: 80,
            9: 130,
            10: 120,
            11: 140,
            12: 320,
        }

        for column in range(self.main_table.columnCount()):
            current = self.main_table.columnWidth(column)
            current = max(current, min_widths.get(column, 80))
            current = min(current, max_widths.get(column, 520))
            self.main_table.setColumnWidth(column, current)

        self.main_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Interactive)

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

        if self.expiry_table.rowCount() > 0:
            if self.expiry_highlighted_column < 0:
                self.expiry_highlighted_column = 0
            self.set_expiry_column_focus(self.expiry_highlighted_column, 0)

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
                    if col_index == self.expiry_highlighted_column:
                        if status == "":
                            hc = QColor("#cce2ff") if self.current_theme == "light" else QColor("#2a3b5c")
                            item.setBackground(hc)
                        else:
                            hc = background.darker(115) if self.current_theme == "light" else background.lighter(135)
                            item.setBackground(hc)
                    else:
                        item.setBackground(background)

    def handle_expiry_cell_clicked(self, row, col):
        if self.expiry_table_updating:
            return
        self.set_expiry_column_focus(col, row)

    def handle_expiry_current_cell_changed(self, current_row, current_col, _previous_row, _previous_col):
        if self.expiry_table_updating:
            return
        if current_col < 0 or current_row < 0:
            return
        self.set_expiry_column_focus(current_col, current_row)

    def set_expiry_column_focus(self, column, row=None):
        if column < 0:
            return
        if row is None or row < 0:
            row = self.expiry_table.currentRow()
        if row < 0:
            row = 0

        blocker = QSignalBlocker(self.expiry_table)
        self.expiry_table_updating = True

        self.expiry_table.setCurrentCell(row, column)

        self.expiry_table_updating = False
        del blocker

        self.expiry_highlighted_column = column
        self.apply_expiry_row_styles(self.collect_expiry_rows_from_table())

    def add_expiry_row(self):
        default_row = ["", "", "", "1 YR", "", "", "", "", "", "", "", "", "FOR RENEW / ESCALATION"]
        dialog = AddBranchDialog(
            self.EXPIRY_HEADERS,
            default_row,
            self,
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

        if not self.confirm_delete_action():
            return

        previous_rows = self.collect_expiry_rows_from_table()
        self.expiry_table_updating = True
        for row_index in selected_rows:
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
            padding: 12px 10px;
            border: 1px solid #22324a;
            font-weight: 800;
        }

        QTableWidget::item {
            padding: 10px 8px;
            border-right: 1px solid #243750;
            border-bottom: 1px solid #243750;
        }

        QTableWidget QLineEdit {
            background-color: #15253a;
            color: #f8fafc;
            border: 2px solid #3277ea;
            border-radius: 6px;
            padding: 6px 8px;
            selection-background-color: #3277ea;
        }

        QTableWidget QLineEdit:focus {
            background-color: #1a2f48;
            border: 2px solid #60a5fa;
        }

        QTableWidget QComboBox {
            background-color: #15253a;
            color: #f8fafc;
            border: 2px solid #3277ea;
            border-radius: 6px;
            padding: 4px 8px;
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
            padding: 12px 10px;
            border: 1px solid #d4dfec;
            font-weight: 800;
        }

        QTableWidget::item {
            padding: 10px 8px;
            border-right: 1px solid #d7e3ef;
            border-bottom: 1px solid #d7e3ef;
        }

        QTableWidget QLineEdit {
            background-color: #ffffff;
            color: #17304d;
            border: 2px solid #3180fb;
            border-radius: 6px;
            padding: 6px 8px;
            selection-background-color: #bfdbfe;
        }

        QTableWidget QLineEdit:focus {
            background-color: #f7fbff;
            border: 2px solid #60a5fa;
        }

        QTableWidget QComboBox {
            background-color: #ffffff;
            color: #17304d;
            border: 2px solid #3180fb;
            border-radius: 6px;
            padding: 4px 8px;
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
