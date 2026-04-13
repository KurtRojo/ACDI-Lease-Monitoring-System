from PyQt6.QtCore import Qt, QSignalBlocker
from PyQt6.QtGui import QAction, QKeySequence
from PyQt6.QtWidgets import (
    QAbstractItemView,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from Backend.shared_data import store


class SheetTableWidget(QTableWidget):
    def keyPressEvent(self, event):
        if event.key() in (Qt.Key.Key_Delete, Qt.Key.Key_Backspace):
            parent_window = self.window()
            should_delete = True
            if hasattr(parent_window, "confirm_delete_action"):
                should_delete = parent_window.confirm_delete_action()

            if should_delete:
                for item in self.selectedItems():
                    item.setText("")
            event.accept()
            return
        super().keyPressEvent(event)


class ContractExpiryWindow(QWidget):
    HEADERS = [
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

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Contract Expiry")
        self.resize(1700, 850)

        self.table = None
        self.legend_group = None
        self.legend_labels = []
        self.summary_label = None
        self.updating = False
        self.dirty = False

        self.create_shortcuts()
        self.build_ui()
        self.load_data()

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
        self.save_action.triggered.connect(self.save_rows)
        self.addAction(self.save_action)

        self.revert_action = QAction("Revert", self)
        self.revert_action.setShortcut("Ctrl+Shift+R")
        self.revert_action.triggered.connect(self.revert_rows)
        self.addAction(self.revert_action)

    def build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(18, 18, 18, 18)
        root.setSpacing(14)

        self.setStyleSheet(
            """
            QWidget {
                background-color: #eef4fb;
                color: #17304d;
                font-family: Segoe UI, Arial, sans-serif;
                font-size: 13px;
            }

            QGroupBox {
                background-color: white;
                border: 1px solid #d4dfec;
                border-radius: 16px;
                margin-top: 12px;
                padding-top: 12px;
                font-weight: 700;
            }

            QGroupBox::title {
                subcontrol-origin: margin;
                left: 14px;
                padding: 0 6px;
                color: #1d4ed8;
            }

            QPushButton {
                background-color: #1f6feb;
                color: white;
                border: none;
                border-radius: 12px;
                padding: 11px 18px;
                font-weight: 700;
            }

            QPushButton:hover {
                background-color: #3180fb;
            }

            QPushButton:pressed {
                background-color: #1859c5;
            }

            QTableWidget {
                background-color: white;
                alternate-background-color: #f7fbff;
                border: 1px solid #d4dfec;
                border-radius: 16px;
                gridline-color: #d7e3ef;
                selection-background-color: #bfdbfe;
                selection-color: #0f172a;
                padding: 10px;
            }

            QHeaderView::section {
                background-color: #edf4fb;
                color: #173b67;
                border: 1px solid #d4dfec;
                padding: 12px 10px;
                font-weight: 800;
            }

            QTableWidget::item {
                padding: 10px 8px;
                border-right: 1px solid #d7e3ef;
                border-bottom: 1px solid #d7e3ef;
            }

            QTableCornerButton::section {
                background-color: #edf4fb;
                border: 1px solid #d4dfec;
            }
            """
        )

        top_row = QHBoxLayout()
        top_row.setSpacing(14)

        title = QLabel("CONTRACT EXPIRY")
        title.setStyleSheet("font-size: 24px; font-weight: 900; color: #0f3054;")
        top_row.addWidget(title)
        top_row.addStretch()

        self.legend_group = QGroupBox("LEGEND")
        legend_layout = QGridLayout(self.legend_group)
        legend_layout.setContentsMargins(8, 8, 8, 8)
        legend_layout.setHorizontalSpacing(10)
        legend_layout.setVerticalSpacing(4)
        top_row.addWidget(self.legend_group)

        root.addLayout(top_row)

        action_row = QHBoxLayout()
        action_row.setSpacing(10)

        add_btn = QPushButton("Add Row")
        add_btn.clicked.connect(self.add_row)

        remove_btn = QPushButton("Remove Selected Row")
        remove_btn.clicked.connect(self.remove_row)

        refresh_btn = QPushButton("Refresh Reminders")
        refresh_btn.clicked.connect(self.load_data)

        save_btn = QPushButton("Save")
        save_btn.clicked.connect(self.save_rows)

        revert_btn = QPushButton("Revert")
        revert_btn.clicked.connect(self.revert_rows)

        self.summary_label = QLabel()
        self.summary_label.setWordWrap(True)
        self.summary_label.setStyleSheet(
            "background: white; border: 1px solid #d4dfec; border-radius: 14px; padding: 12px 14px; color: #42556b;"
        )

        action_row.addWidget(add_btn)
        action_row.addWidget(remove_btn)
        action_row.addWidget(refresh_btn)
        action_row.addWidget(save_btn)
        action_row.addWidget(revert_btn)
        action_row.addWidget(self.summary_label, 1)

        root.addLayout(action_row)

        self.table = SheetTableWidget()
        self.table.setColumnCount(len(self.HEADERS))
        self.table.setHorizontalHeaderLabels(self.HEADERS)
        self.table.verticalHeader().setVisible(False)
        self.table.setWordWrap(True)
        self.table.setAlternatingRowColors(True)
        self.table.setShowGrid(True)
        self.table.setGridStyle(Qt.PenStyle.SolidLine)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectItems)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.table.setEditTriggers(
            QAbstractItemView.EditTrigger.DoubleClicked
            | QAbstractItemView.EditTrigger.EditKeyPressed
            | QAbstractItemView.EditTrigger.AnyKeyPressed
        )

        widths = [180, 100, 100, 140, 170, 145, 140, 95, 95, 90, 180, 120, 300]
        for i, width in enumerate(widths):
            self.table.setColumnWidth(i, width)

        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Fixed)
        self.table.itemChanged.connect(self.handle_item_changed)
        root.addWidget(self.table)

    def load_data(self):
        self.load_legend()
        rows = store.get_expiry_rows()
        blocker = QSignalBlocker(self.table)
        self.updating = True

        self.table.setRowCount(len(rows))
        for row_index, row_data in enumerate(rows):
            self.table.setRowHeight(row_index, 30)
            row_data[6] = store.reminder_date_for(row_data)
            for col_index, value in enumerate(row_data):
                item = self.table.item(row_index, col_index)
                if item is None:
                    item = QTableWidgetItem()
                    item.setTextAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
                    self.table.setItem(row_index, col_index, item)
                item.setText(str(value))

        del blocker
        self.updating = False
        self.dirty = False
        self.update_window_title()
        self.load_summary()

    def load_legend(self):
        layout = self.legend_group.layout()
        while layout.count():
            child = layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()

        self.legend_labels = []
        for i, (count, text) in enumerate(store.get_legend_rows()):
            count_label = QLabel(str(count))
            text_label = QLabel(text)
            layout.addWidget(count_label, i, 0)
            layout.addWidget(text_label, i, 1)
            self.legend_labels.append((count_label, text_label))

    def load_summary(self):
        notices = []
        from datetime import date

        today = date.today()
        for row in self.collect_rows():
            branch = row[0].strip()
            end_date = store.parse_date(row[8])
            reminder_date = store.parse_date(store.reminder_date_for(row))
            if not branch or not end_date or not reminder_date:
                continue
            if reminder_date <= today <= end_date:
                notices.append(
                    {
                        "branch": branch,
                        "officer": row[4].strip() or "Officer not assigned",
                        "contact": row[5].strip() or "No contact number",
                        "end_date": store.format_date(end_date),
                    }
                )
        notices.sort(key=lambda item: store.parse_date(item["end_date"]))
        if notices:
            first = notices[0]
            self.summary_label.setText(
                f"Notify {first['officer']} / {first['contact']} for {first['branch']} before {first['end_date']}."
            )
        else:
            self.summary_label.setText("No contracts are inside the 2-month reminder window.")

    def collect_rows(self):
        rows = []
        for row_index in range(self.table.rowCount()):
            row_data = []
            for col_index in range(self.table.columnCount()):
                item = self.table.item(row_index, col_index)
                row_data.append(item.text().strip() if item else "")
            row_data[6] = store.reminder_date_for(row_data)
            rows.append(row_data)
        return rows

    def handle_item_changed(self, _item):
        if self.updating:
            return
        self.dirty = True
        self.update_window_title()
        self.load_summary()

    def add_row(self):
        self.updating = True
        self.table.insertRow(self.table.rowCount())
        row = self.table.rowCount() - 1
        self.table.setRowHeight(row, 30)
        values = ["NEW BRANCH", "", "", "1 YR", "", "", "", "", "", "", "", "", "FOR RENEW / ESCALATION"]
        for col, value in enumerate(values):
            item = QTableWidgetItem(str(value))
            item.setTextAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
            self.table.setItem(row, col, item)
        self.updating = False
        self.dirty = True
        self.update_window_title()
        self.load_summary()

    def remove_row(self):
        selected_rows = sorted({index.row() for index in self.table.selectedIndexes()}, reverse=True)
        if not selected_rows:
            QMessageBox.information(self, "Remove Row", "Select at least one row to remove.")
            return

        if not self.confirm_delete_action():
            return

        self.updating = True
        for row_index in selected_rows:
            self.table.removeRow(row_index)
        self.updating = False
        self.dirty = True
        self.update_window_title()
        self.load_summary()

    def refresh_data(self):
        self.load_data()

    def confirm_delete_action(self):
        result = QMessageBox.question(
            self,
            "Confirm Delete",
            "Do you confirm in deleting info?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        return result == QMessageBox.StandardButton.Yes

    def save_rows(self):
        store.set_expiry_rows(self.collect_rows())
        self.load_data()

    def revert_rows(self):
        if self.dirty and not self.confirm_revert_action():
            return
        store.load()
        self.load_data()

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
        if not self.dirty:
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
            self.save_rows()
            return "save"
        if clicked == discard_button:
            self.revert_rows_silent()
            return "discard"
        return "cancel"

    def revert_rows_silent(self):
        store.load()
        self.load_data()

    def update_window_title(self):
        suffix = " *" if self.dirty else ""
        self.setWindowTitle(f"Contract Expiry{suffix}")

    def copy_selected_cells(self):
        selected_indexes = self.table.selectedIndexes()
        if not selected_indexes:
            return
        rows = sorted({index.row() for index in selected_indexes})
        cols = sorted({index.column() for index in selected_indexes})
        lines = []
        for row in rows:
            values = []
            for col in cols:
                item = self.table.item(row, col)
                values.append(item.text() if item else "")
            lines.append("\t".join(values))
        from PyQt6.QtWidgets import QApplication
        QApplication.clipboard().setText("\n".join(lines))

    def cut_selected_cells(self):
        selected_items = self.table.selectedItems()
        if not selected_items:
            return
        self.copy_selected_cells()
        if not self.confirm_delete_action():
            return
        self.updating = True
        for item in selected_items:
            item.setText("")
        self.updating = False
        self.dirty = True
        self.update_window_title()
        self.load_summary()

    def paste_cells(self):
        from PyQt6.QtWidgets import QApplication
        start_row = self.table.currentRow()
        start_col = self.table.currentColumn()
        if start_row < 0 or start_col < 0:
            return
        clipboard_text = QApplication.clipboard().text()
        if not clipboard_text:
            return
        self.updating = True
        for row_offset, line in enumerate(clipboard_text.splitlines()):
            values = line.split("\t")
            target_row = start_row + row_offset
            while target_row >= self.table.rowCount():
                self.table.insertRow(self.table.rowCount())
                self.table.setRowHeight(self.table.rowCount() - 1, 30)
            for col_offset, value in enumerate(values):
                target_col = start_col + col_offset
                if target_col >= self.table.columnCount():
                    continue
                item = self.table.item(target_row, target_col)
                if item is None:
                    item = QTableWidgetItem()
                    item.setTextAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
                    self.table.setItem(target_row, target_col, item)
                item.setText(value)
        self.updating = False
        self.dirty = True
        self.update_window_title()
        self.load_summary()

    def closeEvent(self, event):
        choice = self.prompt_unsaved_changes()
        if choice == "cancel":
            event.ignore()
            return
        event.accept()
