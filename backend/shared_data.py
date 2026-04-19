import json
import sqlite3
from copy import deepcopy
from datetime import date, datetime, timedelta
from pathlib import Path

import mysql.connector
from mysql.connector import Error


class LeaseDataStore:
    DASHBOARD_STAGE_FLOW = [
        ("LEGAL", "LEGAL"),
        ("VLG HEAD", "VLG H"),
        ("GSD OFFICER", "GSD"),
        ("AD OFFICER", "AD"),
        ("OD OFFICER", "OD"),
        ("VP-ASSIGNED OTD", "VP-ASSIGNED OTD"),
        ("EVPO-EVPA", "EVPO-EVPA"),
        ("PRESIDENT", "PRESIDENT"),
    ]

    DATE_FORMATS = (
        "%d-%b-%y",
        "%d-%b-%Y",
        "%d-%B-%y",
        "%d-%B-%Y",
        "%d/%m/%Y",
        "%m/%d/%Y",
        "%Y-%m-%d",
        "%B %d, %Y",
        "%b %d, %Y",
    )

    def __init__(self):
        self.config_path = Path(__file__).with_name("mysql_config.json")
        self.schema_path = Path(__file__).with_name("mysql_schema.sql")
        self.legacy_sqlite_path = Path(__file__).with_name("lease_data.db")
        self.legacy_json_path = Path(__file__).with_name("lease_data.json")
        self.defaults = {
            "ui_settings": {
                "theme": "dark",
                "contract_documents": "{}",
                "dashboard_stage_statuses": "{}",
                "dashboard_remark_overrides": "{}",
            },
            "main_dashboard_rows": [
                [
                    "September 09, 2025",
                    "REVISED REQUEST FOR MEMORANDUM ON THE APPROVED CONTRACT OF LEASE EXTENSION OFFICE",
                    "",
                    "",
                    "OK 09.10.25",
                    "OK 9.10.25",
                    "OK 9.22.25",
                    "OK 9.22.2025",
                    "OK 9.22.25",
                    "OK 10.08.25",
                    "OK 10.7.25",
                    "OK 10.07.25",
                    "SENT TO MANOLO EO 10.8.25",
                ],
                [
                    "October 06, 2025",
                    "REQUEST FOR APPROVAL OF VARIOUS VISMIN STOCK ROOM FOR 2025 (LOT 2S / BATCH 1)",
                    "",
                    "",
                    "OK 10.06.25",
                    "OK 10.06.25",
                    "OK 10.08.25",
                    "OK 10.08.25",
                    "OK 10.08.25",
                    "OK 10.09.25",
                    "OK 10.08.25",
                    "OK 10.08.25",
                    "SENT TO VARIOUS OFFICES 10.09.25",
                ],
                [
                    "November 17, 2025",
                    "REVISED REQUEST FOR APPROVAL RE LEASE CONTRACT FOR RENEWAL OF ACDI VISMIN TRANSIENT-STOCK ROOM FOR PR",
                    "",
                    "",
                    "OK 11.17.25",
                    "OK 11.17.25",
                    "OK 11.25.25",
                    "OK 11.27.25",
                    "OK 11.27.25",
                    "OK 11.27.25",
                    "OK 12.04.25",
                    "OK 12.09.25",
                    "SENT TO VARIOUS OFFICES 12.09.25",
                ],
            ],
            "expiry_rows": [
                ["VISMIN", "", "11-Apr-25", "1 YR", "", "", "", "1-Jun-23", "31-May-24", "80 sqm", "CMSG GSD MEMO 008-2023", "", "DONE"],
                ["CVAO", "", "", "", "", "", "", "", "", "", "", "", ""],
                ["MBEAB", "18-Nov-24", "21-Jan-25", "1 YR", "MARY CRIS S PERALES", "", "", "1-Jan-25", "31-Dec-25", "221", "GS MEMO 026-2025", "31-Jan-25", "DONE"],
                ["LAPULAPU EO", "26-Jan-24", "30-Jan-24", "1 YR", "", "", "", "1-Feb-24", "1-Feb-25", "24.5", "GS MEMO 117-2024", "7-Oct-24", "DONE"],
                ["LAPULAPU EO", "20-Jun-25", "", "1 YR", "", "", "28-Feb-26", "1-May-25", "30-Apr-26", "24.5", "", "", "DONE - MAY PLANS FOR FINDING ANOTHER AREA"],
                ["TAGBILARAN", "24-Jan-24", "20-Dec-24", "1 YR", "", "", "", "1-Jun-24", "31-May-25", "60", "GSD MEMO 081-2025", "14-Jul-25", "DONE"],
                ["TAGBILARAN", "15-Jul-25", "", "1 YR", "", "", "31-Mar-26", "1-Jun-25", "31-May-26", "60", "VLG GS MEMO 004-2025", "22-Aug-25", "DONE - MAY PLANS FOR FINDING ANOTHER AREA"],
                ["CEBU", "28-Oct-24", "19-Nov-24", "5 YRS", "Milane C Fernandez- CVAO", "0917-629-4632", "31-Oct-29", "1-Jan-25", "31-Dec-29", "110", "GS MEMO 012-2025", "21-Jan-25", "Waiting for Notarized Contract"],
                ["DAVAO", "12-Jan-26", "", "1 YR", "Milane C Fernandez- CVAO", "0917-629-4632", "1-Sep-26", "1-Nov-25", "1-Nov-26", "20 sqm", "", "", "Waiting for Notarized Contract"],
                ["BOGO EO", "8-Jan-24", "12-Jan-24", "2 YRS", "", "", "", "1-Feb-24", "1-Feb-26", "32", "LUZON GSD MEMO 005-2024", "6-Feb-24", "DONE"],
                ["PGIAO", "", "", "", "", "", "", "", "", "", "", "", ""],
                ["ILOILO TRANSIENT", "16-Jul-25", "26-Aug-25", "3 YRS", "MARGIE HERAS FERRO", "", "30-Jun-28", "1-Sep-25", "31-Aug-28", "75 sqm", "VLG GSS MEMO NO. 007-2025", "4-Nov-25", "DONE"],
                ["SAN JOSE ANTIQUE MO (ILOILO)", "", "", "2 YRS", "", "", "31-May-26", "1-Aug-24", "31-Jul-26", "33 sqm", "LUZON GSD MEMO 068-2024", "28-May-24", "DONE"],
                ["KALIBO MO", "26-Feb-25", "3-Mar-25", "1 YR", "", "", "28-Feb-26", "1-May-25", "30-Apr-26", "30", "GS MEMO 075-2025", "10-Mar-25", "WILL LIPAT NG"],
                ["KALIBO MO", "4-Mar-24", "13-Mar-24", "1 YR", "", "", "", "1-May-24", "30-Apr-25", "30", "LUZON GSD MEMO 055-2024", "4-Jul-24", "DONE"],
                ["ROXAS EO", "22-Aug-24", "6-Sep-24", "5 YRS", "", "", "31-Aug-29", "1-Nov-24", "31-Oct-29", "100", "GS MEMO 137-2024", "17-Oct-24", "DONE"],
                ["ROXAS EO TRANSIENT", "24-Jun-24", "19-Jul-24", "1 YR", "", "", "", "15-Aug-24", "14-Aug-25", "28", "GS MEMO 118-2024", "29-Oct-24", "DONE"],
                ["ROXAS EO TRANSIENT", "16-Jul-25", "", "1 YR", "", "", "14-Jun-26", "15-Aug-25", "14-Aug-26", "28", "", "", ""],
            ],
        }
        self._ensure_config_file()
        self.config = self._load_config()
        self.main_dashboard_rows = []
        self.expiry_rows = []
        self.ui_settings = {}
        self.backend = "mysql"
        self.last_connection_error = ""
        self._initialize_storage()
        self.load()

    def _ensure_config_file(self):
        if self.config_path.exists():
            return
        default_config = {
            "host": "127.0.0.1",
            "port": 3306,
            "user": "root",
            "password": "",
            "database": "acdi_lease_monitoring",
        }
        self.config_path.write_text(json.dumps(default_config, indent=2), encoding="utf-8")

    def _load_config(self):
        config = json.loads(self.config_path.read_text(encoding="utf-8"))
        return {
            "host": config.get("host", "127.0.0.1"),
            "port": int(config.get("port", 3306)),
            "user": config.get("user", "root"),
            "password": config.get("password", ""),
            "database": config.get("database", "acdi_lease_monitoring"),
        }

    def _connect_server(self):
        return mysql.connector.connect(
            host=self.config["host"],
            port=self.config["port"],
            user=self.config["user"],
            password=self.config["password"],
        )

    def _connect_database(self):
        return mysql.connector.connect(
            host=self.config["host"],
            port=self.config["port"],
            user=self.config["user"],
            password=self.config["password"],
            database=self.config["database"],
        )

    def _initialize_storage(self):
        try:
            self._initialize_mysql_database()
            self.backend = "mysql"
        except RuntimeError as exc:
            self.backend = "sqlite"
            self.last_connection_error = str(exc)
            self._initialize_sqlite_database()

    def _initialize_mysql_database(self):
        try:
            with self._connect_server() as connection:
                cursor = connection.cursor()
                cursor.execute(
                    f"CREATE DATABASE IF NOT EXISTS `{self.config['database']}` CHARACTER SET utf8mb4 COLLATE utf8mb4_unicode_ci"
                )
                connection.commit()

            with self._connect_database() as connection:
                cursor = connection.cursor()
                for statement in self._schema_statements():
                    cursor.execute(statement)
                connection.commit()

                cursor.execute("SELECT COUNT(*) FROM main_dashboard_rows")
                dashboard_count = cursor.fetchone()[0]
                cursor.execute("SELECT COUNT(*) FROM expiry_rows")
                expiry_count = cursor.fetchone()[0]
                cursor.execute("SELECT COUNT(*) FROM ui_settings")
                settings_count = cursor.fetchone()[0]

                if dashboard_count == 0 and expiry_count == 0 and settings_count == 0:
                    seed_data = self._read_legacy_sqlite() or self._read_legacy_json() or deepcopy(self.defaults)
                    self._write_all_to_database(connection, seed_data)
        except Error as exc:
            raise RuntimeError(
                "MySQL connection failed. Update mysql_config.json with a running MySQL server before starting the app.\n"
                f"Original error: {exc}"
            ) from exc

    def _initialize_sqlite_database(self):
        with sqlite3.connect(self.legacy_sqlite_path) as connection:
            cursor = connection.cursor()
            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS ui_settings (
                    key TEXT PRIMARY KEY,
                    value TEXT NOT NULL
                )
                """
            )
            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS main_dashboard_rows (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    sort_order INTEGER NOT NULL,
                    col0 TEXT, col1 TEXT, col2 TEXT, col3 TEXT, col4 TEXT, col5 TEXT, col6 TEXT,
                    col7 TEXT, col8 TEXT, col9 TEXT, col10 TEXT, col11 TEXT, col12 TEXT
                )
                """
            )
            cursor.execute(
                """
                CREATE TABLE IF NOT EXISTS expiry_rows (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    sort_order INTEGER NOT NULL,
                    col0 TEXT, col1 TEXT, col2 TEXT, col3 TEXT, col4 TEXT, col5 TEXT, col6 TEXT,
                    col7 TEXT, col8 TEXT, col9 TEXT, col10 TEXT, col11 TEXT, col12 TEXT
                )
                """
            )
            cursor.execute("SELECT COUNT(*) FROM main_dashboard_rows")
            dashboard_count = cursor.fetchone()[0]
            cursor.execute("SELECT COUNT(*) FROM expiry_rows")
            expiry_count = cursor.fetchone()[0]
            cursor.execute("SELECT COUNT(*) FROM ui_settings")
            settings_count = cursor.fetchone()[0]

            if dashboard_count == 0 and expiry_count == 0 and settings_count == 0:
                seed_data = self._read_legacy_json() or deepcopy(self.defaults)
                self._write_all_to_sqlite(connection, seed_data)
            connection.commit()

    def _schema_statements(self):
        schema = f"""
        CREATE TABLE IF NOT EXISTS ui_settings (
            setting_key VARCHAR(100) PRIMARY KEY,
            setting_value TEXT NOT NULL
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

        CREATE TABLE IF NOT EXISTS main_dashboard_rows (
            id INT AUTO_INCREMENT PRIMARY KEY,
            sort_order INT NOT NULL,
            col0 TEXT, col1 TEXT, col2 TEXT, col3 TEXT, col4 TEXT, col5 TEXT, col6 TEXT,
            col7 TEXT, col8 TEXT, col9 TEXT, col10 TEXT, col11 TEXT, col12 TEXT
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;

        CREATE TABLE IF NOT EXISTS expiry_rows (
            id INT AUTO_INCREMENT PRIMARY KEY,
            sort_order INT NOT NULL,
            col0 TEXT, col1 TEXT, col2 TEXT, col3 TEXT, col4 TEXT, col5 TEXT, col6 TEXT,
            col7 TEXT, col8 TEXT, col9 TEXT, col10 TEXT, col11 TEXT, col12 TEXT
        ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
        """
        self.schema_path.write_text(schema.strip() + "\n", encoding="utf-8")
        return [statement.strip() for statement in schema.split(";") if statement.strip()]

    def _read_legacy_sqlite(self):
        if not self.legacy_sqlite_path.exists():
            return None
        try:
            connection = sqlite3.connect(self.legacy_sqlite_path)
            cursor = connection.cursor()

            cursor.execute("SELECT key, value FROM ui_settings")
            ui_settings = dict(cursor.fetchall())

            cursor.execute(
                "SELECT col0, col1, col2, col3, col4, col5, col6, col7, col8, col9, col10, col11, col12 "
                "FROM main_dashboard_rows ORDER BY sort_order, id"
            )
            main_rows = [list(row) for row in cursor.fetchall()]

            cursor.execute(
                "SELECT col0, col1, col2, col3, col4, col5, col6, col7, col8, col9, col10, col11, col12 "
                "FROM expiry_rows ORDER BY sort_order, id"
            )
            expiry_rows = [list(row) for row in cursor.fetchall()]
            connection.close()

            return {
                "ui_settings": ui_settings or deepcopy(self.defaults["ui_settings"]),
                "main_dashboard_rows": main_rows or deepcopy(self.defaults["main_dashboard_rows"]),
                "expiry_rows": expiry_rows or deepcopy(self.defaults["expiry_rows"]),
            }
        except sqlite3.Error:
            return None

    def _read_legacy_json(self):
        if not self.legacy_json_path.exists():
            return None
        try:
            return json.loads(self.legacy_json_path.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            return None

    def _normalize_rows(self, rows):
        normalized = []
        for row in rows:
            current = [str(value) for value in row[:13]]
            if len(current) < 13:
                current.extend([""] * (13 - len(current)))
            normalized.append(current)
        return normalized

    def _write_rows(self, connection, table_name, rows):
        normalized_rows = self._normalize_rows(rows)
        cursor = connection.cursor()
        cursor.execute(f"DELETE FROM `{table_name}`")
        insert_sql = (
            f"INSERT INTO `{table_name}` ("
            "sort_order, col0, col1, col2, col3, col4, col5, col6, col7, col8, col9, col10, col11, col12"
            ") VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)"
        )
        for index, row in enumerate(normalized_rows):
            cursor.execute(insert_sql, (index, *row))

    def _write_settings(self, connection, settings):
        cursor = connection.cursor()
        cursor.execute("DELETE FROM `ui_settings`")
        insert_sql = "INSERT INTO `ui_settings` (setting_key, setting_value) VALUES (%s, %s)"
        for key, value in settings.items():
            cursor.execute(insert_sql, (str(key), str(value)))

    def _write_rows_to_sqlite(self, connection, table_name, rows):
        normalized_rows = self._normalize_rows(rows)
        cursor = connection.cursor()
        cursor.execute(f"DELETE FROM {table_name}")
        insert_sql = (
            f"INSERT INTO {table_name} ("
            "sort_order, col0, col1, col2, col3, col4, col5, col6, col7, col8, col9, col10, col11, col12"
            ") VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"
        )
        for index, row in enumerate(normalized_rows):
            cursor.execute(insert_sql, (index, *row))

    def _write_settings_to_sqlite(self, connection, settings):
        cursor = connection.cursor()
        cursor.execute("DELETE FROM ui_settings")
        for key, value in settings.items():
            cursor.execute(
                "INSERT INTO ui_settings (key, value) VALUES (?, ?)",
                (str(key), str(value)),
            )

    def _write_all_to_sqlite(self, connection, data):
        settings = deepcopy(data.get("ui_settings", self.defaults["ui_settings"]))
        main_rows = deepcopy(data.get("main_dashboard_rows", self.defaults["main_dashboard_rows"]))
        expiry_rows = deepcopy(data.get("expiry_rows", self.defaults["expiry_rows"]))
        self._write_settings_to_sqlite(connection, settings)
        self._write_rows_to_sqlite(connection, "main_dashboard_rows", main_rows)
        self._write_rows_to_sqlite(connection, "expiry_rows", expiry_rows)
        connection.commit()

    def _write_all_to_database(self, connection, data):
        settings = deepcopy(data.get("ui_settings", self.defaults["ui_settings"]))
        main_rows = deepcopy(data.get("main_dashboard_rows", self.defaults["main_dashboard_rows"]))
        expiry_rows = deepcopy(data.get("expiry_rows", self.defaults["expiry_rows"]))
        self._write_settings(connection, settings)
        self._write_rows(connection, "main_dashboard_rows", main_rows)
        self._write_rows(connection, "expiry_rows", expiry_rows)
        connection.commit()

    def _read_rows(self, connection, table_name):
        cursor = connection.cursor()
        cursor.execute(
            f"SELECT col0, col1, col2, col3, col4, col5, col6, col7, col8, col9, col10, col11, col12 "
            f"FROM `{table_name}` ORDER BY sort_order, id"
        )
        return [list(row) for row in cursor.fetchall()]

    def _read_settings(self, connection):
        cursor = connection.cursor()
        cursor.execute("SELECT setting_key, setting_value FROM `ui_settings`")
        settings = dict(cursor.fetchall())
        if "theme" not in settings:
            settings["theme"] = self.defaults["ui_settings"]["theme"]
        return settings

    def _read_rows_from_sqlite(self, connection, table_name):
        cursor = connection.cursor()
        cursor.execute(
            f"SELECT col0, col1, col2, col3, col4, col5, col6, col7, col8, col9, col10, col11, col12 "
            f"FROM {table_name} ORDER BY sort_order, id"
        )
        return [list(row) for row in cursor.fetchall()]

    def _read_settings_from_sqlite(self, connection):
        cursor = connection.cursor()
        cursor.execute("SELECT key, value FROM ui_settings")
        settings = dict(cursor.fetchall())
        if "theme" not in settings:
            settings["theme"] = self.defaults["ui_settings"]["theme"]
        return settings

    def load(self):
        if self.backend == "mysql":
            with self._connect_database() as connection:
                self.main_dashboard_rows = self._read_rows(connection, "main_dashboard_rows")
                self.expiry_rows = self._read_rows(connection, "expiry_rows")
                self.ui_settings = self._read_settings(connection)
        else:
            with sqlite3.connect(self.legacy_sqlite_path) as connection:
                self.main_dashboard_rows = self._read_rows_from_sqlite(connection, "main_dashboard_rows")
                self.expiry_rows = self._read_rows_from_sqlite(connection, "expiry_rows")
                self.ui_settings = self._read_settings_from_sqlite(connection)

        if not self.main_dashboard_rows:
            self.main_dashboard_rows = deepcopy(self.defaults["main_dashboard_rows"])
        if not self.expiry_rows:
            self.expiry_rows = deepcopy(self.defaults["expiry_rows"])

        self.main_dashboard_rows = self._normalize_rows(self.main_dashboard_rows)
        self.expiry_rows = self._normalize_rows(self.expiry_rows)

    def save(self):
        data = {
            "ui_settings": self.ui_settings,
            "main_dashboard_rows": self.main_dashboard_rows,
            "expiry_rows": self.expiry_rows,
        }
        if self.backend == "mysql":
            with self._connect_database() as connection:
                self._write_all_to_database(connection, data)
        else:
            with sqlite3.connect(self.legacy_sqlite_path) as connection:
                self._write_all_to_sqlite(connection, data)

    def parse_date(self, value):
        raw_value = str(value).strip()
        if not raw_value:
            return None
        for fmt in self.DATE_FORMATS:
            try:
                return datetime.strptime(raw_value, fmt).date()
            except ValueError:
                continue
        return None

    def format_date(self, value):
        return value.strftime("%d-%b-%y") if value else ""

    def reminder_date_for(self, row):
        end_date = self.parse_date(row[8])
        if not end_date:
            return ""
        return self.format_date(end_date - timedelta(days=60))

    def days_until_expiry(self, row):
        end_date = self.parse_date(row[8])
        if not end_date:
            return None
        return (end_date - date.today()).days

    def reminder_windows_for(self, row):
        days = self.days_until_expiry(row)
        if days is None:
            return []
        windows = []
        for threshold in (90, 60, 30):
            if 0 <= days <= threshold:
                windows.append(f"{threshold}-day alert")
        return windows

    def manual_status_for(self, row):
        value = str(row[12]).strip()
        normalized = value.lower()
        if "done" in normalized:
            return "Done"
        if "approved" in normalized:
            return "Approved"
        allowed = {
            "for legal review": "For Legal Review",
            "for vlg review": "For VLG Head Review",
            "for vlg head review": "For VLG Head Review",
            "for gsd review": "For GSD Review",
            "for ad review": "For AD Review",
            "for od review": "For OD Review",
            "for vp-assigned otd review": "For VP-Assigned OTD Review",
            "for vp assigned otd review": "For VP-Assigned OTD Review",
            "for evpo-evpa review": "For EVPO-EVPA Review",
            "for evpo evpa review": "For EVPO-EVPA Review",
            "for evp approval": "For EVP Approval",
            "for president approval": "For President Approval",
            "for president review": "For President Approval",
            "approved": "Approved",
            "done": "Done",
        }
        return allowed.get(normalized, value)

    def pending_stage_for(self, row):
        status = self.manual_status_for(row).lower()
        mapping = {
            "for legal review": "LEGAL",
            "for vlg head review": "VLG HEAD",
            "for gsd review": "GSD OFFICER",
            "for ad review": "AD OFFICER",
            "for od review": "OD OFFICER",
            "for vp-assigned otd review": "VP-ASSIGNED OTD",
            "for evpo-evpa review": "EVPO-EVPA",
            "for evp approval": "EVPO-EVPA",
            "for president approval": "PRESIDENT",
            "approved": "COMPLETED",
            "done": "COMPLETED",
        }
        if status in mapping:
            return mapping[status]
        contract_state = self.contract_status_for(row)
        if contract_state == "done":
            return "COMPLETED"
        if contract_state == "expired":
            return "EXPIRED"
        return "LEGAL"

    def dashboard_stage_cells_for(self, contract_status, pending_stage):
        cells = []
        pending_index = next(
            (index for index, (stage_key, _label) in enumerate(self.DASHBOARD_STAGE_FLOW) if stage_key == pending_stage),
            None,
        )

        for index, (_stage_key, label) in enumerate(self.DASHBOARD_STAGE_FLOW):
            if contract_status == "done" or pending_stage == "COMPLETED":
                cells.append(f"COMPLETE {label}")
                continue

            if contract_status == "expired":
                if pending_index is not None and index < pending_index:
                    cells.append(f"COMPLETE {label}")
                elif pending_index is not None and index == pending_index:
                    cells.append(f"EXPIRED {label}")
                else:
                    cells.append("")
                continue

            if pending_index is None:
                cells.append("")
                continue

            if index < pending_index:
                cells.append(f"COMPLETE {label}")
            elif index == pending_index:
                cells.append(f"FOR ACTION {label}" if contract_status == "due" else "PENDING")
            else:
                cells.append("")

        return cells

    def contract_status_for(self, row):
        remarks = str(row[12]).strip().lower()
        end_date = self.parse_date(row[8])
        days_until = self.days_until_expiry(row)
        today = date.today()

        if remarks == "done" or "done" in remarks:
            return "done"
        if end_date and end_date < today:
            return "expired"
        if days_until is not None and 0 <= days_until <= 90:
            return "due"
        if any(str(cell).strip() for cell in row):
            return "active"
        return "blank"

    def notification_rows(self):
        notices = []
        today = date.today()

        for row in self.expiry_rows:
            branch = row[0].strip()
            end_date = self.parse_date(row[8])
            if not branch or not end_date:
                continue

            reminder_date = self.parse_date(row[6]) or self.parse_date(self.reminder_date_for(row))
            reminder_windows = self.reminder_windows_for(row)
            if not reminder_date and not reminder_windows:
                continue

            if 0 <= (end_date - today).days <= 90:
                notices.append(
                    {
                        "branch": branch,
                        "officer": row[4].strip() or "Officer not assigned",
                        "contact": row[5].strip() or "No contact number",
                        "end_date": self.format_date(end_date),
                        "reminder_date": self.format_date(reminder_date),
                        "window": ", ".join(reminder_windows) if reminder_windows else "Expiry alert",
                        "status": self.manual_status_for(row) or self.pending_stage_for(row),
                        "action": "Renew / Escalation or Increase",
                    }
                )

        notices.sort(key=lambda item: self.parse_date(item["end_date"]) or date.max)
        return notices

    def get_main_dashboard_rows(self):
        combined_rows = deepcopy(self.main_dashboard_rows)
        combined_rows.extend(self.build_dashboard_rows_from_expiry())
        return combined_rows

    def get_theme(self):
        return self.ui_settings.get("theme", "dark")

    def set_theme(self, theme):
        self.ui_settings["theme"] = theme
        self.save()

    def get_contract_documents(self):
        raw_value = self.ui_settings.get("contract_documents", "{}")
        try:
            return json.loads(raw_value)
        except json.JSONDecodeError:
            return {}

    def get_dashboard_stage_statuses(self):
        raw_value = self.ui_settings.get("dashboard_stage_statuses", "{}")
        try:
            statuses = json.loads(raw_value)
        except json.JSONDecodeError:
            return {}
        return statuses if isinstance(statuses, dict) else {}

    def dashboard_stage_state_from_text(self, value):
        normalized = str(value).strip().lower()
        if not normalized:
            return {"status": "new", "completed_on": ""}
        if "pending" in normalized or "for action" in normalized:
            return {"status": "pending_action", "completed_on": ""}
        if "progress" in normalized:
            return {"status": "in_progress", "completed_on": ""}
        if (
            "ok" in normalized
            or "complete" in normalized
            or "completed" in normalized
            or "approved" in normalized
            or "done" in normalized
        ):
            return {"status": "complete", "completed_on": ""}
        if "new" in normalized:
            return {"status": "new", "completed_on": ""}
        return {"status": "new", "completed_on": ""}

    def main_dashboard_row_key(self, row):
        return " | ".join(str(row[index]).strip() for index in (0, 1, 2, 3))

    def get_dashboard_stage_state(self, row, column_index):
        row_key = self.main_dashboard_row_key(row)
        statuses = self.get_dashboard_stage_statuses()
        row_statuses = statuses.get(row_key, {})
        state = row_statuses.get(str(column_index), {})
        if isinstance(state, dict) and state.get("status"):
            return {
                "status": str(state.get("status", "new")),
                "completed_on": str(state.get("completed_on", "")).strip(),
            }
        if column_index < len(row):
            return self.dashboard_stage_state_from_text(row[column_index])
        return {"status": "new", "completed_on": ""}

    def set_dashboard_stage_state(self, row, column_index, status, completed_on=""):
        row_key = self.main_dashboard_row_key(row)
        statuses = self.get_dashboard_stage_statuses()
        row_statuses = statuses.get(row_key, {})
        row_statuses[str(column_index)] = {
            "status": str(status),
            "completed_on": str(completed_on).strip(),
        }
        statuses[row_key] = row_statuses
        self.ui_settings["dashboard_stage_statuses"] = json.dumps(statuses)
        self.save()

    def clear_dashboard_stage_states_after(self, row, column_index):
        row_key = self.main_dashboard_row_key(row)
        statuses = self.get_dashboard_stage_statuses()
        row_statuses = statuses.get(row_key, {})
        filtered = {}
        for key, value in row_statuses.items():
            try:
                current_column = int(key)
            except (TypeError, ValueError):
                continue
            if current_column <= column_index:
                filtered[key] = value
        statuses[row_key] = filtered
        self.ui_settings["dashboard_stage_statuses"] = json.dumps(statuses)
        self.save()

    def get_dashboard_remark_overrides(self):
        raw_value = self.ui_settings.get("dashboard_remark_overrides", "{}")
        try:
            overrides = json.loads(raw_value)
        except json.JSONDecodeError:
            return {}
        return overrides if isinstance(overrides, dict) else {}

    def get_dashboard_remark_override(self, row):
        return str(self.get_dashboard_remark_overrides().get(self.main_dashboard_row_key(row), "")).strip()

    def set_dashboard_remark(self, row, remark):
        row_key = self.main_dashboard_row_key(row)
        remark_text = str(remark).strip()

        for index, current_row in enumerate(self.main_dashboard_rows):
            if self.main_dashboard_row_key(current_row) == row_key:
                self.main_dashboard_rows[index][12] = remark_text
                self.save()
                return

        overrides = self.get_dashboard_remark_overrides()
        overrides[row_key] = remark_text
        self.ui_settings["dashboard_remark_overrides"] = json.dumps(overrides)
        self.save()

    def contract_document_key(self, row):
        return " | ".join(
            [
                str(row[0]).strip(),
                str(row[3]).strip(),
                str(row[7]).strip(),
                str(row[8]).strip(),
            ]
        )

    def get_contract_document(self, row):
        return self.get_contract_documents().get(self.contract_document_key(row), "")

    def set_contract_document(self, row, file_path):
        documents = self.get_contract_documents()
        documents[self.contract_document_key(row)] = str(file_path)
        self.ui_settings["contract_documents"] = json.dumps(documents)
        self.save()

    def get_expiry_rows(self):
        return deepcopy(self.expiry_rows)

    def set_expiry_rows(self, rows):
        self.expiry_rows = self._normalize_rows(rows)
        self.save()

    def add_expiry_row(self, row=None):
        new_row = [""] * 13 if row is None else list(row[:13])
        if len(new_row) < 13:
            new_row.extend([""] * (13 - len(new_row)))
        self.expiry_rows.append([str(value) for value in new_row])
        self.save()

    def remove_expiry_row(self, index):
        if 0 <= index < len(self.expiry_rows):
            self.expiry_rows.pop(index)
            self.save()

    def get_legend_rows(self):
        counts = {
            "Done Lease Contracts": 0,
            "For Reminder / Action": 0,
            "Active Contracts": 0,
            "Expired Contracts": 0,
        }
        for row in self.expiry_rows:
            status = self.contract_status_for(row)
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

    def dashboard_summary(self):
        summary = {
            "expired": 0,
            "expiring_90": 0,
            "pending_ho": 0,
            "completed": 0,
        }
        for row in self.expiry_rows:
            if not row[0].strip():
                continue
            status = self.contract_status_for(row)
            if status == "expired":
                summary["expired"] += 1
            if status == "due":
                summary["expiring_90"] += 1
            if self.manual_status_for(row).lower() not in ("done", "approved") and status not in ("blank", "expired"):
                summary["pending_ho"] += 1
            if self.manual_status_for(row).lower() == "done" or status == "done":
                summary["completed"] += 1
        return summary

    def daily_report_rows(self):
        report_rows = []
        for row in self.expiry_rows:
            branch = row[0].strip()
            if not branch:
                continue
            manual_status = self.manual_status_for(row)
            status = self.contract_status_for(row)
            if manual_status.lower() == "done" or status == "done":
                continue
            report_rows.append(
                {
                    "branch": branch,
                    "pending_stage": self.pending_stage_for(row),
                    "status": manual_status or "Pending",
                    "officer": row[4].strip() or "Officer not assigned",
                    "contact": row[5].strip() or "No contact number",
                    "expiry_date": row[8].strip(),
                    "remarks": row[12].strip(),
                }
            )
        report_rows.sort(key=lambda item: item["branch"].lower())
        return report_rows

    def routing_text_for(self, status, label):
        if status == "done":
            return f"COMPLETE {label}"
        if status == "due":
            return f"FOR ACTION {label}"
        if status == "expired":
            return f"EXPIRED {label}"
        if status == "active":
            return f"IN PROGRESS {label}"
        return ""

    def dashboard_title_for(self, row):
        branch = row[0].strip()
        term = row[3].strip()
        from_date = row[7].strip()
        to_date = row[8].strip()
        parts = [f"{branch} LEASE CONTRACT"]
        if term:
            parts.append(term)
        if from_date or to_date:
            parts.append(f"{from_date} TO {to_date}".strip())
        return " | ".join([part for part in parts if part.strip()])

    def dashboard_remark_for(self, row, status):
        override = self.get_dashboard_remark_override(row)
        if override:
            return override
        officer = row[4].strip() or "Officer not assigned"
        contact = row[5].strip() or "No contact"
        reminder = self.reminder_date_for(row)
        remarks = row[12].strip()
        document_path = self.get_contract_document(row)
        remark_parts = [
            f"Officer: {officer}",
            f"Contact: {contact}",
            f"Status: {self.manual_status_for(row) or self.pending_stage_for(row)}",
        ]
        if reminder:
            remark_parts.append(f"Reminder: {reminder}")
        if remarks:
            remark_parts.append(remarks)
        if document_path:
            remark_parts.append("PDF Attached")
        if status == "due":
            remark_parts.append("Action: Renew / Escalation or Increase")
        return " | ".join(remark_parts)

    def build_dashboard_rows_from_expiry(self):
        generated_rows = []
        for row in self.expiry_rows:
            branch = row[0].strip()
            if not branch:
                continue
            if not any(str(cell).strip() for cell in row[1:]):
                continue
            status = self.contract_status_for(row)
            pending_stage = self.pending_stage_for(row)
            stage_cells = self.dashboard_stage_cells_for(status, pending_stage)
            generated_rows.append(
                [
                    row[1].strip() or row[2].strip() or row[7].strip(),
                    self.dashboard_title_for(row),
                    "",
                    row[10].strip(),
                    *stage_cells,
                    self.dashboard_remark_for(row, status),
                ]
            )

        generated_rows.sort(key=lambda current: self.parse_date(current[0]) or date.max)
        return generated_rows


store = LeaseDataStore()
