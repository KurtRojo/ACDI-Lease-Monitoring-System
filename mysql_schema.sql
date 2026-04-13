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