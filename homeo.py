import sys
import os
import json
from PyQt5 import QtWidgets, QtCore

# Optional libraries
try:
    import pandas as pd
    from rapidfuzz import process
    PANDAS_AVAILABLE = True
except ImportError:
    PANDAS_AVAILABLE = False


class HomeoWindow(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("🏥 Homeopathy Remedy Search")
        self.resize(900, 600)

        self.df = None
        self.excel_file = 'remedies.xlsx'
        self._setup_ui()
        self._load_settings()
        self.ensure_excel_exists()
        self._load_df()

    # ---------------------- UI ----------------------
    def _setup_ui(self):
        layout = QtWidgets.QVBoxLayout(self)

        # Top row: search input and buttons
        top = QtWidgets.QHBoxLayout()
        self.query = QtWidgets.QLineEdit()
        self.query.setPlaceholderText("🔍 Search remedy...")
        self.search_btn = QtWidgets.QPushButton("Search")
        self.load_btn = QtWidgets.QPushButton("Load Excel")
        self.add_btn = QtWidgets.QPushButton("Add New")
        top.addWidget(self.query)
        top.addWidget(self.search_btn)
        top.addWidget(self.load_btn)
        top.addWidget(self.add_btn)
        layout.addLayout(top)

        # Search options
        opts = QtWidgets.QHBoxLayout()
        self.mode = QtWidgets.QComboBox()
        self.mode.addItems(["Contains", "Starts with", "Word prefix", "Fuzzy"])
        self.incremental = QtWidgets.QCheckBox("Incremental")
        opts.addWidget(QtWidgets.QLabel("Search Mode:"))
        opts.addWidget(self.mode)
        opts.addWidget(self.incremental)
        opts.addStretch()
        layout.addLayout(opts)

        # Table
        self.table = QtWidgets.QTableWidget()
        self.table.setColumnCount(2)
        self.table.setHorizontalHeaderLabels(["Common Name", "Latin Name"])
        self.table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(self.table)

        # Status
        self.status = QtWidgets.QLabel("Ready")
        layout.addWidget(self.status)

        # Button actions
        self.search_btn.clicked.connect(self.on_search)
        self.load_btn.clicked.connect(self._load_df)
        self.add_btn.clicked.connect(self.on_add_new)
        self.query.textChanged.connect(self.on_query_changed)

    # ---------------------- SETTINGS ----------------------
    def _load_settings(self):
        if os.path.exists("settings.json"):
            try:
                s = json.load(open("settings.json"))
                self.incremental.setChecked(s.get("incremental", True))
                self.mode.setCurrentText(s.get("mode", "Contains"))
            except:
                pass

    def closeEvent(self, event):
        s = {"incremental": self.incremental.isChecked(), "mode": self.mode.currentText()}
        json.dump(s, open("settings.json", "w"))
        event.accept()

    # ---------------------- EXCEL HANDLING ----------------------
    def ensure_excel_exists(self):
        """Create a default Excel file if missing"""
        if getattr(sys, 'frozen', False):
            base_path = sys._MEIPASS
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))
        self.excel_file_path = os.path.join(base_path, self.excel_file)

        if not os.path.exists(self.excel_file_path):
            if not PANDAS_AVAILABLE:
                self.status.setText("pandas not installed; cannot create Excel.")
                return
            df = pd.DataFrame({
                'common_col': ['Arnica', 'Bryonia', 'Belladonna'],
                'latin_col': ['Arnica montana', 'Bryonia alba', 'Atropa belladonna']
            })
            df.to_excel(self.excel_file_path, index=False, engine="openpyxl")
            self.status.setText("Created default remedies.xlsx")

    def _load_df(self):
        if not PANDAS_AVAILABLE:
            self.status.setText("pandas not installed.")
            return
        if not os.path.exists(self.excel_file_path):
            self.status.setText("Excel file missing.")
            return
        try:
            self.df = pd.read_excel(self.excel_file_path, engine="openpyxl")
            self.df.fillna('', inplace=True)
            self._populate_table(self.df)
            self.status.setText(f"Loaded {len(self.df)} entries.")
        except Exception as e:
            self.status.setText(f"Failed to load Excel: {e}")

    def _populate_table(self, df):
        self.table.setRowCount(len(df))
        for i, (_, row) in enumerate(df.iterrows()):
            self.table.setItem(i, 0, QtWidgets.QTableWidgetItem(str(row.get("common_col", ""))))
            self.table.setItem(i, 1, QtWidgets.QTableWidgetItem(str(row.get("latin_col", ""))))

    # ---------------------- SEARCH ----------------------
    def on_query_changed(self):
        if self.incremental.isChecked():
            self.on_search()

    def on_search(self):
        if not PANDAS_AVAILABLE or self.df is None:
            self.status.setText("No data loaded.")
            return

        q = self.query.text().strip().lower()
        if not q:
            self._populate_table(self.df)
            return

        mode = self.mode.currentText()
        if mode == "Contains":
            f = self.df[self.df.apply(lambda r: q in str(r.common_col).lower() or q in str(r.latin_col).lower(), axis=1)]
        elif mode == "Starts with":
            f = self.df[self.df.apply(lambda r: str(r.common_col).lower().startswith(q) or str(r.latin_col).lower().startswith(q), axis=1)]
        elif mode == "Word prefix":
            f = self.df[self.df.apply(lambda r: any(w.startswith(q) for w in str(r.common_col).lower().split()) or any(w.startswith(q) for w in str(r.latin_col).lower().split()), axis=1)]
        elif mode == "Fuzzy":
            f = self._fuzzy_search(q)
        else:
            f = self.df

        self._populate_table(f)
        self.status.setText(f"Found {len(f)} results for '{q}'")

    def _fuzzy_search(self, q):
        names = list(self.df.common_col) + list(self.df.latin_col)
        matches = process.extract(q, names, limit=20)
        found = set([m[0] for m in matches if m[1] > 50])
        return self.df[self.df.apply(lambda r: r.common_col in found or r.latin_col in found, axis=1)]

    # ---------------------- ADD NEW ----------------------
    def on_add_new(self):
        if not PANDAS_AVAILABLE:
            self.status.setText("pandas not installed.")
            return

        dlg = QtWidgets.QDialog(self)
        dlg.setWindowTitle("Add New Remedy")
        layout = QtWidgets.QFormLayout(dlg)
        common_input = QtWidgets.QLineEdit()
        latin_input = QtWidgets.QLineEdit()
        layout.addRow("Common Name:", common_input)
        layout.addRow("Latin Name:", latin_input)
        btns = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        layout.addWidget(btns)
        btns.accepted.connect(dlg.accept)
        btns.rejected.connect(dlg.reject)

        if dlg.exec_() == QtWidgets.QDialog.Accepted:
            common = common_input.text().strip()
            latin = latin_input.text().strip()
            if not common or not latin:
                QtWidgets.QMessageBox.warning(self, "Error", "Both fields are required.")
                return
            try:
                df = pd.read_excel(self.excel_file_path, engine="openpyxl")
                new_row = {"common_col": common, "latin_col": latin}
                df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
                df.to_excel(self.excel_file_path, index=False, engine="openpyxl")
                self._load_df()
                self.status.setText(f"Added: {common} - {latin}")
            except Exception as e:
                QtWidgets.QMessageBox.critical(self, "Error", f"Failed to add: {e}")


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    w = HomeoWindow()
    w.show()
    sys.exit(app.exec_())
