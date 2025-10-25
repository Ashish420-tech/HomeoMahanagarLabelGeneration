import sys
import os
import json
import logging
from datetime import datetime
import pandas as pd
from PyQt5 import QtWidgets, QtCore
from PyQt5.QtWidgets import QCompleter
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
import win32api

# --- Logging setup ---
logging.basicConfig(filename="error_log.txt", level=logging.ERROR, 
                    format="%(asctime)s - %(levelname)s - %(message)s")

class HomeoLabelApp(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("🏥 Homeopathy Label App")
        self.resize(950, 600)

        # Files
        self.excel_file = "remedies.xlsx"
        self.records_file = "records.xlsx"
        self.autocomplete_file = "autocomplete.json"
        self.df_remedies = None

        # Load autocomplete data
        if os.path.exists(self.autocomplete_file):
            try:
                self.autocomplete_data = json.load(open(self.autocomplete_file))
            except:
                self.autocomplete_data = {}
        else:
            self.autocomplete_data = {}

        # UI setup
        self._setup_ui()
        self.ensure_excel_exists()
        self.load_remedies()
        self.setup_completers()

    # ---------------------- UI ----------------------
    def _setup_ui(self):
        layout = QtWidgets.QVBoxLayout(self)

        # Search & Load buttons
        top = QtWidgets.QHBoxLayout()
        self.query = QtWidgets.QLineEdit()
        self.query.setPlaceholderText("🔍 Search remedy...")
        self.search_btn = QtWidgets.QPushButton("Search")
        self.load_btn = QtWidgets.QPushButton("Reload Excel")
        top.addWidget(self.query)
        top.addWidget(self.search_btn)
        top.addWidget(self.load_btn)
        layout.addLayout(top)

        # Table for remedies
        self.table = QtWidgets.QTableWidget()
        self.table.setColumnCount(2)
        self.table.setHorizontalHeaderLabels(["Common Name", "Latin Name"])
        self.table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(self.table)

        # Input fields
        form = QtWidgets.QFormLayout()
        self.potency_input = QtWidgets.QLineEdit()
        self.dose_input = QtWidgets.QSpinBox()
        self.dose_input.setRange(1,4)
        self.quantity_input = QtWidgets.QLineEdit()
        self.shop_input = QtWidgets.QLineEdit()
        form.addRow("Potency:", self.potency_input)
        form.addRow("Dose (number):", self.dose_input)
        form.addRow("Quantity:", self.quantity_input)
        form.addRow("Shop Name:", self.shop_input)
        layout.addLayout(form)

        # Dynamic times input
        self.time_inputs = []
        self.time_layout = QtWidgets.QHBoxLayout()
        layout.addLayout(self.time_layout)
        self.dose_input.valueChanged.connect(self.update_time_inputs)
        self.update_time_inputs()

        # Print Button
        self.print_btn = QtWidgets.QPushButton("Print Label")
        layout.addWidget(self.print_btn)

        # Status
        self.status = QtWidgets.QLabel("Ready")
        layout.addWidget(self.status)

        # Connect signals
        self.search_btn.clicked.connect(self.search_medicine)
        self.load_btn.clicked.connect(self.load_remedies)
        self.print_btn.clicked.connect(self.print_label)

    # ---------------------- Dynamic Time Inputs ----------------------
    def update_time_inputs(self):
        # Clear old inputs
        for w in self.time_inputs:
            for widget in w:
                self.time_layout.removeWidget(widget)
                widget.deleteLater()
        self.time_inputs = []

        # Add inputs based on dose number
        for i in range(self.dose_input.value()):
            hour_spin = QtWidgets.QSpinBox()
            hour_spin.setRange(0,23)
            hour_spin.setPrefix(f"Time {i+1}: ")
            self.time_layout.addWidget(hour_spin)
            self.time_inputs.append((hour_spin, None, None))  # placeholder

    # ---------------------- Autocomplete ----------------------
    def setup_completers(self):
        # Potency
        pot_list = self.autocomplete_data.get("potency", [])
        self.potency_input.setCompleter(QCompleter(pot_list))

        # Quantity
        qty_list = self.autocomplete_data.get("quantity", [])
        self.quantity_input.setCompleter(QCompleter(qty_list))

        # Shop
        shop_list = self.autocomplete_data.get("shop", [])
        self.shop_input.setCompleter(QCompleter(shop_list))

    # ---------------------- Excel Handling ----------------------
    def ensure_excel_exists(self):
        if not os.path.exists(self.excel_file):
            df = pd.DataFrame({
                "latin_col": ["Arnica montana", "Bryonia alba", "Atropa belladonna"],
                "common_col": ["Arnica", "Bryonia", "Belladonna"]
            })
            df.to_excel(self.excel_file, index=False, engine="openpyxl")

    def load_remedies(self):
        try:
            self.df_remedies = pd.read_excel(self.excel_file, engine="openpyxl")
            self.df_remedies.fillna('', inplace=True)
            self.populate_table(self.df_remedies)
            self.status.setText(f"Loaded {len(self.df_remedies)} remedies.")
        except Exception as e:
            logging.error("Failed to load remedies: %s", str(e))
            self.status.setText(f"Failed to load remedies: {e}")

    # ---------------------- Table ----------------------
    def populate_table(self, df):
        self.table.setRowCount(len(df))
        for i, (_, row) in enumerate(df.iterrows()):
            self.table.setItem(i, 0, QtWidgets.QTableWidgetItem(str(row.get("common_col", ""))))
            self.table.setItem(i, 1, QtWidgets.QTableWidgetItem(str(row.get("latin_col", ""))))

    def search_medicine(self):
        if self.df_remedies is None:
            return
        q = self.query.text().strip().lower()
        if not q:
            self.populate_table(self.df_remedies)
            return
        f = self.df_remedies[self.df_remedies.apply(
            lambda r: q in str(r.common_col).lower() or q in str(r.latin_col).lower(), axis=1)]
        self.populate_table(f)
        self.status.setText(f"Found {len(f)} results for '{q}'")

    # ---------------------- Print Label ----------------------
    def print_label(self):
        try:
            row = self.table.currentRow()
            if row < 0:
                QtWidgets.QMessageBox.warning(self, "No Medicine", "Please select a medicine from the table.")
                return

            # --- Get data ---
            common_name = self.table.item(row, 0).text().upper()
            potency = self.potency_input.text().strip().upper()
            dose = str(self.dose_input.value()).strip()
            quantity = self.quantity_input.text().strip()
            shop = self.shop_input.text().strip().upper()

            times = [str(hour.value()) for hour, _, _ in self.time_inputs]
            times_str = "-".join(times)

            # --- Create PDF Label ---
            label_file = "label.pdf"
            width, height = 2*inch, 1*inch
            c = canvas.Canvas(label_file, pagesize=(width, height))

            # --- Cutting Border ---
            border_margin = 2
            c.setLineWidth(0.8)
            c.rect(border_margin, border_margin, width-2*border_margin, height-2*border_margin)

            # --- Medicine Name ---
            c.setFont("Helvetica-Bold", 8)
            c.drawCentredString(width/2, 0.90*inch, common_name)

            # --- Potency and Dose ---
            c.setFont("Helvetica", 7)
            c.drawString(0.1*inch, 0.75*inch, potency)
            c.drawString(1.1*inch, 0.75*inch, dose)

            # --- Quantity ---
            c.drawString(0.1*inch, 0.60*inch, f"Qty: {quantity}")

            # --- Shop section ---
            c.setFont("Helvetica-Bold", 6)
            c.drawString(0.1*inch, 0.48*inch, "SHOP:")
            c.setFont("Helvetica", 6)
            c.drawString(0.5*inch, 0.48*inch, shop)

            # --- Times ---
            c.setFont("Helvetica", 6)
            c.drawString(0.1*inch, 0.35*inch, f"Time: {times_str}")

            c.showPage()
            c.save()

            # Open PDF
            if os.path.exists(label_file):
                win32api.ShellExecute(0, "open", label_file, None, ".", 1)

            # --- Save Record ---
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            record = {
                "Timestamp": timestamp,
                "Medicine": common_name,
                "Potency": potency,
                "Dose": dose,
                "Quantity": quantity,
                "Times": times_str,
                "Shop": shop
            }
            if os.path.exists(self.records_file):
                df_records = pd.read_excel(self.records_file)
                df_records = pd.concat([df_records, pd.DataFrame([record])], ignore_index=True)
            else:
                df_records = pd.DataFrame([record])
            df_records.to_excel(self.records_file, index=False)

            # --- Update autocomplete ---
            def update_autocomplete(key, value):
                if not value: return
                lst = self.autocomplete_data.get(key, [])
                if value not in lst: lst.append(value)
                self.autocomplete_data[key] = lst

            update_autocomplete("potency", potency)
            update_autocomplete("quantity", quantity)
            update_autocomplete("shop", shop)
            with open(self.autocomplete_file, "w") as f:
                json.dump(self.autocomplete_data, f)

            # Refresh completers
            self.setup_completers()
            self.status.setText("Label created and opened.")

        except Exception as e:
            logging.error("Print/Save failed: %s", str(e))
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed: {e}")
            self.status.setText(f"Error: {e}")

# ---------------------- Run App ----------------------
if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    w = HomeoLabelApp()
    w.show()
    sys.exit(app.exec_())
