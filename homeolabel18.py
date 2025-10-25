import sys
import os
import json
import logging
import pandas as pd
from PyQt5 import QtWidgets, QtCore
from PyQt5.QtWidgets import QCompleter, QTableWidgetItem
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
import win32api

# ------------------ Logging ------------------
os.makedirs("records", exist_ok=True)
logging.basicConfig(filename="records/error_log.txt", level=logging.ERROR,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# ------------------ Configuration ------------------
LABEL_WIDTH_MM = 50
LABEL_HEIGHT_MM = 30
PRINTER_NAME = "SNBC TVSE LP 46 NEO BPLE"


class HomeoLabelApp(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("🏥 Homeopathy Label Generator")
        self.resize(950, 700)

        # Paths
        self.records_folder = "records"
        self.excel_file = os.path.join(self.records_folder, 'records.xlsx')
        self.autocomplete_file = os.path.join(self.records_folder, 'autocomplete.json')
        self.remedies_file = 'remedies.xlsx'

        # Label Settings
        self.top_offset = 0.0
        self.font_size_med = 8

        # Load data
        self.load_remedies()
        self.autocomplete_data = self.load_autocomplete()

        # UI
        self.init_ui()

    # ------------------ Load Remedies ------------------
    def load_remedies(self):
        if not os.path.exists(self.remedies_file):
            df = pd.DataFrame({
                'latin_col': ['Arnica montana', 'Bryonia alba', 'Atropa belladonna'],
                'common_col': ['Arnica', 'Bryonia', 'Belladonna']
            })
            df.to_excel(self.remedies_file, index=False, engine="openpyxl")
        self.df_remedies = pd.read_excel(self.remedies_file, engine="openpyxl").fillna('')

    # ------------------ Autocomplete ------------------
    def load_autocomplete(self):
        if os.path.exists(self.autocomplete_file):
            try:
                return json.load(open(self.autocomplete_file))
            except:
                return {}
        return {}

    # ------------------ UI ------------------
    def init_ui(self):
        layout = QtWidgets.QVBoxLayout(self)

        # Search Section
        self.medicine_search = QtWidgets.QLineEdit()
        self.medicine_search.setPlaceholderText("Search medicine (Latin/Common)...")
        self.medicine_search.textChanged.connect(self.update_suggestions)

        self.suggestion_table = QtWidgets.QTableWidget(0, 2)
        self.suggestion_table.setHorizontalHeaderLabels(["Common Name", "Latin Name"])
        self.suggestion_table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.suggestion_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.suggestion_table.cellClicked.connect(self.on_suggestion_clicked)

        search_layout = QtWidgets.QHBoxLayout()
        search_layout.addWidget(self.medicine_search)
        search_layout.addWidget(self.suggestion_table)
        layout.addLayout(search_layout)

        # Add Medicine
        self.add_new_btn = QtWidgets.QPushButton("➕ Add New Medicine")
        self.add_new_btn.clicked.connect(self.add_new_medicine)
        layout.addWidget(self.add_new_btn)

        # Form Fields
        form = QtWidgets.QFormLayout()
        self.potency_input = self.combo_with_autocomplete("potency")
        self.dose_input = self.combo_with_autocomplete("dose")
        self.quantity_input = QtWidgets.QLineEdit()
        self.time_input = self.combo_with_autocomplete("time")
        self.shop_input = self.combo_with_autocomplete("shop")
        self.branch_phone_input = self.combo_with_autocomplete("branch_phone")

        form.addRow("Potency:", self.potency_input)
        form.addRow("Dose:", self.dose_input)
        form.addRow("Quantity:", self.quantity_input)
        form.addRow("Time:", self.time_input)
        form.addRow("Shop Name:", self.shop_input)
        form.addRow("Branch / Phone:", self.branch_phone_input)
        layout.addLayout(form)

        # Preview Frame
        self.preview_frame = QtWidgets.QFrame()
        self.preview_frame.setFrameShape(QtWidgets.QFrame.Box)
        self.preview_frame.setFixedSize(400, 250)
        preview_layout = QtWidgets.QVBoxLayout(self.preview_frame)

        self.preview_labels = [QtWidgets.QLabel("") for _ in range(5)]
        for lbl in self.preview_labels:
            lbl.setAlignment(QtCore.Qt.AlignCenter)
            lbl.setWordWrap(True)
            preview_layout.addWidget(lbl)

        layout.addWidget(self.preview_frame)

        # Top Offset Slider
        offset_layout = QtWidgets.QHBoxLayout()
        offset_layout.addWidget(QtWidgets.QLabel("Top Offset:"))
        self.top_offset_slider = QtWidgets.QSlider(QtCore.Qt.Horizontal)
        self.top_offset_slider.setMinimum(0)
        self.top_offset_slider.setMaximum(20)
        self.top_offset_slider.valueChanged.connect(self.update_top_offset)
        offset_layout.addWidget(self.top_offset_slider)
        layout.addLayout(offset_layout)

        # Buttons
        btn_layout = QtWidgets.QHBoxLayout()
        self.print_btn = QtWidgets.QPushButton("🖨️ Preview Label")
        self.print_btn.clicked.connect(self.print_label)
        self.direct_print_btn = QtWidgets.QPushButton("⚡ Direct Print (LP46 Neo)")
        self.direct_print_btn.clicked.connect(self.print_direct)
        btn_layout.addWidget(self.print_btn)
        btn_layout.addWidget(self.direct_print_btn)
        layout.addLayout(btn_layout)

        # Status
        self.status = QtWidgets.QLabel("Ready")
        layout.addWidget(self.status)

    # ------------------ Helper for Combo Fields ------------------
    def combo_with_autocomplete(self, key):
        combo = QtWidgets.QComboBox()
        combo.setEditable(True)
        data = self.autocomplete_data.get(key, [])
        combo.addItems(data)
        combo.setCompleter(QCompleter(data))
        combo.currentTextChanged.connect(self.update_preview)
        return combo

    # ------------------ Slider ------------------
    def update_top_offset(self):
        self.top_offset = self.top_offset_slider.value() / 100
        self.update_preview()

    # ------------------ Suggestions ------------------
    def update_suggestions(self):
        text = self.medicine_search.text().lower()
        self.suggestion_table.setRowCount(0)
        for _, row in self.df_remedies.iterrows():
            if text in row['common_col'].lower() or text in row['latin_col'].lower():
                r = self.suggestion_table.rowCount()
                self.suggestion_table.insertRow(r)
                self.suggestion_table.setItem(r, 0, QTableWidgetItem(row['common_col']))
                self.suggestion_table.setItem(r, 1, QTableWidgetItem(row['latin_col']))

    def on_suggestion_clicked(self, row, _):
        med = self.suggestion_table.item(row, 0).text()
        self.medicine_search.setText(med)
        self.update_preview()

    def add_new_medicine(self):
        new, ok = QtWidgets.QInputDialog.getText(self, "Add Medicine", "Enter new medicine:")
        if ok and new.strip():
            self.df_remedies.loc[len(self.df_remedies)] = [new, new]
            self.df_remedies.to_excel(self.remedies_file, index=False, engine="openpyxl")
            self.update_suggestions()

    # ------------------ Preview ------------------
    def update_preview(self):
        med = self.medicine_search.text().upper()
        pot = self.potency_input.currentText().upper()
        dose = self.dose_input.currentText()
        qty = self.quantity_input.text()
        time = self.time_input.currentText()
        shop = self.shop_input.currentText().upper()
        branch = self.branch_phone_input.currentText().upper()

        wrapped_time = " ".join(time.split())
        wrapped_time_lines = "\n".join(self.wrap_text(wrapped_time, 25))

        texts = [
            f"{med} {pot}",
            f"{qty}",
            f"{dose}  {wrapped_time_lines}",
            f"{shop}",
            f"{branch}"
        ]
        for lbl, txt in zip(self.preview_labels, texts):
            lbl.setText(txt)

    def wrap_text(self, text, width):
        return [text[i:i+width] for i in range(0, len(text), width)]

    # ------------------ Print PDF ------------------
    def print_label(self):
        pdf = os.path.join(self.records_folder, "label.pdf")
        self.generate_pdf(pdf)
        os.startfile(pdf)
        self.status.setText("✅ Preview generated successfully.")

    def print_direct(self):
        try:
            pdf = os.path.join(self.records_folder, "label.pdf")
            self.generate_pdf(pdf)
            win32api.ShellExecute(0, "printto", pdf, f'"{PRINTER_NAME}"', ".", 0)
            self.status.setText("🖨️ Sent to LP46 Neo successfully.")
        except Exception as e:
            logging.error(f"Direct print failed: {e}")
            QtWidgets.QMessageBox.critical(self, "Print Error", f"Direct print failed:\n{e}")
            self.status.setText(f"⚠️ Print failed: {e}")

    # ------------------ PDF Generation ------------------
    def generate_pdf(self, path):
        width = LABEL_WIDTH_MM / 25.4 * inch
        height = LABEL_HEIGHT_MM / 25.4 * inch

        c = canvas.Canvas(path, pagesize=(width, height))
        c.setLineWidth(1)
        c.rect(0.03 * inch, 0.03 * inch, width - 0.06 * inch, height - 0.06 * inch)

        y = height - (0.15 * inch + self.top_offset * inch)
        c.setFont("Helvetica-Bold", self.font_size_med)
        c.drawCentredString(width / 2, y, f"{self.medicine_search.text().upper()} {self.potency_input.currentText().upper()}")

        y -= 0.14 * inch
        c.setFont("Helvetica", 8)
        c.drawCentredString(width / 2, y, self.quantity_input.text())

        y -= 0.14 * inch
        dose = self.dose_input.currentText()
        time = self.time_input.currentText()
        wrapped = self.wrap_text(f"{dose} {time}", 25)
        for line in wrapped:
            c.drawCentredString(width / 2, y, line)
            y -= 0.12 * inch

        c.setFont("Helvetica-Bold", 7)
        c.drawCentredString(width / 2, y, self.shop_input.currentText().upper())
        y -= 0.12 * inch
        c.drawCentredString(width / 2, y, self.branch_phone_input.currentText().upper())
        c.save()

        self.save_record()

    # ------------------ Save Record ------------------
    def save_record(self):
        data = {
            "Medicine": self.medicine_search.text(),
            "Potency": self.potency_input.currentText(),
            "Dose": self.dose_input.currentText(),
            "Quantity": self.quantity_input.text(),
            "Time": self.time_input.currentText(),
            "Shop": self.shop_input.currentText(),
            "Branch/Phone": self.branch_phone_input.currentText()
        }

        if os.path.exists(self.excel_file):
            df = pd.read_excel(self.excel_file, engine="openpyxl")
            df = pd.concat([df, pd.DataFrame([data])], ignore_index=True)
        else:
            df = pd.DataFrame([data])
        df.to_excel(self.excel_file, index=False, engine="openpyxl")


# ------------------ Run App ------------------
if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    w = HomeoLabelApp()
    w.show()
    sys.exit(app.exec_())
