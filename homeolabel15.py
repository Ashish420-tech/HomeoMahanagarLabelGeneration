import sys
import os
import json
import logging
import pandas as pd
from PyQt5 import QtWidgets, QtCore
from PyQt5.QtWidgets import QCompleter, QTableWidgetItem
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import inch

# ------------------ Logging ------------------
os.makedirs("records", exist_ok=True)
logging.basicConfig(filename="records/error_log.txt", level=logging.ERROR,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# ------------------ Homeopathy Label App ------------------
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

        # Load remedies & autocomplete
        self.df_remedies = None
        self.load_remedies()
        self.autocomplete_data = self.load_autocomplete()

        # Label defaults
        self.top_offset = 0.1  # inch
        self.font_size_med = 8

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
        try:
            self.df_remedies = pd.read_excel(self.remedies_file, engine="openpyxl")
            self.df_remedies.fillna('', inplace=True)
        except Exception as e:
            logging.error(f"Failed to load remedies.xlsx: {e}")
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to load remedies.xlsx:\n{e}")

    # ------------------ Load autocomplete ------------------
    def load_autocomplete(self):
        if os.path.exists(self.autocomplete_file):
            try:
                return json.load(open(self.autocomplete_file))
            except:
                return {}
        return {}

    # ------------------ UI Setup ------------------
    def init_ui(self):
        layout = QtWidgets.QVBoxLayout(self)

        # Top: Selected Medicine
        self.selected_medicine_label = QtWidgets.QLabel("MEDICINE: ")
        self.selected_medicine_label.setStyleSheet("font-weight:bold; font-size:12pt;")
        layout.addWidget(self.selected_medicine_label)

        # Search + Suggestion Table
        search_layout = QtWidgets.QHBoxLayout()
        layout.addLayout(search_layout)

        self.medicine_search = QtWidgets.QLineEdit()
        self.medicine_search.setPlaceholderText("Type medicine name (Latin or Common)")
        self.medicine_search.textChanged.connect(self.update_suggestions)
        search_layout.addWidget(self.medicine_search)

        self.suggestion_table = QtWidgets.QTableWidget()
        self.suggestion_table.setColumnCount(2)
        self.suggestion_table.setHorizontalHeaderLabels(["Common Name", "Latin Name"])
        self.suggestion_table.setMinimumWidth(400)
        self.suggestion_table.setMinimumHeight(200)
        self.suggestion_table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.suggestion_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.suggestion_table.cellClicked.connect(self.on_suggestion_clicked)
        search_layout.addWidget(self.suggestion_table)

        # Add New Medicine Button
        self.add_new_btn = QtWidgets.QPushButton("Add New Medicine")
        self.add_new_btn.clicked.connect(self.add_new_medicine)
        layout.addWidget(self.add_new_btn)

        # Form Inputs
        form = QtWidgets.QFormLayout()
        layout.addLayout(form)

        # Potency
        self.potency_input = QtWidgets.QComboBox()
        self.potency_input.setEditable(True)
        pot_list = self.autocomplete_data.get("potency", [])
        self.potency_input.addItems(pot_list)
        self.potency_input.setCompleter(QCompleter(pot_list))
        self.potency_input.currentTextChanged.connect(self.update_preview)
        form.addRow("Potency:", self.potency_input)

        # Dose
        self.dose_input = QtWidgets.QComboBox()
        self.dose_input.setEditable(True)
        dose_list = self.autocomplete_data.get("dose", [])
        self.dose_input.addItems(dose_list)
        self.dose_input.setCompleter(QCompleter(dose_list))
        self.dose_input.currentTextChanged.connect(self.update_preview)
        form.addRow("Dose:", self.dose_input)

        # Quantity
        self.quantity_input = QtWidgets.QLineEdit()
        self.quantity_input.textChanged.connect(self.update_preview)
        form.addRow("Quantity:", self.quantity_input)

        # Time
        self.time_input = QtWidgets.QComboBox()
        self.time_input.setEditable(True)
        time_list = self.autocomplete_data.get("time", [])
        self.time_input.addItems(time_list)
        self.time_input.setCompleter(QCompleter(time_list))
        self.time_input.currentTextChanged.connect(self.update_preview)
        form.addRow("Time:", self.time_input)

        # Shop
        self.shop_input = QtWidgets.QComboBox()
        self.shop_input.setEditable(True)
        shop_list = self.autocomplete_data.get("shop", [])
        self.shop_input.addItems(shop_list)
        self.shop_input.setCompleter(QCompleter(shop_list))
        self.shop_input.currentTextChanged.connect(self.update_preview)
        form.addRow("Shop Name:", self.shop_input)

        # Branch/Phone
        self.branch_phone_input = QtWidgets.QComboBox()
        self.branch_phone_input.setEditable(True)
        branch_list = self.autocomplete_data.get("branch_phone", [])
        self.branch_phone_input.addItems(branch_list)
        self.branch_phone_input.setCompleter(QCompleter(branch_list))
        self.branch_phone_input.currentTextChanged.connect(self.update_preview)
        form.addRow("Branch/Phone:", self.branch_phone_input)

        # Label Preview
        self.preview_frame = QtWidgets.QFrame()
        self.preview_frame.setFrameShape(QtWidgets.QFrame.Box)
        self.preview_frame.setFixedSize(400, 180)
        preview_layout = QtWidgets.QVBoxLayout(self.preview_frame)
        self.preview_line1 = QtWidgets.QLabel("")
        self.preview_line1.setAlignment(QtCore.Qt.AlignCenter)
        self.preview_line2 = QtWidgets.QLabel("")
        self.preview_line2.setAlignment(QtCore.Qt.AlignCenter)
        self.preview_line3 = QtWidgets.QLabel("")
        self.preview_line3.setAlignment(QtCore.Qt.AlignCenter)
        self.preview_line4 = QtWidgets.QLabel("")
        self.preview_line4.setAlignment(QtCore.Qt.AlignCenter)
        self.preview_line5 = QtWidgets.QLabel("")
        self.preview_line5.setAlignment(QtCore.Qt.AlignCenter)
        for lbl in [self.preview_line1, self.preview_line2, self.preview_line3, self.preview_line4, self.preview_line5]:
            preview_layout.addWidget(lbl)

        # Top Offset Slider
        settings_group = QtWidgets.QGroupBox("Label Settings")
        settings_layout = QtWidgets.QHBoxLayout(settings_group)
        self.top_offset_slider = QtWidgets.QSlider(QtCore.Qt.Horizontal)
        self.top_offset_slider.setRange(0, 50)
        self.top_offset_slider.setValue(int(self.top_offset*100))
        self.top_offset_slider.valueChanged.connect(self.update_label_settings)
        settings_layout.addWidget(QtWidgets.QLabel("Top Offset:"))
        settings_layout.addWidget(self.top_offset_slider)
        layout.addWidget(settings_group)

        # Preview + Print Button
        preview_btn_layout = QtWidgets.QHBoxLayout()
        preview_btn_layout.addWidget(self.preview_frame)
        self.print_btn = QtWidgets.QPushButton("Print Label")
        self.print_btn.clicked.connect(self.print_label)
        preview_btn_layout.addWidget(self.print_btn)
        layout.addLayout(preview_btn_layout)

        # Status
        self.status = QtWidgets.QLabel("Ready")
        layout.addWidget(self.status)

    # ------------------ Update Label Settings ------------------
    def update_label_settings(self):
        self.top_offset = self.top_offset_slider.value()/100
        self.update_preview()

    # ------------------ Suggestions ------------------
    def update_suggestions(self):
        text = self.medicine_search.text().lower()
        self.suggestion_table.setRowCount(0)
        for _, row in self.df_remedies.iterrows():
            common = str(row['common_col'])
            latin = str(row['latin_col'])
            if text in common.lower() or text in latin.lower():
                row_idx = self.suggestion_table.rowCount()
                self.suggestion_table.insertRow(row_idx)
                self.suggestion_table.setItem(row_idx, 0, QTableWidgetItem(common))
                self.suggestion_table.setItem(row_idx, 1, QTableWidgetItem(latin))

    def on_suggestion_clicked(self, row, column):
        item = self.suggestion_table.item(row, column)
        if item:
            self.medicine_search.setText(item.text())
            self.update_selected_medicine()

    def add_new_medicine(self):
        new_name, ok = QtWidgets.QInputDialog.getText(self, "Add New Medicine", "Enter medicine name:")
        if ok and new_name.strip():
            self.save_new_medicine(new_name.strip())
            self.medicine_search.setText(new_name.strip())
            self.update_suggestions()

    def update_selected_medicine(self):
        med_name = self.medicine_search.text().upper()
        self.selected_medicine_label.setText(f"MEDICINE: {med_name}")
        self.update_preview()

    def update_preview(self):
        med_name = self.medicine_search.text().upper()
        potency = self.potency_input.currentText().upper()
        dose = self.dose_input.currentText()
        quantity = self.quantity_input.text()
        time_val = self.time_input.currentText()
        shop = self.shop_input.currentText().upper()
        branch_phone = self.branch_phone_input.currentText().upper()
        self.preview_line1.setText(f"{med_name} {potency}")
        self.preview_line2.setText(f"{quantity}")
        self.preview_line3.setText(f"{dose}   {time_val}")
        self.preview_line4.setText(f"{shop}")
        self.preview_line5.setText(f"{branch_phone}")

    def save_new_medicine(self, med_name):
        med_name = med_name.strip()
        exists = ((self.df_remedies['common_col'].str.lower() == med_name.lower()) |
                  (self.df_remedies['latin_col'].str.lower() == med_name.lower())).any()
        if not exists:
            new_row = {'common_col': med_name, 'latin_col': med_name}
            self.df_remedies = pd.concat([self.df_remedies, pd.DataFrame([new_row])], ignore_index=True)
            self.df_remedies.to_excel(self.remedies_file, index=False, engine='openpyxl')
            logging.info(f"New medicine added: {med_name}")

    # ------------------ Print Label ------------------
    def print_label(self):
        try:
            # Collect fields
            med_name = self.medicine_search.text().upper()
            potency = self.potency_input.currentText().upper()
            dose = self.dose_input.currentText()
            quantity = self.quantity_input.text()
            time_val = self.time_input.currentText()
            shop = self.shop_input.currentText().upper()
            branch_phone = self.branch_phone_input.currentText().upper()
            if not med_name:
                QtWidgets.QMessageBox.warning(self, "Error", "Medicine name is required")
                return

            # Save record Excel
            record = {"Medicine": med_name, "Potency": potency, "Dose": dose,
                      "Quantity": quantity, "Time": time_val, "Shop": shop,
                      "Branch/Phone": branch_phone}
            if os.path.exists(self.excel_file):
                df_records = pd.read_excel(self.excel_file, engine="openpyxl")
                df_records = pd.concat([df_records, pd.DataFrame([record])], ignore_index=True)
            else:
                df_records = pd.DataFrame([record])
            df_records.to_excel(self.excel_file, index=False, engine="openpyxl")

            # Update autocomplete
            for field, value in [("potency", potency), ("dose", dose), ("time", time_val),
                                 ("shop", shop), ("branch_phone", branch_phone)]:
                if value:
                    lst = self.autocomplete_data.setdefault(field, [])
                    if value not in lst:
                        lst.append(value)
            with open(self.autocomplete_file, "w") as f:
                json.dump(self.autocomplete_data, f)

            # Generate PDF
            pdf_file = os.path.join(self.records_folder, "label.pdf")
            c = canvas.Canvas(pdf_file, pagesize=(1.89*inch, 1.10*inch))  # 48x28 mm
            c.setLineWidth(1)
            c.rect(0.05*inch, 0.05*inch, 1.79*inch, 1.0*inch-0.05*inch)
            y = 1.10*inch - self.top_offset*inch
            c.setFont("Helvetica-Bold", self.font_size_med)
            c.drawCentredString(0.945*inch, y, f"{med_name} {potency}")
            y -= 0.15*inch
            c.setFont("Helvetica", 8)
            c.drawCentredString(0.945*inch, y, f"{quantity}")
            y -= 0.15*inch
            c.drawCentredString(0.945*inch, y, f"{dose}   {time_val}")
            y -= 0.15*inch
            c.setFont("Helvetica-Bold", 7)
            c.drawCentredString(0.945*inch, y, f"{shop}")
            y -= 0.12*inch
            c.drawCentredString(0.945*inch, y, f"{branch_phone}")
            c.save()

            # Open PDF
            if sys.platform == "win32":
                os.startfile(pdf_file)
            else:
                os.system(f"open {pdf_file}")

            self.status.setText("Label generated, PDF opened, and record saved.")

        except Exception as e:
            logging.error(f"Print failed: {e}")
            QtWidgets.QMessageBox.critical(self, "Error", f"Print failed: {e}")
            self.status.setText(f"Error: {e}")

# ------------------ Run App ------------------
if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    w = HomeoLabelApp()
    w.show()
    sys.exit(app.exec_())
