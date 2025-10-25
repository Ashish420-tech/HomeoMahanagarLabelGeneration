import sys
import os
import json
import logging
from PyQt5 import QtWidgets, QtCore
from PyQt5.QtWidgets import QCompleter, QTableWidgetItem
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm

# ------------------ Logging ------------------
logging.basicConfig(filename="error_log.txt", level=logging.ERROR,
                    format='%(asctime)s - %(levelname)s - %(message)s')

# ------------------ Homeopathy Label App ------------------
class HomeoLabelApp(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("🏥 Homeopathy Label Generator")
        self.resize(900, 700)

        # Paths
        self.excel_file = 'remedies.xlsx'
        self.records_file = 'records.xlsx'
        self.autocomplete_file = 'autocomplete.json'

        # Load remedies
        self.df_remedies = None
        self.load_remedies()

        # Autocomplete data
        self.autocomplete_data = self.load_autocomplete()

        # UI
        self.init_ui()

    # ------------------ Load Remedies ------------------
    def load_remedies(self):
        if not os.path.exists(self.excel_file):
            df = pd.DataFrame({
                'latin_col': ['Arnica montana', 'Bryonia alba', 'Atropa belladonna'],
                'common_col': ['Arnica', 'Bryonia', 'Belladonna']
            })
            df.to_excel(self.excel_file, index=False, engine="openpyxl")
        try:
            self.df_remedies = pd.read_excel(self.excel_file, engine="openpyxl")
            self.df_remedies.fillna('', inplace=True)
        except Exception as e:
            logging.error("Failed to load remedies.xlsx: %s", str(e))
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed to load remedies.xlsx:\n{e}")

    # ------------------ Autocomplete ------------------
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

        # ----- Top: Selected Medicine Display -----
        self.selected_medicine_label = QtWidgets.QLabel("MEDICINE: ")
        self.selected_medicine_label.setStyleSheet("font-weight:bold; font-size:10pt;")
        layout.addWidget(self.selected_medicine_label)

        # ----- Horizontal layout: Search + Suggestions -----
        search_layout = QtWidgets.QHBoxLayout()
        layout.addLayout(search_layout)

        # Medicine Search Input
        self.medicine_search = QtWidgets.QLineEdit()
        self.medicine_search.setPlaceholderText("Type medicine name (Latin or Common)")
        self.medicine_search.textChanged.connect(self.update_suggestions)
        search_layout.addWidget(self.medicine_search)

        # Suggestion Table (2 columns: Common, Latin)
        self.suggestion_table = QtWidgets.QTableWidget()
        self.suggestion_table.setColumnCount(2)
        self.suggestion_table.setHorizontalHeaderLabels(["Common Name", "Latin Name"])
        self.suggestion_table.setMinimumWidth(400)
        self.suggestion_table.setMinimumHeight(200)
        self.suggestion_table.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.suggestion_table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectRows)
        self.suggestion_table.cellClicked.connect(self.on_suggestion_clicked)
        search_layout.addWidget(self.suggestion_table)

        # ----- Add New Medicine Button -----
        self.add_new_btn = QtWidgets.QPushButton("Add New Medicine")
        self.add_new_btn.clicked.connect(self.add_new_medicine)
        layout.addWidget(self.add_new_btn)

        # ----- Form Inputs -----
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
        self.dose_input = QtWidgets.QLineEdit()
        self.dose_input.setPlaceholderText("Enter dose")
        self.dose_input.textChanged.connect(self.update_preview)
        form.addRow("Dose:", self.dose_input)

        # Quantity
        self.quantity_input = QtWidgets.QLineEdit()
        self.quantity_input.setPlaceholderText("Enter quantity")
        self.quantity_input.textChanged.connect(self.update_preview)
        form.addRow("Quantity:", self.quantity_input)

        # Time
        self.time_input = QtWidgets.QLineEdit()
        self.time_input.setPlaceholderText("Enter times like 8-12-4")
        self.time_input.textChanged.connect(self.update_preview)
        form.addRow("Time:", self.time_input)

        # Shop Name
        self.shop_input = QtWidgets.QLineEdit()
        self.shop_input.setPlaceholderText("Enter Shop Name")
        self.shop_input.textChanged.connect(self.update_preview)
        form.addRow("Shop Name:", self.shop_input)

        # Branch/Phone
        self.branch_phone_input = QtWidgets.QLineEdit()
        self.branch_phone_input.setPlaceholderText("Enter Branch and Phone")
        self.branch_phone_input.textChanged.connect(self.update_preview)
        form.addRow("Branch/Phone:", self.branch_phone_input)

        # Label Adjustment Inputs
        self.top_margin_input = QtWidgets.QSpinBox()
        self.top_margin_input.setRange(0, 20)
        self.top_margin_input.setValue(4)
        self.top_margin_input.setSuffix(" mm")
        self.top_margin_input.valueChanged.connect(self.update_preview)
        form.addRow("Top Margin:", self.top_margin_input)

        self.line_spacing_input = QtWidgets.QDoubleSpinBox()
        self.line_spacing_input.setRange(1, 10)
        self.line_spacing_input.setValue(4)
        self.line_spacing_input.setSuffix(" mm")
        self.line_spacing_input.valueChanged.connect(self.update_preview)
        form.addRow("Line Spacing:", self.line_spacing_input)

        # ----- Label Preview -----
        self.preview_frame = QtWidgets.QFrame()
        self.preview_frame.setFrameShape(QtWidgets.QFrame.Box)
        self.preview_frame.setFixedSize(400, 180)
        preview_layout = QtWidgets.QVBoxLayout(self.preview_frame)

        self.preview_line1 = QtWidgets.QLabel("")  # Medicine + Potency
        self.preview_line1.setStyleSheet("font-weight:bold; font-size:10pt;")
        self.preview_line1.setAlignment(QtCore.Qt.AlignCenter)
        preview_layout.addWidget(self.preview_line1)

        self.preview_line2 = QtWidgets.QLabel("")  # Quantity
        self.preview_line2.setAlignment(QtCore.Qt.AlignCenter)
        preview_layout.addWidget(self.preview_line2)

        self.preview_line3 = QtWidgets.QLabel("")  # Dose + Time
        self.preview_line3.setAlignment(QtCore.Qt.AlignCenter)
        preview_layout.addWidget(self.preview_line3)

        self.preview_line4 = QtWidgets.QLabel("")  # Shop
        self.preview_line4.setAlignment(QtCore.Qt.AlignCenter)
        preview_layout.addWidget(self.preview_line4)

        self.preview_line5 = QtWidgets.QLabel("")  # Branch/Phone
        self.preview_line5.setAlignment(QtCore.Qt.AlignCenter)
        preview_layout.addWidget(self.preview_line5)

        layout.addWidget(self.preview_frame)

        # Generate Label Button
        self.print_btn = QtWidgets.QPushButton("Generate Label")
        self.print_btn.clicked.connect(self.print_label)
        layout.addWidget(self.print_btn)

        # Status
        self.status = QtWidgets.QLabel("Ready")
        layout.addWidget(self.status)

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

    # ------------------ Add New Medicine ------------------
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

    # ------------------ Update Preview -----
    def update_preview(self):
        med_name = self.medicine_search.text().upper()
        potency = self.potency_input.currentText().upper()
        dose = self.dose_input.text()
        quantity = self.quantity_input.text()
        time_val = self.time_input.text()
        shop = self.shop_input.text().upper()
        branch_phone = self.branch_phone_input.text().upper()

        self.preview_line1.setText(f"{med_name} {potency}")
        self.preview_line2.setText(f"{quantity}")
        self.preview_line3.setText(f"{dose}   {time_val}")  # removed "pills"
        self.preview_line4.setText(f"{shop}")
        self.preview_line5.setText(f"{branch_phone}")

    # ------------------ Save New Medicine -----
    def save_new_medicine(self, med_name):
        med_name = med_name.strip()
        exists = ((self.df_remedies['common_col'].str.lower() == med_name.lower()) |
                  (self.df_remedies['latin_col'].str.lower() == med_name.lower())).any()
        if not exists:
            new_row = {'common_col': med_name, 'latin_col': med_name}
            self.df_remedies = pd.concat([self.df_remedies, pd.DataFrame([new_row])], ignore_index=True)
            self.df_remedies.to_excel(self.excel_file, index=False, engine='openpyxl')
            logging.info(f"New medicine added: {med_name}")

    # ------------------ Generate PDF -----
    def print_label(self):
        try:
            med_name = self.medicine_search.text().upper()
            potency = self.potency_input.currentText().upper()
            dose = self.dose_input.text()
            quantity = self.quantity_input.text()
            time_val = self.time_input.text()
            shop = self.shop_input.text().upper()
            branch_phone = self.branch_phone_input.text().upper()

            top_margin = self.top_margin_input.value() * mm
            spacing = self.line_spacing_input.value() * mm

            label_width = 45 * mm
            label_height = 25 * mm
            pdf_file = "label_45x25.pdf"

            # Save autocomplete for future suggestions
            for field, value in [("medicine", med_name), ("potency", potency),
                                 ("quantity", quantity), ("shop", shop), ("time", time_val)]:
                if value:
                    lst = self.autocomplete_data.setdefault(field, [])
                    if value not in lst:
                        lst.append(value)
            with open(self.autocomplete_file, "w") as f:
                json.dump(self.autocomplete_data, f)

            # Save record
            record = {
                "Medicine": med_name,
                "Potency": potency,
                "Dose": dose,
                "Quantity": quantity,
                "Time": time_val,
                "Shop": shop,
                "Branch/Phone": branch_phone
            }
            if os.path.exists(self.records_file):
                df_records = pd.read_excel(self.records_file, engine="openpyxl")
                df_records = pd.concat([df_records, pd.DataFrame([record])], ignore_index=True)
            else:
                df_records = pd.DataFrame([record])
            df_records.to_excel(self.records_file, index=False, engine="openpyxl")

            # Generate PDF
            c = canvas.Canvas(pdf_file, pagesize=(label_width, label_height))
            c.setLineWidth(1)
            c.rect(2*mm, 2*mm, label_width-4*mm, label_height-4*mm)

            y = label_height - top_margin

            # Line 1: Medicine + Potency
            c.setFont("Helvetica-Bold", 8)
            c.drawCentredString(label_width/2, y, f"{med_name} {potency}")

            # Line 2: Quantity
            y -= spacing
            c.setFont("Helvetica", 8)
            c.drawCentredString(label_width/2, y, f"{quantity}")

            # Line 3: Dose + Time (no "pills")
            y -= spacing
            c.drawCentredString(label_width/2, y, f"{dose}   {time_val}")

            # Line 4: Shop
            y -= spacing
            c.setFont("Helvetica-Bold", 7)
            c.drawCentredString(label_width/2, y, f"{shop}")

            # Line 5: Branch/Phone
            y -= spacing
            c.drawCentredString(label_width/2, y, f"{branch_phone}")

            c.save()

            # Open PDF
            if sys.platform == "win32":
                os.startfile(pdf_file)
            else:
                os.system(f"open {pdf_file}")

            self.status.setText("Label created, opened, and record saved.")

        except Exception as e:
            logging.error("Print/Save failed: %s", str(e))
            QtWidgets.QMessageBox.critical(self, "Error", f"Failed: {e}")
            self.status.setText(f"Error: {e}")

# ------------------ Run App ------------------
if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    w = HomeoLabelApp()
    w.show()
    sys.exit(app.exec_())
